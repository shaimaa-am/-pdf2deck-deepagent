import os
import re
import shutil
import uuid
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

load_dotenv()

from llm_factory import PROVIDERS
from agent import run_agent
from security import get_safe_pdf, get_safe_pptx

MAX_MB = 50
DEBUG = os.getenv("APP_DEBUG", "0") == "1"   # technical details are hidden unless APP_DEBUG=1
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

# --- Page Configuration ---
st.set_page_config(
    page_title="PDF to Presentation Agent",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- Custom Styling ---
st.markdown("""
<style>
    :root {
        --primary-color: #6366F1;
        --primary-hover: #4F46E5;
        --bg-card: #0F172A;
        --border-color: #1E293B;
        --text-muted: #94A3B8;
    }

    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    .gradient-title {
        background: linear-gradient(135deg, #A855F7 0%, #6366F1 50%, #3B82F6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        font-size: 2.3rem;
        margin-bottom: 0.2rem;
    }

    .sub-title {
        color: var(--text-muted);
        font-size: 1rem;
        margin-bottom: 2rem;
    }

    .stButton > button {
        background: linear-gradient(135deg, #6366F1 0%, #4F46E5 100%);
        color: white !important;
        font-weight: 600;
        border-radius: 8px;
        border: none;
        padding: 0.6rem 1.2rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 12px rgba(99, 102, 241, 0.3);
    }

    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(99, 102, 241, 0.4);
    }

    section[data-testid="stSidebar"] {
        border-right: 1px solid var(--border-color);
    }
</style>
""", unsafe_allow_html=True)

# --- Header Section ---
st.markdown('<h1 class="gradient-title">PDF2Deck AI Assistant</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Transform complex technical documents into styled presentation decks automatically</p>', unsafe_allow_html=True)

# --- Sidebar Configuration ---
with st.sidebar:
    st.markdown("### ⚡ Configuration Panel")
    st.markdown("Configure your AI model and environment settings.")

    provider_keys = list(PROVIDERS.keys())
    default_provider = os.getenv("LLM_PROVIDER", provider_keys[0]).lower()
    default_provider_idx = provider_keys.index(default_provider) if default_provider in provider_keys else 0

    selected_provider = st.selectbox(
        "LLM Provider",
        options=provider_keys,
        index=default_provider_idx,
        format_func=lambda x: PROVIDERS[x]["label"],
    )

    spec = PROVIDERS[selected_provider]
    env_key = spec.get("env_key", "")
    pkg_name = spec.get("pkg", "unknown package")

    default_model = os.getenv("LLM_MODEL", "")
    model_name = st.text_input(
        "Model Name",
        value=default_model,
        placeholder="e.g. qwen3:8b, llama-3.3-70b-versatile",
    )

    api_key = ""
    if env_key:
        env_val = os.getenv(env_key, "")
        api_key = st.text_input(
            f"API Key ({env_key})",
            value=env_val if not env_val else "••••••••••••••••",
            type="password",
            help=f"Loaded from environment ({env_key}) if set.",
        )
        if api_key.startswith("•"):
            api_key = env_val

    st.divider()
    st.caption(f"📦 Required Package: `{pkg_name}`")

# --- Main Interface ---
col_upload, col_action = st.columns([1.2, 1], gap="large")

with col_upload:
    st.markdown("##### 📄 Document Input")
    uploaded_file = st.file_uploader("Upload PDF Document", type=["pdf"], label_visibility="collapsed")

    st.markdown("##### 🏷️ Presentation File Name")
    custom_ppt_name = st.text_input(
        "Output Presentation Name",
        value="",
        placeholder="e.g., my_presentation (leave blank to use PDF name)",
        label_visibility="collapsed",
    )
    
with col_action:
    st.markdown("##### 🚀 Execution Control")
    st.write("Click below to start the autonomous agent workflow.")
    run_button = st.button("Generate Presentation Deck ⚡", type="primary", use_container_width=True)


# --- Helpers ---
def mask(text):
    """Never show the API key in any message."""
    t = str(text)
    return t.replace(api_key, "***") if api_key else t


def explain_failure(msgs):
    """Turn the agent's tool output into one clear, safe sentence."""
    text = " ".join(str(getattr(m, "content", "")) for m in msgs
                    if getattr(m, "type", "") == "tool").lower()
    if "timed out" in text or "exit code 124" in text:
        return "The conversion took too long and was stopped. Try a smaller PDF or a faster model."
    if "not allowed" in text or "rejected" in text:
        return "A command was blocked by the security policy."
    if "cannot load the model" in text or "did not respond" in text:
        return "The selected model could not be used. Check the model name and the API key."
    if "pdf not found" in text:
        return "The uploaded PDF could not be found by the converter."
    return "The agent finished without creating the presentation. Please try again."


# --- Agent Processing ---
if run_button:
    st.session_state["result"] = None   # forget the previous result

    if not uploaded_file:
        st.error("⚠️ Please upload a PDF document first.")
    elif not model_name.strip():
        st.error("⚠️ Model name is required. Configure it in the sidebar.")
    elif env_key and not api_key:
        st.error(f"⚠️ An API key is required for {spec['label']}. Enter it in the sidebar or set {env_key} in .env.")
    elif uploaded_file.size > MAX_MB * 1024 * 1024:
        st.error(f"⚠️ The PDF is larger than {MAX_MB} MB.")
    elif bytes(uploaded_file.getbuffer()[:5]) != b"%PDF-":
        st.error("⚠️ This file is not a valid PDF.")
    else:
        os.makedirs("outputs", exist_ok=True)
        os.makedirs("inputs", exist_ok=True)

        job_id = uuid.uuid4().hex
        pdf_path = out_path = None
        work_dir = Path("outputs") / f"_work_{job_id}"   # must match process_pdf.py

        raw = custom_ppt_name.strip() or os.path.splitext(uploaded_file.name)[0]
        raw = re.sub(r"\.pptx$", "", raw, flags=re.IGNORECASE)
        download_name = (re.sub(r"[^\w\-]+", "_", raw).strip("_") or "presentation") + ".pptx"

        try:
            with open(os.path.join("inputs", f"{job_id}.pdf"), "wb") as f:
                f.write(uploaded_file.getbuffer())

            pdf_path = get_safe_pdf(f"{job_id}.pdf")
            out_path = get_safe_pptx(f"{job_id}.pptx")
            pdf_rel = os.path.relpath(pdf_path).replace("\\", "/")
            out_rel = os.path.relpath(out_path).replace("\\", "/")

            msgs = []
            with st.spinner("🤖 AI Agent is processing your document and generating the presentation... Please wait."):
                for step in run_agent(
                    selected_provider,
                    model_name.strip(),
                    api_key or None,
                    pdf_rel,
                    out_rel,
                    True,
                ):
                    msgs.append(step)

            if out_path.exists():
                final_path = out_path.with_name(download_name)
                if final_path.exists():
                    final_path = out_path.with_name(f"{final_path.stem}_{job_id[:6]}.pptx")
                out_path.rename(final_path)
                # keep the result in session_state so the download button survives Streamlit reruns
                st.session_state["result"] = {"name": final_path.name, "data": final_path.read_bytes()}
            else:
                st.warning(f"⚠️ {explain_failure(msgs)}")
                if DEBUG:
                    with st.expander("Technical details (APP_DEBUG=1)"):
                        for m in msgs:
                            for tc in getattr(m, "tool_calls", None) or []:
                                st.write("TOOL:", tc.get("name"), mask(tc.get("args"))[:200])
                            if getattr(m, "type", "") == "tool":
                                st.write("RESULT:", mask(m.content)[:300])
                        if msgs:
                            st.code(mask(getattr(msgs[-1], "content", ""))[-800:])

        except Exception as e:
            st.error(f"Execution Error: {mask(e)[:300]}")
        finally:
            # always clean up: the uploaded PDF and this job's temporary work folder
            if pdf_path is not None:
                Path(pdf_path).unlink(missing_ok=True)
            else:
                Path("inputs", f"{job_id}.pdf").unlink(missing_ok=True)
            if out_path is not None:
                Path(out_path).unlink(missing_ok=True)   # only if it was not renamed
            shutil.rmtree(work_dir, ignore_errors=True)

# --- Result (outside the button block, so it stays after clicking download) ---
result = st.session_state.get("result")
if result:
    st.success(f"🎉 Presentation successfully generated as `{result['name']}`!")
    st.download_button(
        label=f"📥 Download {result['name']}",
        data=result["data"],
        file_name=result["name"],
        mime=PPTX_MIME,
        use_container_width=True,
    )