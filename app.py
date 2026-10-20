import os
import sys
import streamlit as st
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Import your existing factory modules
from llm_factory import PROVIDERS, build_chat_model
from deepagents import create_deep_agent
from deepagents.backends import LocalShellBackend

# --- Page Configuration ---
st.set_page_config(
    page_title="PDF to Presentation Agent",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
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
""", unsafe_allow_html=1)

# --- Header Section ---
st.markdown('<h1 class="gradient-title">PDF2Deck AI Assistant</h1>', unsafe_allow_html=1)
st.markdown('<p class="sub-title">Transform complex technical documents into styled presentation decks automatically</p>', unsafe_allow_html=1)

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
        format_func=lambda x: PROVIDERS[x]["label"]
    )
    
    spec = PROVIDERS[selected_provider]
    
    default_model = os.getenv("LLM_MODEL", "")
    model_name = st.text_input("Model Name", value=default_model, placeholder="e.g. qwen3:8b, llama-3.3-70b-versatile")
    
    api_key = ""
    if spec["env_key"]:
        env_val = os.getenv(spec["env_key"], "")
        api_key = st.text_input(
            f"API Key ({spec['env_key']})",
            value=env_val if not env_val else "••••••••••••••••",
            type="password",
            help=f"Loaded from environment ({spec['env_key']}) if set."
        )
        if api_key.startswith("•"):
            api_key = env_val

    st.divider()
    st.caption(f"📦 Required Package: `{spec['pkg']}`")

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
        label_visibility="collapsed"
    )

with col_action:
    st.markdown("##### 🚀 Execution Control")
    st.write("Click below to start the autonomous agent workflow.")
    run_button = st.button("Generate Presentation Deck ⚡", type="primary", use_container_width=1)

# --- Agent Processing ---
if run_button:
    if not uploaded_file:
        st.error("⚠️ Please upload a PDF document first.")
    elif not model_name:
        st.error("⚠️ Model name is required. Configure it in the sidebar.")
    else:
        # Determine output filename based on user input or uploaded PDF filename
        os.makedirs("outputs", exist_ok=1)
        os.makedirs("inputs", exist_ok=1)
        
        input_pdf_path = os.path.join("inputs", uploaded_file.name)
        with open(input_pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        if custom_ppt_name.strip():
            # Clean name and ensure .pptx extension
            clean_name = custom_ppt_name.strip()
            if not clean_name.endswith(".pptx"):
                clean_name += ".pptx"
            output_filename = os.path.join("outputs", clean_name)
        else:
            # Default to PDF base name
            base_name = os.path.splitext(uploaded_file.name)[0]
            output_filename = os.path.join("outputs", f"{base_name}.pptx")

        try:
            with st.spinner("🤖 AI Agent is processing your document and generating the presentation... Please wait."):
                model = build_chat_model(selected_provider, model_name, api_key=api_key if api_key else None)
                
                backend = LocalShellBackend(
                    root_dir=os.getcwd(),
                    virtual_mode=1,
                    inherit_env=1,
                )

                system_prompt = """You are a presentation expert on Windows PowerShell.
Read skills/pdf-pptx-skill/SKILL.md first and run the command it gives, exactly as written.
Always put file paths between double quotes. Use "python", not "python3".
Do not write your own scripts and do not give manual instructions: execute the commands yourself.
Do not use "mkdir -p" or other Linux-only commands. The folder outputs/ already exists.
If a Python library is missing, run the pip install line from SKILL.md once, then rerun the command.
Keep tool outputs short and finish after the file is created."""

                prompt_text = f"""Convert the PDF into a summarized PowerPoint with images.
Step 1: run exactly this command with the execute tool and wait for it to finish:
python skills/pdf-pptx-skill/scripts/process_pdf.py "{input_pdf_path}" "{output_filename}" --summarize
Step 2: run exactly this check:
python -c "from pptx import Presentation; p=Presentation('{output_filename}'); print(len(p.slides), 'slides,', sum(1 for s in p.slides for sh in s.shapes if sh.shape_type==13), 'pictures')"
Step 3: report the last line of step 1 and the output of step 2. Do not run anything else."""

                agent = create_deep_agent(
                    model=model,
                    backend=backend,
                    skills=["skills/pdf-pptx-skill"],
                    system_prompt=system_prompt,
                )

                for step in agent.stream(
                    {"messages": [{"role": "user", "content": prompt_text}]},
                    stream_mode="values",
                ):
                    pass

            if os.path.exists(output_filename):
                st.success(f"🎉 Presentation successfully generated as `{os.path.basename(output_filename)}`!")
                with open(output_filename, "rb") as file_download:
                    st.download_button(
                        label=f"📥 Download {os.path.basename(output_filename)}",
                        data=file_download,
                        file_name=os.path.basename(output_filename),
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=1
                    )
            else:
                st.warning("⚠️ Workflow completed, but target output file was not detected.")

        except Exception as e:
            st.error(f"Execution Error: {str(e)}")