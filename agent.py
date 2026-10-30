import os
import sys
from dotenv import load_dotenv
from deepagents import create_deep_agent
from llm_factory import build_chat_model
from security import RestrictedShellBackend

SYSTEM = """You convert PDFs into PowerPoint files. You run commands in a Linux shell.
Follow the steps in the user message exactly.
Only these commands exist: process_pdf.py (as given in the steps) and nothing else.
Never use pdftoppm, pandoc, libreoffice, pip or any other tool. If a command is rejected, do not try alternatives: stop and report the error.
Always put file paths between double quotes.
Do not write your own scripts and do not give manual instructions: execute the commands yourself.
The folder outputs/ already exists. Libraries are already installed.
Keep tool outputs short and finish after the file is created."""


def build_prompt(pdf, out, summarize):
    if not summarize:
        return f"""Convert the PDF into a PowerPoint with images.
You MUST call the execute tool. Never claim success without running the command.
Keep the double quotes around both paths exactly as written. Do not remove them.
Step 1: call execute with this exact command and wait for it to finish:
python skills/pdf-pptx-skill/scripts/process_pdf.py "{pdf}" "{out}"
Step 2: report the last line of the output. If it failed, report the error. Do not run anything else."""

    # Same prompt for EVERY provider: the summarizing happens inside the script,
    # using the model the user selected, so the result does not depend on the agent model.
    return f"""Convert the PDF into a summarized PowerPoint with images.
You MUST call the execute tool. Never claim success without running the command.
Keep the double quotes around both paths exactly as written.
Step 1: call execute with this exact command and wait for it to finish:
python skills/pdf-pptx-skill/scripts/process_pdf.py "{pdf}" "{out}" --summarize
Step 2: report the last line of the output. If it failed, report the error. Do not run anything else.
Copy the command character by character. Do NOT add quotes around the script path. Quotes are only around the two file paths, exactly as shown."""


def run_agent(provider, model_name, api_key, pdf, out, summarize=True):
    prompt = build_prompt(pdf, out, summarize)

    # The script (not the agent) needs to know which model to use for summarizing.
    env = {
        "PATH": os.environ["PATH"],
        "SUM_PROVIDER": provider or "",
        "SUM_MODEL": model_name or "",
    }
    for k in ("SYSTEMROOT", "OLLAMA_BASE_URL"):  # SYSTEMROOT is needed by Python/SSL on Windows
        if os.environ.get(k):
            env[k] = os.environ[k]
    if api_key and provider != "ollama":
        env["SUM_API_KEY"] = api_key

    backend = RestrictedShellBackend(
        root_dir=os.getcwd(),
        virtual_mode=True,
        inherit_env=False,
        env=env,
    )
    model = build_chat_model(provider, model_name, api_key=api_key)

    agent = create_deep_agent(
        model=model,
        backend=backend,
        skills=["skills/pdf-pptx-skill"],
        system_prompt=SYSTEM,
    )
    for step in agent.stream(
        {"messages": [{"role": "user", "content": prompt}]},
        stream_mode="values",
    ):
        yield step["messages"][-1]


if __name__ == "__main__":
    load_dotenv()
    if len(sys.argv) < 3:
        sys.exit('Usage: python agent.py "inputs/file.pdf" "outputs/file.pptx" [--summarize]')
    for msg in run_agent(
        os.getenv("LLM_PROVIDER"), os.getenv("LLM_MODEL"), os.getenv("LLM_API_KEY"),
        sys.argv[1], sys.argv[2], summarize="--summarize" in sys.argv,
    ):
        msg.pretty_print()