import os
import sys
from dotenv import load_dotenv
from deepagents import create_deep_agent
from deepagents.backends import LocalShellBackend
from langchain_ollama import ChatOllama
from llm_factory import build_chat_model

load_dotenv()

PDF = sys.argv[1] if len(sys.argv) > 1 else "inputs/Multi-Agent Systems.pdf"
OUT = sys.argv[2] if len(sys.argv) > 2 else "outputs/deep_agent.pptx"

PROMPT = f"""Convert the PDF into a summarized PowerPoint with images.
Step 1: run exactly this command with the execute tool and wait for it to finish:
python skills/pdf-pptx-skill/scripts/process_pdf.py "{PDF}" "{OUT}" --summarize
Step 2: run exactly this check:
python -c "from pptx import Presentation; p=Presentation('{OUT}'); print(len(p.slides), 'slides,', sum(1 for s in p.slides for sh in s.shapes if sh.shape_type==13), 'pictures')"
Step 3: report the last line of step 1 and the output of step 2. Do not run anything else."""

SYSTEM = """You are a presentation expert on Windows PowerShell.
Read skills/pdf-pptx-skill/SKILL.md first and run the command it gives, exactly as written.
Always put file paths between double quotes. Use "python", not "python3".
Do not write your own scripts and do not give manual instructions: execute the commands yourself.
Do not use "mkdir -p" or other Linux-only commands. The folder outputs/ already exists.
If a Python library is missing, run the pip install line from SKILL.md once, then rerun the command.
Keep tool outputs short and finish after the file is created."""

backend = LocalShellBackend(
    root_dir=os.getcwd(),
    virtual_mode=True,
    inherit_env=True,
)


model = build_chat_model(os.getenv("LLM_PROVIDER"), os.getenv("LLM_MODEL"), api_key=os.getenv("LLM_API_KEY"))

agent = create_deep_agent(
    model=model,
    backend=backend,
    skills=["skills/pdf-pptx-skill"],
    system_prompt=SYSTEM,
)


if __name__ == "__main__":
    for step in agent.stream(
        {"messages": [{"role": "user", "content": PROMPT}]},
        stream_mode="values",
    ):
        step["messages"][-1].pretty_print()