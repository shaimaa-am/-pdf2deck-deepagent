# PDF2Deck AI Assistant & Deep Agent

An advanced AI-powered application that transforms complex PDF documents into structured, summarized PowerPoint presentation decks (`.pptx`) with extracted figures, custom section dividers, and a modern Streamlit web interface. It supports multi-LLM providers (Ollama for local execution, plus OpenAI, Anthropic, Google Gemini, and Groq).

## Project Structure

```
project_directory/
  agent.py                      Deep Agent runner script
  app.py                        Streamlit web user interface
  llm_factory.py                Multi-provider LLM factory (Ollama, OpenAI, Anthropic, Google, Groq)
  check_figures.py              Optional check: pages with a caption but no image
  inputs/                       Directory for input PDF files
  outputs/                      Generated presentation decks (.pptx) and working cache
  skills/pdf-pptx-skill/
    SKILL.md                    Skill description and execution commands
    scripts/
      process_pdf.py            Single entry point for PDF processing
      extract_pdf.py            Text, headings, images, vector figures, and captions
      summarize.py              Per-page summaries, caching, and section dividers
      build_pptx.py             PowerPoint builder using python-pptx
```

## Requirements & Dependencies

Create a `requirements.txt` file with the following packages:

```text
deepagents
langchain
langchain-ollama
python-dotenv
streamlit
pypdf
pdfplumber
pypdfium2
pillow
python-pptx
# Install only the provider package(s) you intend to use:
# langchain-anthropic
# langchain-google-genai
# langchain-openai
# langchain-groq
```

## Setup & Installation

1. Clone the repository and set up a virtual environment:
   ```bash
   python -m venv venv
   venv\Scripts\activate
   pip install -r requirements.txt
   ```

2. Place your target PDF documents inside the `inputs/` directory.

---

> **Optional Note (.env Configuration):**
> If you prefer to save your API keys and configuration parameters in advance rather than entering them manually through the UI every time, you can create a file named `.env` in the root directory of the project and add the following variables as needed:
> ```env
> LLM_PROVIDER=ollama
> LLM_MODEL=qwen3:8b
> # OPENAI_API_KEY=your_openai_key_here
> # ANTHROPIC_API_KEY=your_anthropic_key_here
> # GOOGLE_API_KEY=your_google_key_here
> # GROQ_API_KEY=your_groq_key_here
> ```

---

## Running the Application

### 1. Streamlit Web Interface (Recommended)
Launch the interactive web UI to configure your provider, model, upload PDFs, and generate presentations seamlessly:
```bash
streamlit run app.py
```

### 2. Command Line Deep Agent
You can also run the agent directly via the terminal:
```bash
python agent.py
python agent.py "inputs/your_document.pdf" "outputs/custom_name.pptx"
```

### 3. Direct Script Execution
```bash
python skills/pdf-pptx-skill/scripts/process_pdf.py "inputs/Multi-Agent Systems.pdf" "outputs/deep_agent.pptx" --summarize
```

---

## Features & Options

| Option / Flag | Description |
|---|---|
| `--summarize` | Summarizes each page into 3 to 5 concise bullets using the selected model |
| `--pages 1-40` | Converts only a specific page range for quick testing |
| `--sections 8` | Number of slides per section divider in summary mode (`0` = none) |
| `--max-bullets 4` | Maximum bullet points per slide |
| `--model <name>` | Specifies a custom model for the chosen provider |
| `--extract-only` | Extracts text and images into `outputs/_work` without building the deck |

## Troubleshooting

- **File Locked Error (`Cannot write the file`):** Ensure the target `.pptx` file is closed in PowerPoint before running a new generation.
- **Ollama Connection Error:** Make sure Ollama is running locally and your selected model is pulled (e.g., `ollama pull qwen3:8b`).