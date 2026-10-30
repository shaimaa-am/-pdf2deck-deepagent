# PDF2Deck

Turn a PDF into a styled PowerPoint deck with an AI agent. You choose the model (Ollama, Claude, Gemini, OpenAI or Groq), upload a PDF in a Streamlit page, and get a `.pptx` with one summarized slide per page, figures, tables, an overview slide and a closing slide.

The agent is built with [deepagents](https://github.com/langchain-ai/deepagents). The agent itself only runs **one fixed command**; the layout, slide count and design come from code, so the result looks the same whichever model you pick. Only the wording of the bullets depends on the model.

---

## What you get

- **Summarized slides:** one slide per PDF page (long pages are summarized by the selected model, short pages are kept as they are).
- **Figures:** embedded images and vector charts are extracted and placed beside the text. A figure is moved to the page whose text mentions it (for example "Figure 8-1"), even when it floats to the next page in the PDF.
- **Tables:** real tables become native, editable PowerPoint tables. Long tables are split over several slides with the header row repeated. Tables are copied as they are, never summarized.
- **Structure:** a cover, section dividers every 8 slides (named by the model), a numbered overview slide after the cover, and a closing slide with one icon per key takeaway.
- **Design:** navy cover with rings, Cambria headings, Calibri body, small teal kicker with an orange square, figures in soft rounded cards, quiet footer with slide number.
- **Any provider:** Ollama (local), Anthropic, Google Gemini, OpenAI, Groq.

---

## How it works

```
Streamlit (app.py)
   │ uploads the PDF to inputs/, picks provider + model + key
   ▼
agent.py ── deepagents agent ── RestrictedShellBackend (security.py)
   │ runs ONE allowed command:
   │ python skills/pdf-pptx-skill/scripts/process_pdf.py "in.pdf" "out.pptx" --summarize
   ▼
process_pdf.py
   1. extract_pdf.py   text, headings, images, vector figures, tables  → content.json
   2. figures.py       move each figure to the page that mentions it
   3. build_pptx.py    auto_outline(): one slide per page (+ table slides)
   4. summarize.py     the SELECTED model summarizes every page (parallel for cloud models),
                       names the sections, writes the overview and the closing slide
   5. build_pptx.py    draws the .pptx
```

The model chosen in the sidebar is passed to the script through environment variables (`SUM_PROVIDER`, `SUM_MODEL`, `SUM_API_KEY`). It is never written in the prompt or in the command, so the agent model cannot see or change it.

---

## Project structure

```
PDF2Deck/
├── app.py                  Streamlit interface
├── agent.py                creates the deep agent and its restricted backend
├── llm_factory.py          builds the chat model for each provider
├── security.py             allowed command, safe paths, RestrictedShellBackend
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
├── .env                    your keys (never commit, never copy into the image)
├── .dockerignore
├── .gitignore
├── inputs/                 uploaded PDFs (a tmpfs in Docker, removed after each job)
├── outputs/                generated .pptx files
└── skills/
    └── pdf-pptx-skill/
        ├── SKILL.md        instructions the agent reads
        └── scripts/
            ├── process_pdf.py    entry point (the only allowed command)
            ├── extract_pdf.py    text, images and tables from the PDF
            ├── figures.py        figure → the page that discusses it
            ├── summarize.py      summaries, sections, overview, closing
            ├── build_pptx.py     slide design and auto outline
            └── icons_data.py     38 embedded icons (generated file)
```

`skills/` is mounted **read-only** in Docker, so after editing a script you do not need to rebuild the image. Changes to `app.py`, `agent.py`, `security.py` or `llm_factory.py` need `docker compose up -d --build`.

---

## Run with Docker (recommended)

1. Install Docker Desktop and, for local models, [Ollama](https://ollama.com) (`ollama pull qwen3:8b`).
2. Optional: create `.env` next to `docker-compose.yml`:

   ```env
   LLM_PROVIDER=ollama
   LLM_MODEL=qwen3:8b
   OLLAMA_BASE_URL=http://host.docker.internal:11434
   # keys for cloud providers (only the ones you use)
   GOOGLE_API_KEY=
   ANTHROPIC_API_KEY=
   OPENAI_API_KEY=
   GROQ_API_KEY=
   ```

   Docker does not read `.env` inside the container unless you add `env_file: .env` to the `agent` service. Never `COPY` the file into the image: list `.env` in `.dockerignore`.
3. Start:

   ```bash
   docker compose up -d --build
   ```
4. Open <http://localhost:8501>, choose the provider and model in the sidebar, upload a PDF and click **Generate Presentation Deck**.

On Linux, Ollama must listen on all interfaces so the container can reach it: `OLLAMA_HOST=0.0.0.0 ollama serve`.

## Run without Docker

```bash
python -m venv venv
venv\Scripts\activate          # Windows   (Linux/macOS: source venv/bin/activate)
pip install -r requirements.txt
streamlit run app.py
```

Command line (uses `LLM_PROVIDER`, `LLM_MODEL`, `LLM_API_KEY` from `.env`):

```bash
python agent.py "inputs/file.pdf" "outputs/file.pptx" --summarize
```

---

## Configuration

| Variable | Where | Meaning |
|---|---|---|
| `LLM_PROVIDER`, `LLM_MODEL` | `.env` | default provider and model shown in the sidebar |
| `LLM_API_KEY` | `.env` | key used by the command-line `agent.py` only |
| `ANTHROPIC_API_KEY`, `GOOGLE_API_KEY`, `OPENAI_API_KEY`, `GROQ_API_KEY` | `.env` | keys the sidebar reads automatically |
| `OLLAMA_BASE_URL` | `.env` / compose | Ollama address (default `http://host.docker.internal:11434`) |
| `APP_DEBUG=1` | `.env` | show technical details when a run fails (off by default) |

Options of `process_pdf.py` (run by hand; the agent uses the defaults):

| Option | Default | Meaning |
|---|---|---|
| `--summarize` | off | summarize every page with the selected model |
| `--max-bullets` | 5 | bullets per summarized slide |
| `--max-slides` | 0 | upper limit of slides (0 = no limit, nothing is removed) |
| `--sections` | 8 | content slides per section divider (0 = none) |
| `--pages` | all | page range, e.g. `1-40` |
| `--extract-only` | off | only write `content.json` and the images |

Summaries are cached per model in `outputs/_work_<job>/outline_cache_<provider>_<model>.json`, so two models never reuse each other's wording. The app deletes the work folder after every job.

---

## Security

- **One command only.** `security.py` accepts exactly `python skills/pdf-pptx-skill/scripts/process_pdf.py "<pdf>" "<pptx>" [--extract-only|--summarize]`. Any other command, `..`, or any of `; & | $ < > \` and backtick is rejected. No PowerShell, no shell tools.
- **Safe paths.** The PDF must be inside `inputs/`, the result inside `outputs/`, relative paths only, PDF up to 50 MB and starting with `%PDF-`.
- **Isolated environment.** The backend runs with `inherit_env=False`: the command only receives `PATH` and the `SUM_*` variables. The script reads the key once and removes it from its environment; error messages hide it.
- **No tools for the summarizer.** The text of the PDF is only sent to the summarizing model, which has no tools. Instructions hidden inside a PDF cannot run anything.
- **Docker hardening.** `read_only: true`, `cap_drop: ALL`, `no-new-privileges`, memory/CPU/process limits, port bound to `127.0.0.1`, `skills/` mounted read-only, `inputs/` kept in memory (`tmpfs`).
- **Clean errors.** The interface shows a short reason; commands and tool output appear only with `APP_DEBUG=1`.

---

## Limitations

- A table without ruling lines (text aligned with spaces only) is not detected and stays as plain paragraphs.
- A figure without a numbered caption ("Figure 3-2 …") stays on the page where it appears.
- Slides are one per page. Very long PDFs take time, and local models are slower than cloud models (cloud models summarize four pages in parallel).
- The agent tool stops a command after 120 seconds. If a large PDF with a local model hits this limit, use a cloud model or split the PDF.
- Right-to-left layout is not configured in the slides.
- Icons come from [Feather](https://feathericons.com) (MIT license).

---

## Troubleshooting

| Message | Cause and fix |
|---|---|
| `Command timed out after 120 seconds` | the document is too large for the model speed: use a cloud model or split the PDF |
| `Command not allowed` | the agent tried a command outside the allowed one: report it, do not widen `security.py` |
| `ERROR: the model … did not respond` | wrong model name, wrong key, or Ollama not reachable (check `OLLAMA_BASE_URL`) |
| `Missing API key` | enter the key in the sidebar or set it in `.env` and use `env_file` |
| `target /app/inputs already mounted` | `inputs` must be a `tmpfs` only, do not also add it as a volume |
| Same wording for two models | an old cache: delete `outputs/_work_*` |
