---
name: pdf-pptx-skill
description: Convert a PDF into a polished PowerPoint (.pptx) with text and embedded images extracted and placed on the slides, section dividers, and speaker notes. Use whenever the user asks to turn a PDF into slides or a presentation.
---

# PDF to PPTX

Works on Windows, macOS and Linux. Python only (pypdf, pdfplumber, pypdfium2, Pillow, python-pptx). No Node, no Poppler, no LibreOffice.

## Setup (once)

```
pip install pypdf pdfplumber pypdfium2 pillow python-pptx
```

## Default: one command

Always quote paths (file names may contain spaces):

```
python skills/pdf-pptx-skill/scripts/process_pdf.py "inputs/FILE.pdf" "outputs/deep_agent.pptx"
```

It extracts text and embedded images, then builds the deck. When it finishes it prints the slide count and the image count. Read that output, tell the user where the file is, and stop. Do not write other scripts. Do not give manual install instructions.

## Better quality (optional, 3 steps)

Use this only if the user asks for a summarized deck, or the default result is too wordy.

1. Extract only:
   ```
   python skills/pdf-pptx-skill/scripts/process_pdf.py "inputs/FILE.pdf" "outputs/deep_agent.pptx" --extract-only
   ```
   This writes `outputs/_work/content.json` and `outputs/_work/images/`, and prints a short per-page summary.
2. Write `outputs/_work/outline.json` (format below). Keep bullets under 15 words, 3 to 6 per slide, and put the longer explanation in `notes`.
3. Build:
   ```
   python skills/pdf-pptx-skill/scripts/process_pdf.py "inputs/FILE.pdf" "outputs/deep_agent.pptx" --outline outputs/_work/outline.json
   ```

### outline.json format

```json
{
  "title": "Deck title",
  "slides": [
    {"type": "cover", "title": "Multi-Agent Systems", "subtitle": "When agents work together", "images": ["images/p1_1.png"]},
    {"type": "section", "number": "01", "title": "What is a multi-agent system?", "subtitle": "One line"},
    {"type": "content", "title": "From a single agent to many", "bullets": ["Point one", "Point two"], "images": ["images/p4_1.png"], "caption": "Figure 8-1", "notes": "Speaker notes"},
    {"type": "cards", "title": "Benefits and costs", "cards": [{"title": "Benefits", "text": "..."}, {"title": "Costs", "text": "..."}]}
  ]
}
```

Image paths are relative to `outputs/_work/`. Use only images that exist in `content.json`.

## Rules

- Output goes to `outputs/`. The folder already exists.
- Use `python`, not `python3`. No `mkdir -p`, no Linux-only commands.
- If a library is missing, run the `pip install` line above once, then rerun the command.
