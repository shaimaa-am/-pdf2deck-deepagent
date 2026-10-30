---
name: pdf-pptx-skill
description: Convert a PDF into a polished PowerPoint (.pptx) with text and embedded images extracted and placed on the slides, section dividers, and speaker notes. Use whenever the user asks to turn a PDF into slides or a presentation.
---

# PDF to PPTX

Works on Windows, macOS and Linux. Python only (pypdf, pdfplumber, pypdfium2, Pillow, python-pptx). No Node, no Poppler, no LibreOffice.

## Normal deck: one command

Always quote paths (file names may contain spaces):

```
python skills/pdf-pptx-skill/scripts/process_pdf.py "inputs/FILE.pdf" "outputs/deep_agent.pptx"
```

It extracts text and embedded images, then builds the deck. When it finishes it prints the slide count and the image count. Read that output, tell the user where the file is, and stop.

## Summarized deck: one command

Use this only if the user asks for a summarized deck:

```
python skills/pdf-pptx-skill/scripts/process_pdf.py "inputs/FILE.pdf" "outputs/deep_agent.pptx" --summarize
```

The script itself summarizes every page with the model the user selected. You do not summarize anything yourself, you do not read or write any JSON file. Run the command, read the last line, and stop.

## Rules

- Output goes to `outputs/`. The folder already exists.
- Use `python`, not `python3`.
- All libraries are already installed. Never run pip or any command that is not listed in this file.
- Do not write your own scripts. Do not give manual instructions.
- If a command is rejected, do not try alternatives: stop and report the error.
