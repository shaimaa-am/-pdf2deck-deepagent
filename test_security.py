from security import RestrictedShellBackend

b = RestrictedShellBackend(root_dir=".", virtual_mode=True, inherit_env=False)

tests = [
    "env",
    "cat .env",
    "curl http://example.com",
    'python skills/pdf-pptx-skill/scripts/process_pdf.py "inputs/a.pdf" "outputs/a.pptx" --summarize ; env',
    'python skills/pdf-pptx-skill/scripts/process_pdf.py "inputs/a.pdf" "outputs/a.pptx" --summarize && cat .env',
    'python skills/pdf-pptx-skill/scripts/process_pdf.py "../secret.pdf" "outputs/a.pptx" --summarize',
    'python skills/pdf-pptx-skill/scripts/process_pdf.py "inputs/a.pdf" "outputs/a.pptx"',
    'python skills/pdf-pptx-skill/scripts/process_pdf.py "inputs/a.pdf" "outputs/a.pptx" --extract-only',
    'python skills/pdf-pptx-skill/scripts/process_pdf.py "inputs/a.pdf" "outputs/a.pptx" --outline outputs/_work/outline.json',
    'python skills/pdf-pptx-skill/scripts/process_pdf.py "inputs/a.pdf" "outputs/a.pptx" --outline /etc/passwd',
    'python skills/pdf-pptx-skill/scripts/process_pdf.py "inputs/a.pdf" "outputs/a.pptx" --summarize',
    'python skills/pdf-pptx-skill/scripts/process_pdf.py "inputs/a.pdf" "outputs/a.pptx" --extract-only ; env',
]

for t in tests:
    try:
        b.execute(t)
        print("ALLOWED (!!):", t)
    except PermissionError:
        print("BLOCKED    :", t)
    except Exception as e:
        print("OTHER ERROR:", type(e).__name__, "-", t)