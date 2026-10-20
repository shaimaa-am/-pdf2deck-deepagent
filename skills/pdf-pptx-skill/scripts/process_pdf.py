"""PDF -> PPTX. Usage:
python process_pdf.py "in.pdf" "out.pptx" [--extract-only] [--outline outline.json]
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

try:
    from extract_pdf import extract, summary
    from build_pptx import build, auto_outline
except ImportError as e:
    sys.exit(f"Missing library: {e}. Run: pip install pypdf pdfplumber pypdfium2 pillow python-pptx")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("pptx")
    ap.add_argument("--outline")
    ap.add_argument("--extract-only", action="store_true")
    ap.add_argument("--summarize", action="store_true", help="summarize each page with a local Ollama model")
    ap.add_argument("--model", help="Ollama model name (default qwen3:8b or $OLLAMA_MODEL)")
    ap.add_argument("--max-bullets", type=int, default=5)
    ap.add_argument("--sections", type=int, default=8, help="slides per section divider in summary mode (0 = none)")
    ap.add_argument("--pages", help="page range to convert, e.g. 1-40")
    a = ap.parse_args()

    pdf, out = Path(a.pdf), Path(a.pptx)
    if not pdf.exists():
        sys.exit(f"PDF not found: {pdf}")
    work = out.parent / "_work"
    data = extract(pdf, work)
    if a.pages:
        lo, _, hi = a.pages.partition("-")
        lo, hi = int(lo or 1), int(hi or lo or 10**6)
        data["pages"] = [q for q in data["pages"] if lo <= q["page"] <= hi]
    n_img = sum(len(p["images"]) for p in data["pages"])

    if a.extract_only:
        print(summary(data))
        print(f"\nSaved {work / 'content.json'} and {n_img} images in {work / 'images'}")
        return

    if a.outline:
        spec = json.loads(Path(a.outline).read_text(encoding="utf-8"))
    else:
        spec = auto_outline(data)
        if a.summarize:
            from summarize import summarize_outline, add_sections
            print("Summarizing pages...")
            spec = summarize_outline(spec, data, model=a.model, max_bullets=a.max_bullets,
                                     cache_path=work / "summary_cache.json")
            spec = add_sections(spec, model=a.model, size=a.sections, cache_path=work / "summary_cache.json")
            (work / "outline_summarized.json").write_text(json.dumps(spec, ensure_ascii=False, indent=1), encoding="utf-8")
        (work / "outline_auto.json").write_text(json.dumps(spec, ensure_ascii=False, indent=1), encoding="utf-8")
    try:
        n = build(spec, work, out)
    except PermissionError:
        out = out.with_name(out.stem + "_new" + out.suffix)
        print(f"Cannot write the file (is it open in PowerPoint?). Saving as {out} instead.")
        n = build(spec, work, out)
    print(f"OK: {out} | slides: {n} | images extracted: {n_img}")


if __name__ == "__main__":
    main()
