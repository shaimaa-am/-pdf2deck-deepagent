"""PDF -> PPTX. Usage:
python process_pdf.py "in.pdf" "out.pptx" [--extract-only] [--summarize] [--outline outline.json]
"""
import argparse
import json
import os
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
# project root (scripts -> pdf-pptx-skill -> skills -> root) so llm_factory.py can be imported
sys.path.insert(0, str(HERE.parents[2]))
sys.path.insert(0, os.getcwd())
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

try:
    from extract_pdf import extract, summary
    from build_pptx import build, auto_outline
except ImportError as e:
    sys.exit(f"Missing library: {e}. Run: pip install pypdf pdfplumber pypdfium2 pillow python-pptx")

try:
    from figures import relocate_figures
except ImportError:
    relocate_figures = lambda d: d

try:
    import summarize as _sum
except ImportError:
    _sum = None

_KEY = os.environ.pop("SUM_API_KEY", None)   # read once, then hidden from the environment


def _safe(msg):
    return str(msg).replace(_KEY, "***") if _KEY else str(msg)


def cap_slides(spec, max_slides):
    """Keep the deck under max_slides: drop '(cont.)' slides first, then the shortest content slides."""
    if not max_slides or max_slides <= 0:
        return spec
    s = list(spec["slides"])
    while len(s) > max_slides:
        i = next((k for k, x in enumerate(s)
                  if x.get("type") == "content" and str(x.get("title", "")).endswith("(cont.)")), None)
        if i is None:
            cand = [k for k, x in enumerate(s) if x.get("type") == "content"]
            if not cand:
                break
            i = min(cand, key=lambda k: sum(len(b) for b in s[k].get("bullets", [])))
        s.pop(i)
    return dict(spec, slides=s)


def setup_model_for_summary():
    """Connect summarize.py to the model selected in the app (passed through SUM_* env vars)."""
    provider = os.environ.get("SUM_PROVIDER")
    model_name = os.environ.get("SUM_MODEL")
    if not provider or not model_name:
        print("Note: SUM_PROVIDER/SUM_MODEL not set; using local Ollama fallback.")
        return
    try:
        if provider == "ollama":
            from langchain_ollama import ChatOllama
            m = ChatOllama(model=model_name, temperature=0, num_ctx=8192,
                           reasoning=False, format="json",
                           base_url=os.environ.get("OLLAMA_BASE_URL", "http://localhost:11434"))
        else:
            from llm_factory import build_chat_model
            m = build_chat_model(provider, model_name, api_key=_KEY)
    except Exception as e:
        sys.exit(f"ERROR: cannot load the model for summarizing: {_safe(e)}")

    def call(prompt):
        r = m.invoke(prompt + "\n\nReturn ONLY valid JSON. No markdown, no explanation.")
        c = r.content
        if isinstance(c, str):
            return c
        return "".join(x.get("text", "") for x in c if isinstance(x, dict))

    try:  # fail fast if the model/key does not work
        call('Reply with {"ok": true}')
    except Exception as e:
        sys.exit(f"ERROR: the model '{model_name}' ({provider}) did not respond: {_safe(e)}")
    _sum.set_chat(call)
    print(f"Summarizing with {provider}:{model_name}", flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("pptx")
    ap.add_argument("--outline")
    ap.add_argument("--extract-only", action="store_true")
    ap.add_argument("--summarize", action="store_true", help="summarize each page with the selected model")
    ap.add_argument("--model", help="model name (only used by the old Ollama fallback)")
    ap.add_argument("--max-bullets", type=int, default=5)
    ap.add_argument("--max-slides", type=int, default=0, help="upper limit of slides (0 = no limit)")
    ap.add_argument("--sections", type=int, default=8, help="slides per section divider in summary mode (0 = none)")
    ap.add_argument("--pages", help="page range to convert, e.g. 1-40")
    a = ap.parse_args()

    pdf, out = Path(a.pdf), Path(a.pptx)
    if not pdf.exists():
        sys.exit(f"PDF not found: {pdf}")
    work = out.parent / f"_work_{out.stem}"     # one work folder per job
    data = extract(pdf, work)
    if a.pages:
        lo, _, hi = a.pages.partition("-")
        lo, hi = int(lo or 1), int(hi or lo or 10**6)
        data["pages"] = [q for q in data["pages"] if lo <= q["page"] <= hi]

    data = relocate_figures(data)               # a figure goes to the page that discusses it
    n_img = sum(len(p["images"]) for p in data["pages"])
    n_tbl = sum(len(p.get("tables") or []) for p in data["pages"])

    if a.extract_only:
        print(summary(data))
        print(f"\nSaved {work / 'content.json'}, {n_img} images and {n_tbl} tables")
        return

    if a.outline:
        spec = json.loads(Path(a.outline).read_text(encoding="utf-8"))
    else:
        spec = auto_outline(data)
        spec = cap_slides(spec, a.max_slides)
        if a.summarize:
            if _sum is None:
                sys.exit("summarize.py not found.")
            setup_model_for_summary()
            tag = re.sub(r"[^A-Za-z0-9_.-]", "_", f"{os.environ.get('SUM_PROVIDER')}_{os.environ.get('SUM_MODEL')}")
            cache = str(work / f"outline_cache_{tag}.json")
            workers = 1 if os.environ.get("SUM_PROVIDER") == "ollama" else 4
            spec = _sum.summarize_outline(spec, data, model=a.model, max_bullets=a.max_bullets,
                                          cache_path=cache, max_workers=workers)
            spec = _sum.add_sections(spec, model=a.model, size=a.sections, cache_path=cache)
            spec = _sum.add_intro_outro(spec, model=a.model, cache_path=cache)
            (work / "outline_auto.json").write_text(json.dumps(spec, ensure_ascii=False, indent=1), encoding="utf-8")
    try:
        n = build(spec, work, out)
    except PermissionError:
        out = out.with_name(out.stem + "_new" + out.suffix)
        print(f"Cannot write the file (is it open in PowerPoint?). Saving as {out} instead.")
        n = build(spec, work, out)
    print(f"OK: {out} | slides: {n} | images extracted: {n_img} | tables: {n_tbl}")


if __name__ == "__main__":
    main()