"""Move a figure to the page whose text actually discusses it.

In a PDF a figure often floats to the top of the NEXT page. This module reads the figure
number from its caption ("Figure 8-1 ...") and, when the current page never mentions that
figure but one of the previous pages does, moves the image to that earlier page.
Figures without a numbered caption are left where they are.
"""
import re

_REF = re.compile(r"\b(Figure|Fig|Table)\.?\s*(\d+(?:[.\-\u2013]\d+)*)", re.I)
MAX_IMAGES = 2   # a slide shows at most two figures


def _key(kind, num):
    return ("table" if kind.lower() == "table" else "figure", num.replace("\u2013", "-"))


def _keys(text):
    return {_key(k, n) for k, n in _REF.findall(text)}


def _text(p):
    return " ".join([p.get("heading", "")] + list(p.get("paragraphs", [])))


def relocate_figures(data, look_back=2):
    pages = data["pages"]
    moved = 0
    for i, p in enumerate(pages):
        if i == 0 or not p.get("images"):      # i == 0 is the cover
            continue
        cap = (p.get("caption") or "").strip()
        m = _REF.match(cap)
        if not m:
            continue
        key = _key(*m.groups())
        if key in _keys(_text(p)):             # this page itself discusses the figure
            continue
        for j in range(i - 1, max(0, i - look_back - 1), -1):
            if key not in _keys(_text(pages[j])):
                continue
            tgt = pages[j]
            if len(tgt.get("images", [])) + len(p["images"]) > MAX_IMAGES:
                break
            # figure captions are kept apart from the page's own (possibly table) caption
            own = tgt.get("fig_caption") or (tgt.get("caption") if tgt.get("images") else "")
            tgt["fig_caption"] = (own + " \u00b7 " + cap)[:200] if own else cap
            tgt["images"] = tgt.get("images", []) + p["images"]
            p["images"], p["caption"] = [], ""
            moved += 1
            break
    # a page that only carried a moved figure is now empty: drop it
    data["pages"] = [p for k, p in enumerate(pages)
                     if k == 0 or p.get("images") or p.get("paragraphs") or p.get("heading") or p.get("tables")]
    if moved:
        print(f"  moved {moved} figure(s) to the page that mentions them", flush=True)
    return data