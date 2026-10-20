"""Build a PPTX from content.json (auto) or from an outline.json (curated)."""
import json
import math
import re
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

NAVY, TEAL, INK, MUTED, LIGHT, WHITE = "1F2A44", "0E9F9A", "1F2937", "6B7280", "F5F7FA", "FFFFFF"
FONT = "Calibri"
W, H = 13.333, 7.5
MARGIN = 0.6
BODY_TOP, BODY_BOTTOM = 1.65, 6.75


def rgb(h):
    return RGBColor.from_string(h)


def _bg(slide, color):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = rgb(color)


def _text(slide, x, y, w, h, text, size, color=INK, bold=False, align=PP_ALIGN.LEFT,
          anchor=MSO_ANCHOR.TOP, name=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = 0
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size, r.font.bold, r.font.name = Pt(size), bold, FONT
    r.font.color.rgb = rgb(color)
    return tb


def _rect(slide, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(color)
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def _est_height(bullets, size, width):
    cpl = max(10, int(width * 72 / (size * 0.50)))
    total = 0
    for b in bullets:
        total += math.ceil(len(b) / cpl) * size * 1.2 + size * 0.55
    return total / 72


def _bullets(slide, x, y, w, h, items, size=None):
    if size is None:
        for size in (26, 24, 22, 20, 18, 16, 14):
            if _est_height(items, size, w - 0.4) <= h:
                break
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.name = "Body"
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    for i, t in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(size * 0.55)
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(int(Inches(0.32))))
        pPr.set("indent", str(-int(Inches(0.32))))
        for tag in ("a:buNone", "a:buChar"):
            pass
        from lxml import etree
        ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        clr = etree.SubElement(pPr, f"{{{ns}}}buClr")
        etree.SubElement(clr, f"{{{ns}}}srgbClr").set("val", TEAL)
        etree.SubElement(pPr, f"{{{ns}}}buFont").set("typeface", "Arial")
        etree.SubElement(pPr, f"{{{ns}}}buChar").set("char", "\u2022")
        r = p.add_run()
        r.text = t
        r.font.size, r.font.name = Pt(size), FONT
        r.font.color.rgb = rgb(INK)
    return tb


def _picture(slide, path, x, y, w, h):
    """Fit image inside the box keeping aspect ratio, centered."""
    try:
        with Image.open(path) as im:
            iw, ih = im.size
    except Exception:
        return False
    scale = min(w / iw, h / ih)
    pw, ph = iw * scale, ih * scale
    pic = slide.shapes.add_picture(str(path), Inches(x + (w - pw) / 2), Inches(y + (h - ph) / 2),
                                   Inches(pw), Inches(ph))
    pic.name = "Figure"
    return True


def _images_box(slide, paths, x, y, w, h, caption=""):
    paths = [p for p in paths if p and Path(p).exists()]
    if not paths:
        return False
    cap_h = 0.4 if caption else 0
    h -= cap_h
    n = min(len(paths), 3)
    gap = 0.2
    each = (h - gap * (n - 1)) / n
    for i, p in enumerate(paths[:n]):
        _picture(slide, p, x, y + i * (each + gap), w, each)
    if caption:
        _text(slide, x, y + h + 0.05, w, cap_h - 0.05, caption, 12, MUTED, align=PP_ALIGN.CENTER, name="Caption")
    return True


def _footer(slide, deck_title, num):
    _text(slide, MARGIN, 7.0, 8, 0.3, deck_title[:80], 10, MUTED, name="Footer")
    _text(slide, W - MARGIN - 1, 7.0, 1, 0.3, str(num), 10, MUTED, align=PP_ALIGN.RIGHT, name="Slide number")


def _title(slide, title):
    size = 32 if len(title) <= 45 else 26 if len(title) <= 70 else 22
    _text(slide, MARGIN, 0.45, W - 2 * MARGIN, 0.95, title, size, NAVY, bold=True,
          anchor=MSO_ANCHOR.MIDDLE, name="Title")
    _rect(slide, MARGIN, 1.42, 1.2, 0.07, TEAL)


def _notes(slide, text):
    if text:
        slide.notes_slide.notes_text_frame.text = text


# ---------- slide types ----------
def cover(prs, s, deck_title, base):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, NAVY)
    imgs = [base / p for p in s.get("images", [])][:1]
    text_w = 7.2 if imgs and imgs[0].exists() else W - 2 * MARGIN - 1
    _rect(sl, MARGIN, 2.2, 0.12, 2.3, TEAL)
    _text(sl, MARGIN + 0.4, 2.1, text_w, 1.9, s.get("title", deck_title), 44, WHITE, bold=True,
          anchor=MSO_ANCHOR.BOTTOM, name="Title")
    if s.get("subtitle"):
        _text(sl, MARGIN + 0.4, 4.15, text_w, 1.2, s["subtitle"], 22, "BFD7EA", name="Subtitle")
    if imgs:
        _images_box(sl, imgs, 8.3, 1.4, 4.4, 4.6)
    _notes(sl, s.get("notes", ""))


def section(prs, s, deck_title, base, num):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, NAVY)
    _text(sl, MARGIN + 0.3, 1.7, 4, 1.6, s.get("number", f"{num:02d}"), 96, TEAL, bold=True, name="Section number")
    _text(sl, MARGIN + 0.3, 3.4, W - 2 * MARGIN - 1, 1.2, s.get("title", ""), 40, WHITE, bold=True,
          anchor=MSO_ANCHOR.TOP, name="Title")
    if s.get("subtitle"):
        _text(sl, MARGIN + 0.3, 4.9, W - 2 * MARGIN - 1, 1, s["subtitle"], 20, "BFD7EA", name="Subtitle")
    _notes(sl, s.get("notes", ""))


def content(prs, s, deck_title, base, num):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, WHITE)
    _title(sl, s.get("title", ""))
    imgs = [base / p for p in s.get("images", []) if (base / p).exists()]
    bullets = s.get("bullets", [])
    avail_h = BODY_BOTTOM - BODY_TOP
    if imgs:
        tw = 6.5
        _bullets(sl, MARGIN, BODY_TOP, tw, avail_h, bullets) if bullets else None
        bx = MARGIN + tw + 0.5 if bullets else MARGIN
        bw = W - MARGIN - bx
        _images_box(sl, imgs, bx, BODY_TOP, bw, avail_h, s.get("caption", ""))
    elif bullets:
        _bullets(sl, MARGIN, BODY_TOP, W - 2 * MARGIN, avail_h, bullets)
    _footer(sl, deck_title, num)
    _notes(sl, s.get("notes", ""))


def cards(prs, s, deck_title, base, num):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, WHITE)
    _title(sl, s.get("title", ""))
    items = s.get("cards", [])[:4]
    n = max(len(items), 1)
    gap = 0.35
    cw = (W - 2 * MARGIN - gap * (n - 1)) / n
    ch = BODY_BOTTOM - BODY_TOP
    for i, c in enumerate(items):
        x = MARGIN + i * (cw + gap)
        _rect(sl, x, BODY_TOP, cw, ch, LIGHT, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.04
        _rect(sl, x, BODY_TOP, cw, 0.09, TEAL)
        _text(sl, x + 0.3, BODY_TOP + 0.3, cw - 0.6, 0.8, c.get("title", ""), 22, NAVY, bold=True, name="Card title")
        txt = c.get("text", "")
        size = 18 if len(txt) < 220 else 16 if len(txt) < 360 else 14
        _text(sl, x + 0.3, BODY_TOP + 1.2, cw - 0.6, ch - 1.5, txt, size, INK, name="Card text")
    _footer(sl, deck_title, num)
    _notes(sl, s.get("notes", ""))


BUILDERS = {"cover": cover, "section": section, "content": content, "cards": cards}


# ---------- auto outline from content.json ----------
def _shorten(t, limit=150):
    t = re.sub(r"\s+", " ", t).strip()
    if len(t) <= limit:
        return t
    cut = t[:limit]
    m = re.search(r"[.!?;:]\s", cut[60:])
    if m:
        return cut[:60 + m.start() + 1]
    return cut.rsplit(" ", 1)[0] + "\u2026"


def _sentence(t, limit):
    """First full sentence(s) up to limit chars; never cut mid-sentence."""
    t = re.sub(r"\s+", " ", t).strip()
    out = ""
    for s in re.split(r"(?<=[.!?])\s+", t):
        if not re.search(r"[.!?]$", s) or len(out) + len(s) + 1 > limit:
            break
        out = (out + " " + s).strip()
    return out or _shorten(t, limit).rstrip("-")


CAP_RE = re.compile(r"^(Figure|Fig\.|Table)\s*[\d.\-\u2013]+\s*[.:]?\s*", re.I)


def _page_title(p):
    if p["heading"]:
        return p["heading"]
    for x in p["paragraphs"]:
        if not CAP_RE.match(x):
            return _shorten(x, 60)
    if p["paragraphs"]:
        return _shorten(CAP_RE.sub("", p["paragraphs"][0]), 60)
    return f"Page {p['page']}"


def auto_outline(data):
    pages = data["pages"]
    slides = []
    sec_no = 0
    first = pages[0] if pages else None
    if first:
        head = first["heading"] or data["title"]
        left, sep, right = head.partition("|")
        title = re.sub(r"^\s*(chapter\s+)?\d+[.:)]?\s+", "", left, flags=re.I).strip() or head
        sub = right.strip() if sep else ""
        if not sub and first["paragraphs"]:
            sub = _sentence(first["paragraphs"][0], 110)
        slides.append({"type": "cover", "title": title, "subtitle": sub,
                       "images": [i["path"] for i in first["images"]][:1],
                       "notes": " ".join(first["paragraphs"])})
    for p in pages[1:]:
        paras = p["paragraphs"]
        chars = sum(len(x) for x in paras)
        title = _page_title(p)
        if p["heading"] and not p["images"] and len(paras) <= 2 and chars < 180:
            sec_no += 1
            slides.append({"type": "section", "number": f"{sec_no:02d}", "title": p["heading"],
                           "subtitle": " ".join(paras), "notes": ""})
            continue
        bullets = [_shorten(x) for x in paras]
        if not p["heading"] and len(paras) == 1 and CAP_RE.match(paras[0]) and p["images"]:
            bullets = []  # figure-only page: the caption is shown under the image
        notes = " ".join(paras)
        imgs = [i["path"] for i in p["images"]]
        chunk, chunks = [], []
        for b in bullets:
            if len(chunk) >= 6 or _est_height(chunk + [b], 16, 5.9 if imgs else 11.7) > BODY_BOTTOM - BODY_TOP:
                chunks.append(chunk)
                chunk = []
            chunk.append(b)
        chunks.append(chunk)
        for k, ch in enumerate(chunks):
            slides.append({"type": "content", "page": p["page"], "title": title if k == 0 else f"{title} (cont.)",
                           "bullets": ch, "images": imgs if k == 0 else [],
                           "caption": p.get("caption", "") if k == 0 else "",
                           "notes": notes if k == 0 else ""})
    return {"title": slides[0]["title"] if slides else data["title"], "slides": slides}


def build(spec, base_dir, out_path):
    base = Path(base_dir)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    deck_title = spec.get("title", "")
    sec = 0
    for i, s in enumerate(spec["slides"], 1):
        t = s.get("type", "content")
        if t == "cover":
            cover(prs, s, deck_title, base)
        elif t == "section":
            sec += 1
            section(prs, s, deck_title, base, sec)
        else:
            BUILDERS.get(t, content)(prs, s, deck_title, base, i)
    prs.core_properties.title = deck_title
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return len(prs.slides)
