"""Build a PPTX from content.json (auto) or from an outline.json (curated).

Design: dark navy cover with rings, serif (Cambria) titles, small teal kicker with an
orange square, clean bullets, figures inside a soft rounded card with a caption,
and a quiet footer with page number.
"""
import base64
import io
import math
import re
from pathlib import Path

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

try:
    from icons_data import ICONS
except ImportError:      # icons_data.py missing: slides fall back to plain bullets
    ICONS = {}

# ---------- palette & fonts (taken from the reference deck) ----------
NAVY = "1C2331"        # cover / section background
RING = "36435A"        # decorative circles on dark slides
ORANGE = "C97A2B"      # small accent square, cover kicker
TEAL = "3E7C8A"        # kicker text
INK = "222732"         # titles and body text
MUTED = "6B7280"       # captions, footer
PALE = "D8DEE7"        # big numbers (agenda)
CARD = "F5F6F8"       # figure card fill
CARD_LINE = "E1E4E9"
SHADOW = "9AA3B2"
SUB = "C7CEDA"        # subtitle on dark
SUB2 = "A9B2C3"       # footnote on dark
WHITE = "FFFFFF"
HEAD_FONT = "Cambria"
FONT = "Calibri"

W, H = 13.333, 7.5
MARGIN = 0.55
BODY_TOP, BODY_BOTTOM = 1.85, 6.7
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

_CTX = {"section": ""}   # title of the last section slide, used as the kicker


def rgb(h):
    return RGBColor.from_string(h)


def _bg(slide, color):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = rgb(color)


def _text(slide, x, y, w, h, text, size, color=INK, bold=False, italic=False, align=PP_ALIGN.LEFT,
          anchor=MSO_ANCHOR.TOP, name=None, font=FONT, spc=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size, r.font.bold, r.font.italic, r.font.name = Pt(size), bold, italic, font
    r.font.color.rgb = rgb(color)
    if spc:
        r._r.get_or_add_rPr().set("spc", str(spc))
    return tb


def _rect(slide, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(color)
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def _ring(slide, x, y, d):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    s.fill.background()
    s.line.color.rgb = rgb(RING)
    s.line.width = Pt(1.25)
    s.shadow.inherit = False
    return s


def _soft_card(slide, x, y, w, h):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.adjustments[0] = 0.026
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(CARD)
    s.line.color.rgb = rgb(CARD_LINE)
    s.line.width = Pt(1)
    spPr = s._element.spPr
    for old in spPr.findall(f"{{{A_NS}}}effectLst"):
        spPr.remove(old)
    spPr.append(etree.fromstring(
        f'<a:effectLst xmlns:a="{A_NS}"><a:outerShdw blurRad="76200" dist="25400" dir="5400000" '
        f'algn="bl" rotWithShape="0"><a:srgbClr val="{SHADOW}"><a:alpha val="28000"/></a:srgbClr>'
        f'</a:outerShdw></a:effectLst>'))
    return s


def _est_height(bullets, size, width):
    cpl = max(10, int(width * 72 / (size * 0.50)))
    total = 0
    for b in bullets:
        total += math.ceil(len(b) / cpl) * size * 1.2 + size * 0.55
    return total / 72


def _bullets(slide, x, y, w, h, items, size=None, color=INK):
    if size is None:
        for size in (16, 15, 14.5, 14, 13, 12, 11):
            if _est_height(items, size, w - 0.3) <= h:
                break
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.name = "Body"
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, t in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.12
        p.space_after = Pt(12)
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", "203200")
        pPr.set("indent", "-203200")
        etree.SubElement(pPr, f"{{{A_NS}}}buSzPct").set("val", "100000")
        etree.SubElement(pPr, f"{{{A_NS}}}buChar").set("char", "\u2022")
        r = p.add_run()
        r.text = t
        r.font.size, r.font.name = Pt(size), FONT
        r.font.color.rgb = rgb(color)
    return tb


def _fit(path, w, h):
    try:
        with Image.open(path) as im:
            iw, ih = im.size
    except Exception:
        return None
    sc = min(w / iw, h / ih)
    return iw * sc, ih * sc


def _figure_card(slide, paths, x, y, w, max_h, caption=""):
    """Figure(s) inside a soft rounded card, with an italic caption under it."""
    paths = [p for p in paths if p and Path(p).exists()][:2]
    if not paths:
        return False
    pad, gap = 0.18, 0.15
    cap_h = 0.32 if caption else 0
    inner_w = w - 2 * pad
    each_h = (max_h - cap_h - 2 * pad - gap * (len(paths) - 1)) / len(paths)
    sizes = [(p, _fit(p, inner_w, each_h)) for p in paths]
    sizes = [(p, s) for p, s in sizes if s]
    if not sizes:
        return False
    card_h = sum(s[1] for _, s in sizes) + gap * (len(sizes) - 1) + 2 * pad
    _soft_card(slide, x, y, w, card_h).name = "Figure card"
    cy = y + pad
    for p, (pw, ph) in sizes:
        pic = slide.shapes.add_picture(str(p), Inches(x + (w - pw) / 2), Inches(cy), Inches(pw), Inches(ph))
        pic.name = "Figure"
        cy += ph + gap
    if caption:
        _text(slide, x, y + card_h + 0.08, w, 0.3, caption, 10.5, MUTED, italic=True,
              align=PP_ALIGN.CENTER, name="Caption")
    return True


def _footer(slide, deck_title, num):
    _text(slide, MARGIN, 7.08, 7, 0.3, (deck_title or "").upper()[:80], 8.5, MUTED, name="Footer", spc=80)
    _text(slide, W - MARGIN - 0.6, 7.08, 0.6, 0.3, str(num), 8.5, MUTED, align=PP_ALIGN.RIGHT,
          name="Slide number")


def _kicker(slide, text, y=0.53):
    text = (text or "").strip()
    if not text:
        return
    _rect(slide, MARGIN, y + 0.11, 0.13, 0.13, ORANGE)
    _text(slide, MARGIN + 0.24, y, 9, 0.32, text.upper(), 12, TEAL, bold=True,
          anchor=MSO_ANCHOR.MIDDLE, name="Kicker", spc=150)


def _title(slide, title, y=0.82, big=False):
    n = len(title)
    size = (32 if big else 27) if n <= 45 else 24 if n <= 70 else 20
    _text(slide, MARGIN, y, W - 2 * MARGIN, 0.85, title, size, INK, bold=True, font=HEAD_FONT, name="Title")


def _notes(slide, text):
    if text:
        slide.notes_slide.notes_text_frame.text = text


# ---------- slide types ----------
def cover(prs, s, deck_title, base):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, NAVY)
    imgs = [base / p for p in s.get("images", [])][:1]
    has_img = bool(imgs) and imgs[0].exists()
    if has_img:  # smaller rings in the corner, so they never cross the figure card
        _ring(sl, 11.1, -1.9, 3.4)
        _ring(sl, 10.5, -2.5, 4.6)
    else:
        _ring(sl, 9.13, -1.0, 5.2)
        _ring(sl, 8.43, -1.7, 6.6)
    text_w = 7.4 if has_img else 10.6
    kicker = s.get("kicker", "")
    if kicker:
        _rect(sl, MARGIN, 2.55, 0.16, 0.16, ORANGE)
        _text(sl, MARGIN + 0.28, 2.47, 8, 0.32, kicker.upper(), 13, ORANGE, bold=True,
              anchor=MSO_ANCHOR.MIDDLE, name="Kicker", spc=200)
    title = s.get("title", deck_title)
    size = 48 if len(title) <= 28 else 40 if len(title) <= 50 else 32
    _text(sl, MARGIN, 2.9, text_w, 1.5, title, size, WHITE, bold=True, font=HEAD_FONT, name="Title")
    if s.get("subtitle"):
        _text(sl, MARGIN, 4.45, text_w, 1.0, s["subtitle"], 19, SUB, name="Subtitle")
    if has_img:
        _figure_card(sl, imgs, 8.7, 2.6, 4.1, 3.6)
    if s.get("footnote"):
        _text(sl, MARGIN, 6.8, 10, 0.35, s["footnote"], 10.5, SUB2, italic=True, name="Footnote")
    _notes(sl, s.get("notes", ""))


def section(prs, s, deck_title, base, num):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, NAVY)
    _ring(sl, 9.13, -1.0, 5.2)
    _ring(sl, 8.43, -1.7, 6.6)
    _CTX["section"] = s.get("title", "")
    _text(sl, MARGIN, 2.0, 4, 1.2, s.get("number", f"{num:02d}"), 66, ORANGE, bold=True, font=HEAD_FONT,
          name="Section number")
    _text(sl, MARGIN, 3.35, 10.6, 1.3, s.get("title", ""), 40, WHITE, bold=True, font=HEAD_FONT, name="Title")
    if s.get("subtitle"):
        _text(sl, MARGIN, 4.85, 10.0, 1.0, s["subtitle"], 19, SUB, name="Subtitle")
    _notes(sl, s.get("notes", ""))


def content(prs, s, deck_title, base, num):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, WHITE)
    _kicker(sl, s["kicker"] if "kicker" in s else _CTX["section"])
    _title(sl, s.get("title", ""))
    imgs = [base / p for p in s.get("images", []) if (base / p).exists()]
    bullets = s.get("bullets", [])
    avail_h = BODY_BOTTOM - BODY_TOP
    if imgs and bullets:
        _bullets(sl, MARGIN, BODY_TOP, 5.35, avail_h, bullets)
        _figure_card(sl, imgs, 6.35, BODY_TOP, W - MARGIN - 6.35, avail_h - 0.1, s.get("caption", ""))
    elif imgs:
        bw = 9.0
        _figure_card(sl, imgs, (W - bw) / 2, BODY_TOP, bw, avail_h - 0.1, s.get("caption", ""))
    elif bullets:
        _bullets(sl, MARGIN, BODY_TOP, 10.5, avail_h, bullets)
    _footer(sl, deck_title, num)
    _notes(sl, s.get("notes", ""))


def cards(prs, s, deck_title, base, num):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, WHITE)
    _kicker(sl, s["kicker"] if "kicker" in s else _CTX["section"])
    _title(sl, s.get("title", ""))
    items = s.get("cards", [])[:4]
    n = max(len(items), 1)
    gap = 0.3
    cw = (W - 2 * MARGIN - gap * (n - 1)) / n
    ch = BODY_BOTTOM - BODY_TOP - 0.3
    for i, c in enumerate(items):
        x = MARGIN + i * (cw + gap)
        _soft_card(sl, x, BODY_TOP, cw, ch)
        _rect(sl, x + 0.3, BODY_TOP + 0.32, 0.13, 0.13, ORANGE)
        _text(sl, x + 0.3, BODY_TOP + 0.6, cw - 0.6, 0.8, c.get("title", ""), 18, INK, bold=True,
              font=HEAD_FONT, name="Card title")
        txt = c.get("text", "")
        size = 15 if len(txt) < 220 else 14 if len(txt) < 360 else 12.5
        _text(sl, x + 0.3, BODY_TOP + 1.45, cw - 0.6, ch - 1.7, txt, size, INK, name="Card text")
    _footer(sl, deck_title, num)
    _notes(sl, s.get("notes", ""))


def agenda(prs, s, deck_title, base, num):
    """Numbered roadmap: items = [{"title": ..., "text": ...}, ...] (up to 6)."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, WHITE)
    _kicker(sl, s["kicker"] if "kicker" in s else _CTX["section"], y=0.69)
    _title(sl, s.get("title", ""), y=0.95, big=True)
    for i, it in enumerate(s.get("items", [])[:6]):
        y = 2.05 + i * 0.98
        _text(sl, MARGIN, y, 1.0, 0.72, f"{i + 1:02d}", 36, ORANGE, bold=True, font=HEAD_FONT,
              anchor=MSO_ANCHOR.MIDDLE)
        _rect(sl, 1.6, y + 0.06, 0.012, 0.6, PALE)
        _text(sl, 1.9, y, 3.6, 0.72, it.get("title", ""), 18, NAVY, bold=True, font=HEAD_FONT,
              anchor=MSO_ANCHOR.MIDDLE)
        _text(sl, 5.55, y, 7.2, 0.72, it.get("text", ""), 14, MUTED, anchor=MSO_ANCHOR.MIDDLE)
    _footer(sl, deck_title, num)
    _notes(sl, s.get("notes", ""))


# ---------- tables ----------
def _table_size(rows):
    n, c = len(rows), len(rows[0])
    longest = max(len(x) for r in rows for x in r)
    if c <= 3 and n <= 7 and longest < 80:
        return 15
    if c <= 5 and n <= 10:
        return 13
    return 11


def _col_widths(rows, total_w):
    n = len(rows[0])
    lens = [max(min(len(r[j]), 60) for r in rows) + 4 for j in range(n)]
    ws = [max(total_w * l / sum(lens), 0.8) for l in lens]
    k = total_w / sum(ws)
    return [w * k for w in ws]


def _row_heights(rows, widths, size):
    hs = []
    for r in rows:
        lines = 1
        for j, t in enumerate(r):
            cpl = max(4, int((widths[j] - 0.2) * 72 / (size * 0.5)))
            lines = max(lines, math.ceil(len(t) / cpl) if t else 1)
        hs.append(lines * size * 1.2 / 72 + 0.18)
    return hs


def _split_table(rows, avail_h=BODY_BOTTOM - BODY_TOP - 0.45):
    """Cut a long table into chunks that fit one slide; the header row is repeated."""
    widths = _col_widths(rows, W - 2 * MARGIN)
    size = _table_size(rows)
    hs = _row_heights(rows, widths, size)
    head, body, bh = rows[0], rows[1:], hs[1:]
    chunks, cur, used = [], [], hs[0]
    for r, h in zip(body, bh):
        if cur and used + h > avail_h:
            chunks.append([head] + cur)
            cur, used = [], hs[0]
        cur.append(r)
        used += h
    if cur:
        chunks.append([head] + cur)
    return chunks or [rows]


def table(prs, s, deck_title, base, num):
    rows = s.get("rows") or []
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, WHITE)
    _kicker(sl, s["kicker"] if "kicker" in s else _CTX["section"])
    _title(sl, s.get("title", ""))
    if rows:
        ncol = len(rows[0])
        tw = W - 2 * MARGIN
        widths = _col_widths(rows, tw)
        size = _table_size(rows)
        hs = _row_heights(rows, widths, size)
        gf = sl.shapes.add_table(len(rows), ncol, Inches(MARGIN), Inches(BODY_TOP), Inches(tw), Inches(sum(hs)))
        gf.name = "Table"
        tbl = gf.table
        tbl.first_row = True
        tbl.horz_banding = False
        for j, w in enumerate(widths):
            tbl.columns[j].width = Inches(w)
        for i, r in enumerate(rows):
            tbl.rows[i].height = Inches(hs[i])
            for j in range(ncol):
                c = tbl.cell(i, j)
                c.margin_left = c.margin_right = Inches(0.1)
                c.margin_top = c.margin_bottom = Inches(0.05)
                c.vertical_anchor = MSO_ANCHOR.MIDDLE
                c.fill.solid()
                c.fill.fore_color.rgb = rgb(NAVY if i == 0 else (CARD if i % 2 == 0 else WHITE))
                tf = c.text_frame
                tf.word_wrap = True
                run = tf.paragraphs[0].add_run()
                run.text = r[j] if j < len(r) else ""
                run.font.size, run.font.name = Pt(size), FONT
                run.font.bold = (i == 0)
                run.font.color.rgb = rgb(WHITE if i == 0 else INK)
        if s.get("caption"):
            _text(sl, MARGIN, min(BODY_TOP + sum(hs) + 0.1, 6.75), tw, 0.3, s["caption"], 10.5, MUTED,
                  italic=True, align=PP_ALIGN.CENTER, name="Caption")
    _footer(sl, deck_title, num)
    _notes(sl, s.get("notes", ""))


_ICON_WORDS = [
    (r"memory|data|storage|database", "database"), (r"tool|action|engineer|build", "tool"),
    (r"plan|goal|strateg|objective", "target"), (r"learn|knowledge|educat|concept|definition|read", "book"),
    (r"secur|safe|risk|privacy|ethic|responsib|trust", "shield"), (r"cost|price|budget|financ|money|revenue|profit|econom", "dollar"),
    (r"chart|metric|analy|statistic|measure|evaluat|performance|kpi", "chart"), (r"growth|trend|scale|improv", "trending"),
    (r"team|user|people|human|customer|client|stakeholder", "users"), (r"market|global|world|network|internet|web", "globe"),
    (r"time|schedul|timeline|history|roadmap", "clock"), (r"setting|config|parameter|control|process|workflow", "settings"),
    (r"code|program|software|api|implement", "code"), (r"search|research|discover|explor|find", "search"),
    (r"message|communicat|language|llm|chat|prompt", "message"), (r"model|intelligen|brain|reason|agent|\bai\b", "cpu"),
    (r"layer|architect|structur|framework|component|system", "layers"), (r"result|conclusion|outcome|success|takeaway", "check"),
    (r"idea|innovat|creativ|energy|fast|power", "zap"), (r"document|report|file|text|writ", "file"),
    (r"product|package|deliver|deploy|launch", "box"), (r"loop|cycle|iterat|repeat|feedback", "refresh"),
    (r"direction|guide|navigat|path|journey", "compass"), (r"award|quality|best|excellen", "award"),
    (r"design|visual|image|figure|table", "eye"),
]
_DEFAULT_ICONS = ["layers", "target", "compass", "book", "zap", "check", "star", "flag"]


def _pick_icon(item, i, used):
    """The model's icon if valid, else a keyword match, else a default; avoid repeating an icon."""
    cands = [str(item.get("icon", "")).strip().lower()]
    text = f"{item.get('title', '')} {item.get('text', '')}".lower()
    cands += [nm for pat, nm in _ICON_WORDS if re.search(pat, text)]
    cands += _DEFAULT_ICONS[i % len(_DEFAULT_ICONS):] + _DEFAULT_ICONS
    for nm in cands:
        if nm in ICONS and nm not in used:
            return nm
    return next(iter(ICONS), None)


def _icon_badge(sl, name, x, y, d=0.64):
    """Orange circle with a white icon inside."""
    _rect(sl, x, y, d, d, ORANGE, MSO_SHAPE.OVAL).name = "Icon badge"
    pic = sl.shapes.add_picture(io.BytesIO(base64.b64decode(ICONS[name])),
                                Inches(x + d * 0.22), Inches(y + d * 0.22), Inches(d * 0.56), Inches(d * 0.56))
    pic.name = f"Icon {name}"


def closing(prs, s, deck_title, base, num):
    """Dark closing slide: one icon per key takeaway."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, NAVY)
    _ring(sl, 9.13, -1.0, 5.2)
    _ring(sl, 8.43, -1.7, 6.6)
    _rect(sl, MARGIN, 1.0, 0.13, 0.13, ORANGE)
    _text(sl, MARGIN, 1.25, 11.5, 0.9, s.get("title", ""), 36, WHITE, bold=True, font=HEAD_FONT, name="Title")
    items = s.get("items") or [{"text": b} for b in s.get("bullets", [])]
    if items and ICONS and s.get("items"):
        used = set()
        for i, it in enumerate(items[:5]):
            y = 2.55 + i * 0.9
            name = _pick_icon(it, i, used)
            used.add(name)
            _icon_badge(sl, name, MARGIN, y)
            _text(sl, MARGIN + 0.95, y - 0.03, 7.1, 0.7, it.get("text", ""), 17, "E5E9F0",
                  anchor=MSO_ANCHOR.MIDDLE, name="Takeaway")
    elif items:
        _bullets(sl, MARGIN, 2.6, 8.0, 3.8, [it.get("text", "") for it in items], color="E5E9F0")
    _notes(sl, s.get("notes", ""))


BUILDERS = {"cover": cover, "section": section, "content": content, "cards": cards, "agenda": agenda,
            "table": table, "closing": closing}


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
    if p.get("caption"):
        return _shorten(CAP_RE.sub("", p["caption"]), 60) or f"Page {p['page']}"
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
        tables = p.get("tables") or []
        chars = sum(len(x) for x in paras)
        title = _page_title(p)
        if not paras and not p["images"] and not tables and not p["heading"]:
            continue  # blank page
        cap = p.get("caption", "")
        tcap = cap if (tables and re.match(r"^Table", cap, re.I)) else ""
        fcap = p.get("fig_caption") or ("" if tcap else cap)
        if p["heading"] and not p["images"] and not tables and len(paras) <= 2 and chars < 180:
            sec_no += 1
            slides.append({"type": "section", "number": f"{sec_no:02d}", "title": p["heading"],
                           "subtitle": " ".join(paras), "notes": ""})
            continue
        bullets = [_shorten(x) for x in paras]
        if not p["heading"] and len(paras) == 1 and CAP_RE.match(paras[0]) and p["images"]:
            bullets = []  # figure-only page: the caption is shown under the image
        notes = " ".join(paras)
        imgs = [i["path"] for i in p["images"]]
        if tables and not paras and not imgs:
            bullets = []
        chunk, chunks = [], []
        for b in bullets:
            if len(chunk) >= 6 or _est_height(chunk + [b], 16, 5.0 if imgs else 10.2) > BODY_BOTTOM - BODY_TOP:
                chunks.append(chunk)
                chunk = []
            chunk.append(b)
        chunks.append(chunk)
        if not (tables and not paras and not imgs):   # a table-only page gets only table slides
            for k, ch in enumerate(chunks):
                slides.append({"type": "content", "page": p["page"], "title": title if k == 0 else f"{title} (cont.)",
                               "bullets": ch, "images": imgs if k == 0 else [],
                               "caption": fcap if k == 0 else "",
                               "notes": notes if k == 0 else ""})
        has_text_slide = bool(paras or imgs)
        for t_no, rows in enumerate(tables):
            if has_text_slide or t_no > 0:   # avoid two slides with the same title
                base_title = _shorten(CAP_RE.sub("", tcap), 60) if (tcap and t_no == 0) else f"{title} (table)"
            else:
                base_title = title
            for k, part in enumerate(_split_table(rows)):
                first = (t_no == 0 and k == 0)
                slides.append({"type": "table", "page": p["page"],
                               "title": base_title if k == 0 else f"{base_title} (cont.)",
                               "rows": part, "caption": tcap if first else "",
                               "notes": ""})
    return {"title": slides[0]["title"] if slides else data["title"], "slides": slides}


def build(spec, base_dir, out_path):
    base = Path(base_dir)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    deck_title = spec.get("title", "")
    _CTX["section"] = ""
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