"""Extract text (with headings) and embedded images from a PDF into content.json."""
import hashlib
import io
import json
import re
import statistics
import sys
from collections import Counter
from pathlib import Path

import pdfplumber
from pypdf import PdfReader

MIN_IMG_SIDE = 100      # px, skip icons/bullets smaller than this
MIN_IMG_BYTES = 2500
BULLET_RE = re.compile(r"^\s*([\u2022\u25CF\u25AA\u25E6\u2023\u2043\-\*\u2013]|\d+[.)])\s+")


def _norm(t):
    return re.sub(r"\d+", "#", t.strip())


def _lines(page):
    """Group words into lines with their font size."""
    words = page.extract_words(extra_attrs=["size"], keep_blank_chars=False, use_text_flow=False)
    words.sort(key=lambda w: (round(w["top"] / 3), w["x0"]))
    lines, cur = [], []
    for w in words:
        if cur and abs(w["top"] - cur[-1]["top"]) > max(3, cur[-1]["size"] * 0.5):
            lines.append(cur)
            cur = []
        cur.append(w)
    if cur:
        lines.append(cur)
    out = []
    for ln in lines:
        ln.sort(key=lambda w: w["x0"])
        out.append({
            "text": " ".join(w["text"] for w in ln).strip(),
            "size": round(statistics.median(w["size"] for w in ln), 1),
            "top": min(w["top"] for w in ln),
            "bottom": max(w["bottom"] for w in ln),
        })
    return [l for l in out if l["text"]]


def _paragraphs(lines, body):
    paras, cur, prev = [], "", None
    for l in lines:
        t = l["text"]
        new = False
        if prev is not None:
            gap = l["top"] - prev["bottom"]
            if gap > body * 0.6 or BULLET_RE.match(t):
                new = True
        if new and cur:
            paras.append(cur)
            cur = ""
        t = BULLET_RE.sub("", t) if BULLET_RE.match(t) else t
        cur = (cur + " " + t).strip()
        prev = l
    if cur:
        paras.append(cur)
    return paras


def _save_images(reader, pdf, doc_dir, hash_pages):
    """Return {page_index: [image dicts]} for raster images actually drawn on each page."""
    img_dir = doc_dir / "images"
    img_dir.mkdir(parents=True, exist_ok=True)
    result = {}
    for pi, page in enumerate(reader.pages):
        pl = pdf.pages[pi]
        drawn = {}
        for d in pl.images:
            wpt, hpt = d["x1"] - d["x0"], d["bottom"] - d["top"]
            if wpt < 40 or hpt < 40:
                continue  # icon / bullet
            if wpt > pl.width * 0.95 and hpt > pl.height * 0.95:
                continue  # full-page background
            drawn.setdefault(d.get("name"), (round(d["top"]), round(d["x0"])))
        try:
            imgs = list(page.images)
        except Exception:
            imgs = []
        items, used = [], set()
        for im in imgs:
            key = Path(im.name).stem
            if key not in drawn or key in used:
                continue
            used.add(key)
            try:
                data, pil = im.data, im.image
                w, h = pil.size
            except Exception:
                continue
            if w < MIN_IMG_SIDE or h < MIN_IMG_SIDE or len(data) < MIN_IMG_BYTES:
                continue
            digest = hashlib.md5(data).hexdigest()
            hash_pages.setdefault(digest, set()).add(pi)
            items.append((drawn[key], pil, w, h, digest))
        items.sort(key=lambda t: t[0])
        out = []
        for n, (_, pil, w, h, digest) in enumerate(items, 1):
            name = f"p{pi + 1}_{n}.png"
            try:
                if pil.mode not in ("RGB", "RGBA", "L"):
                    pil = pil.convert("RGB")
                pil.save(img_dir / name)
            except Exception:
                continue
            out.append({"path": f"images/{name}", "w": w, "h": h, "hash": digest})
        result[pi] = out
    return result


def _render_vector_figure(page, doc_dir, pno):
    """Fallback: crop vector drawings (charts/diagrams) to a PNG."""
    objs = list(page.rects) + list(page.curves) + list(page.lines)
    if len(objs) < 25:
        return None
    x0 = max(min(o["x0"] for o in objs), 0)
    x1 = min(max(o["x1"] for o in objs), page.width)
    y0 = max(min(o["top"] for o in objs), 0)
    y1 = min(max(o["bottom"] for o in objs), page.height)
    if (x1 - x0) < 120 or (y1 - y0) < 80:
        return None
    if (x1 - x0) > page.width * 0.97 and (y1 - y0) > page.height * 0.97:
        return None  # full-page background, not a figure
    name = f"p{pno}_vec.png"
    page.crop((x0, y0, x1, y1)).to_image(resolution=150).save(doc_dir / "images" / name)
    return {"path": f"images/{name}", "w": int(x1 - x0), "h": int(y1 - y0), "hash": name}


def extract(pdf_path, work_dir):
    pdf_path, work_dir = Path(pdf_path), Path(work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)
    reader = PdfReader(str(pdf_path))
    hash_pages = {}
    with pdfplumber.open(str(pdf_path)) as pdf_img:
        images = _save_images(reader, pdf_img, work_dir, hash_pages)
    n_pages = len(reader.pages)

    pages, all_lines = [], []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for pi, page in enumerate(pdf.pages):
            lines = _lines(page)
            all_lines.append((page, lines))
    # repeated header/footer lines
    counter = Counter()
    for _, lines in all_lines:
        for t in {_norm(l["text"]) for l in lines}:
            counter[t] += 1
    repeated = {t for t, c in counter.items() if n_pages >= 4 and c > max(3, n_pages * 0.4)}

    with pdfplumber.open(str(pdf_path)) as pdf:
        for pi, (_, lines) in enumerate(all_lines):
            page = pdf.pages[pi]
            lines = [l for l in lines if _norm(l["text"]) not in repeated and not re.fullmatch(r"\d{1,3}", l["text"])]
            sizes = [l["size"] for l in lines for _ in range(max(1, len(l["text"]) // 10))]
            body = statistics.median(sizes) if sizes else 12
            heading_lines = []
            for l in lines:
                if l["size"] >= body * 1.25 and len(heading_lines) < 3 and (not heading_lines or l["size"] >= max(h["size"] for h in heading_lines) * 0.85):
                    heading_lines.append(l)
                elif heading_lines:
                    break
                elif l["top"] > page.height * 0.5:
                    break
            heading = " ".join(h["text"] for h in heading_lines)
            rest = [l for l in lines if l not in heading_lines]
            paras = _paragraphs(rest, body)
            if re.match(r"^(Figure|Fig\.|Table)\s*\d", heading):
                paras.insert(0, heading)  # a caption in a big font is not a title
                heading = ""
            imgs = images.get(pi, [])
            if not imgs:
                try:
                    v = _render_vector_figure(page, work_dir, pi + 1)
                except Exception:
                    v = None
                if v:
                    imgs = [v]
            caption = next((p for p in paras if re.match(r"^(Figure|Fig\.|Table)\s*\d", p)), "")
            if caption:
                paras.remove(caption)
            pages.append({"page": pi + 1, "heading": heading, "paragraphs": paras,
                          "caption": caption[:120], "images": imgs})

    # drop logos/watermarks that repeat on many pages
    for p in pages:
        p["images"] = [i for i in p["images"] if len(hash_pages.get(i["hash"], ())) <= max(3, n_pages * 0.3)]
        for i in p["images"]:
            i.pop("hash", None)

    meta = reader.metadata
    title = (meta.title if meta and meta.title else "") or (pages[0]["heading"] if pages else pdf_path.stem)
    data = {"title": title.strip() or pdf_path.stem, "source": pdf_path.name, "pages": pages}
    (work_dir / "content.json").write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    return data


def summary(data, max_chars=110):
    out = [f"Title: {data['title']}  | pages: {len(data['pages'])}"]
    for p in data["pages"]:
        first = (p["paragraphs"][0][:max_chars] if p["paragraphs"] else "")
        imgs = ",".join(Path(i["path"]).name for i in p["images"])
        out.append(f"p{p['page']}: [{p['heading'][:60]}] {first} | imgs: {imgs or '-'}")
    return "\n".join(out)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    d = extract(sys.argv[1], sys.argv[2])
    print(summary(d))
