"""Summarize each page into short slide bullets using the model selected by the user.

process_pdf.py calls set_chat(fn) with a function that talks to the chosen model.
If set_chat was not called, it falls back to a local Ollama model (old behaviour).
"""
import json
import os
import re
import threading
import urllib.request
from concurrent.futures import ThreadPoolExecutor

try:
    from icons_data import ICON_NAMES
except ImportError:
    ICON_NAMES = []

HOST = os.environ.get("OLLAMA_HOST", "http://localhost:11434")
MODEL = os.environ.get("OLLAMA_MODEL", "qwen3:8b")

_CHAT_FN = None


def set_chat(fn):
    """fn(prompt: str) -> str. Set by process_pdf.py with the selected model."""
    global _CHAT_FN
    _CHAT_FN = fn


PROMPT = (
    "You turn page text into slide content.\n"
    "Write {n} bullet points that summarize the text. Each bullet: at most 14 words, "
    "a complete idea, keep technical terms and numbers, same language as the text. "
    "Also write a slide title of at most 7 words.\n"
    'Return ONLY JSON: {{"title": "...", "bullets": ["...", "..."]}}\n\n'
    "Page heading: {heading}\nText:\n{text}"
)


def _chat(prompt, model=None, host=None, timeout=300):
    if _CHAT_FN:
        return _CHAT_FN(prompt)
    body = json.dumps({
        "model": model or MODEL, "stream": False, "format": "json", "think": False,
        "options": {"temperature": 0, "num_ctx": 8192},
        "messages": [{"role": "user", "content": prompt}],
    }).encode("utf-8")
    req = urllib.request.Request((host or HOST).rstrip("/") + "/api/chat", body,
                                 {"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))["message"]["content"]


def _clean(raw):
    raw = re.sub(r"<think>.*?</think>", "", raw, flags=re.S)
    raw = re.sub(r"```(?:json)?", "", raw)
    return raw.strip()


def _parse(raw):
    raw = _clean(raw)
    m = re.search(r"\{.*\}", raw, flags=re.S)
    if not m:
        return None
    try:
        d = json.loads(m.group(0))
    except Exception:
        return None
    if not isinstance(d, dict) or not isinstance(d.get("bullets", []), list):
        return None
    b = [str(x).strip() for x in d.get("bullets", []) if str(x).strip()]
    return {"title": str(d.get("title", "")).strip(), "bullets": b} if b else None


def _load_cache(cache_path):
    if cache_path and os.path.exists(cache_path):
        try:
            return json.load(open(cache_path, encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _save_cache(cache_path, cache):
    if cache_path:
        json.dump(cache, open(cache_path, "w", encoding="utf-8"), ensure_ascii=False)


def summarize_outline(spec, data, model=None, host=None, max_bullets=5, chat=_chat,
                      cache_path=None, max_workers=1):
    cache = _load_cache(cache_path)
    lock = threading.Lock()
    pages = {p["page"]: p for p in data["pages"]}

    jobs = {}  # page -> text (None = short page, kept as is)
    for s in spec["slides"]:
        pg = s.get("page")
        if s.get("type") != "content" or pg is None or pg in jobs or pg not in pages:
            continue
        text = " ".join(pages[pg]["paragraphs"])
        jobs[pg] = text if len(text) >= 250 else None

    def work(pg):
        text = jobs[pg]
        if text is None:
            return pg, None
        res = cache.get(str(pg))
        if res:
            return pg, res
        try:
            res = _parse(chat(PROMPT.format(n=max_bullets, heading=pages[pg]["heading"], text=text[:8000])))
        except Exception as e:
            print(f"  page {pg}: model call failed ({str(e)[:120]}); kept original text", flush=True)
            return pg, None
        if res:
            with lock:
                cache[str(pg)] = res
                _save_cache(cache_path, cache)
        return pg, res

    results = {}
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        for k, (pg, res) in enumerate(ex.map(work, list(jobs)), 1):
            results[pg] = res
            if res:
                print(f"  summarized page {pg} ({k}/{len(jobs)})", flush=True)

    out, done = [], set()
    for s in spec["slides"]:
        pg = s.get("page")
        if s.get("type") != "content" or pg is None:
            out.append(s)
            continue
        if pg in done:
            continue  # continuation slide of a page already summarized into one slide
        done.add(pg)
        res = results.get(pg)
        if res:
            s = dict(s, bullets=res["bullets"][:max_bullets + 1],
                     title=s["title"] if pages[pg]["heading"] else (res["title"] or s["title"]))
        out.append(s)
    return dict(spec, slides=out)


SEC_PROMPT = (
    "These are consecutive slide titles from one part of a presentation:\n{titles}\n"
    'Name this part. Return ONLY JSON: {{"title": "at most 6 words", "subtitle": "at most 12 words"}}'
)


def add_sections(spec, chat=_chat, model=None, host=None, size=8, cache_path=None):
    """Insert a section divider before every group of `size` content slides."""
    cache = _load_cache(cache_path)
    slides = spec["slides"]
    if any(x.get("type") == "section" for x in slides):
        return spec  # the PDF already has real section pages
    idx = [i for i, s in enumerate(slides) if s.get("type") == "content"]
    if size <= 0 or len(idx) < size:
        return spec
    groups = [idx[i:i + size] for i in range(0, len(idx), size)]
    if len(groups) > 1 and len(groups[-1]) < 3:  # avoid a tiny last section
        last = groups.pop()
        groups[-1] = groups[-1] + last
    inserts = {}
    for g_no, g in enumerate(groups, 1):
        titles = [slides[i]["title"] for i in g]
        key = f"sec_{g_no}_{slides[g[0]].get('page')}_{len(g)}"
        res = cache.get(key)
        if not res:
            try:
                raw = _clean(chat(SEC_PROMPT.format(titles="\n".join("- " + t for t in titles)), model, host))
                d = json.loads(re.search(r"\{.*\}", raw, flags=re.S).group(0))
                res = {"title": str(d["title"]).strip(), "subtitle": str(d.get("subtitle", "")).strip()}
                cache[key] = res
                _save_cache(cache_path, cache)
            except Exception as e:
                print(f"  section {g_no}: model call failed ({str(e)[:120]}); using a generic title")
                res = {"title": f"Part {g_no}", "subtitle": ""}
        print(f"  section {g_no}: {res['title']}", flush=True)
        inserts[g[0]] = {"type": "section", "number": f"{g_no:02d}", **res, "notes": ""}
    out = []
    for i, s in enumerate(slides):
        if i in inserts:
            out.append(inserts[i])
        out.append(s)
    return dict(spec, slides=out)


INTRO_PROMPT = (
    "Here are the slides of a presentation (title, then the first bullet):\n{outline}\n\n"
    "Write two extra slides in the SAME LANGUAGE as the slides.\n"
    "1) overview: a title of at most 4 words (like 'What Is Inside') and 4 to 5 items. "
    "Each item has a short name (at most 4 words) and a one-line description (at most 10 words). "
    "Together the items must show the main parts of the presentation in order.\n"
    "2) closing: a title of at most 4 words (like 'Key Takeaways') and 4 to 5 items with the main conclusions. "
    "Each item has a \"text\" (at most 14 words, only facts found in the slides) and an \"icon\" "
    "chosen from this list: {icons}.\n"
    'Return ONLY JSON: {{"overview": {{"title": "...", "items": [{{"title": "...", "text": "..."}}]}}, '
    '"closing": {{"title": "...", "items": [{{"text": "...", "icon": "..."}}]}}}}'
)


def _items(raw_items):
    out = []
    for it in raw_items or []:
        if isinstance(it, dict):
            t, x = str(it.get("title", "")).strip(), str(it.get("text", "")).strip()
        else:  # a model that returned plain strings
            t, x = str(it).strip(), ""
        if t:
            out.append({"title": t, "text": x})
    return out[:6]


def _closing_items(raw_items):
    out = []
    for it in raw_items or []:
        if isinstance(it, dict):
            t, ic = str(it.get("text") or it.get("title") or "").strip(), str(it.get("icon", "")).strip().lower()
        else:
            t, ic = str(it).strip(), ""
        if t:
            out.append({"text": t, "icon": ic})
    return out[:5]


def add_intro_outro(spec, chat=_chat, model=None, host=None, cache_path=None):
    """Add a numbered overview (agenda) slide after the cover and a key-takeaways slide at the end."""
    slides = spec["slides"]
    body = [s for s in slides if s.get("type") in ("content", "section", "cards", "table")]
    if len(body) < 4:  # tiny deck: nothing to introduce
        return spec
    lines = []
    for s in body:
        first = (s.get("bullets") or [""])[0] if s.get("type") == "content" else ""
        lines.append(f"- {s.get('title', '')}" + (f": {first[:100]}" if first else ""))
    cache = _load_cache(cache_path)
    res = cache.get("intro_outro_v3")
    if not res:
        try:
            raw = _clean(chat(INTRO_PROMPT.format(outline="\n".join(lines)[:7000], icons=", ".join(ICON_NAMES)), model, host))
            d = json.loads(re.search(r"\{.*\}", raw, flags=re.S).group(0))
            ov, cl = d["overview"], d["closing"]
            res = {"overview": {"title": str(ov["title"]).strip(),
                                "items": _items(ov.get("items") or ov.get("bullets"))},
                   "closing": {"title": str(cl["title"]).strip(),
                               "items": _closing_items(cl.get("items") or cl.get("bullets"))}}
            if not (res["overview"]["title"] and len(res["overview"]["items"]) >= 2
                    and res["closing"]["title"] and res["closing"]["items"]):
                raise ValueError("empty overview or closing")
            cache["intro_outro_v3"] = res
            _save_cache(cache_path, cache)
        except Exception as e:
            print(f"  overview/closing: model call failed ({str(e)[:120]}); skipped", flush=True)
            return spec
    intro = {"type": "agenda", "title": res["overview"]["title"], "items": res["overview"]["items"],
             "kicker": "", "notes": ""}
    outro = {"type": "closing", "title": res["closing"]["title"], "items": res["closing"]["items"], "notes": ""}
    out = list(slides)
    ci = next((i for i, s in enumerate(out) if s.get("type") == "cover"), -1)
    out.insert(ci + 1, intro)
    out.append(outro)
    print("  added overview and closing slides", flush=True)
    return dict(spec, slides=out)