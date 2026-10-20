"""Summarize each page into short slide bullets using a local Ollama model."""
import json
import os
import re
import urllib.request

HOST = os.environ.get("OLLAMA_HOST", "http://localhost:11434")
MODEL = os.environ.get("OLLAMA_MODEL", "qwen3:8b")

PROMPT = (
    "You turn page text into slide content.\n"
    "Write {n} bullet points that summarize the text. Each bullet: at most 14 words, "
    "a complete idea, keep technical terms and numbers, same language as the text. "
    "Also write a slide title of at most 7 words.\n"
    'Return ONLY JSON: {{"title": "...", "bullets": ["...", "..."]}}\n\n'
    "Page heading: {heading}\nText:\n{text}"
)


def _chat(prompt, model, host, timeout=300):
    body = json.dumps({
        "model": model, "stream": False, "format": "json", "think": False,
        "options": {"temperature": 0, "num_ctx": 8192},
        "messages": [{"role": "user", "content": prompt}],
    }).encode("utf-8")
    req = urllib.request.Request(host.rstrip("/") + "/api/chat", body, {"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))["message"]["content"]


def _parse(raw):
    raw = re.sub(r"<think>.*?</think>", "", raw, flags=re.S).strip()
    m = re.search(r"\{.*\}", raw, flags=re.S)
    if not m:
        return None
    try:
        d = json.loads(m.group(0))
    except Exception:
        return None
    b = [str(x).strip() for x in d.get("bullets", []) if str(x).strip()]
    return {"title": str(d.get("title", "")).strip(), "bullets": b} if b else None


def summarize_outline(spec, data, model=None, host=None, max_bullets=5, chat=_chat, cache_path=None):
    model, host = model or MODEL, host or HOST
    cache = {}
    if cache_path and os.path.exists(cache_path):
        try:
            cache = json.load(open(cache_path, encoding="utf-8"))
        except Exception:
            cache = {}
    pages = {p["page"]: p for p in data["pages"]}
    out, done, total = [], set(), sum(1 for s in spec["slides"] if s.get("type") == "content")
    k = 0
    for s in spec["slides"]:
        pg = s.get("page")
        if s.get("type") != "content" or pg is None:
            out.append(s)
            continue
        if pg in done:
            continue  # continuation slide of a page already summarized into one slide
        done.add(pg)
        k += 1
        p = pages[pg]
        text = " ".join(p["paragraphs"])
        if len(text) < 250:  # already short, keep as is
            out.append(s)
            continue
        res = cache.get(str(pg))
        if res:
            print(f"  page {pg}: from cache ({k}/{total})")
        else:
            try:
                res = _parse(chat(PROMPT.format(n=max_bullets, heading=p["heading"], text=text[:4000]), model, host))
            except Exception as e:
                print(f"  page {pg}: model call failed ({e}); kept original text")
                res = None
            if res and cache_path:
                cache[str(pg)] = res
                json.dump(cache, open(cache_path, "w", encoding="utf-8"), ensure_ascii=False)
        if res:
            s = dict(s, bullets=res["bullets"][:max_bullets + 1],
                     title=s["title"] if p["heading"] else (res["title"] or s["title"]))
        if res:
            print(f"  summarized page {pg} ({k}/{total})", flush=True)
        out.append(s)
    return dict(spec, slides=out)


SEC_PROMPT = (
    "These are consecutive slide titles from one part of a presentation:\n{titles}\n"
    'Name this part. Return ONLY JSON: {{"title": "at most 6 words", "subtitle": "at most 12 words"}}'
)


def add_sections(spec, chat=_chat, model=None, host=None, size=8, cache_path=None):
    """Insert a section divider before every group of `size` content slides."""
    model, host = model or MODEL, host or HOST
    cache = {}
    if cache_path and os.path.exists(cache_path):
        try:
            cache = json.load(open(cache_path, encoding="utf-8"))
        except Exception:
            cache = {}
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
                raw = chat(SEC_PROMPT.format(titles="\n".join("- " + t for t in titles)), model, host)
                raw = re.sub(r"<think>.*?</think>", "", raw, flags=re.S)
                d = json.loads(re.search(r"\{.*\}", raw, flags=re.S).group(0))
                res = {"title": str(d["title"]).strip(), "subtitle": str(d.get("subtitle", "")).strip()}
                if cache_path:
                    cache[key] = res
                    json.dump(cache, open(cache_path, "w", encoding="utf-8"), ensure_ascii=False)
            except Exception as e:
                print(f"  section {g_no}: model call failed ({e}); using a generic title")
                res = {"title": f"Part {g_no}", "subtitle": ""}
        print(f"  section {g_no}: {res['title']}", flush=True)
        inserts[g[0]] = {"type": "section", "number": f"{g_no:02d}", **res, "notes": ""}
    out = []
    for i, s in enumerate(slides):
        if i in inserts:
            out.append(inserts[i])
        out.append(s)
    return dict(spec, slides=out)
