import json
d = json.load(open("outputs/_work/content.json", encoding="utf-8"))
for p in d["pages"]:
    if p["caption"] and not p["images"]:
        print("page", p["page"], "caption without image:", p["caption"][:70])
print("pages with images:", sum(1 for p in d["pages"] if p["images"]),
      "| pages with captions:", sum(1 for p in d["pages"] if p["caption"]))
