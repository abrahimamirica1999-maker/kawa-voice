# -*- coding: utf-8 -*-
"""Build items.json: the "fix all typos" review (every row of the typo workbook's Corrections sheet).

Two groups, chosen on the first screen:
  ar  - Arabic (Fusha meaning in the Kurdish courses, Iraqi Arabic in the Arabic course) + English
  ku  - Kurdish (Standard Sorani, Hawleri)
Each item says what the text must be written in (`L`), shows the English meaning (`en`) found in the course where the
text is used, the text now (`t`), the suggested fix (`f`, may be empty) and why (`w`, English, as in the workbook).

  python build_items.py        # from the app repo's others/kawa-voice

Inputs (app repo, others/): typo-review-all-courses-2026-10-21.xlsx (the rows) and typo-hunt-2026-10-20/
strings.json + courses.json (the 2026-10-20 snapshot the workbook was made from; where each text is used).
Ids are `x` + sha1(lang|text)[:8], so a rebuild keeps the reviewer's saved progress. `r` = the workbook row number
(the # column), so answers join back to the workbook too. The Wrong letters sheet is NOT here (owner, 2026-10-24:
fix those automatically).
"""
import collections, hashlib, json, os, re, sys

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
OTHERS = os.path.dirname(HERE)
BOOK = os.path.join(OTHERS, "typo-review-all-courses-2026-10-21.xlsx")
HUNT = os.path.join(OTHERS, "typo-hunt-2026-10-20")
OUT = os.path.join(HERE, "items.json")

# workbook "Language" -> (strings.json lang, group, section key)
LANG = {"Arabic (Fusha meaning)": ("ar-fusha", "ar", "fusha"), "Arabic (Iraqi)": ("ar-iraqi", "ar", "iraqi"),
        "English": ("en", "ar", "en"), "Kurdish (Sorani)": ("ku-std", "ku", "std"),
        "Kurdish (Hawleri)": ("ku-hawleri", "ku", "hawleri")}
# sections in the order the reviewer meets them; titles in the group's own language
SECTIONS = [("ar", "fusha", "العربية الفصحى", "Fusha (Standard Arabic)"),
            ("ar", "iraqi", "اللهجة العراقية", "Iraqi Arabic"),
            ("ar", "en", "الإنكليزية", "English"),
            ("ku", "std", "کوردیی ستاندارد (سۆرانی)", "Standard Kurdish (Sorani)"),
            ("ku", "hawleri", "هەولێری", "Hawleri")]
CHUNK = 40

C = json.load(open(os.path.join(HUNT, "courses.json"), encoding="utf-8"))
S = json.load(open(os.path.join(HUNT, "strings.json"), encoding="utf-8"))
TESTS = {t["id"]: t for c in ("standard_kurdish", "hawleri_kurdish", "iraqi_arabic") for t in C[c]["tests"]}
CARDS = {f["id"]: f for c in ("standard_kurdish", "hawleri_kurdish", "iraqi_arabic") for f in C[c]["flashcards"]}
WHERE = collections.defaultdict(list)
for s in S:
    WHERE[(s["lang"], s["text"])] += [dict(w, course=s["course"]) for w in s["where"]]

EN_KEYS = ("en", "caption_en", "target_en")
OTHER_KEYS = ("ar", "ku", "caption_ar", "caption_ku", "target_ar", "target_ku")


def holds(v, text):
    return v == text or (isinstance(v, list) and text in v)


def find_dict(v, text):
    """The innermost dict of a test's answers that has `text` as one of its values (an option, a pair, labels)."""
    if isinstance(v, dict):
        for x in v.values():
            d = find_dict(x, text)
            if d is not None:
                return d
        if any(holds(x, text) for x in v.values()):
            return v
    elif isinstance(v, list):
        for x in v:
            d = find_dict(x, text)
            if d is not None:
                return d
    return None


def first(v):
    return (v[0] if isinstance(v, list) and v else v) if v else ""


def quoted(prompt):
    m = re.search(r'"([^"]+)"', prompt or "")
    return m.group(1) if m else ""


def norm(x):
    return " ".join(re.sub(r"[.?!,;:؟،«»\"]", " ", x or "").split())


# a whole sentence (Arabic or Kurdish, any course) -> its English, from the tests that carry both
SENT_EN = {}
for t in TESTS.values():
    va = t.get("valid_answers") or {}
    en = first(va.get("target_en")) or first(va.get("valid_en")) or quoted(t.get("prompt"))
    if not en:
        continue
    for k in ("target_ar", "target_ku", "valid_ar", "valid_ku", "answers"):
        for x in (va.get(k) if isinstance(va.get(k), list) else [va.get(k)]):
            if isinstance(x, str) and x:
                SENT_EN.setdefault(norm(x), en)
TITLES = {}
for c in ("standard_kurdish", "hawleri_kurdish", "iraqi_arabic"):
    for x in C[c]["sections"] + C[c]["units"] + C[c]["steps"]:
        for f in ("title_ar", "title_ku"):
            if x.get(f):
                TITLES[x[f]] = x.get("title") or ""
        if x.get("title"):
            TITLES.setdefault("en:" + x["title"], x.get("title_ku") or x.get("title_ar") or "")


def sentences(lang, text):
    """For a word of an arrange exercise: the sentence it is part of."""
    got = collections.Counter()
    for w in WHERE[(lang, text)]:
        t = TESTS.get(w.get("test"))
        if t and t["exercise_type"] == "arrange" and w.get("field") == "order":
            got[" ".join((t.get("valid_answers") or {}).get("order") or [])] += 1
    return [x for x, _ in got.most_common(1)]


def meanings(lang, text):
    """English meaning(s) of a text (for an English text: the Arabic/Kurdish shown with it), most used first."""
    got = collections.Counter()
    for w in WHERE[(lang, text)]:
        if "flashcard" in w:
            f = CARDS.get(w["flashcard"]) or {}
            if lang == "en":
                got[f.get("target_text") or ""] += 1
            else:
                got[f.get("english_text") or ""] += 1
            continue
        t = TESTS.get(w.get("test"))
        if not t:
            continue
        va = t.get("valid_answers") or {}
        if w.get("field", "").startswith("prompt"):          # the instruction itself, e.g. fill_blank 'for "..."'
            got[(quoted(t.get("prompt")) or t.get("prompt") or "") if lang != "en"
                else (quoted(t.get("prompt_ku") or t.get("prompt_ar")) or t.get("prompt_ku") or t.get("prompt_ar") or "")] += 1
            continue
        if t["exercise_type"] == "arrange" and w.get("field") == "order":
            got[SENT_EN.get(norm(" ".join(va.get("order") or [])), "")] += 1
            continue
        d = find_dict(va, text)
        if lang == "en":
            hit = d and next((d[k] for k in OTHER_KEYS if d.get(k) and d.get(k) != text), "")
            got[hit or first(va.get("target_ar")) or first(va.get("target_ku")) or ""] += 1
        else:
            hit = d and next((first(d[k]) for k in EN_KEYS if d.get(k)), "")
            if not hit and d and isinstance(d.get("labels"), dict):
                hit = d["labels"].get("en", "")
            # a whole answer of the test (translate / fill_blank / arrange): the test's own English
            hit = hit or first(va.get("target_en")) or first(va.get("valid_en")) or quoted(t.get("prompt"))
            got[hit or ""] += 1
    if not got and "title" in " ".join(w.get("kind", "") for w in WHERE[(lang, text)]):
        got[TITLES.get(text) or TITLES.get("en:" + text) or ""] += 1
    got.pop("", None)
    return [m for m, _ in got.most_common(2)]


wb = openpyxl.load_workbook(BOOK, read_only=True)
rows = list(wb["Corrections"].iter_rows(min_row=2, values_only=True))
items, stats = [], collections.Counter()
for r in rows:
    if r[0] is None:
        continue
    no, courses, language, found, sure, text, fix, why, used, where = r[:10]
    lang, group, key = LANG[language]
    m = meanings(lang, text)
    stats[(group, key, "with meaning" if m else "NO meaning")] += 1
    items.append({"id": "x" + hashlib.sha1((lang + "|" + text).encode("utf-8")).hexdigest()[:8], "r": no,
                  "g": group, "k": key, "co": courses, "t": text, "f": fix or "", "w": why or "",
                  "by": found, "s": sure or "", "n": used, "wh": where or "", "en": m, "in": sentences(lang, text)})

# sections: per language, chunks of CHUNK, most-used texts first
sections, ordered = [], []
for g, k, ti, en in SECTIONS:
    its = sorted([i for i in items if i["k"] == k], key=lambda i: (-(i["n"] or 0), i["r"]))
    parts = max(1, -(-len(its) // CHUNK))
    for p in range(parts):
        chunk = its[p * CHUNK:(p + 1) * CHUNK]
        if not chunk:
            continue
        si = len(sections)
        sections.append({"g": g, "k": k, "title": ti + (" %d" % (p + 1) if parts > 1 else ""),
                         "name": en + (" %d" % (p + 1) if parts > 1 else "")})
        for it in chunk:
            it["c"] = si
            ordered.append(it)

assert len({i["id"] for i in ordered}) == len(ordered) == len(items), "duplicate ids"
json.dump({"batch": "typos", "sections": sections, "items": ordered},
          open(OUT, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
for k, v in sorted(stats.items()):
    print(k, v)
print(len(ordered), "items,", len(sections), "sections ->", OUT)
