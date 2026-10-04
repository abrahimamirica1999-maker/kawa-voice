# -*- coding: utf-8 -*-
"""Build items.json: the Standard Kurdish flashcard deck, for conversion to Hawleri.

Reads the LIVE standard_kurdish flashcards (read-only) and the native speaker's
answers from the hawleri-tests round. A card whose Standard word she already
converted there, WITH THE SAME MEANING, is pre-filled with her Hawleri and
flagged `h`; she only confirms it. Every other card is pre-filled with the
Standard word for her to convert (only word-initial ر is respelt ڕ). Where the same Kurdish word was in
the tests with a different meaning, her test answer is shown as a hint (`t`)
but NOT applied — that is a guess, and guesses are hers to make.

  python build_items.py          # from the app repo's others/kawa-voice

Ids are `f` + the card's uuid[:8], so a rebuild keeps her saved progress.
"""
import collections, io, json, os, re, sys, urllib.request

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))
SHEET = os.path.join(REPO, "others", "Hawleri-Sorani", "hawleri-tests all corrected.xlsx")
OUT = os.path.join(HERE, "items.json")

# Same Kurdish word, but the tests used it for another meaning: not applied.
SENSE_DIFFERS = {"interview", "earth", "sharp", "wife", "needle", "person",
                 "than", "pay", "difficult", "heavy", "above"}

env = {}
for line in io.open(os.path.join(REPO, ".env.local"), encoding="utf-8"):
    line = line.strip()
    if line and not line.startswith("#") and "=" in line:
        k, _, v = line.partition("=")
        env[k.strip()] = v.strip().strip('"').strip("'")
URL = env["SUPABASE_URL"].rstrip("/")
KEY = env["SUPABASE_SERVICE_ROLE_KEY"]
H = {"apikey": KEY, "Authorization": "Bearer " + KEY, "User-Agent": "curl/8"}


def fetch(table, flt):
    out, off = [], 0
    while True:  # PostgREST caps at 1000; total order ending in id
        r = urllib.request.Request(
            URL + "/rest/v1/%s?select=*&%s&order=id&limit=1000&offset=%d" % (table, flt, off), headers=H)
        c = json.loads(urllib.request.urlopen(r, timeout=180).read())
        out += c
        if len(c) < 1000:
            return out
        off += 1000


def letters(s):
    """Arabic kaf/yeh (the dictionary's) -> Kurdish ک/ی. Spelling only, no word change."""
    s = (s or "").replace("ك", "ک").replace("ي", "ی").replace("ى", "ی")
    s = re.sub("[​‍‎‏ـ]", "", s)
    return " ".join(s.split())


# ------------------------------------------------------------ her answers
wb = openpyxl.load_workbook(SHEET, read_only=True)
rows = [r for r in list(wb["answers"].iter_rows(values_only=True))[1:] if r[0] == "tests-all"]
rows.sort(key=lambda r: r[13])  # latest submission wins for a repeated id
HER, TEST_EN = {}, {}
for r in rows:
    std, her = letters(r[7]), " ".join((r[9] or "").split())
    if std and her:
        HER[std] = her
        TEST_EN[std] = (r[5] or "").strip()

# ------------------------------------------------------------ deck
cards = fetch("flashcards", "course_id=eq.standard_kurdish")
cats = sorted(fetch("flashcard_categories", "course_id=eq.standard_kurdish"),
              key=lambda c: (c["sort_order"], c["name"]))
ci = {c["id"]: i for i, c in enumerate(cats)}

# The deck (from a printed dictionary) often writes ر for ڕ and ل for ڵ, so a
# word she did convert can miss an exact match: رەش / ڕەش. Two looser keys:
#   initial  - word-initial ر is always ڕ in Sorani, so this is a spelling fix;
#   folded   - ڕ->ر and ڵ->ل everywhere, which CAN join different words
#              (کەڕ deaf / کەر donkey), so it is only trusted with the meaning.
# Either way her word is applied only when the English meaning agrees.
def initial(s):
    return re.sub(r"(^|\s)ر", lambda m: m.group(1) + "ڕ", s)


def folded(s):
    return s.replace("ڕ", "ر").replace("ڵ", "ل")


SAME_SENSE = {"aeroplane": "plane"}


def senses(s):
    s = (s or "").lower()
    return {re.sub(r"^(to|the|a|an) ", "", x.strip()) for x in re.split(r"[/,]", s) if x.strip()}


def same_meaning(card_en, test_en):
    e = senses(card_en)
    e |= {SAME_SENSE[x] for x in e if x in SAME_SENSE}
    return bool(e & senses(test_en))


BY_INITIAL = {initial(k): k for k in HER}
BY_FOLDED = collections.defaultdict(list)
for k in HER:
    BY_FOLDED[folded(k)].append(k)

items, stats = [], collections.Counter()
for c in sorted(cards, key=lambda c: (ci.get(c["category_id"], 999), (c["english_text"] or "").lower())):
    p = letters(c["target_text"])
    en = (c["english_text"] or "").strip()
    it = {"id": "f" + c["id"].replace("-", "")[:8], "c": ci.get(c["category_id"], -1),
          "en": en, "ar": (c["arabic_text"] or "").strip(),
          "s": c["target_text"], "p": p}
    if p in HER and en.lower() not in SENSE_DIFFERS:
        it["h"] = HER[p]
        stats["to confirm: her word from the tests" + (" (changed)" if HER[p] != p else " (same as Standard)")] += 1
    elif p in HER:
        it["t"] = [HER[p], TEST_EN[p]]
        stats["to review, with a hint from the tests"] += 1
    elif initial(p) in BY_INITIAL:
        k = BY_INITIAL[initial(p)]
        if same_meaning(en, TEST_EN[k]):
            it["h"] = HER[k]
            stats["to confirm: her word from the tests (ر/ڕ spelling)"] += 1
        else:
            it["t"] = [HER[k], TEST_EN[k]]
            stats["to review, with a hint from the tests"] += 1
    else:
        hits = [k for k in BY_FOLDED.get(folded(p), []) if same_meaning(en, TEST_EN[k])]
        if len(hits) == 1:
            it["h"] = HER[hits[0]]
            stats["to confirm: her word from the tests (ر/ڕ ل/ڵ spelling)"] += 1
        else:
            stats["to review"] += 1
    # What she starts from: the Standard word with word-initial ر written ڕ,
    # the one spelling fix that can't change a word (ڕ is the rule there).
    if "h" in it:
        it["p"] = it["h"]
    elif initial(p) != p:
        it["p"] = initial(p)
        stats["  of which pre-filled with the ر -> ڕ spelling fix"] += 1
    items.append(it)

ids = [i["id"] for i in items]
assert len(ids) == len(set(ids)), "id collision"
secs = [{"name": c["name"], "ku": initial(letters(c["name_ku"])), "ar": c["name_ar"]} for c in cats]

data = {
    "batch": "flashcards",
    "label": "وشەکانی فلاشکارت — بۆ هەولێری",
    "sections": secs,
    "items": items,
}
json.dump(data, io.open(OUT, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
print("cards %d in %d categories" % (len(items), len(secs)))
for k, v in sorted(stats.items()):
    print("  %-45s %d" % (k, v))
print("->", OUT)
