"""
GSMPH localization-candidate extractor (v2)

Pipeline
  1. Every tool emits Candidate objects with CHARACTER OFFSETS (not bare strings).
  2. Overlaps are resolved by category priority + span length (no string-based dedup).
  3. Names/places/orgs are VOTED on by several independent sources (spaCy NER,
     a transformer NER, a first-name gazetteer) -> agreement = Tier 1, single vote = Tier 2.
  4. Generic nouns are filtered with WordNet (only artifact/food/animal/plant senses)
     and land in Tier 3 so they can't drown out the high-value candidates.
  5. Excel output has NO merged cells (so filter/sort work), a Decision dropdown,
     and a --precision-report mode that reads your reviewed workbook back and tells
     you which tool/category is producing false positives.

Usage
  python gsmph_candidates_v2.py --csv gsm8k.csv --limit 100 --out sample.xlsx
  (review sample.xlsx: set the Decision column to Keep / Reject)
  python gsmph_candidates_v2.py --precision-report sample.xlsx
"""
import argparse
import csv
import os
import re
from collections import defaultdict
from dataclasses import dataclass
from functools import lru_cache

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

try:
    from transformers import pipeline as hf_pipeline
except ImportError:
    hf_pipeline = None

try:
    import nltk
    from nltk.corpus import names as nltk_names, wordnet as wn
except ImportError:
    nltk = nltk_names = wn = None


# ============================================================================
# CONFIG
# ============================================================================
SPACY_MODEL = "en_core_web_trf"          # falls back to en_core_web_sm
HF_NER_MODEL = "dslim/bert-base-NER"     # PER / LOC / ORG / MISC (MISC ignored)
FOOD_MODEL = "Dizex/FoodBaseBERT-NER"
NER_MIN_SCORE = 0.85
FOOD_MIN_SCORE = 0.80
WN_SENSES = 2                            # how many WordNet senses to consult per noun
HF_BATCH = 32
GSM8K_CSV_PATH = "gsm8k.csv"
QUESTION_COLUMN_CANDIDATES = ("question", "problem", "text")


# ============================================================================
# WORD LISTS  (extend these as you review more of the dataset)
# ============================================================================
MONTHS = {"january", "february", "march", "april", "may", "june", "july",
          "august", "september", "october", "november", "december"}
DAYS = {"monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"}

KINSHIP = {"friend", "parent", "grandparent", "brother", "sister", "mother", "father",
           "mom", "dad", "mommy", "daddy", "son", "daughter", "cousin", "aunt", "uncle",
           "husband", "wife", "family", "kid", "child", "children", "classmate",
           "neighbor", "neighbour", "people", "person", "man", "woman", "boy", "girl"}
ABSTRACT = {"half", "quarter", "third", "purpose", "amount", "quantity", "total", "sum",
            "difference", "remainder", "percentage", "percent", "rate", "value", "number",
            "average", "ratio", "portion", "fraction"}
TIME_WORDS = {"yesterday", "today", "tomorrow", "hour", "minute", "day", "week", "month",
              "year", "shift", "overtime", "time", "second", "weekend", "morning", "night"}
GENERIC = {"thing", "item", "stuff", "piece", "type", "kind", "set", "pair", "lot",
           "bunch", "row", "side", "part", "end", "way", "rest", "dollar", "cent"}
EXCLUDED_NOUNS = KINSHIP | ABSTRACT | TIME_WORDS | GENERIC

# Lexicons: category -> (terms, tier). Case-insensitive, optional plural 's'.
LEXICONS = {
    "Event": ([
        "Christmas", "Christmas Eve", "Halloween", "Thanksgiving", "Easter",
        "Valentine's Day", "Valentine's", "New Year's Eve", "New Year's Day", "New Year",
        "Independence Day", "Fourth of July", "July 4th", "Memorial Day", "Labor Day",
        "Black Friday", "Cyber Monday", "Hanukkah", "Super Bowl", "Mother's Day",
        "Father's Day", "St. Patrick's Day", "Groundhog Day", "trick-or-treat",
        "trick or treat", "Easter egg hunt"], 2),
    "Culture": ([
        "bake sale", "garage sale", "yard sale", "lemonade stand", "prom", "homecoming",
        "spelling bee", "science fair", "talent show", "PTA", "varsity", "freshman",
        "sophomore", "kindergarten", "recess", "cafeteria", "locker", "baseball",
        "softball", "football", "hockey", "touchdown", "home run", "quarterback",
        "slumber party", "sleepover", "boy scouts", "girl scouts", "summer camp",
        "road trip", "babysitting", "babysitter"], 2),
    "Brand": ([
        "Walmart", "Costco", "McDonald's", "Starbucks", "Amazon", "eBay", "Netflix",
        "Disney", "Lego", "iPhone", "Xbox", "PlayStation", "Nintendo", "Barbie",
        "Pokemon", "Pokémon", "Subway", "Home Depot", "Whole Foods", "Walgreens", "CVS",
        "Kroger", "Trader Joe's", "Dunkin", "Pizza Hut", "KFC", "Domino's",
        "Burger King"], 2),
}


def _word_regex(terms):
    parts = sorted({re.escape(t).replace("'", "['’]") for t in terms}, key=len, reverse=True)
    return re.compile(r"\b(?:" + "|".join(parts) + r")s?\b", re.I)


LEXICON_RES = {cat: (_word_regex(terms), tier) for cat, (terms, tier) in LEXICONS.items()}

HONORIFIC_RE = re.compile(r"\b(?:Mr|Mrs|Ms|Mx|Dr|Prof|Miss)\b\.?(?=\s+[A-Z])")  # case-sensitive

CURRENCY_RES = [
    re.compile(r"\$\s?(?:\d[\d,]*(?:\.\d+)?|\.\d+)(?:\s?(?:thousand|million|billion))?", re.I),
    re.compile(r"(?<![\w$])\d[\d,]*(?:\.\d+)?\s?(?:dollars?|cents?|bucks)\b", re.I),
    re.compile(r"(?<![\w$])\d+(?:\.\d+)?\s?¢"),
]
CURRENCY_WORD_RE = re.compile(r"\b(?:dollars?|cents?|bucks?|nickels?|dimes?|pennies|penny)\b", re.I)

# US-customary units (a localization target when the Philippine version is metric).
UNIT_SAFE_RE = re.compile(
    r"\b(?:miles?|mph|gallons?|ounces?|inch(?:es)?|yards?|fahrenheit|acres?|"
    r"square feet|square miles)\b", re.I)
UNIT_DEG_RE = re.compile(r"°\s?F\b")
_NUMWORDS = ("a|an|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|"
             "twenty|thirty|forty|fifty|hundred|half")
# Ambiguous words ("foot", "pound", "pint"...) only count when attached to a quantity.
UNIT_NEEDS_NUM_RE = re.compile(
    r"(?:(?<![\w$.])\d[\d,./]*[\s-]?|\b(?:" + _NUMWORDS + r")[\s-]+)"
    r"(feet|foot|ft|pounds?|lbs?|oz|pints?|quarts?)\b", re.I)

SPACY_MAP = {"PERSON": "Proper Name", "GPE": "Place", "LOC": "Place", "FAC": "Place",
             "ORG": "Organization", "EVENT": "Event", "NORP": "Nationality/Group",
             "PRODUCT": "Product", "WORK_OF_ART": "Work of Art"}
HF_MAP = {"PER": "Proper Name", "LOC": "Place", "ORG": "Organization"}  # MISC deliberately dropped

OBJECT_LEXNAMES = {"noun.food": "Food", "noun.artifact": "Object",
                   "noun.animal": "Animal", "noun.plant": "Plant"}
BAD_ENTS = {"PERSON", "DATE", "TIME", "MONEY", "CARDINAL", "QUANTITY", "PERCENT", "ORDINAL"}
QUANTIFIERS = {"how", "what", "which", "many", "much", "more", "most", "less", "least",
               "twice", "enough", "several", "few", "remaining", "only", "final", "other",
               "another", "additional", "extra", "entire", "whole", "same", "each",
               "every", "any", "no", "all", "both", "new", "old"}

PRIORITY = {"Currency": 0, "Honorific": 1, "Proper Name": 2, "Event": 3, "Unit": 4,
            "Brand": 5, "Place": 5, "Organization": 5, "Product": 5,
            "Nationality/Group": 5, "Work of Art": 5, "Food": 6, "Culture": 6,
            "Object": 7, "Animal": 7, "Plant": 7}


# ============================================================================
# DATA
# ============================================================================
@dataclass
class Cand:
    start: int
    end: int
    text: str
    category: str
    sources: set
    tier: int            # 1 = high confidence, 2 = review, 3 = low-value / generic

    @property
    def span(self):
        return (self.start, self.end)


def load_problems(path=GSM8K_CSV_PATH, column=None):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Could not find '{path}'.")
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise ValueError(f"'{path}' has no header row.")
        col = column
        if col is None:
            lower = {n.lower(): n for n in reader.fieldnames}
            col = next((lower[c] for c in QUESTION_COLUMN_CANDIDATES if c in lower),
                       reader.fieldnames[0])
        problems = [(r.get(col) or "").strip() for r in reader]
    problems = [p for p in problems if p]
    if not problems:
        raise ValueError(f"No problem text found in '{path}' (column '{col}').")
    print(f"Loaded {len(problems)} problems from '{path}' (column '{col}').")
    return problems


# ============================================================================
# SPAN HELPERS
# ============================================================================
def expand_to_word(text, s, e):
    """Snap a span to full word boundaries (fixes '##' / mid-word fragments)."""
    while s > 0 and text[s - 1].isalnum():
        s -= 1
    while e < len(text) and text[e].isalnum():
        e += 1
    return s, e


def cluster(spans):
    """[(start, end, source)] -> [[start, end, {sources}]], merging overlapping/touching spans."""
    out = []
    for s, e, src in sorted(spans):
        if out and s <= out[-1][1]:
            out[-1][1] = max(out[-1][1], e)
            out[-1][2].add(src)
        else:
            out.append([s, e, {src}])
    return out


_HONOR_PREFIX = re.compile(r"^(?:Mr|Mrs|Ms|Mx|Dr|Prof|Miss)\b\.?\s+")
_POSSESSIVE = re.compile(r"['’]s$")


def trim_person(text, s, e):
    """Drop a leading honorific ('Dr. Smith' -> 'Smith') and trailing possessive."""
    m = _HONOR_PREFIX.match(text[s:e])
    if m:
        s += m.end()
    m = _POSSESSIVE.search(text[s:e])
    if m:
        e -= len(m.group())
    return s, e


def resolve(cands):
    """Overlap resolution: highest-priority category wins, then longest span.
    Identical spans+category from different tools are merged (their sources union)."""
    cands = sorted(cands, key=lambda c: (PRIORITY.get(c.category, 8), -(c.end - c.start), c.start))
    kept = []
    for c in cands:
        clash = [k for k in kept if c.start < k.end and k.start < c.end]
        if not clash:
            kept.append(c)
        elif len(clash) == 1 and clash[0].span == c.span and clash[0].category == c.category:
            clash[0].sources |= c.sources
            clash[0].tier = min(clash[0].tier, c.tier)
    return sorted(kept, key=lambda c: c.start)


# ============================================================================
# WORDNET OBJECT FILTER
# ============================================================================
@lru_cache(maxsize=None)
def wn_category(lemma):
    if wn is None:
        return None
    for syn in wn.synsets(lemma, pos="n")[:WN_SENSES]:
        cat = OBJECT_LEXNAMES.get(syn.lexname())
        if cat:
            return cat
    return None


def is_object_noun(tok):
    if tok.is_stop or tok.ent_type_ in BAD_ENTS or not tok.text.isalpha():
        return False
    return not ({tok.text.lower(), tok.lemma_.lower()} & EXCLUDED_NOUNS)


# ============================================================================
# CANDIDATE COLLECTION
# ============================================================================
def collect(text, doc, hf_ents, food_ents, names=frozenset()):
    cands = []

    def add(s, e, cat, src, tier):
        if s < e:
            cands.append(Cand(s, e, text[s:e], cat, {src}, tier))

    # ---- rule-based, high precision ----
    for m in HONORIFIC_RE.finditer(text):
        add(m.start(), m.end(), "Honorific", "regex", 1)
    for rx in CURRENCY_RES:
        for m in rx.finditer(text):
            add(m.start(), m.end(), "Currency", "regex", 1)
    for m in CURRENCY_WORD_RE.finditer(text):
        add(m.start(), m.end(), "Currency", "regex-word", 2)
    for rx in (UNIT_SAFE_RE, UNIT_DEG_RE):
        for m in rx.finditer(text):
            add(m.start(), m.end(), "Unit", "regex", 1)
    for m in UNIT_NEEDS_NUM_RE.finditer(text):
        add(m.start(1), m.end(1), "Unit", "regex", 1)
    for cat, (rx, tier) in LEXICON_RES.items():
        for m in rx.finditer(text):
            add(m.start(), m.end(), cat, "lexicon", tier)

    # ---- NER with voting (spaCy + transformer + gazetteer) ----
    spans = defaultdict(list)
    for ent in doc.ents:
        cat = SPACY_MAP.get(ent.label_)
        if cat:
            spans[cat].append((ent.start_char, ent.end_char, "spacy"))
    for ent in hf_ents:
        cat = HF_MAP.get(ent.get("entity_group"))
        if cat and ent["score"] >= NER_MIN_SCORE:
            s, e = expand_to_word(text, ent["start"], ent["end"])
            spans[cat].append((s, e, "hf"))
    if names:
        for tok in doc:
            if tok.pos_ == "PROPN" and tok.text.istitle() and tok.text in names:
                spans["Proper Name"].append((tok.idx, tok.idx + len(tok.text), "gazetteer"))

    for cat, items in spans.items():
        for s, e, srcs in cluster(items):
            if cat == "Proper Name":
                s, e = trim_person(text, s, e)
                if s >= e:
                    continue
            word = text[s:e]
            has_model = bool(srcs & {"spacy", "hf"})
            # "May", "April", "Monday"... are names only if a real NER model says so.
            if word.lower() in (MONTHS | DAYS) and not (cat == "Proper Name" and has_model):
                continue
            tier = 1 if len(srcs) >= 2 else 2
            cands.append(Cand(s, e, word, cat, set(srcs), tier))

    # ---- food model ----
    food_spans = []
    for ent in food_ents:
        if ent["score"] >= FOOD_MIN_SCORE:
            s, e = expand_to_word(text, ent["start"], ent["end"])
            food_spans.append((s, e, "foodbert"))
    for s, e, _ in cluster(food_spans):
        add(s, e, "Food", "foodbert", 2)

    # ---- generic objects via WordNet (low tier) ----
    for tok in doc:
        if tok.pos_ != "NOUN" or not is_object_noun(tok):
            continue
        cat = wn_category(tok.lemma_.lower())
        if not cat:
            continue
        start = tok.i
        # extend left over noun-noun compounds only ("care package", "boarding school")
        while (start > 0 and doc[start - 1].dep_ == "compound"
               and doc[start - 1].head.i >= start
               and doc[start - 1].pos_ in ("NOUN", "VERB", "ADJ")
               and doc[start - 1].text.lower() not in QUANTIFIERS):
            start -= 1
        add(doc[start].idx, tok.idx + len(tok.text), cat, "wordnet", 2 if cat == "Food" else 3)

    return resolve(cands)


# ============================================================================
# MODEL LOADING
# ============================================================================
def load_spacy():
    import spacy
    for name in (SPACY_MODEL, "en_core_web_sm"):
        try:
            nlp = spacy.load(name)
            print(f"spaCy model: {name}")
            return nlp
        except Exception as exc:
            print(f"[warn] could not load spaCy model {name}: {exc}")
    raise SystemExit("No spaCy model available. Run: python -m spacy download en_core_web_sm")


def load_hf(model_name):
    if hf_pipeline is None:
        print(f"[warn] transformers not installed; skipping {model_name}")
        return None
    try:
        import torch
        device = 0 if torch.cuda.is_available() else -1
    except ImportError:
        device = -1
    return hf_pipeline("token-classification", model=model_name,
                       aggregation_strategy="simple", device=device)


def run_hf(pipe, texts):
    if pipe is None:
        return [[] for _ in texts]
    try:
        return pipe(texts, batch_size=HF_BATCH)
    except Exception as exc:
        print(f"[warn] HF pipeline failed ({exc}); continuing without it")
        return [[] for _ in texts]


def load_nltk_resources():
    names, ok_wn = set(), False
    if nltk is None:
        print("[warn] nltk not installed: no name gazetteer, no WordNet object filter")
        return names
    for pkg in ("wordnet", "names"):
        try:
            wn.synsets("dog") if pkg == "wordnet" else nltk_names.words()
        except LookupError:
            nltk.download(pkg, quiet=True)
    try:
        names = set(nltk_names.words())
    except LookupError:
        print("[warn] NLTK 'names' corpus unavailable")
    return names


# ============================================================================
# REPORTING
# ============================================================================
HEADERS = ["Problem #", "Candidate", "Context", "Category", "Tier", "Sources",
           "Count", "Decision", "Final Category", "Localized Replacement"]
TIER_FILL = {1: "E2EFDA", 2: "FFF2CC", 3: "EDEDED"}


def context_snippet(text, c, width=35):
    s, e = max(0, c.start - width), min(len(text), c.end + width)
    return ("…" if s else "") + text[s:c.start] + "[" + text[c.start:c.end] + "]" + \
           text[c.end:e] + ("…" if e < len(text) else "")


def build_rows(problems, results):
    rows = []
    for i, (text, cands) in enumerate(zip(problems, results), 1):
        groups = {}
        for c in cands:  # one row per unique (term, category) per problem
            key = (c.text.lower(), c.category)
            if key in groups:
                g = groups[key]
                g[1] += 1
                g[0].sources |= c.sources
                g[0].tier = min(g[0].tier, c.tier)
            else:
                groups[key] = [c, 1]
        for c, count in sorted(groups.values(), key=lambda g: (g[0].tier, g[0].start)):
            rows.append((i, c.text, context_snippet(text, c), c.category, c.tier,
                         ", ".join(sorted(c.sources)), count))
    return rows


def write_report(problems, results, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Candidates"
    ws.append(HEADERS)
    for r in build_rows(problems, results):
        ws.append(list(r) + ["", "", ""])

    head_fill = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = head_fill
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        t = row[4].value
        if t in TIER_FILL:
            row[4].fill = PatternFill("solid", start_color=TIER_FILL[t], end_color=TIER_FILL[t])
    for col, w in zip("ABCDEFGHIJ", (10, 22, 60, 18, 6, 22, 7, 11, 18, 24)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    if ws.max_row >= 2:
        dv = DataValidation(type="list", formula1='"Keep,Reject,Unsure"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"H2:H{ws.max_row}")

    ws2 = wb.create_sheet("Per Problem (Tier 1-2)")
    ws2.append(["Problem #", "Problem Text", "Candidates (Tier 1-2)"])
    for cell in ws2[1]:
        cell.font = Font(bold=True)
        cell.fill = head_fill
    by_problem = defaultdict(lambda: defaultdict(list))
    for i, term, _, cat, tier, _, _ in build_rows(problems, results):
        if tier <= 2:
            by_problem[i][cat].append(term)
    for i, text in enumerate(problems, 1):
        lines = [f"{cat}: {', '.join(terms)}" for cat, terms in sorted(by_problem[i].items())]
        ws2.append([i, text, "\n".join(lines) or "(none)"])
        r = ws2.max_row
        for c in range(1, 4):
            ws2.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical="top")
        ws2.row_dimensions[r].height = max(30, 15 * max(1, len(lines)))
    ws2.column_dimensions["A"].width = 10
    ws2.column_dimensions["B"].width = 60
    ws2.column_dimensions["C"].width = 60
    ws2.freeze_panes = "A2"

    wb.save(path)
    print(f"Saved: {path}")


def precision_report(path):
    """Read a reviewed workbook (Decision = Keep/Reject) and print precision
    per category, tier, and source tool. This is how you find which tool to
    tighten or drop. (Recall needs a hand-labeled gold set; see notes.)"""
    ws = load_workbook(path, read_only=True)["Candidates"]
    rows = ws.iter_rows(values_only=True)
    ix = {h: i for i, h in enumerate(next(rows))}
    tables = {"category": defaultdict(lambda: [0, 0]),
              "tier": defaultdict(lambda: [0, 0]),
              "source": defaultdict(lambda: [0, 0])}
    for r in rows:
        d = r[ix["Decision"]]
        if d not in ("Keep", "Reject"):
            continue
        k = 0 if d == "Keep" else 1
        tables["category"][r[ix["Category"]]][k] += 1
        tables["tier"][r[ix["Tier"]]][k] += 1
        for s in str(r[ix["Sources"]]).split(", "):
            tables["source"][s][k] += 1
    for name, table in tables.items():
        print(f"\nPrecision by {name}")
        print(f"{'':24}{'keep':>6}{'reject':>8}{'precision':>11}")
        for key, (keep, rej) in sorted(table.items(), key=lambda kv: str(kv[0])):
            print(f"{str(key):24}{keep:>6}{rej:>8}{keep / (keep + rej):>11.0%}")


# ============================================================================
# MAIN
# ============================================================================
def extract(problems):
    nlp = load_spacy()
    names = load_nltk_resources()
    hf_ner = load_hf(HF_NER_MODEL)
    food_ner = load_hf(FOOD_MODEL)

    print("Running spaCy...")
    docs = list(nlp.pipe(problems, batch_size=64))
    print("Running transformer NER...")
    hf_out = run_hf(hf_ner, problems)
    print("Running food NER...")
    food_out = run_hf(food_ner, problems)
    return [collect(t, d, h, f, names) for t, d, h, f in zip(problems, docs, hf_out, food_out)]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", default=GSM8K_CSV_PATH)
    ap.add_argument("--column")
    ap.add_argument("--out", default="gsmph_candidates.xlsx")
    ap.add_argument("--limit", type=int, help="only process the first N problems")
    ap.add_argument("--precision-report", metavar="REVIEWED_XLSX")
    args = ap.parse_args()

    if args.precision_report:
        precision_report(args.precision_report)
        return
    problems = load_problems(args.csv, args.column)
    if args.limit:
        problems = problems[:args.limit]
    write_report(problems, extract(problems), args.out)


if __name__ == "__main__":
    main()
