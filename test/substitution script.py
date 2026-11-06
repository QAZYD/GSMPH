#!/usr/bin/env python3
"""
GSMPH pipeline v2 -- full-workbook build, aligned with the updated pool taxonomy.

Replaces gsmph_pipeline_demo_prototype.py. The old flow produced a 2-sheet demo
("Review - Object?" categories, no pools). This one produces the same 8-sheet
workbook as full_run_enhanced_npscan_cleaned.xlsx and routes every candidate
through the pool workbook's Subcategory Map:

    candidate -> Specific Subcategory -> Parent Group -> Handling + Pool/rule
              -> Group ID (one replacement per concept per problem)
              -> suggested replacement(s) drawn from the Pool sheet

Commands
  build-lexicon  Learn term -> subcategory data from a reviewed full-run workbook
                 (honours 'Corrected subcategory' filled in the Review Queue).
  run            Build the full workbook.
                   --candidates-from FULL.xlsx   reuse its candidate rows/groups (no NLP stack needed)
                   (omit it)                     run Stage 1 extraction (gsmph_extract.py) on --problems
  selftest       Hold-out check of the classifier (lexicon built on 80% of problems,
                 scored on the other 20%).

Examples
  python gsmph_pipeline.py build-lexicon --seed full_run_enhanced_npscan_cleaned.xlsx --out gsmph_lexicon.json
  python gsmph_pipeline.py run --pool pool.xlsx --lexicon gsmph_lexicon.json \
        --candidates-from full_run_enhanced_npscan_cleaned.xlsx --out gsmph_full_run_v2.xlsx
  python gsmph_pipeline.py run --pool pool.xlsx --lexicon gsmph_lexicon.json \
        --problems gsm8k.csv --out gsmph_full_run_v2.xlsx
  python gsmph_pipeline.py selftest --seed full_run_enhanced_npscan_cleaned.xlsx
"""
import argparse
import csv
import hashlib
import json
import random
import re
from collections import Counter, defaultdict

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
HDR = PatternFill("solid", start_color="D9E1F2")
YELLOW = PatternFill("solid", start_color="FFF2CC")
GREY = PatternFill("solid", start_color="D9D9D9")
CURRENCY_GROUP_LABEL = "USD currency: all $ amounts and 'dollars' (convert together)"
CURRENCY_FAMILY_SUBCAT = "currency_usd_dollar (family)"

# Amount bands (derived from the cleaned run: small <20, medium 20-199, large 200-1999, very large >=2000)
CURRENCY_BANDS = [(20, "small"), (200, "medium"), (2000, "large")]
CURRENCY_SCALE = 1.0          # peso amount = USD amount * scale. 1.0 keeps the arithmetic/answers unchanged.

COIN_WORDS = {"cent": "currency_usd_coin_cent_penny", "cents": "currency_usd_coin_cent_penny",
              "penny": "currency_usd_coin_cent_penny", "pennies": "currency_usd_coin_cent_penny",
              "nickel": "currency_usd_coin_nickel_dime_quarter", "nickels": "currency_usd_coin_nickel_dime_quarter",
              "dime": "currency_usd_coin_nickel_dime_quarter", "dimes": "currency_usd_coin_nickel_dime_quarter",
              "quarter": "currency_usd_coin_nickel_dime_quarter", "quarters": "currency_usd_coin_nickel_dime_quarter"}
UNIT_WORDS = {
    **dict.fromkeys(["pound", "pounds", "lb", "lbs", "ounce", "ounces", "oz"], "unit_weight_pound_ounce"),
    **dict.fromkeys(["foot", "feet", "ft", "inch", "inches", "yard", "yards", "yd"], "unit_length_foot_inch_yard"),
    **dict.fromkeys(["mile", "miles"], "unit_length_mile"),
    **dict.fromkeys(["gallon", "gallons", "quart", "quarts", "pint", "pints"], "unit_volume_gallon_quart_pint"),
    **dict.fromkeys(["acre", "acres"], "unit_area_acre_sqft"),
    **dict.fromkeys(["mph"], "unit_speed_mph"),
    **dict.fromkeys(["fahrenheit"], "unit_temperature_fahrenheit"),
}
# Rule: imperial -> metric WORD. The numbers must be re-derived by a person (answers change).
METRIC_WORD = {"pound": "kilogram", "lb": "kg", "ounce": "gram", "oz": "g", "foot": "meter", "feet": "meters",
               "ft": "m", "inch": "centimeter", "yard": "meter", "mile": "kilometer", "gallon": "liter",
               "quart": "liter", "pint": "liter", "acre": "hectare", "mph": "kph", "fahrenheit": "Celsius"}
HONORIFIC_SUBCAT = {"mr": "honorific_mr_male_formal", "mrs": "honorific_mrs_female_married",
                    "ms": "honorific_ms_miss_female", "miss": "honorific_ms_miss_female",
                    "dr": "honorific_doctor_professional"}
HONORIFIC_ENGLISH = {"honorific_mr_male_formal": {"Mr.", "Mr. (abbreviation)", "polite address to a man"},
                     "honorific_mrs_female_married": {"Mrs.", "Mrs. (abbreviation)"},
                     "honorific_ms_miss_female": {"Miss/Ms.", "Miss/Ms. (abbreviation)", "polite address to a woman"},
                     "honorific_doctor_professional": {"Doctor", "woman doctor", "attorney", "engineer"}}

# Subcategories whose pools hold SPECIFIC titles/events, not generic words. For these, a generic
# noun ("book", "video", "babysitting") is only ever replaced by (a) a curated generic Tagalog word
# below, or (b) a pool term whose English/description actually matches the word. Otherwise the group
# goes to review with no suggestion, instead of drawing a random title.
SEMANTIC_ONLY_PREFIXES = ("media_", "culture_")
# CURATED BY THE PIPELINE AUTHOR, NOT FROM THE POOL: have a Tagalog speaker verify before use.
GENERIC_TAGALOG = {
    "book": "aklat", "novel": "nobela", "magazine": "magasin", "comic book": "komiks", "comic": "komiks",
    "coloring book": "librong pangkulay", "diary": "talaarawan", "bible": "bibliya", "textbook": "aklat-aralin",
    "video": "bidyo", "clip": "klip", "album": "album", "dvd": "DVD", "record": "plaka", "movie": "pelikula",
    "film": "pelikula", "song": "kanta", "music": "musika", "channel": "istasyon", "sci fi movie": "pelikulang sci-fi",
    "picture": "larawan", "photo": "litrato", "painting": "pinta", "drawing": "guhit", "portrait": "retrato",
    "canvas": "kanbas", "mural": "mural", "sculpture": "eskultura", "art": "sining", "photo album": "album ng larawan",
    "babysitting": "pag-aalaga ng bata", "babysitter": "yaya",
}
SYNONYMS = {"babysitter": ["nanny"], "babysitting": ["nanny", "caregiver"]}
STOP_CUE = set("the a an and of to in on at for with by from is are was were he she it they his her their "
               "has have had that this these those as be been how many much what which each per".split())
GENERIC_BROAD = {"vehicle", "animal", "pet", "food", "fruit", "vegetable", "tool", "toy", "drink", "snack",
                 "dessert", "flower", "tree", "building", "store", "shop", "container", "device", "machine",
                 "plant", "bird", "fish", "meal", "sport", "instrument", "furniture", "appliance"}
TIER_RANK = {"Curated": 0, "High": 0, "Medium": 1, "Low": 2, "Dictionary (unrated)": 3}


# ---------------------------------------------------------------------------
# Small text helpers
# ---------------------------------------------------------------------------
def norm(term):
    t = re.sub(r"[\"'“”‘’\[\]()]", "", str(term)).lower().strip()
    return re.sub(r"\s+", " ", t)


def lemma(word):
    w = word.lower()
    if len(w) > 4 and w.endswith("ies"):
        return w[:-3] + "y"
    if len(w) > 4 and w.endswith(("ches", "shes", "sses", "xes", "zes")):
        return w[:-2]
    if len(w) > 3 and w.endswith("s") and not w.endswith(("ss", "us", "is")):
        return w[:-1]
    return w


def lemma_phrase(term):
    parts = norm(term).split()
    return " ".join(parts[:-1] + [lemma(parts[-1])]) if parts else ""


def stable_rng(key):
    return random.Random(int(hashlib.md5(str(key).encode()).hexdigest()[:8], 16))


# ---------------------------------------------------------------------------
# Taxonomy (Subcategory Map in the pool workbook is the single source of truth)
# ---------------------------------------------------------------------------
def load_taxonomy(pool_xlsx):
    m = pd.read_excel(pool_xlsx, sheet_name="Subcategory Map")
    tax = {}
    for _, r in m.iterrows():
        tax[r["Specific Subcategory"]] = {
            "parent": r["Parent Group"] if pd.notna(r["Parent Group"]) else "",
            "handling": r["Handling"], "pool": r["Pool / rule"]}
    return tax


def load_pool(pool_xlsx):
    p = pd.read_excel(pool_xlsx, sheet_name="Pool")
    p = p[~p["Reviewer status"].astype(str).str.lower().isin(["reject", "rejected"])].copy()
    p["_rank"] = p["Recognition tier"].map(TIER_RANK).fillna(3)
    p["_score"] = pd.to_numeric(p["Score"], errors="coerce").fillna(0)
    p = p.sort_values(["Pool", "_rank", "_score"], ascending=[True, True, False])
    p["_serves"] = p["Serves subcategories (cleaned run)"].fillna("").map(lambda s: {x.strip() for x in s.split(",") if x.strip()})
    p["_binomial"] = (p["Term (use)"].astype(str).str.match(r"^[A-Z][a-z]+ [a-z]+$")
                      & (p["Term (use)"].astype(str) == p["English"].astype(str)))
    d = p["Description"].fillna("").str.lower()
    p["_gender"] = "unknown"
    p.loc[d.str.contains(r"\bmale\b|\bman\b|\bboy\b") & ~d.str.contains("female"), "_gender"] = "male"
    p.loc[d.str.contains(r"female|woman|girl"), "_gender"] = "female"
    return p


# ---------------------------------------------------------------------------
# Lexicon (learned from a reviewed full-run workbook)
# ---------------------------------------------------------------------------
def build_lexicon(cand, review_queue=None, tax=None):
    """cand: Candidates dataframe. Only High/Medium rows teach the lexicon; a
    'Corrected subcategory' typed in the Review Queue overrides the row's label."""
    cand = cand.copy()
    if review_queue is not None and "Corrected subcategory" in review_queue:
        fix = review_queue.dropna(subset=["Corrected subcategory"])
        fixmap = {(int(r["Problem #"]), norm(r["Candidate"])): r["Corrected subcategory"] for _, r in fix.iterrows()}
        cand["Specific Subcategory"] = [fixmap.get((int(p), norm(c)), s) for p, c, s in
                                        zip(cand["Problem #"], cand["Candidate"], cand["Specific Subcategory"])]
    good = cand[cand["Subcat Confidence"].isin(["High", "Medium"]) & ~cand["Specific Subcategory"].str.endswith("_review")]
    terms = defaultdict(Counter)
    heads = defaultdict(Counter)
    names = defaultdict(Counter)
    cues = defaultdict(Counter)
    for term, sub, ctx in zip(good["Candidate"], good["Specific Subcategory"], good["Context"]):
        if sub.startswith(("name_", "currency_", "unit_", "honorific_")):
            if sub.startswith("name_"):
                names[norm(term)][sub] += 1
            continue
        n = lemma_phrase(term)
        terms[n][sub] += 1
        heads[n.split()[-1]][sub] += 1
        for w in re.findall(r"[a-z]{3,}", str(ctx).lower()):
            if w not in STOP_CUE and w != n:
                cues[sub][w] += 1
    sub_cat = (cand.groupby("Specific Subcategory")["Suggested Final Category"].agg(lambda s: s.value_counts().index[0]).to_dict())
    return {
        "terms": {t: dict(c) for t, c in terms.items()},
        "heads": {h: dict(c) for h, c in heads.items()},
        "names": {n: dict(c) for n, c in names.items()},
        "cues": {s: [w for w, _ in c.most_common(80)] for s, c in cues.items()},
        "subcat_category": sub_cat,
    }


def cue_pick(alts, text, lex):
    words = set(re.findall(r"[a-z]{3,}", str(text).lower()))
    scored = sorted(((len(words & set(lex["cues"].get(s, []))), n, s) for s, n in alts.items()), reverse=True)
    return scored[0][2], scored[0][0] > 0


try:
    import gender_guesser.detector as _gg
    _GG = _gg.Detector()
except Exception:                                          # noqa: BLE001
    _GG = None


def name_subcat(term, lex):
    toks = norm(term).split()
    if not toks:
        return None
    first = lex["names"].get(toks[0]) or lex["names"].get(" ".join(toks))
    gender = None
    if not first and _GG is not None:                      # optional: same library the cleaned run used
        g = _GG.get_gender(toks[0].capitalize())
        if g in ("male", "mostly_male"):
            return ("name_fullname_male" if len(toks) > 1 else "name_given_male"), True
        if g in ("female", "mostly_female"):
            return ("name_fullname_female" if len(toks) > 1 else "name_given_female"), True
        if g == "andy":
            return "name_given_unisex", True
    if first:
        g = max(first, key=first.get)
        gender = g.split("_")[-1] if g.split("_")[-1] in ("male", "female", "unisex") else "unknown"
        if "gender_unknown" in g:
            gender = "unknown"
    if len(toks) >= 2:
        return f"name_fullname_{'gender_unknown' if gender in (None, 'unisex', 'unknown') else gender}", bool(first)
    if gender is None:
        return "name_given_gender_unknown", False
    return ("name_given_gender_unknown" if gender == "unknown" else f"name_given_{gender}"), True


def classify(term, tag, context, lex, tax):
    """Return (specific_subcategory, confidence, method)."""
    t = str(term).strip()
    nt = norm(t)
    low = nt.rstrip(".")
    # ---- patterns: currency / units / honorifics -------------------------------------------
    if "$" in t:
        m = re.search(r"\d[\d,]*\.?\d*|\.\d+", t)
        if m:
            raw = m.group().replace(",", "")
            val = float(raw)
            if raw.startswith(".") or (float(raw) % 1 != 0):
                return "currency_usd_amount_decimal", "High", "pattern"
            for cap, band in CURRENCY_BANDS:
                if val < cap:
                    return f"currency_usd_amount_whole_{band}", "High", "pattern"
            return "currency_usd_amount_whole_very_large", "High", "pattern"
    words = nt.split()
    if words and words[-1] in COIN_WORDS and re.search(r"\d", nt):
        return COIN_WORDS[words[-1]], "High", "pattern"
    if words and words[-1] in ("dollar", "dollars") :
        return "currency_usd_word_dollar", "High", "pattern"
    if low in HONORIFIC_SUBCAT:
        return HONORIFIC_SUBCAT[low], "High", "pattern"
    if len(words) <= 2 and words and words[-1] in UNIT_WORDS and low not in lex["terms"]:
        return UNIT_WORDS[words[-1]], "High", "pattern"
    # ---- lexicon: exact phrase, then head noun -----------------------------------------------
    lp = lemma_phrase(t)
    if lp in lex["terms"]:
        alts = lex["terms"][lp]
        total = sum(alts.values())
        top = max(alts, key=alts.get)
        if alts[top] / total >= 0.8:
            return top, "High" if total >= 2 else "Medium", "lexicon:phrase"
        sub, hit = cue_pick(alts, context, lex)
        return sub, "Medium" if hit else "Low", "cue:problem-text"
    # ---- person names -------------------------------------------------------------------------
    if t[:1].isupper() and (tag == "PERSON" or norm(t).split()[0] in lex["names"]):
        sub, known = name_subcat(t, lex)
        return sub, "High" if known else "Medium", "name-list" if known else "name-unknown"
    head = lp.split()[-1] if lp else ""
    if head in lex["heads"]:
        alts = lex["heads"][head]
        total = sum(alts.values())
        top = max(alts, key=alts.get)
        if alts[top] / total >= 0.8:
            return top, "Medium", "lexicon:head"
        sub, hit = cue_pick(alts, context, lex)
        return sub, "Medium" if hit else "Low", "cue:problem-text"
    # ---- fallbacks -------------------------------------------------------------------------------
    if tag == "LOCATION":
        return "place_geo_other_review", "Low", "ner-tag"
    if tag == "ORGANIZATION" or t[:1].isupper():
        return "org_proper_unclassified_review", "Low", "none"
    return "unclassified_review", "Low", "none"


# ---------------------------------------------------------------------------
# Grouping: one Group ID per concept per problem (ordered by first appearance)
# ---------------------------------------------------------------------------
def group_key(term, sub, tax):
    parent = tax.get(sub, {}).get("parent", "")
    if parent == "Currency":
        return ("CUR",)
    if sub.startswith("name_"):
        return ("NAME", norm(term).split()[-1] if sub.startswith("name_fullname") else norm(term))
    if sub.startswith("honorific_"):
        return ("HON", norm(term).rstrip("."))
    if sub.startswith("unit_"):
        return ("UNIT", lemma(norm(term).split()[-1]))
    lp = lemma_phrase(term)
    return ("N", lp.split()[-1] if lp else norm(term), sub)


def build_groups(rows, text, tax):
    """rows: list of dicts with candidate/sub. Adds 'group_id' / 'group_label'; returns group dicts."""
    pid = rows[0]["problem"]
    for r in rows:
        i = text.lower().find(norm(r["candidate"]).strip(" $"))
        r["_pos"] = i if i >= 0 else 10 ** 6
    order, keymap = [], {}
    for r in sorted(rows, key=lambda r: r["_pos"]):
        k = group_key(r["candidate"], r["sub"], tax)
        if k not in keymap:
            keymap[k] = f"{pid}-G{len(keymap) + 1}"
            order.append(k)
        r["group_id"] = keymap[k]
    groups = {}
    for r in rows:
        g = groups.setdefault(r["group_id"], {"id": r["group_id"], "problem": pid, "rows": [], "key": None})
        g["rows"].append(r)
    for k, gid in keymap.items():
        groups[gid]["key"] = k
    for g in groups.values():
        k = g["key"]
        g["label"] = CURRENCY_GROUP_LABEL if k[0] == "CUR" else (k[1] if k[0] != "N" else k[1])
        for r in g["rows"]:
            r["group_label"] = g["label"]
    return groups


# ---------------------------------------------------------------------------
# Replacement suggestions from the Pool sheet
# ---------------------------------------------------------------------------
class Suggester:
    def __init__(self, tax, pool_df):
        self.tax, self.pool = tax, pool_df
        self.by_pool = {p: d for p, d in pool_df.groupby("Pool")}

    def _eligible(self, pool_name, sub):
        d = self.by_pool.get(pool_name)
        if d is None or d.empty:
            return d
        served = d[d["_serves"].map(lambda s: sub in s)]
        return served if len(served) >= 3 else d

    def suggest(self, sub, forms, gid, used, label=""):
        """Return (primary, alt1, alt2, basis) for one group."""
        info = self.tax.get(sub.replace(" (family)", ""), None)
        if info is None:
            return "", "", "", "Subcategory not in Subcategory Map: manual"
        h, pool_name = info["handling"], info["pool"]
        forms_l = [norm(f) for f in forms]
        plural = any(f.endswith("s") and lemma(f) != f for f in forms_l)
        if h.startswith("Keep"):
            return "", "", "", "KEEP: no substitution"
        if h == "Manual review":
            return "", "", "", "REVIEW: classify manually before choosing a replacement"
        if h == "Rule-based":
            return self._rule(sub, forms), "", "", str(pool_name)
        d = self._eligible(pool_name, sub)
        if d is None or d.empty:
            return "", "", "", f"No usable terms in pool '{pool_name}'"
        if sub.startswith(SEMANTIC_ONLY_PREFIXES) or pool_name == "Books & media":
            return self._semantic_only(sub, label, forms, gid, d, used, pool_name)
        if sub.startswith("name_"):
            g = "male" if sub.endswith("male") and "female" not in sub else ("female" if "female" in sub else None)
            exact = d[d["_gender"] == g] if g else d
            d = exact if len(exact) >= 3 else d
        if sub.startswith("honorific_"):
            eng = HONORIFIC_ENGLISH.get(sub, set())
            d2 = d[d["English"].isin(eng)]
            d = d2 if len(d2) else d
            abbr = any("." in str(f) for f in forms)
            pref = d["Term (use)"].astype(str).str.contains(r"\.$") == abbr
            d = pd.concat([d[pref], d[~pref]])
        if sub == "currency_usd_word_dollar":
            d2 = d[d["Term (use)"].isin(["pesos" if plural or any(re.search(r"\d", f) and not re.fullmatch(r"1 dollar", f) for f in forms_l) else "piso"])]
            d = d2 if len(d2) else d
        if not sub.startswith(("name_", "honorific_")):
            nb = d[~d["_binomial"]]
            d = nb if len(nb) >= 5 else d                       # Latin binomials are poor stand-ins
        sem = []
        lab = re.sub(r"[^a-z ]", "", str(label).lower()).strip()
        if lab and not sub.startswith(("name_", "honorific_", "currency_")):
            pat = r"\b" + re.escape(lemma(lab.split()[-1])) + r"s?\b"
            hit = d[d["English"].astype(str).str.lower().str.contains(pat) | d["Description"].astype(str).str.lower().str.contains(pat)]
            sem = [t for t in hit["Term (use)"].astype(str) if t not in used][:3]
        top = d.head(30)
        terms = [t for t in top["Term (use)"].astype(str) if t not in used and t not in sem]
        if not terms and not sem:
            terms = list(top["Term (use)"].astype(str))
        stable_rng(gid).shuffle(terms)
        picks = (sem + terms + ["", "", ""])[:3]
        if picks[0]:
            used.add(picks[0])
        tier = d.iloc[0]["Recognition tier"] if len(d) else ""
        basis = (f"Pool '{pool_name}' ({len(d)} usable terms; "
                 + ("first pick matches the word in the pool's English/description; " if sem else "")
                 + "others: top-30 by recognition, seeded pick)")
        if plural:
            basis += "; plural in problem -> prefix 'mga'"
        return picks[0], picks[1], picks[2], basis

    def _semantic_only(self, sub, label, forms, gid, d, used, pool_name):
        lab = re.sub(r"[^a-z ]", "", str(label).lower()).strip()
        key = lemma_phrase(lab) if lab else ""
        words = [key, lab] + SYNONYMS.get(key, []) + SYNONYMS.get(lab, [])   # full phrase only: no last-word match
        picks, basis = [], ""
        last = lab.split()[-1] if lab else ""
        generic = (GENERIC_TAGALOG.get(key) or GENERIC_TAGALOG.get(lab)
                   or GENERIC_TAGALOG.get(last) or GENERIC_TAGALOG.get(lemma(last)))
        if generic:
            picks.append(generic)
            basis = "curated generic Tagalog word (not from the pool; verify)"
        for w in dict.fromkeys(w for w in words if w):
            pat = r"\b" + re.escape(w) + r"s?\b"
            hit = d[d["English"].astype(str).str.lower().str.contains(pat)]
            picks += [t for t in hit["Term (use)"].astype(str) if t not in picks and t not in used]
        picks = (picks + ["", "", ""])[:3]
        if not picks[0]:
            return "", "", "", f"REVIEW: pool '{pool_name}' has no generic term for '{label}' (it holds specific titles/events)"
        used.add(picks[0])
        plural = any(norm(f).endswith("s") and lemma(norm(f)) != norm(f) for f in forms)
        return picks[0], picks[1], picks[2], basis + ("; pool terms matched on English meaning" if len(picks) > 1 and picks[1] else "") + ("; plural in problem -> prefix 'mga'" if plural else "")

    @staticmethod
    def _rule(sub, forms):
        if sub.startswith("currency") or "family" in sub:
            out = []
            for f in forms:
                s = str(f).replace("$", "₱")
                s = re.sub(r"\bdollars?\b", lambda m: "piso" if m.group().lower() == "dollar" else "pesos", s, flags=re.I)
                s = re.sub(r"\bcents?\b", "sentimo", s, flags=re.I)
                out.append(s)
            note = "" if CURRENCY_SCALE == 1 else f" (x{CURRENCY_SCALE})"
            return " | ".join(out) + note + "  [symbol/word swap only; amounts NOT rescaled]"
        out = []
        for f in forms:
            w = norm(f).split()[-1] if norm(f) else ""
            out.append(METRIC_WORD.get(w, METRIC_WORD.get(lemma(w), "?")))
        return " | ".join(dict.fromkeys(out)) + "  [metric word only; numbers must be re-derived]"


# ---------------------------------------------------------------------------
# Loading inputs
# ---------------------------------------------------------------------------
def load_problems(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        df = pd.read_excel(path)
        col = next(c for c in df.columns if str(c).lower() in ("problem text", "question", "problem", "text"))
        idc = next((c for c in df.columns if "problem #" in str(c).lower()), None)
        return {int(i if idc else k + 1): t for k, (i, t) in enumerate(zip(df[idc] if idc else range(len(df)), df[col]))}
    with open(path, newline="", encoding="utf-8-sig") as f:
        rd = csv.DictReader(f)
        col = next((c for c in rd.fieldnames if c.lower() in ("question", "problem", "text")), rd.fieldnames[0])
        return {i: (r[col] or "").strip() for i, r in enumerate(rd, 1) if (r[col] or "").strip()}


def extract_rows(problems, lex, tax):
    import gsmph_extract as ex
    rows = []
    for i, (pid, text) in enumerate(problems.items(), 1):
        if i % 100 == 0:
            print(f"  extracting {i}/{len(problems)}")
        for c in ex.extract_candidates(pid, text):
            sub, conf, meth = classify(c["candidate"], c["tag"], text, lex, tax)
            rows.append({**c, "sub": sub, "conf": conf, "method": meth})
    return rows


# ---------------------------------------------------------------------------
# Workbook writing
# ---------------------------------------------------------------------------
def write_sheet(wb, title, df, widths, header_fill=HDR, yellow_cols=(), wrap_cols=(), filt=True):
    ws = wb.create_sheet(title)
    ws.append(list(df.columns))
    for r in df.itertuples(index=False):
        ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in r])
    for c in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.fill = YELLOW if df.columns[c - 1] in yellow_cols else header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = widths.get(df.columns[c - 1], 18)
    for name in wrap_cols:
        if name in df.columns:
            ci = list(df.columns).index(name) + 1
            for row in ws.iter_rows(min_row=2, min_col=ci, max_col=ci):
                row[0].alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    if filt:
        ws.auto_filter.ref = ws.dimensions
    return ws


def pool_checks(tax, pool_all):
    rows = []
    used = {v["pool"] for v in tax.values() if v["handling"] == "Term pool"}
    cnt = pool_all["Pool"].value_counts()
    for p in sorted(set(cnt.index) - used):
        rows.append(["Pool not routed by any subcategory", p, int(cnt[p]),
                     "Terms exist but no Specific Subcategory points here; they are never suggested. Re-point a subcategory or retire the pool."])
    for s, v in tax.items():
        if v["handling"] == "Term pool" and cnt.get(v["pool"], 0) < 10:
            rows.append(["Thin pool (<10 terms)", f"{s} -> {v['pool']}", int(cnt.get(v["pool"], 0)),
                         "Too few options for varied replacements."])
        if v["handling"] == "Term pool" and v["pool"] not in cnt.index:
            rows.append(["Pool missing", f"{s} -> {v['pool']}", 0, "Pool name in Subcategory Map not found on Pool sheet."])
    seen = {"food_seafood": "Seafood pool has 54 terms but food_seafood routes to 'Aquatic animals'",
            "place_road_path_bridge": "'Street/road name' pool has 183 terms but this routes to 'Places & buildings'",
            "animal_aquarium_fish": "Parent group is Aquatic animals but it routes to the 'Pets' pool"}
    for s, note in seen.items():
        if s in tax:
            rows.append(["Possible routing mismatch", s, "", note])
    return pd.DataFrame(rows, columns=["Check", "Item", "Count", "Note"])


def build_workbook(cand, vg, per_problem, rq, removed, tax, pool_all, out, note_extra=""):
    wb = Workbook()
    wb.remove(wb.active)
    n_c, n_g = len(cand), len(vg)

    # ---- README ---------------------------------------------------------------------------
    ws = wb.create_sheet("README")
    ws.column_dimensions["A"].width = 120
    ws.column_dimensions["B"].width = 14
    lines = [
        ("GSMPH localization candidates -- full run, routed through the updated pool taxonomy (pipeline v2)", True),
        ("", False),
        ("WHAT THIS FILE IS", True),
        ("Same structure as full_run_enhanced_npscan_cleaned.xlsx. Original columns A-L (Decision / Final Category / Localized Replacement) and Variant Groups 'Replacement (fill once)' remain yours to fill.", False),
        ("", False),
        ("WHAT PIPELINE v2 ADDED", True),
        ("1. Candidates col I 'Suggested Pool' now holds the pool or rule from the pool workbook's Subcategory Map (the old pool names are kept in 'Previous Suggested Pool (superseded)').", False),
        ("2. Candidates 'Handling' says what to do: Term pool (pick from the Pool sheet), Rule-based (currency / imperial units), Keep (no substitution) or Manual review.", False),
        ("3. Variant Groups cols P-U: Handling, Pool / rule, Suggested replacement (auto), Alt 1, Alt 2, Basis. These are SUGGESTIONS drawn from the Pool sheet (top-30 by recognition, deterministic seeded pick, no repeats inside a problem). They are not written into 'Replacement (fill once)': the pool README says terms are context-fitting candidates to approve, not equivalents.", False),
        ("4. Candidates 'Suggested Replacement (auto)' looks the suggestion up by Group ID (formula), so every form in a concept shows the same suggestion.", False),
        ("5. 'Pool Checks' sheet lists pools no subcategory routes to, thin pools and possible routing mismatches in the Subcategory Map.", False),
        ("6. Review Queue gains Priority 6: Term-pool groups for which no suggestion could be made.", False),
        ("7. Freeze panes reset to row 1 on every sheet (the previous file froze panes at the bottom of several sheets).", False),
        ("", False),
        ("RULES AND LIMITS", True),
        ("Rule-based currency: '$' -> '₱' and dollar(s) -> piso/pesos; amounts are NOT rescaled (CURRENCY_SCALE = 1.0 in gsmph_pipeline.py). Realistic PH prices need a decision and change the answers.", False),
        ("Rule-based imperial units: only the metric WORD is suggested; the numbers must be converted by a person, which changes the answers.", False),
        ("Person-name suggestions use gender in the Pool 'Description' column; about 1 in 7 pool names have no gender recorded, so those can come back for either gender.", False),
        ("Tagalog nouns do not inflect; plural forms are flagged in the Basis column with the 'mga' prefix.", False),
        ("Classification is still rule-based (learned lexicon + cues). 'High' means a rule matched cleanly, not that a person verified it. See the self-test numbers in the build log / chat summary.", False),
        (note_extra, False),
        ("", False),
        ("SUMMARY (live formulas)", True),
    ]
    for t, b in lines:
        ws.append([t])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=b)
        ws.cell(row=ws.max_row, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    summ = [
        ("Candidate rows", f"=COUNTA(Candidates!A2:A{n_c + 1})"),
        ("Concept groups (Variant Groups rows)", f"=COUNTA('Variant Groups'!A2:A{n_g + 1})"),
        ("Rows: Term pool", f'=COUNTIF(Candidates!Z2:Z{n_c + 1},"Term pool")'),
        ("Rows: Rule-based", f'=COUNTIF(Candidates!Z2:Z{n_c + 1},"Rule-based")'),
        ("Rows: Keep (no substitution)", f'=COUNTIF(Candidates!Z2:Z{n_c + 1},"Keep*")'),
        ("Rows: Manual review", f'=COUNTIF(Candidates!Z2:Z{n_c + 1},"Manual review")'),
        ("Rows with High confidence subcategory", f'=COUNTIF(Candidates!Q2:Q{n_c + 1},"High")'),
        ("Rows with Medium confidence", f'=COUNTIF(Candidates!Q2:Q{n_c + 1},"Medium")'),
        ("Rows with Low confidence (review)", f'=COUNTIF(Candidates!Q2:Q{n_c + 1},"Low")'),
        ("Groups with a suggested replacement", f"=COUNTIF('Variant Groups'!R2:R{n_g + 1},\"?*\")"),
        ("Review Queue items", f"=COUNTA('Review Queue'!A2:A{len(rq) + 1})"),
    ]
    for lab, f in summ:
        ws.append([lab, f])

    # ---- Candidates -------------------------------------------------------------------------
    cols = ["Problem #", "Candidate", "Context", "Category", "Tier", "Sources", "Score", "Count", "Suggested Pool",
            "Decision", "Final Category", "Localized Replacement", "Specific Subcategory", "Parent Group",
            "Suggested Final Category", "Category Check", "Subcat Confidence", "Subcat Method", "Group ID",
            "Group Label", "Form in Problem", "All forms in this group", "Related Groups (same problem)",
            "Group Replacement (auto from Variant Groups)", "Review Flag", "Handling",
            "Previous Suggested Pool (superseded)", "Suggested Replacement (auto)"]
    ws = wb.create_sheet("Candidates")
    ws.append(cols)
    gl = f"'Variant Groups'!$A$2:$A${n_g + 1}"
    for i, r in enumerate(cand.itertuples(index=False), start=2):
        d = r._asdict() if hasattr(r, "_asdict") else {}
        ws.append([
            r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], None, None, None, r[12], r[13], r[14], r[15], r[16],
            r[17], r[18], r[19], r[20], r[21], r[22],
            f"=IFERROR(INDEX('Variant Groups'!$N$2:$N${n_g + 1},MATCH($S{i},{gl},0))&\"\",\"\")", r[24],
            r[25], r[26],
            f"=IFERROR(INDEX('Variant Groups'!$R$2:$R${n_g + 1},MATCH($S{i},{gl},0))&\"\",\"\")"])
    widths = [10, 22, 60, 18, 6, 26, 7, 8, 26, 11, 22, 24, 36, 22, 16, 26, 10, 24, 10, 28, 16, 40, 30, 28, 28, 16, 26, 34]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.fill = HDR
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for c in (10, 11, 12):
        ws.cell(row=1, column=c).fill = YELLOW
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{n_c + 1}"
    dv = DataValidation(type="list", formula1='"Keep,Reject,Unsure"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"J2:J{n_c + 1}")

    # ---- other sheets --------------------------------------------------------------------------
    write_sheet(wb, "Per Problem (Tier 1-2)", per_problem,
                {"Problem #": 10, "Problem Text": 60, "Candidates (Tier 1-2)": 40, "Repeated concepts (one replacement each)": 48},
                wrap_cols=("Problem Text", "Candidates (Tier 1-2)", "Repeated concepts (one replacement each)"), filt=False)
    write_sheet(wb, "Variant Groups", vg,
                {"Group ID": 11, "Problem #": 10, "Group Label": 34, "Concept Type": 22, "Specific Subcategory (group)": 34,
                 "Parent Group": 22, "Suggested Category": 16, "Forms in this problem": 44, "Rows": 7, "Distinct forms": 9,
                 "Redundancy type": 28, "Related groups (same problem)": 50, "Example context": 60,
                 "Replacement (fill once)": 28, "Notes": 30, "Handling": 14, "Pool / rule": 30,
                 "Suggested replacement (auto)": 34, "Alt 1": 20, "Alt 2": 20, "Basis": 60},
                yellow_cols=("Replacement (fill once)", "Notes"))
    tx = (cand.groupby("Specific Subcategory", dropna=False)
          .agg(parent=("Parent Group", "first"), cat=("Suggested Final Category", lambda s: s.value_counts().index[0]),
               rows=("Problem #", "size"), probs=("Problem #", "nunique"),
               high=("Subcat Confidence", lambda s: round((s == "High").mean(), 2)),
               terms=("Candidate", lambda s: ", ".join(norm(x) for x in s.value_counts().index[:8])),
               handling=("Handling", "first"), pool=("Suggested Pool", "first")).reset_index()
          .sort_values("rows", ascending=False))
    pool_n = pool_all["Pool"].value_counts()
    tx["Pool terms"] = tx["pool"].map(lambda p: int(pool_n.get(p, 0)) if p in pool_n.index else None)
    tx.columns = ["Specific Subcategory", "Parent Group", "Suggested Category", "Candidate rows", "Problems",
                  "Share High confidence", "Most frequent terms", "Handling", "Pool / rule (Subcategory Map)", "Pool terms"]
    write_sheet(wb, "Subcategory Taxonomy", tx, {"Specific Subcategory": 40, "Parent Group": 24, "Suggested Category": 18,
                                                  "Most frequent terms": 70, "Handling": 20, "Pool / rule (Subcategory Map)": 38})
    write_sheet(wb, "Review Queue", rq, {"Priority": 8, "Reason": 44, "Problem #": 10, "Group ID": 11, "Candidate": 22,
                                         "Context": 60, "Specific Subcategory": 36, "Confidence": 10, "Method": 26, "Note": 50,
                                         "Reviewer decision": 26, "Corrected subcategory": 28},
                yellow_cols=("Reviewer decision", "Corrected subcategory"))
    write_sheet(wb, "Removed np-scan rows", removed, {"Problem #": 10, "Candidate": 24, "Context": 70, "Category": 14,
                                                       "Specific Subcategory": 30, "Group ID": 12, "Reason removed": 60},
                header_fill=GREY, filt=False)
    write_sheet(wb, "Pool Checks", pool_checks(tax, pool_all), {"Check": 34, "Item": 48, "Count": 8, "Note": 90}, filt=False)
    wb.save(out)
    print(f"Saved {out}")


# ---------------------------------------------------------------------------
# run
# ---------------------------------------------------------------------------
def pool_label(sub, tax):
    i = tax.get(sub)
    return (i["pool"] if i else "REVIEW"), (i["handling"] if i else "Manual review")


def run_from_seed(seed_path, pool_path, out, lex_path=None):
    tax = load_taxonomy(pool_path)
    pool = load_pool(pool_path)
    pool_all = pd.read_excel(pool_path, sheet_name="Pool")
    cand = pd.read_excel(seed_path, sheet_name="Candidates")
    vg = pd.read_excel(seed_path, sheet_name="Variant Groups")
    pp = pd.read_excel(seed_path, sheet_name="Per Problem (Tier 1-2)")
    rq = pd.read_excel(seed_path, sheet_name="Review Queue")
    removed = pd.read_excel(seed_path, sheet_name="Removed np-scan rows")

    # --- re-route every candidate through the updated Subcategory Map ---
    missing = sorted(set(cand["Specific Subcategory"]) - set(tax))
    if missing:
        print("WARNING subcategories missing from Subcategory Map:", missing)
    routed = cand["Specific Subcategory"].map(lambda s: pool_label(s, tax))
    cand["Previous Suggested Pool (superseded)"] = cand["Suggested Pool"]
    cand["Suggested Pool"] = [p for p, _ in routed]
    cand["Handling"] = [h for _, h in routed]
    cand["Parent Group"] = cand["Specific Subcategory"].map(lambda s: tax[s]["parent"] if s in tax else "")
    cand["Group Replacement (auto from Variant Groups)"] = None
    cand["Suggested Replacement (auto)"] = None
    cand = cand[["Problem #", "Candidate", "Context", "Category", "Tier", "Sources", "Score", "Count", "Suggested Pool",
                 "Decision", "Final Category", "Localized Replacement", "Specific Subcategory", "Parent Group",
                 "Suggested Final Category", "Category Check", "Subcat Confidence", "Subcat Method", "Group ID",
                 "Group Label", "Form in Problem", "All forms in this group", "Related Groups (same problem)",
                 "Group Replacement (auto from Variant Groups)", "Review Flag", "Handling",
                 "Previous Suggested Pool (superseded)", "Suggested Replacement (auto)"]]

    # --- suggestions per group (problem order so 'used' is per problem) ---
    sg = Suggester(tax, pool)
    vg = vg.copy()
    handling, poolcol, p1, a1, a2, basis = [], [], [], [], [], []
    used_by_problem = defaultdict(set)
    for r in vg.itertuples(index=False):
        sub = str(r[4])
        base = sub.replace(" (family)", "")
        pname, h = pool_label(base, tax)
        if sub.endswith("(family)"):
            pname, h = "RULE: currency amount (scale + PH prices)", "Rule-based"
        forms = [f.strip() for f in str(r[7]).split("|")] if pd.notna(r[7]) else []
        if h == "Rule-based" and sub.endswith("(family)"):
            s, b = sg._rule("currency", forms), "RULE: currency amount (scale + PH prices) + 'Currency word' pool for dollar words"
            x1 = x2 = ""
        else:
            s, x1, x2, b = sg.suggest(base, forms, r[0], used_by_problem[r[1]], label=r[2])
        handling.append(h); poolcol.append(pname); p1.append(s); a1.append(x1); a2.append(x2); basis.append(b)
    vg["Handling"], vg["Pool / rule"] = handling, poolcol
    vg["Suggested replacement (auto)"], vg["Alt 1"], vg["Alt 2"], vg["Basis"] = p1, a1, a2, basis

    # --- review queue: add Priority 6 for term-pool groups with no suggestion ---
    ctx = dict(zip(vg["Group ID"], vg["Example context"]))
    add = vg[(vg["Handling"] == "Term pool") & (vg["Suggested replacement (auto)"].astype(str).str.strip() == "")]
    new = pd.DataFrame({"Priority": 6, "Reason": "Term-pool group with no pool suggestion", "Problem #": add["Problem #"],
                        "Group ID": add["Group ID"], "Candidate": add["Group Label"], "Context": add["Group ID"].map(ctx),
                        "Specific Subcategory": add["Specific Subcategory (group)"], "Confidence": "", "Method": "pool-routing",
                        "Note": add["Basis"], "Reviewer decision": None, "Corrected subcategory": None})
    rq = pd.concat([rq, new], ignore_index=True)
    note = (f"Built from {seed_path}: {len(cand)} candidate rows, {len(vg)} groups; {len(new)} term-pool groups had no suggestion.")
    build_workbook(cand, vg, pp, rq, removed, tax, pool_all, out, note)
    return cand, vg


def run_extract(problems_path, pool_path, lex_path, out):
    tax = load_taxonomy(pool_path)
    pool = load_pool(pool_path)
    pool_all = pd.read_excel(pool_path, sheet_name="Pool")
    lex = json.load(open(lex_path, encoding="utf-8"))
    problems = load_problems(problems_path)
    rows = extract_rows(problems, lex, tax)
    byp = defaultdict(list)
    for r in rows:
        byp[r["problem"]].append(r)
    sg = Suggester(tax, pool)
    cand_rows, vg_rows, pp_rows, rq_rows = [], [], [], []
    for pid, text in problems.items():
        rs = byp.get(pid, [])
        used = set()
        if rs:
            groups = build_groups(rs, text, tax)
            for gid, g in groups.items():
                sub = Counter(r["sub"] for r in g["rows"]).most_common(1)[0][0]
                forms = list(dict.fromkeys(norm(r["candidate"]) for r in g["rows"]))
                pname, h = pool_label(sub, tax)
                s, x1, x2, b = sg.suggest(sub, forms, gid, used, label=g["label"])
                if h == "Rule-based":
                    pname = tax[sub]["pool"]
                ctype = ("person name" if sub.startswith("name_") else "currency system" if sub.startswith("currency_")
                         else "imperial unit" if sub.startswith("unit_") else "honorific" if sub.startswith("honorific_")
                         else "common-noun concept")
                vg_rows.append([gid, pid, g["label"], ctype, sub, tax.get(sub, {}).get("parent", ""),
                                lex["subcat_category"].get(sub, "Object"), " | ".join(forms), len(g["rows"]), len(forms),
                                None, None, g["rows"][0]["context"], None, None, h, pname, s, x1, x2, b])
                for r in g["rows"]:
                    pl, hh = pool_label(r["sub"], tax)
                    cand_rows.append([pid, r["candidate"], r["context"], lex["subcat_category"].get(r["sub"], "Object"),
                                      r["tier"], r["sources"], None, r["count"], pl, None, None, None, r["sub"],
                                      tax.get(r["sub"], {}).get("parent", ""), lex["subcat_category"].get(r["sub"], "Object"),
                                      "agree", r["conf"], r["method"], gid, g["label"],
                                      "plural" if norm(r["candidate"]) != lemma_phrase(r["candidate"]) else "singular",
                                      " | ".join(forms), None, None, "LOW CONFIDENCE" if r["conf"] == "Low" else None, hh, None, None])
                    if r["conf"] == "Low":
                        rq_rows.append([1, "Low-confidence or unclassified subcategory", pid, gid, r["candidate"], r["context"],
                                        r["sub"], r["conf"], r["method"], None, None, None])
        lines = defaultdict(list)
        for r in rs:
            if r["tier"] <= 2:
                lines[lex["subcat_category"].get(r["sub"], "Object")].append(r["candidate"])
        pp_rows.append([pid, text, "\n".join(f"{k}: {', '.join(v)}" for k, v in lines.items()), None])
    cand = pd.DataFrame(cand_rows, columns=["Problem #", "Candidate", "Context", "Category", "Tier", "Sources", "Score", "Count",
                                            "Suggested Pool", "Decision", "Final Category", "Localized Replacement",
                                            "Specific Subcategory", "Parent Group", "Suggested Final Category", "Category Check",
                                            "Subcat Confidence", "Subcat Method", "Group ID", "Group Label", "Form in Problem",
                                            "All forms in this group", "Related Groups (same problem)",
                                            "Group Replacement (auto from Variant Groups)", "Review Flag", "Handling",
                                            "Previous Suggested Pool (superseded)", "Suggested Replacement (auto)"])
    vg = pd.DataFrame(vg_rows, columns=["Group ID", "Problem #", "Group Label", "Concept Type", "Specific Subcategory (group)",
                                        "Parent Group", "Suggested Category", "Forms in this problem", "Rows", "Distinct forms",
                                        "Redundancy type", "Related groups (same problem)", "Example context",
                                        "Replacement (fill once)", "Notes", "Handling", "Pool / rule",
                                        "Suggested replacement (auto)", "Alt 1", "Alt 2", "Basis"])
    pp = pd.DataFrame(pp_rows, columns=["Problem #", "Problem Text", "Candidates (Tier 1-2)", "Repeated concepts (one replacement each)"])
    rq = pd.DataFrame(rq_rows, columns=["Priority", "Reason", "Problem #", "Group ID", "Candidate", "Context",
                                        "Specific Subcategory", "Confidence", "Method", "Note", "Reviewer decision",
                                        "Corrected subcategory"])
    removed = pd.DataFrame(columns=["Problem #", "Candidate", "Context", "Category", "Specific Subcategory", "Group ID", "Reason removed"])
    build_workbook(cand, vg, pp, rq, removed, tax, pool_all, out, f"Built from {problems_path} via Stage 1 extraction.")


# ---------------------------------------------------------------------------
# selftest: honest hold-out evaluation (lexicon never sees the test problems)
# ---------------------------------------------------------------------------
def selftest(seed_path, pool_path=None):
    cand = pd.read_excel(seed_path, sheet_name="Candidates")
    tax = load_taxonomy(pool_path) if pool_path else {}
    test_mask = cand["Problem #"] % 5 == 0
    train, test = cand[~test_mask], cand[test_mask]
    lex = build_lexicon(train)
    test = test[test["Subcat Confidence"].isin(["High", "Medium"])]       # score only rows the seed itself trusts
    ok = pk = n = 0
    kinds = Counter()
    wrong = []
    for t, tag_src, ctx, sub in zip(test["Candidate"], test["Sources"], test["Context"], test["Specific Subcategory"]):
        got, conf, meth = classify(t, "PERSON" if sub.startswith("name_") else "", ctx, lex, tax)
        n += 1
        ok += got == sub
        if tax:
            pk += pool_label(got, tax) == pool_label(sub, tax)
        kinds[(meth.split(":")[0], got == sub)] += 1
        if got != sub and len(wrong) < 15:
            wrong.append((t, sub, got, meth))
    print(f"Hold-out rows scored: {n}")
    print(f"Exact subcategory accuracy: {ok / n:.1%}")
    if tax:
        print(f"Same pool/handling (what matters for replacements): {pk / n:.1%}")
    print("By method (correct / total):")
    for m in sorted({k[0] for k in kinds}):
        c, w = kinds[(m, True)], kinds[(m, False)]
        print(f"  {m:14s} {c}/{c + w}")
    print("Sample errors (term, seed, predicted, method):")
    for w in wrong:
        print("  ", w)
    return ok / n


# ---------------------------------------------------------------------------
# localize: apply the suggested replacements to the problem text (DRAFT output)
# ---------------------------------------------------------------------------
def localize_text(text, subs, currency):
    """subs: list of (form, replacement, plural). Longest forms first; placeholders stop double replacement."""
    applied, holders = [], {}
    for form, rep, plural in sorted(subs, key=lambda x: -len(x[0])):
        pat = re.compile(r"(?<![\w])" + re.escape(form) + r"(?![\w])", re.I)
        def sub(m, rep=rep, plural=plural, form=form):
            out = ("mga " if plural else "") + rep
            if m.group()[:1].isupper() and not plural:
                out = out[:1].upper() + out[1:]
            key = f"\x00{len(holders)}\x00"
            holders[key] = out
            return key
        new, n = pat.subn(sub, text)
        if n:
            text = new
            applied.append(f"{form} -> {('mga ' if plural else '') + rep}")
    if currency:
        text = text.replace("$", "₱")
        text = re.sub(r"\bdollars\b", "pesos", text, flags=re.I)
        text = re.sub(r"\bdollar\b", "piso", text, flags=re.I)
        applied.append("$ -> ₱ (amounts unchanged)")
    for k, v in holders.items():
        text = text.replace(k, v)
    return text, applied


def localize(full_xlsx, out):
    cand = pd.read_excel(full_xlsx, sheet_name="Candidates")
    vg = pd.read_excel(full_xlsx, sheet_name="Variant Groups")
    pp = pd.read_excel(full_xlsx, sheet_name="Per Problem (Tier 1-2)")
    sugg = dict(zip(vg["Group ID"], vg["Suggested replacement (auto)"]))
    human = dict(zip(vg["Group ID"], vg["Replacement (fill once)"]))
    byp = defaultdict(list)
    for r in cand.itertuples(index=False):
        byp[r[0]].append(r)
    rows = []
    for pid, text in zip(pp["Problem #"], pp["Problem Text"]):
        subs, left, cur = [], [], False
        for r in byp.get(pid, []):
            handling, gid, form = r[25], r[18], str(r[1]).strip("\"'[]() ")
            rep = human.get(gid) if pd.notna(human.get(gid)) else sugg.get(gid)
            if handling == "Rule-based" and str(r[12]).startswith("currency"):
                cur = True
            elif handling == "Rule-based":
                left.append(f"{form} (unit: convert numbers by hand)")
            elif handling == "Term pool" and pd.notna(rep) and len(form) > 1:
                subs.append((form, str(rep), "plural" in str(r[20])))
            elif handling in ("Manual review", "Term pool"):
                left.append(f"{form} (needs review)")
        subs = list(dict.fromkeys(subs))
        new, applied = localize_text(str(text), subs, cur)
        rows.append([pid, text, new, "; ".join(applied), "; ".join(dict.fromkeys(left))])
    df = pd.DataFrame(rows, columns=["Problem #", "Original problem", "Localized problem (auto draft)",
                                     "Replacements applied", "Left unchanged / needs a person"])
    wb = Workbook(); wb.remove(wb.active)
    write_sheet(wb, "Localized Problems", df, {"Problem #": 10, "Original problem": 60, "Localized problem (auto draft)": 60,
                "Replacements applied": 50, "Left unchanged / needs a person": 40},
                wrap_cols=tuple(df.columns[1:]))
    wb.save(out)
    print(f"Saved {out}: {len(df)} problems, {(df['Replacements applied'] != '').sum()} changed")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    a = sp.add_parser("build-lexicon")
    a.add_argument("--seed", required=True); a.add_argument("--out", default="gsmph_lexicon.json")
    b = sp.add_parser("run")
    b.add_argument("--pool", required=True); b.add_argument("--out", default="gsmph_full_run_v2.xlsx")
    b.add_argument("--lexicon", default="gsmph_lexicon.json")
    b.add_argument("--candidates-from"); b.add_argument("--problems")
    d = sp.add_parser("localize")
    d.add_argument("--full", required=True); d.add_argument("--out", default="gsmph_localized_problems.xlsx")
    c = sp.add_parser("selftest")
    c.add_argument("--seed", required=True); c.add_argument("--pool")
    args = ap.parse_args()
    if args.cmd == "build-lexicon":
        cand = pd.read_excel(args.seed, sheet_name="Candidates")
        rq = pd.read_excel(args.seed, sheet_name="Review Queue")
        lex = build_lexicon(cand, rq)
        json.dump(lex, open(args.out, "w", encoding="utf-8"), ensure_ascii=False)
        print(f"Wrote {args.out}: {len(lex['terms'])} terms, {len(lex['names'])} names")
    elif args.cmd == "run":
        if args.candidates_from:
            run_from_seed(args.candidates_from, args.pool, args.out, args.lexicon)
        elif args.problems:
            run_extract(args.problems, args.pool, args.lexicon, args.out)
        else:
            ap.error("give --candidates-from FULL.xlsx or --problems FILE")
    elif args.cmd == "localize":
        localize(args.full, args.out)
    else:
        selftest(args.seed, args.pool)


if __name__ == "__main__":
    main()
