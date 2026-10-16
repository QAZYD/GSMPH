import csv
import os
import re
import spacy
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from nltk.tag import StanfordNERTagger
from nltk.tokenize import word_tokenize
from nltk.internals import config_java
from transformers import AutoTokenizer, AutoModelForTokenClassification, pipeline
from concurrent.futures import ThreadPoolExecutor
import copy


# JAVA SETUP
JAVA_PATH = r"C:\Program Files\Common Files\Oracle\Java\javapath\java.exe"
config_java(bin=JAVA_PATH)

# STANFORD NER SETUP (Tool 1 -- the real 7-class MUC model)
STANFORD_NER_JAR = r"C:\nltk_data\stanford-ner-2020-11-17\stanford-ner.jar"
STANFORD_NER_MODEL = r"C:\nltk_data\stanford-ner-2020-11-17\classifiers\english.muc.7class.distsim.crf.ser.gz"

stanford_tagger = StanfordNERTagger(
    STANFORD_NER_MODEL,
    STANFORD_NER_JAR,
    encoding="utf-8",
)

_ner_cache = {}
_orig_tag = stanford_tagger.tag

def _cached_tag(tokens):
    key = tuple(tokens)
    if key not in _ner_cache:          # fallback: still works if a text wasn't prefetched
        _ner_cache[key] = _orig_tag(tokens)
    return _ner_cache[key]

stanford_tagger.tag = _cached_tag   

def prefetch_stanford(texts, chunk_size=200, workers=4):
    """Tag all texts in a few big Java launches instead of one per problem."""
    token_lists = [word_tokenize(t) for t in texts]
    chunks = [token_lists[i:i + chunk_size] for i in range(0, len(token_lists), chunk_size)]
    print(f"Prefetching Stanford NER: {len(token_lists)} problems in {len(chunks)} batches...")
    with ThreadPoolExecutor(max_workers=workers) as ex:
        for chunk, tagged_sents in zip(chunks, ex.map(stanford_tagger.tag_sents, chunks)):
            for toks, tagged in zip(chunk, tagged_sents):
                _ner_cache[tuple(toks)] = tagged
    print("Prefetch done.")




nlp = spacy.load("en_core_web_sm")



# FOOD NER SETUP (Tool 5 -- FoodBaseBERT-NER)
FOOD_NER_MODEL_NAME = "Dizex/FoodBaseBERT-NER"
food_ner_tokenizer = AutoTokenizer.from_pretrained(FOOD_NER_MODEL_NAME)
food_ner_model = AutoModelForTokenClassification.from_pretrained(FOOD_NER_MODEL_NAME)
food_ner_pipe = pipeline(
    "ner",
    model=food_ner_model,
    tokenizer=food_ner_tokenizer,
    aggregation_strategy="simple",
)


# MISC NER SETUP (Tool 6 -- dslim/distilbert-NER, used only for its MISC
# tag as a low-confidence Event signal; PER/ORG/LOC output is ignored since
# Stanford NER (Tool 1) already covers those categories)
MISC_NER_MODEL_NAME = "dslim/distilbert-NER"
misc_ner_tokenizer = AutoTokenizer.from_pretrained(MISC_NER_MODEL_NAME)
misc_ner_model = AutoModelForTokenClassification.from_pretrained(MISC_NER_MODEL_NAME)
misc_ner_pipe = pipeline(
    "ner",
    model=misc_ner_model,
    tokenizer=misc_ner_tokenizer,
    aggregation_strategy="simple",
)

MUC_7_LABELS = {"PERSON", "ORGANIZATION", "LOCATION", "MONEY", "PERCENT"}


MONTHS = {
    "january","february","march","april","may","june",
    "july","august","september","october","november","december",
}
DAYS = {
    "monday","tuesday","wednesday","thursday","friday","saturday","sunday",
}


# KEYWORD LISTS (Tool 4) -- word lists for categories no NER/POS/transformer
# model here reliably detects, because these are common (lowercase) nouns,
# not proper nouns. Expand these as you review more of the real dataset.
HONORIFIC_WORDS = {
    "mr.", "mr", "mrs.", "mrs", "ms.", "ms", "dr.", "dr", "miss",
}

# KINSHIP / RELATIONSHIP WORDS -- excluded from Tool 7 / 7b candidates.
# These are relationship nouns (who someone is to the narrator), not
# physical/cultural objects, so they never belong on a localization
# candidate list even though spaCy correctly tags them as concrete NOUNs.
# Added after manual review flagged "friends", "parents", "grandparents"
# as noise (GSMPH candidate list requirements: objects only).
KINSHIP_WORDS = {
    "friend", "friends", "parent", "parents", "grandparent", "grandparents",
    "brother", "brothers", "sister", "sisters", "mother", "father", "mom",
    "dad", "mommy", "daddy", "son", "sons", "daughter", "daughters",
    "cousin", "cousins", "aunt", "aunts", "uncle", "uncles", "husband",
    "wife", "family", "kid", "kids", "child", "children", "classmate",
    "classmates", "neighbor", "neighbors", "neighbour", "neighbours",
}

# ABSTRACT / QUANTITY NOUNS -- excluded from Tool 7 / 7b candidates.
# These are common nouns in word problems that spaCy tags as concrete
# NOUNs, but they denote quantities, measures, or abstract concepts
# rather than physical, localizable objects (e.g. "half", "purpose").
# Expand this list as more of the real dataset is reviewed.
ABSTRACT_QUANTITY_WORDS = {
    "half", "quarter", "third", "purpose", "amount", "quantity", "total",
    "sum", "difference", "remainder", "percentage", "rate", "value",
    "number", "average", "ratio", "portion", "fraction",
}

TIME_WORDS={
"yesterday","today","tomorrow","hour","hours","minute","minutes",
"day","days","week","weeks","month","months","year","years",
"shift","shifts","overtime"
}


# QUANTIFIER / VAGUE MODIFIER WORDS -- used as a backstop to reject noun
# phrases like "as many clips", "how much more money", "twice as many
# pages", "enough brownies", "remaining pages", "only half", and
# "final weight". These words signal a quantity/comparison being
# discussed, not a fixed compound-noun object like "care package" or
# "boarding school", so any chunk containing one (anywhere, not just as
# the first word) is rejected outright rather than left for manual
# review. This is a deliberate backstop alongside the POS/dependency
# check in tool7_noun_phrases, since the POS tagger doesn't always tag
# these consistently (e.g. "remaining" can come back as VERB or ADJ).
QUANTIFIER_OR_VAGUE_WORDS = {
    "how", "what", "which", "many", "much", "more", "most", "less",
    "least", "twice", "thrice", "enough", "several", "few", "various",
    "remaining", "only", "final", "other", "another", "additional",
    "extra", "entire", "whole", "same", "certain", "each", "every",
    "any", "no", "all", "both", "either", "neither", "new", "old",
}

# GSM8K problems are now read from a CSV file (see load_problems()) instead
# of being hardcoded here. Default filename/column names below; override
# via load_problems(path=..., column=...) if your CSV differs.
GSM8K_CSV_PATH = "gsm8k.csv"
# Column names (case-insensitive) load_problems will look for, in order.
QUESTION_COLUMN_CANDIDATES = ("question", "problem", "text")


def load_problems(path=GSM8K_CSV_PATH, column=None):
    """Load the list of GSM8K problem strings from a CSV file.

    - `column`: exact column name to read. If omitted, auto-detects by
      matching (case-insensitively) against QUESTION_COLUMN_CANDIDATES;
      if none of those match, falls back to the first column in the file.
    - Rows with an empty/whitespace-only value in that column are skipped.
    - Raises FileNotFoundError with a clear message if the CSV is missing,
      and ValueError if the file has no header row / no usable column.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Could not find '{path}'. Place your GSM8K CSV next to this "
            f"script (or pass load_problems(path=...) with the correct path)."
        )

    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise ValueError(f"'{path}' has no header row / no columns.")

        target_column = column
        if target_column is None:
            lower_to_actual = {name.lower(): name for name in reader.fieldnames}
            for candidate in QUESTION_COLUMN_CANDIDATES:
                if candidate in lower_to_actual:
                    target_column = lower_to_actual[candidate]
                    break
            if target_column is None:
                target_column = reader.fieldnames[0]

        problems = []
        for row in reader:
            value = (row.get(target_column) or "").strip()
            if value:
                problems.append(value)

    if not problems:
        raise ValueError(
            f"No problem text found in '{path}' under column '{target_column}'."
        )

    print(f"Loaded {len(problems)} problems from '{path}' (column: '{target_column}').")
    return problems


def _is_punct_token(token):
    """True if a token is pure punctuation that should never continue or
    be swallowed into an entity span (e.g. trailing '?', '.', ',', '!').
    Currency/percent symbols ($, %, #) are deliberately excluded here --
    they're meaningful parts of MONEY/PERCENT entities, not sentence
    punctuation, and stripping them would silently turn "$100" into "100"."""
    return re.match(r'^[^\w$%#]+$', token) is not None


def _join_entity_tokens(tokens):
    """Join the tokens of one entity span back into a single string.
    Plain ' '.join() reintroduces spaces that word_tokenize removed
    (e.g. '$12' -> ['$', '12'] -> '$ 12'). Re-glue a leading currency /
    unit symbol (and remove space before closing punctuation) so the
    output matches how the amount actually appears in the source text."""
    text = " ".join(tokens)
    text = re.sub(r'([$#])\s+(?=\d)', r'\1', text)   # "$ 12" -> "$12"
    text = re.sub(r'\s+([%,.;:])', r'\1', text)       # "80 %" edge case guard
    return text


def tool1_ner_stanford(text):
    """Tool 1 (Stanford NER 7-class MUC model)."""
    tokens = word_tokenize(text)
    tagged = stanford_tagger.tag(tokens)

    entities = []
    current_tokens = []
    current_label = None

    for token, label in tagged:
        # Never let a pure-punctuation token extend an entity span, even if
        # the tagger reports the same label for it (fixes artifacts like
        # the trailing "?" being fused onto "May" -> "May ?").
        if _is_punct_token(token):
            if current_tokens and current_label in MUC_7_LABELS:
                entities.append((_join_entity_tokens(current_tokens), current_label))
            current_tokens = []
            current_label = None
            continue

        if label == current_label and label != "O":
            current_tokens.append(token)
        else:
            if current_tokens and current_label in MUC_7_LABELS:
                entities.append((_join_entity_tokens(current_tokens), current_label))
            current_tokens = [token] if label != "O" else []
            current_label = label

    if current_tokens and current_label in MUC_7_LABELS:
        entities.append((_join_entity_tokens(current_tokens), current_label))

    return entities


def tool2_pos(doc, already_flagged_texts, stanford_person_tokens):
    """Tool 2 (spaCy POS tagger) -- proper nouns NER missed.

    Two guards added:
    1. Capitalization check: a genuine proper noun is capitalized. Without
       this, spaCy mistags on lowercase common nouns (e.g. "box") slip
       through as NNP/PROPN false positives.
    2. Local dedup: track terms already flagged *within this call* (not
       just the NER set passed in) so a name that legitimately repeats in
       the same problem (e.g. "Ken ... Ken ...") is reported once per
       problem, consistent with how tool4's keyword pass dedupes.
    """
    flagged = []
    seen_this_problem = set()
    for token in doc:
        if token.text in stanford_person_tokens:
            continue
        if token.ent_type_=="DATE" and token.text.lower() in (MONTHS|DAYS):
            continue
        if token.tag_ in ("NNP", "NNPS") or token.pos_ == "PROPN":
            if not token.text[:1].isupper():
                continue
            if token.text in already_flagged_texts or token.text in seen_this_problem:
                continue
            flagged.append((token.text, token.tag_))
            seen_this_problem.add(token.text)
    return flagged


def tool3_regex(text, already_flagged_texts):
    """Tool 3 (regex pass) -- remaining mid-sentence capitals."""
    flagged = []
    sentences = re.split(r'(?<=[.?!])\s+', text)
    for sent in sentences:
        words = sent.split()
        for w in words[1:]:
            clean = re.sub(r'[^\w]', '', w)
            if clean.lower() in MONTHS or clean.lower() in DAYS:
                continue
            if clean.lower() in {h.rstrip(".") for h in HONORIFIC_WORDS}:
                continue
            if clean and clean[0].isupper() and clean not in already_flagged_texts:
                flagged.append(clean)
    return flagged


def tool4_keyword_list(text, already_flagged_texts):
    """Tool 4 (keyword list pass) -- catches Unit and Honorific terms,
    which are common lowercase words that NER/POS/regex structurally
    cannot detect (they aren't proper nouns and aren't a regex pattern)."""
    flagged = []
    text_lower = text.lower()

    all_keyword_lists = HONORIFIC_WORDS
    for phrase in sorted(all_keyword_lists, key=len, reverse=True):
        pattern = r'\b' + re.escape(phrase) + r'\b'
        for match in re.finditer(pattern, text_lower):
            matched_text = text[match.start():match.end()]
            if matched_text not in already_flagged_texts:
                flagged.append(matched_text)
                already_flagged_texts.add(matched_text)

    return flagged


def tool4b_currency_regex(text, already_flagged_texts):
    """Tool 4b (regex pass) -- catches currency amounts like '$12' via a
    literal pattern match, not a word list. Kept separate from tool4 (and
    labeled 'Regex' rather than 'Keyword List' in the report) since no
    keyword list is involved here -- it's the same kind of raw pattern
    matching as tool3, just for currency symbols instead of capitalized
    words. Acts as a backstop: if tool1 (Stanford NER) already tagged the
    full '$12'-style span, this is skipped via already_flagged_texts; if
    Stanford dropped the symbol or missed the amount entirely, this tool
    supplies the correct, symbol-inclusive form."""
    flagged = []
    for match in re.finditer(r'\$\d+(?:\.\d+)?', text):
        matched_text = match.group()
        if matched_text not in already_flagged_texts:
            flagged.append(matched_text)
            already_flagged_texts.add(matched_text)
    return flagged


def _clean_ner_span(text, ent):
    """Reconstruct an aggregated entity's surface text directly from the
    original string using the pipeline's character offsets, instead of
    trusting ent['word']. With aggregation_strategy='simple', ent['word']
    is rebuilt by the tokenizer and can leak raw WordPiece continuation
    fragments (e.g. 'j' + '##elly beans', 'brown' + '##ies') when a span
    breaks across sub-tokens. Slicing text[start:end] always yields the
    real substring and can't produce '##' artifacts."""
    start, end = ent.get("start"), ent.get("end")
    if start is None or end is None:
        return ent["word"].strip()
    return text[start:end].strip()


def _merge_adjacent_ner_spans(text, results):
    """Re-join NER entities that got split mid-word before they're ever
    turned into candidate strings.

    aggregation_strategy='simple' is supposed to glue wordpiece sub-tokens
    back into one entity, but if the model's per-token confidence dips (or
    the entity_group genuinely flips) on an interior sub-token, the
    aggregator emits two *separate* entity dicts for what is really one
    lexical word -- e.g. 'goodies' -> [{'word': 'good', ...},
    {'word': '##ies', ...}] or 'jelly beans' -> [{'word': 'j', ...},
    {'word': '##elly beans', ...}]. Left alone, each fragment becomes its
    own row in the report ('good' / 'ies', 'j' / 'elly beans').

    Fix: if two consecutive entities (sorted by start offset) touch or
    overlap in the *original* string with no whitespace between them,
    they are part of the same word/phrase and must be merged into a
    single span before anything downstream (scoring, text extraction)
    sees them. A real word boundary always has whitespace or punctuation
    with a space; a wordpiece split never does.
    """
    if not results:
        return results

    ordered = sorted(results, key=lambda e: (e.get("start") or 0))
    merged = [dict(ordered[0])]

    for ent in ordered[1:]:
        prev = merged[-1]
        prev_end = prev.get("end")
        ent_start = ent.get("start")
        if prev_end is None or ent_start is None:
            merged.append(dict(ent))
            continue

        gap = text[prev_end:ent_start] if ent_start >= prev_end else ""
        touches_no_space = ent_start <= prev_end or gap == ""

        if touches_no_space:
            # Same lexical word/phrase split across entities -- merge.
            prev["end"] = max(prev_end, ent.get("end", prev_end))
            # Keep the entity_group of whichever fragment the model was
            # more confident about; conservative min() for the score so a
            # low-confidence fragment doesn't get silently upgraded.
            if ent.get("score", 0) > prev.get("score", 0):
                prev["entity_group"] = ent.get("entity_group", prev.get("entity_group"))
            prev["score"] = min(prev.get("score", 0), ent.get("score", 0))
        else:
            merged.append(dict(ent))

    return merged


def tool5_food_ner(text, already_flagged_texts):
    """Tool 5 (FoodBaseBERT-NER) -- detects FOOD entities via a fine-tuned
    BERT model, generalizing beyond any fixed word list."""
    flagged = []
    try:
        results = food_ner_pipe(text)
    except Exception as e:
        print(f"  [Tool 5 warning] Food NER failed on this problem: {e}")
        return flagged

    results = _merge_adjacent_ner_spans(text, results)

    for ent in results:
        matched_text = _clean_ner_span(text, ent)
        if not matched_text or "#" in matched_text or ent.get("score", 0) < 0.5:
            continue
        if matched_text not in already_flagged_texts:
            flagged.append(matched_text)
            already_flagged_texts.add(matched_text)

    return flagged

def make_cached_pipe(pipe):
    cache = {}
    def wrapped(text, *args, **kwargs):
        if text not in cache:               # fallback: unbatched path still works
            cache[text] = pipe(text, *args, **kwargs)
        return copy.deepcopy(cache[text])   # callers can't corrupt the cache
    wrapped.cache = cache
    wrapped.raw = pipe
    return wrapped

food_ner_pipe = make_cached_pipe(food_ner_pipe)
misc_ner_pipe = make_cached_pipe(misc_ner_pipe)

def prefetch_hf(wrapped_pipe, texts, name, chunk_size=256, batch_size=32):
    texts = sorted(set(texts), key=len)     # similar lengths per batch = less padding
    for start in range(0, len(texts), chunk_size):
        chunk = texts[start:start + chunk_size]
        try:
            outs = wrapped_pipe.raw(chunk, batch_size=batch_size)
            for t, o in zip(chunk, outs):
                wrapped_pipe.cache[t] = o
        except Exception as e:
            print(f"  [{name} prefetch warning] chunk failed, will run unbatched: {e}")
        print(f"{name}: {min(start + chunk_size, len(texts))}/{len(texts)}")

_doc_cache = {}

def prefetch_spacy(texts, batch_size=64):
    print(f"spaCy: processing {len(texts)} problems...")
    for text, doc in zip(texts, nlp.pipe(texts, batch_size=batch_size)):
        _doc_cache[text] = doc
    



def tool6_misc_ner(text, already_flagged_texts):
    """Tool 6 (dslim/distilbert-NER, MISC tag only) -- a low-confidence
    signal for possible Event terms (e.g. holidays). Only MISC-tagged
    output is used; PER/ORG/LOC from this model are ignored since Tool 1
    (Stanford NER) already covers those categories. MISC is a noisy
    catch-all in CoNLL-2003, so these hits are NOT auto-classified as
    Event -- they're flagged for manual review only."""
    flagged = []
    try:
        results = misc_ner_pipe(text)
    except Exception as e:
        print(f"  [Tool 6 warning] MISC NER failed on this problem: {e}")
        return flagged

    results = _merge_adjacent_ner_spans(text, results)

    for ent in results:
        if ent.get("entity_group") != "MISC":
            continue
        matched_text = _clean_ner_span(text, ent)
        if not matched_text or "#" in matched_text or ent.get("score", 0) < 0.5:
            continue
        if matched_text not in already_flagged_texts:
            flagged.append(matched_text)
            already_flagged_texts.add(matched_text)

    return flagged


def _is_compound_modifier(token):
    """True if `token` can stand as a modifier inside a fixed noun-noun
    compound object like "care package" or "boarding school".

    Only NOUN/PROPN modifiers (or tokens the dependency parser itself
    labels "compound") qualify. This deliberately excludes NUM, DET,
    ADJ, ADV, and quantifier/participle modifiers -- so "care package"
    and "boarding school" pass, while "new wallet", "12 pages",
    "enough brownies", "twice as many pages", and "remaining pages" do
    not: those aren't fixed compound objects, they're a generic
    adjective/quantifier attached to a noun that's already captured on
    its own by Tool 7b (e.g. "wallet", "pages")."""
    if token.text.lower() in QUANTIFIER_OR_VAGUE_WORDS:
        return False
    if token.dep_ == "compound":
        return True
    return token.pos_ in ("NOUN", "PROPN")


def tool7_noun_phrases(doc, already_flagged_texts):
    """Tool 7 (spaCy noun chunks) -- restricted to fixed noun-noun
    compound objects (e.g. "care package", "boarding school",
    "department store"). Chunks led or interrupted by a number,
    determiner, quantifier, or generic adjective (e.g. "new wallet",
    "12 pages", "enough brownies", "twice as many pages", "remaining
    pages", "final weight", "only half", "how much more money") are
    rejected -- those are quantity/comparison phrases, not localizable
    compound objects, and their head noun is already captured on its
    own by Tool 7b.

    Returns (flagged_phrases, consumed_token_indices) -- the latter lets
    the caller stop Tool 7b from re-flagging "care" and "package"
    separately once "care package" has already been captured whole.
    """
    flagged=[]
    consumed_indices=set()
    already_lower={t.lower() for t in already_flagged_texts}
    seen=set()
    for chunk in doc.noun_chunks:
        tokens=[t for t in chunk if not t.is_punct]
        while tokens and (tokens[0].pos_=="DET" or tokens[0].dep_=="poss"):
            tokens=tokens[1:]
        if len(tokens)<2:
            continue
        if not _is_concrete_candidate(tokens[-1]):
            continue
        if not all(_is_compound_modifier(t) for t in tokens[:-1]):
            continue
        clean=" ".join(t.text for t in tokens)
        key=clean.lower()
        if key in already_lower or key in seen:
            continue
        flagged.append(clean)
        seen.add(key)
        consumed_indices.update(t.i for t in tokens)
    return flagged, consumed_indices


def tool7b_single_noun_objects(doc, already_flagged_texts, consumed_indices=frozenset()):
    """Tool 7b (spaCy single-token nouns). Skips any token already
    consumed by a Tool 7 compound-noun phrase (e.g. "care"/"package"
    once "care package" is captured), so fragments of a phrase don't
    also show up as separate single-noun rows."""
    flagged=[]
    seen=set()
    already_lower={t.lower() for t in already_flagged_texts}
    for token in doc:
        if token.i in consumed_indices:
            continue
        if not _is_concrete_candidate(token):
            continue
        key=token.text.lower()
        if key in already_lower or key in seen:
            continue
        flagged.append(token.text)
        seen.add(key)
    return flagged



def _is_concrete_candidate(token):
    """
    Heuristic for candidate concrete nouns -- i.e. single nouns and noun
    phrase heads that could be objects worth flagging as culturally
    specific / localization candidates (e.g. "page", "box", "scale").

    Deliberately excludes categories that are grammatically nouns but are
    NOT objects, per manual review of the GSMPH candidate list
    requirements:
      - kinship/relationship words ("friends", "grandparents") -- people,
        not objects
      - abstract/quantity words ("half", "purpose") -- concepts, not
        objects
      - symbol-only tokens (e.g. a stray "%") -- tokenizer artifacts
    Still intentionally permissive beyond that, since all extracted
    candidates are manually reviewed.
    """
    if token.pos_ != "NOUN":
        return False
    if token.is_stop:
        return False
    if token.ent_type_ == "PERSON":
        return False
    if not any(c.isalpha() for c in token.text):
        return False
    lemma_lower = token.lemma_.lower()
    text_lower = token.text.lower()
    if lemma_lower in KINSHIP_WORDS or text_lower in KINSHIP_WORDS:
        return False
    if lemma_lower in ABSTRACT_QUANTITY_WORDS or text_lower in ABSTRACT_QUANTITY_WORDS:
        return False
    if lemma_lower in TIME_WORDS or text_lower in TIME_WORDS:
        return False
    if token.ent_type_ in ("DATE","TIME"):
        return False
    abstract_suffixes=("ness","tion","sion","ment","ship","ity","ance","ence")
    if lemma_lower.endswith(abstract_suffixes):
        return False
    return True

def classify_category(term, source_tool, tag_or_label):
    """Maps a raw candidate into one of the lookup table's categories
    (Table 4.1), or a short "Review - ..." bucket flagged for manual
    review. This is a proposed category only -- final assignment is
    always a manual decision per the methodology.

    Category strings are kept short and human-scannable on purpose
    (e.g. "Review - Date" rather than a full sentence): the "why does
    this need review" detail lives in classify_note() instead, so the
    Category column stays skimmable in a spreadsheet and the combined
    candidate list per problem doesn't turn into a wall of text.
    """
    term_lower = term.lower().strip()

    if source_tool == "NER (Stanford)":
        if tag_or_label == "PERSON":
            return "Proper Name"
        if tag_or_label == "MONEY":
            return "Currency"
        if tag_or_label == "DATE":
            return "Review - Date"
        if tag_or_label in ("ORGANIZATION", "LOCATION", "TIME", "PERCENT"):
            return f"Review - {tag_or_label.title()}"

    if term.startswith("$"):
        return "Currency"

    if source_tool == "Food NER (FoodBaseBERT)":
        return "Review - Food"

    if source_tool == "MISC NER (DistilBERT)":
        return "Review - Event?"

    if term_lower.rstrip(".") in {h.rstrip(".") for h in HONORIFIC_WORDS}:
        return "Honorific"

    if source_tool in ("POS (spaCy)", "Regex"):
        return "Review - Proper Noun?"

    if source_tool in ("Noun Phrase (spaCy)", "Single Noun (spaCy)"):
        return "Review - Object?"

    return "Review - Other"


def classify_note(category):
    """One short, plain-language explanation per category, shown in its
    own 'Notes' column instead of being crammed into the category label
    itself. Keeps `classify_category`'s output short while still telling
    a reviewer exactly what to check."""
    notes = {
        "Proper Name": "Auto-classified from NER; localize as a name.",
        "Currency": "Auto-classified; convert to local currency/amount.",
        "Honorific": "Auto-classified; localize title (Mr./Mrs./Dr./etc.).",
        "Review - Date": "NER tagged this DATE; confirm it isn't a month/day already excluded.",
        "Review - Organization": "NER tagged this an ORGANIZATION; confirm and classify.",
        "Review - Location": "NER tagged this a LOCATION; confirm and classify.",
        "Review - Time": "NER tagged this a TIME expression; confirm and classify.",
        "Review - Percent": "NER tagged this a PERCENT value; confirm and classify.",
        "Review - Food": "Food-NER hit; decide whether to keep or localize.",
        "Review - Event?": "Low-confidence MISC hit; check if this is a holiday/event.",
        "Review - Proper Noun?": "Capitalized word POS/regex caught; likely a name, confirm.",
        "Review - Object?": "Candidate concrete noun/noun phrase; confirm if culturally specific.",
        "Review - Other": "Doesn't fit an existing bucket; classify manually.",
    }
    return notes.get(category, "Classify manually.")


NEEDS_REVIEW_PREFIX = "Review"

PROBLEM_BAND_COLORS = ("FFFFFF", "F2F2F2")  # alternate white / light grey per problem


def _remove_fragment_duplicates(noun_phrases, single_nouns):
    kept=[]
    for n in single_nouns:
        if any(re.search(r"\b"+re.escape(n.lower())+r"\b", p.lower()) for p in noun_phrases):
            continue
        kept.append(n)
    return noun_phrases, kept

def process_problem(text):
    ner_hits = tool1_ner_stanford(text)
    ner_texts = {t for t, _ in ner_hits}
    stanford_person_tokens=set()
    for t,lbl in ner_hits:
        if lbl=="PERSON":
            stanford_person_tokens.update(t.split())

    doc = _doc_cache.get(text)
    if doc is None:
        doc = nlp(text)
    pos_hits = tool2_pos(doc, ner_texts, stanford_person_tokens)
    pos_texts = {t for t, _ in pos_hits}

    regex_hits = tool3_regex(text, ner_texts |  pos_texts)
    regex_texts = set(regex_hits)

    already_flagged = ner_texts | pos_texts | regex_texts
    keyword_hits = tool4_keyword_list(text, already_flagged)
    already_flagged |= set(keyword_hits)

    currency_hits = tool4b_currency_regex(text, already_flagged)
    already_flagged |= set(currency_hits)

    food_hits = tool5_food_ner(text, already_flagged)
    already_flagged |= set(food_hits)

    misc_hits = tool6_misc_ner(text, already_flagged)
    already_flagged |= set(misc_hits)

    noun_phrase_hits, consumed_indices = tool7_noun_phrases(doc, already_flagged)
    already_flagged |= set(noun_phrase_hits)

    single_noun_hits = tool7b_single_noun_objects(doc, already_flagged, consumed_indices)
    noun_phrase_hits, single_noun_hits = _remove_fragment_duplicates(noun_phrase_hits, single_noun_hits)

    return {
        "ner": ner_hits,
        "pos": pos_hits,
        "regex": regex_hits,
        "keyword_list": keyword_hits,
        "currency_regex": currency_hits,
        "food": food_hits,
        "misc": misc_hits,
        "noun_phrases": noun_phrase_hits,
        "single_noun_objects": single_noun_hits,
    }


def build_excel_report(output_path="gsmph_candidates.xlsx", csv_path=GSM8K_CSV_PATH, problems=None):
    """Build the Excel report. Problems come from `problems` if given,
    otherwise are loaded from `csv_path` via load_problems()."""
    if problems is None:
        problems = load_problems(csv_path)


    prefetch_stanford(problems)
    prefetch_spacy(problems)
    prefetch_hf(food_ner_pipe, problems, "Food NER")
    prefetch_hf(misc_ner_pipe, problems, "MISC NER")
    wb = Workbook()

    # ---------- Sheet 1: Candidates (Detailed) ----------
    ws1 = wb.active
    ws1.title = "Candidates (Detailed)"
    headers1 = ["Problem #", "Problem Text", "Candidate Term", "Source Tool",
                "Tag/Label", "Category", "Needs Review", "Notes"]
    ws1.append(headers1)

    # ---------- Sheet 2: Candidates (Per Problem) ----------
    ws2 = wb.create_sheet("Candidates (Per Problem)")
    headers2 = ["Problem #", "Problem Text", "Candidates (one per line)"]
    ws2.append(headers2)

    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    header_font = Font(bold=True)
    wrap_top = Alignment(wrap_text=True, vertical="top")
    wrap_center = Alignment(wrap_text=True, vertical="center", horizontal="center")
    review_fill = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
    thin = Side(style="thin", color="BFBFBF")
    top_border = Border(top=thin)

    for ws, headers in ((ws1, headers1), (ws2, headers2)):
        for col_idx, _ in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 20

    for i, problem in enumerate(problems, 1):
        print(f"Processing problem {i}/{len(problems)}...")
        result = process_problem(problem)

        rows = []
        for text, label in result["ner"]:
            rows.append((text, "NER (Stanford)", label))
        for text, tag in result["pos"]:
            rows.append((text, "POS (spaCy)", tag))
        for text in result["regex"]:
            rows.append((text, "Regex", ""))
        for text in result["keyword_list"]:
            rows.append((text, "Keyword List", ""))
        for text in result["currency_regex"]:
            rows.append((text, "Regex", ""))
        for text in result["food"]:
            rows.append((text, "Food NER (FoodBaseBERT)", "FOOD"))
        for text in result["misc"]:
            rows.append((text, "MISC NER (DistilBERT)", "MISC"))
        for text in result["noun_phrases"]:
            rows.append((text, "Noun Phrase (spaCy)", ""))
        for text in result["single_noun_objects"]:
            rows.append((text, "Single Noun (spaCy)", ""))

        band_color = PROBLEM_BAND_COLORS[i % 2]
        band_fill = PatternFill(start_color=band_color, end_color=band_color, fill_type="solid")

        # ---- ws1: one block of rows per problem, with Problem #/Text merged
        # down the block instead of repeated on every line, banding to make
        # each problem's block visually distinct, and a top border to mark
        # where the next problem starts. ----
        first_row = ws1.max_row + 1

        if not rows:
            ws1.append([i, problem, "(no candidates found)", "", "", "", "", ""])
        else:
            for text, tool, tag in rows:
                category = classify_category(text, tool, tag)
                needs_review = "Yes" if category.startswith(NEEDS_REVIEW_PREFIX) else "No"
                note = classify_note(category)
                ws1.append([i, problem, text, tool, tag, category, needs_review, note])

        last_row = ws1.max_row

        # Merge Problem # and Problem Text down the whole block for this
        # problem so the text isn't repeated on every candidate row.
        if last_row > first_row:
            ws1.merge_cells(start_row=first_row, start_column=1, end_row=last_row, end_column=1)
            ws1.merge_cells(start_row=first_row, start_column=2, end_row=last_row, end_column=2)

        for r in range(first_row, last_row + 1):
            ws1.cell(row=r, column=1).alignment = wrap_center
            ws1.cell(row=r, column=2).alignment = wrap_top
            for c in range(3, 9):
                cell = ws1.cell(row=r, column=c)
                cell.alignment = wrap_top
                cell.fill = band_fill
                if c == 7 and cell.value == "Yes":
                    cell.fill = review_fill
            if r == first_row:
                for c in range(1, 9):
                    ws1.cell(row=r, column=c).border = top_border

        # ---- ws2: readable one-term-per-line combined list, grouped by
        # category instead of one long semicolon-joined string. ----
        by_category = {}
        seen = set()
        for text, tool, tag in rows:
            if text in seen:
                continue
            seen.add(text)
            category = classify_category(text, tool, tag)
            by_category.setdefault(category, []).append(text)

        if by_category:
            lines = []
            for category in sorted(by_category, key=lambda c: (c.startswith(NEEDS_REVIEW_PREFIX), c)):
                terms = ", ".join(by_category[category])
                lines.append(f"{category}: {terms}")
            combined_text = "\n".join(lines)
            n_lines = len(lines)
        else:
            combined_text = "(none)"
            n_lines = 1

        ws2.append([i, problem, combined_text])
        row2 = ws2.max_row
        ws2.cell(row=row2, column=1).alignment = wrap_center
        ws2.cell(row=row2, column=2).alignment = wrap_top
        ws2.cell(row=row2, column=3).alignment = wrap_top
        for c in range(1, 4):
            ws2.cell(row=row2, column=c).fill = band_fill
            ws2.cell(row=row2, column=c).border = top_border
        # rough auto-height so multi-line cells aren't clipped
        ws2.row_dimensions[row2].height = max(15 * n_lines, 30)

    # ---------- column widths ----------
    ws1.column_dimensions["A"].width = 10
    ws1.column_dimensions["B"].width = 45
    ws1.column_dimensions["C"].width = 16
    ws1.column_dimensions["D"].width = 20
    ws1.column_dimensions["E"].width = 11
    ws1.column_dimensions["F"].width = 18
    ws1.column_dimensions["G"].width = 12
    ws1.column_dimensions["H"].width = 42

    ws2.column_dimensions["A"].width = 10
    ws2.column_dimensions["B"].width = 45
    ws2.column_dimensions["C"].width = 70

    ws1.freeze_panes = "A2"
    ws2.freeze_panes = "A2"
    ws1.auto_filter.ref = ws1.dimensions

    wb.save(output_path)
    print(f"\nSaved: {output_path}")


if __name__ == "__main__":
    all_problems = load_problems(GSM8K_CSV_PATH)
    sample = all_problems
    build_excel_report(output_path="test.xlsx", problems=sample)