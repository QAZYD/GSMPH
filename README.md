# extract_pool.py

Builds the pool of Filipino replacement terms for localizing GSM8K (GSMPH). For each tagged term subcategory in the term bank, it collects culturally relevant Tagalog/Filipino terms, brand names and person names, and checks that each pool is big enough (tagged unique terms + 10% excess).

|  |  |
|---|---|
| **Input** | `GSMPH_Cleaned_Terms.xlsx` (sheet `Lookup Table (Cleaned)`) |
| **Output** | `pool.xlsx` |

## Setup

```
pip install requests openpyxl
```

## Run

Run the commands one line at a time (works in PowerShell and bash).

**1. Download the Tagalog dictionary** (large file, one-time)

```
python extract_pool.py fetch-dict --contact you@example.com
```

**2. Check the real category/topic names, and tune `DICT_RULES` if needed**

```
python extract_pool.py discover-dict kaikki_tagalog.jsonl
python extract_pool.py discover --contact you@example.com
```

**3. Full run**

```
python extract_pool.py run --contact you@example.com --dict kaikki_tagalog.jsonl --names-file names.csv --out pool.xlsx
```

PowerShell users can split a long command with a backtick `` ` `` at the end of each line. The bash `\` does not work there.

### Other commands

| Command | What it does |
|---|---|
| `python extract_pool.py targets` | Prints how many terms each pool needs (works offline) |
| `python extract_pool.py list-pools` | Lists every pool and its sources (Wikipedia, dictionary, curated) |
| `python extract_pool.py run --pools "Pets,Fruit" --no-views` | Quick test of a few pools |

### Useful `run` options

| Option | Default | Meaning |
|---|---|---|
| `--dict PATH` | none | kaikki.org Tagalog dictionary, used for generic-noun pools |
| `--names-file PATH` | none | Names list for the Person name pool: `name,count,gender` (CSV, or an Excel file) |
| `--targets PATH` | `GSMPH_Cleaned_Terms.xlsx` | Term bank used for the Coverage sheet |
| `--buffer` | `0.10` | Excess over tagged unique terms (10%) |
| `--min-score` | `3.0` | Wikipedia candidates below this recognition score go to `Rejected (auto)`; `0` disables |
| `--dict-cap` | `400` | Max dictionary rows per pool |
| `--no-views` | off | Skip pageview lookups (faster, no recognition filter) |
| `--allow-english-loans` | off | Keep dictionary words Wiktionary marks as borrowed from English |

## Where terms come from

| Source | Used for |
|---|---|
| Wikipedia / Wikidata category trees | Foods, animals, clothing, places, brands, events |
| kaikki.org Tagalog Wiktionary (`--dict`) | Generic nouns: places & buildings, containers, household, pets, staples, dairy, meat, and more |
| `CURATED` lists in the script | Everyday Filipino terms, brands, institutions, streets, first names (need review) |
| `--names-file` | Person names (the term bank needs about 2,400) |

Dictionary filter: words Wiktionary marks as borrowed from English are dropped, as are archaic, dialectal, rare and slang senses, and multi-word phrases. Spanish loans (silya, kutsara) are kept.

`DICT_RULES` in the script maps each pool to two regex patterns: `cats` (matched against Wiktionary categories/topics) and `gloss` (matched against the English meaning). Use `discover-dict` output to tune them.

## Output workbook (`pool.xlsx`)

| Sheet | Contents |
|---|---|
| `README` | Run notes, sources, licenses |
| `Coverage` | Per pool: tagged terms, target (with buffer), candidates, shortfall, status. The buffer cell is editable and the sheet uses formulas |
| `Pool` | All candidates with English, Tagalog, description, recognition tier, source. Set **Reviewer status** to Approve / Reject / Maybe |
| `Sources` | Categories and files used, with counts |
| `Run Log` | Warnings, such as empty categories or missing files |
| `Generic Terms` | Each tagged generic term with its literal dictionary translation (what not to use blindly) |
| `Rejected (auto)` | Candidates dropped for low recognition, kept for audit |

## Notes

- Currency amount pools are rule-based conversions, not term pools. Coverage marks them as such.
- Many "check context" terms (box, party, milk, dog) will probably end up Keep, so the real need can be lower than the Coverage targets.
- Category names in the script are seeds. Check the `Run Log` sheet for categories that returned nothing.
- Wikipedia/Wiktionary text is CC BY-SA 4.0 and Wikidata is CC0. Keep the Source and URL columns if you redistribute the pool. Cite Ylonen, *Wiktextract*, LREC 2022, for the kaikki.org data.
- Don't commit large or regenerable files. Add to `.gitignore`: `kaikki_tagalog.jsonl`, `.pool_cache/`, `pool.xlsx`, `names.csv`.
