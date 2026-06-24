# FEC Donor Lookup

A batch-processing tool that takes a spreadsheet of names, matches each one
against the FEC's public campaign finance database, resolves duplicate and
ambiguous records (same name, different address; pass-through donations via
platforms like ActBlue), and applies configurable eligibility rules to flag
which entries are worth a closer look — with the reasoning for every call
written out in plain English.

Originally built for political campaign donor research; the matching engine
and config system are domain-agnostic and reusable for any "look up a list
of names against a public record source" problem.

**Technical highlights:**
- Tiered entity resolution against a fuzzy-matched public API (name → state →
  employer/address disambiguation → fallback employer search), handling
  common-name collisions and incomplete input data
- Concurrent processing (`ThreadPoolExecutor`) with thread-safe caching,
  automatic retry/backoff, and per-task timeouts so one slow record can't
  stall a batch run
- Clean separation between configuration (`config.json`) and business logic
  (`evaluate_donor()`), so the eligibility rules are swappable without
  touching the data-fetching engine
- Deduplication logic for a real-world data quirk (donations routed through
  third-party platforms get double-recorded by the source API)

## What it actually does (plain-language version)

1. You give it a spreadsheet — one row per person, with at least a first and
   last name.
2. For each person, it asks the FEC's public database: "what federal
   campaign donations has this person made?"
3. It cleans that up — merging duplicate records, figuring out if two
   donations from slightly different addresses are the same person, etc.
4. It applies your rules ("don't bother with anyone who works at X," "only
   flag people who've given $Y+ recently," etc.) and writes a plain-English
   explanation of why each person was flagged or not.
5. You get one summary spreadsheet (everyone, one row each) and one detailed
   text file per person (their full donation history + the reasoning).


It does **not** make final decisions for you — it surfaces information and a
recommendation, but every "CONSIDER" or "DO NOT CONSIDER" call comes with the
reasoning spelled out so a human can sanity-check it.

## Setup

### 1. Get a free FEC API key
Go to https://api.data.gov/signup/ — takes about a minute, the key arrives
by email instantly. This is free and has no usage cost; it's a public
government API.

### 2. Install Python dependencies
```bash
pip3 install requests pandas openpyxl certifi
```

### 3. Set up your config
Copy the example config and fill in your own values:
```bash
cp config.example.json config.json
```
Open `config.json` and:
- Paste your API key into `"api_key"`
- Set `"input_file"` to your spreadsheet's filename
- Check the `"columns"` section matches your spreadsheet's actual headers
- **Replace the example eligibility criteria** — see "Customizing your
  criteria" below. The shipped values (`"example corp"`, state `"XX"`, etc.)
  are placeholders and won't do anything useful as-is.

`config.json` is in `.gitignore` so your API key and your organization's
actual criteria never get committed to version control by accident.

### 4. Edit the eligibility logic
Open `fec_donor_lookup.py` and find the function `evaluate_donor()` — it's
clearly marked with a big comment block. This is where the actual judgment
calls happen. The placeholder logic shipped here shows the *shape* real
criteria usually take, but you'll want to replace it with your own rules.
See **PROMPTS.md** for ready-to-use prompts you can hand to an AI coding
assistant to do this.

### 5. Prepare your spreadsheet
CSV or Excel (.xlsx), one row per donor. Required columns: First, Last.
Strongly recommended: City, State, Company, Occupation, Address (the more
of these you have, the more accurately the tool can tell two people with
the same name apart).

### 6. Run it
```bash
python3 fec_donor_lookup.py
```

You'll see live progress in the terminal. When it's done, you'll have:
- `fec_results_summary.csv` — one row per donor, with the final call
- `fec_results/` folder — one detailed `.txt` file per donor

## Customizing your criteria

Almost everything you'd want to change lives in one of two places:

| If you want to change...                                | Edit this                          |
|-----------------------------------------------------------|-------------------------------------|
| Dollar thresholds, date cutoffs, donation counts         | `config.json`                       |
| Disqualifying employers or occupations                   | `config.json`                       |
| Which state/region gets special-cased                    | `config.json`                       |
| Spreadsheet column name mapping                          | `config.json`                       |
| The actual *logic* connecting those rules (the "if X and not Y but Z overrides" branching) | `fec_donor_lookup.py` → `evaluate_donor()` |

Real-world eligibility criteria are almost never a single threshold — they're
usually a handful of "yes, but only if..." branches. That's why the branching
logic lives in code rather than being squeezed into a config file. The
config file handles the numbers and lists; the Python function handles how
they fit together.

**You do not need to know how to code to do this.** Open PROMPTS.md, find the
prompt that matches what you want to change, fill in your specifics, and
hand it to an AI coding assistant (Claude, ChatGPT, etc.) along with your
copy of `fec_donor_lookup.py`.

## How the matching works (so you can trust the output)

A name search against FEC records runs into a basic problem: lots of people
share a name. This tool handles that with a tiered approach:

1. Searches by name + state.
2. If everyone in the results shares an employer, treats them as the same
   person and stops there.
3. If employers diverge, checks street address first, then employer name,
   then city as a last resort — in that priority order, because street
   address is the strongest signal and city alone is the weakest.
4. If results are still thin, tries a search by name + employer instead.
5. If results are *still* thin, flags the donor for manual review with a
   direct link to search FEC.gov yourself.

It also automatically removes duplicate records that happen when someone
donates through a platform like ActBlue — the FEC records that as two
transactions (one to the platform, one to the actual campaign), and this
tool collapses them back into one so donation counts aren't inflated.

None of this is perfect. Common names, donors who've moved, and incomplete
spreadsheet data will sometimes produce a wrong or uncertain match — that's
why every donor gets a `SEARCH NOTES` section in their output file
explaining exactly what the tool found and any uncertainty involved, so a
human can spot-check before relying on it.

## Performance notes

- Runs donors in parallel (`max_workers` in config.json, default 15). Lower
  this if you're hitting FEC rate-limit errors; raise it if you have a fast
  connection and want to go faster.
- A 90-second per-donor timeout prevents one slow lookup from blocking the
  whole batch — timed-out donors get flagged in the summary for manual
  follow-up rather than crashing the run.
- Results are cached in-memory during a run, so re-querying the same
  name+criteria combination (which happens during the tiered search) doesn't
  cost extra API calls.

## Limitations

- This only covers **federal** campaign contributions (House, Senate,
  President, PACs registered federally). It has no visibility into state or
  local political giving.
- FEC data is reported by committees and isn't always clean — names,
  employers, and occupations are free-text fields filled in by the donor or
  committee, so typos and inconsistent formatting exist in the underlying
  data itself.
- This tool flags and explains; it doesn't replace human judgment on
  ambiguous matches.

## Contributing

Issues and pull requests welcome — especially around improving the entity
matching logic, since name/employer disambiguation is the hardest and most
imperfect part of this tool. See `PROMPTS.md` if you're adapting the
eligibility rules for your own use case rather than contributing back.

## License

MIT — see `LICENSE`.
