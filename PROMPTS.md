# Prompts for customizing this tool with AI assistance

You don't need to know Python to adapt this tool to your own criteria. The
prompts below are templates — copy one, fill in the bracketed parts with
your specifics, and give it to an AI coding assistant (Claude, ChatGPT,
etc.) along with your copy of `fec_donor_lookup.py` and `config.json`.

**General tips for getting good results:**
- Always attach or paste the current `fec_donor_lookup.py`, not just describe
  it from memory — the AI needs to see the actual code.
- Ask for the change to `config.json` AND `evaluate_donor()` together if
  your rule touches both — they need to stay in sync.
- Ask the AI to show you a before/after diff or to clearly mark what it
  changed, so you can review it.
- After any change, re-run the script on a small test spreadsheet (5-10 known
  donors where you already know the right answer) before running it on your
  full list.
- If a change doesn't work right, paste the error message back and ask for a
  fix — don't guess.

---

## 1. Add a new disqualifying employer or occupation

> I'm using a Python donor-vetting script (attached: `fec_donor_lookup.py`
> and `config.json`). I want to add [EMPLOYER NAME / list of employer names]
> to the disqualifying employer list. Please update `config.json` only —
> don't touch the Python code, since `is_disqualifying_employer()` already
> reads from config.

> Same idea but for occupations: I want anyone in [STATE CODE] with the
> occupation [OCCUPATION] to be excluded entirely. Update
> `disqualifying_occupations_by_state` in `config.json` for that state.

## 2. Change a dollar threshold or date cutoff

> In `config.json`, change `min_single_donation` to $[AMOUNT]. Don't change
> anything else.

> I want the "high value" donation threshold in the standard eligibility
> path to be $[AMOUNT] instead of the current value, and the "recency
> window" to be [N] years instead of the current value. These are in the
> `standard_path` section of `config.json`.

## 3. Add or change the "special state" rules

> Right now `special_state.code` in `config.json` is set to [CURRENT STATE
> OR EMPTY]. I want to change this to [NEW STATE CODE], and give it these
> rules: [describe — e.g. "donors need 15+ donations in the last 8 years to
> be a top-tier match, or one donation of $7,500+ to be a mid-tier match"].
> Update both `config.json` and the special-state block inside
> `evaluate_donor()` in `fec_donor_lookup.py` to match.

> I want to disable the special-state logic entirely and treat every donor
> the same way regardless of state. Set `special_state.code` to an empty
> string in `config.json` and confirm `evaluate_donor()` correctly skips
> that block when it's empty (it should already — just verify).

## 4. Add a brand new tier or rule that doesn't exist yet

This is the most common real-world request, because real criteria are
rarely simple. Be as specific as possible about the *exact* condition.

> I want to add a new rule to `evaluate_donor()` in `fec_donor_lookup.py`.
> Here's the rule in plain English: [describe precisely — e.g. "if a donor
> has given to both a Democratic and a Republican committee in the same
> election cycle, flag them for manual review instead of auto-approving or
> rejecting, regardless of dollar amount"].
>
> Please:
> 1. Show me where in the function this check should go relative to the
>    existing rules (before or after which step, and why)
> 2. Write the actual code
> 3. Explain in plain English what changed and what would trigger it
> 4. Tell me if this needs any new fields in `config.json` or if it can be
>    self-contained in the function

## 5. Change which spreadsheet columns it reads

> My spreadsheet uses different column headers than the default. My columns
> are: [list your actual headers, e.g. "FNAME, LNAME, HOMETOWN, ST,
> EMPLOYER, JOB_TITLE"]. Update the `"columns"` section of `config.json` to
> map to these.

## 6. Add a brand new disqualifying or flagging category that isn't employer/occupation/state

> I want to add a new kind of disqualifying check that isn't covered by the
> existing employer/occupation/state logic. Here's what I want excluded:
> [describe — e.g. "anyone whose listed address matches a known PO box list
> I maintain separately" or "anyone under 18" or "donors flagged on this
> separate CSV I have of people we've already contacted this cycle"].
>
> Walk me through whether this is simple enough for `config.json` (a list or
> threshold) or whether it needs new logic in `evaluate_donor()`, and
> implement whichever is the better fit.

## 7. Adjust how aggressively the tool matches donors with the same name

> The tool is currently flagging too many / too few [pick one] contributions
> as "may be a different person" when employer or address doesn't match.
> Here's an example of a case it got wrong: [paste the relevant SEARCH NOTES
> from a donor's .txt file, and explain what the correct answer should have
> been]. Can you adjust the matching logic in `fetch_contributions()` to
> handle this better, and explain what tradeoff you're making (since being
> more permissive about matches means more false positives, and being
> stricter means more false negatives)?

## 8. Export or reformat the output differently

> I want the output changed: [describe — e.g. "add a column to the summary
> CSV showing the donor's total dollar amount across all donations, not just
> their highest single gift" or "I want the individual .txt files renamed
> using format X instead"]. Update `save_donor_file()` and/or the
> `process_donor()` function inside `main()` in `fec_donor_lookup.py`
> accordingly.

---

## A note on judgment calls

Some things are genuinely ambiguous and don't have a clean rule — for
example, deciding whether two donation records with slightly different
addresses are the same person, or whether someone's stated occupation is
disqualifying when it's phrased unusually. For cases like that, it's often
better to let the tool flag the donor for manual review (it already does
this in several places — look for "VERIFY" and "MANUAL REVIEW" in the
output) rather than trying to write a rule that handles every edge case
automatically. A human spending 30 seconds on a flagged edge case usually
beats a rule that's wrong some percentage of the time across your whole
list.
