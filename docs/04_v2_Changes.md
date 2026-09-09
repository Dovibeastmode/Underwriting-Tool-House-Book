# v2 changes (built by `build/build_model_v2.py`; `demo` argument = 8 sample deals + populated Deal sheet)

**History can now loosen the cash-flow sizing.** The modifier table on Controls (section 5) has a new column, *Sizing flex*: 0 = size at the green lines (the tool exactly as built), 1 = size at the red lines. Recommended values: Strong positive 90%, Mild positive 45%, everything else 0. Section C on the Deal sheet shows the sizing line in use in column H (purple); section E sizes the max offer at those lines, so a supportive book produces a larger offer whose metrics may show YELLOW. Flex is capped below 100%, so RED is impossible. Negative evidence keeps the green lines and applies the fund multiplier, band step, factor step and weekly rule as before.

Full-book test (renewal restaurant via Westwood): green-line max $85,500 → sized max $105,500 at 45% flex; holdback 21.1% and fund/revenue 0.73 turn YELLOW, nothing red; every safety proof OK.

**Section C extended (C2, rows 49–57).** Same Green / Red / Status layout: primary cohort, ISO, industry, new/renewal, position, state (informational), and the seasoned-deal count behind the primary cohort. Green/red lines are the modifier table's 0.85 / 1.15 index bounds expressed as lost rates. These rows are not counted in the cash-flow score or red flags.

**Removed:** confidence level and its cap, stipulations (Decision Summary and Controls section 8), the written conclusion paragraph, the proposed-position input and override (matching reads C42). State is optional; a blank state never blocks an offer.

**Decision Summary** is now a bulleted read-out: decision, offer, sizing line used, the four offer-side metrics with status, modifier, cohort numbers, largest impact, limiting metric, triggers, one evidence line per dimension, warnings, reds and greens at the requested offer.

**Demo file minimums** are lowered (3 seasoned deals / $30,000 / rate shown at 2 / indicator green at 5) so eight sample deals produce cohorts. The production values (10 / $150,000 / 5 / 30) are noted beside each cell on Controls.

---

# v3 changes (`build/build_model_v3.py`)

- **Credibility and blend removed.** Index = primary-cohort principal lost rate ÷ seasoned-book principal lost rate. Side notes on Controls (row 42) and Historical Score (row 2) record that a blend rate and credibility factor are under consideration.
- **Historical Score cut to eight columns:** Tier, Cohort, Seasoned deals, Funded (seasoned), Meets min?, Principal lost rate, Return on funded, Index. Single-dimension rows kept (position, new/renewal, industry, ISO) for the C2 block and the largest-impact test; state removed. RESULT block: primary tier (lowest tier meeting both minimums), its PLR / ROF / index, modifier row, largest-impact variable, warnings.
- **Largest-impact variable:** column I holds |PLR − book PLR| for each single-dimension row; MAX picks the biggest gap, MATCH finds its row, INDEX returns the name.
- **Seq removed.** Comparable Deals is now a filterable copy of the book with "Best tier" and "In primary cohort" flag columns (filter column V to 1).
- **State removed** from the Deal profile, matching columns, C2, and the summary. Credit score input moved to K15; new/renewal to K14.
- **C2 block** shows PLR, green/red lines, status, seasoned n, and return on funded for the primary cohort and the four single dimensions.
- Demo consequence: with no credibility, 5 New-Deal sample deals at 50.8% PLR vs a 41% book give index 1.24 → Mild negative (fund × 0.85, factor 1.49). Full-book Westwood renewal restaurant: T3 (23 seasoned, PLR 2.5%) → index 0.12 → Strong positive, 90% flex, $85,500 → $121,500.

---

# v4 changes (`build/build_model_v4.py`)

- **Sheets removed:** Decision Summary, Offer Engine, Checks. Workbook is now README, Deal, Historical Score, Comparable Deals, Historical Data, Controls, Ref.
- **Section F is an outline** (planned rows with plain-English logic, no formulas). Sections A–E stay live; E sizes the max offer at the sizing line.
- **Three-outcome house-book modifier** (Controls section 5): Supportive when index ≤ 0.85 → sizing flex 50%; Negative when index > 1.15 → fund cut 20% (planned, in F); otherwise Neutral = the tool as built. Manual-review triggers removed.
- **Historical Score results** cut to: primary tier, cohort, seasoned n, PLR, ROF, index, outcome, sizing flex, largest-impact variable.
- **Colour-blind Controls:** no blue. Editable cells are bold black on grey with a thick border; links to Controls are italic grey. Statuses remain written as words.

---

# v5 changes (`build/build_v5_inplace.py`, applied in place on the user's edited v4 demo so their formatting and text edits are kept)

- **Cash-flow thresholds live on the Deal sheet again** (typed values in D/E/G of section C, C61 red-flag limit, E61 minimum fund, F6 fee, K8:L9 conversion constants), exactly as in the original tool. Controls section 1 removed.
- **Controls** now: 1 pricing bands + max term, 2 conversion constants, 3 historical settings, 4 house-book outcome, 5 status mapping. Removed: factor tiers, tier cut-offs (hard-coded 6.5 / 5 / 3.5 in the Tier label formula), credibility constant, minimum-deals-before-a-rate-is-shown. Rates now always show when the cohort has funded dollars.
- **Historical Score:** no tier column; cohort names without "T1:" prefixes; RESULT = primary cohort (name, with the row number beside it for Historical Data), seasoned deals, PLR, ROF, index, outcome, sizing flex, largest-impact variable.
- **Historical Data:** tier flag headers renamed (NR+Ind+Pos, NR+ISO+Pos, ...). Comparable Deals column U renamed "Best match".
- Helper cells D65/D67 (metrics scored, scaled score) that were deleted in the edited file are folded into the formulas that used them (C62, C66, C72, D72).
- Two files: SAMPLE (8 deals, minimums 3 / $30k / 5) and FULL (1,067 deals, minimums 10 / $150k / 30). Both carry the same populated Deal sheet (now a 2nd-position deal because of the sample position added on row 28).
- **Section F is live (v5.1):** final decision; green-line max (cash flow only) shown beside the max at the sizing line (C73) so the house-book lift is visible; recommended fund = C73, or C73 × (1 − cut) when the outcome is Negative; factor, term, payment, daily-equivalent; the four offer-side metrics re-checked at the final offer with GREEN / YELLOW / RED; house-book outcome with the primary cohort's numbers. Full-book Westwood renewal test: green-line max $85,500 → C73 $105,500 at 50% flex → recommended $105,500, holdback and fund/revenue YELLOW, nothing red.

---

# v6 changes (`build/build_v6.py`, then `build/finalize_xlookup.py`)

1. **Seasoning** = age ≥ multiple × term **OR** the deal reached a terminal status (Closed family, Written-Off, Bankruptcy, Settlement). Full book: 637 → 716 seasoned deals, book PLR 21.1% → 16.3%, ROF +1.7% → +10.3%.
2. **Dead columns cut**: AZ (duplicate of AR) and BA (always 1) removed; BB (Best tier) and Comparable Deals column U removed. Tier flags now AU–AZ + AR, "In primary cohort" moved to BA.
3. **New tier 1: New/renewal + ISO + industry.** Order: ISO+Ind → Ind+Pos → ISO+Pos → Ind → ISO → Pos → NR → Book. No four-way tier (it never qualified on the full book).
4. **Historical Score C36** → `XLOOKUP`, returns the dimension name only. Written after recalculation so LibreOffice cannot rewrite it; it evaluates in Google Sheets and Excel 365, and shows `#NAME?` in LibreOffice or pre-2021 Excel.
5. **Redundant guards stripped** in C2 columns C and H (`IF(x="","",x)` → `x`).
6. **Section E green-line column** (F67:F73) runs the same chain at the green lines, so the house-book lift is visible as C73 − F73.
7. **Section E hidden** (rows 64–74).
8. **Section F restructured**, no detail column: Final Decision, Cash flow only (9pt), RECOMMENDED FUND, Factor, Term (weeks), Term (daily pmts), Weekly payment, Daily payment, House-book outcome. The four metric re-checks moved to columns I:K, off the printed area, still feeding the decision.

---

# v7 fixes (`build/apply_v7_fixes.py`) — blank-deal behaviour

The whole-seasoned-book row always met the minimums, so with no deal profile entered it became the "primary cohort" and printed the book's own PLR, seasoned count and ROF as if they were evidence. Section F then produced a full recommended offer off bank data alone.

- `Historical Score D29` returns blank when new/renewal (C5) is empty, so no cohort is selected at all.
- `C29`–`C33` blank out when D29 is blank **or** 8 (the book row). This restores a guard that was removed in v6 — it was doing real work.
- `C34` returns Neutral in both cases, so sizing flex stays 0.
- `Historical Data BA` guards `CHOOSE` against a blank index (otherwise `#VALUE!` on every row).
- Deal `F51`, `G51`, `C56`, `F56` follow, removing the contradiction where row 51 showed 716 seasoned deals and row 56 showed 0.
- Deal Section F (`C76`, `C77`, `C78`, `C84`) stays blank until new/renewal is entered. Section E still sizes from bank data alone, which is the original tool's behaviour and is unchanged.

Verified: everything blank → F empty. Bank data, no profile → E sizes $51,000, F empty, C2 reads "- no deal entered -". Full profile → cohort selected, offer produced, no regression.
