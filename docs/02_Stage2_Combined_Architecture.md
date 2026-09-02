# Stage 2 — Combined Architecture

Design principle: three layers, in order. Layer 1 (cash flow) sets the ceiling and the knockouts. Layer 2 (history) says how much of that ceiling to use. Layer 3 (offer) applies the adjustment inside the safe range, re-checks every metric, and writes the reasoning.

---

## 1. Workbook structure

The deal-by-deal flow stays on **one sheet that looks like the current `New Template`**. Everything historical sits behind it.

| # | Sheet | Role | Who edits |
|---|---|---|---|
| 1 | `README` | How to run a deal, layer map, limitations | — |
| 2 | `Deal` | = Underwriting Input + Cash Flow Engine. The current template, unchanged in layout, plus a 7-row **Deal Profile** block (ISO, industry, state, new/renewal, credit score, proposed position) where the pricing bands used to sit, a **Credit score** row in the metrics block, and a **Section F — Final Offer** read-out at the bottom. Thresholds are shown but are links to `Controls`. | Underwriter, every deal |
| 3 | `Decision Summary` | Full user-facing output (the 20 items in the brief), stipulations, manual-review triggers, written explanation | Read only |
| 4 | `Offer Engine` | Base cash-flow offer → historical adjustment → adjusted offer, side by side, with post-adjustment metric re-check and the "never exceeds cash-flow max" proof | Read only |
| 5 | `Historical Score` | The 7 matching tiers + 5 single-dimension cohorts, each with the 14 cohort statistics, seasoning test, credibility Z, blended lost rate, relative risk index; primary-tier selection; modifier level; confidence; largest-impact variable | Read only |
| 6 | `Comparable Deals` | The individual house-book deals in the primary cohort (Contract ID, DBA, status, outcome, why included, similarity tier, seasoned flag) — up to 400 rows | Read only |
| 7 | `Historical Data` | 1,067 normalised deals: raw fields as values, status flags via Controls mapping, seasoning helpers, match flags T1–T7 against the current Deal, sequence numbers for the comparable list | Refresh when CRM re-exported |
| 8 | `Controls` | Every editable number: 8 metric thresholds, credit bands, pricing bands, term cap, factor set, fee, conversions, MB constants, status mapping, as-of date and window, seasoning multiple, minimum counts and dollars, credibility constant, modifier table, adjustment limits, confidence rules, stipulation rules, manual-review triggers | Credit policy owner |
| 9 | `Checks` | Missing inputs, formula errors, invalid statuses, small/unseasoned cohorts, conflicting signals, data-quality counts, reconciliation to the original house-book totals, final ≤ cash-flow-max proof | Read only |
| 10 | `Ref` | SIC-2 → industry map, ISO list, state list, Fed holidays, status list (dropdown sources) | Rarely |

Why `Underwriting Input` and `Cash Flow Engine` are one sheet: you asked for the deal flow to stay as simple as the current tool. Splitting inputs from the scorecard would mean tabbing back and forth for every deal. Everything else in the brief's nine-sheet list is present under its own name.

Archiving a deal: copy the `Deal` sheet as today. A copy keeps its Layer 1 numbers (they are self-contained) but its Section F links to the shared historical sheets, which always compute for the sheet named `Deal`. Paste-values on Section F (or on the `Decision Summary`) when archiving. This is a spreadsheet limitation; the CRM version will not have it.

---

## 2. Input dictionary (Deal sheet)

| Cell | Input | Type | Used by |
|---|---|---|---|
| B6 | Requested fund $ | number | Fund/rev metric, requested-offer decision |
| C6 | Requested factor | 1.40/1.45/1.49 | Payback, holdback |
| D6 | Requested term (weeks if weekly, daily payments if daily) | integer | Payment |
| E6 | Frequency | Daily / Weekly | Everything payment-related; final frequency default |
| F6 | Fee % | 5% default (link to Controls) | Net $ |
| B12:F23 | Month, true revenue, avg daily balance, negative days, NSFs — newest first | up to 12 rows | Averages, trend |
| B28:G33 | Existing positions: funder, funded $, payment, freq, start date | up to 6 rows | Leverage, position count |
| K12 | **ISO** | dropdown (house-book ISO list) | Tiers 2, 4; ISO cohort |
| K13 | **Industry (SIC major group)** | dropdown (Ref) | Tiers 1, 3; industry cohort |
| K14 | **State** | dropdown | State cohort (informational) |
| K15 | **New deal or renewal** | New Deal / Renewal | Every tier |
| K16 | **Credit score (FICO)** | number | Metric 8 |
| K17 | Proposed position | auto = existing positions + 1 (override allowed) | Tiers 1, 2, 5; position cohort |

All other numbers on the sheet are formulas. Threshold columns D/E/G in the metrics block are green-font links to `Controls`.

---

## 3. Hard-rule framework (Layer 1, unchanged rules)

1. Seven cash-flow metrics + credit score, each GREEN / YELLOW / RED / KNOCKOUT from `Controls`.
2. Any merchant-side knockout (positions, trend, negative days/NSFs, credit if enabled) → no recommended fund, DECLINE.
3. Offer-side knockout on the *requested* offer (leverage, holdback, fund/rev) → requested offer is DECLINED but the engine still sizes a green-line offer.
4. Red flags (RED + KNOCKOUT) ≥ limit (2) → DECLINE.
5. Recommended fund below the minimum ($20,000) → DECLINE.
6. Cash-flow maximum daily payment = MIN(holdback green line, total-leverage green line, daily-debt/balance green line). Term = MIN(score band cap, 9 months, fund/rev green line). Factor from score band. Fund = FLOOR(payment × term ÷ factor, 500).
7. **Nothing in Layers 2–3 can raise the fund above item 6, remove a knockout, or push a green metric to red.** The Offer Engine recomputes all four offer-side metrics at the adjusted offer and Checks proves it.

Credit score bands (your specification): green ≥ 650, yellow 630–649, red < 630. Knockout line on Controls, default 0 (off). Blank score = metric not scored; the score denominator shrinks to 7 and the sheet flags the missing input.

---

## 4. Historical-comparison framework (Layer 2)

**Data basis.** Seasoned deals only (age in business days ≥ 1.25 × contracted term, as-of date on Controls). Unseasoned deals are counted and shown for every cohort but never enter a rate.

**Matching hierarchy (recommended after the Stage 1 tests).** New-vs-renewal is the strongest surviving signal, so it is in every tier. Industry and ISO out-rank position, which was weak once the others were controlled.

| Tier | Keys | Why |
|---|---|---|
| 1 | New/renewal + industry + position | Closest match |
| 2 | New/renewal + ISO + position | ISO relationship at the same stack depth |
| 3 | New/renewal + industry | Industry survived controls |
| 4 | New/renewal + ISO | ISO survived controls (Westwood, PMF-suggestive) |
| 5 | New/renewal + position | Position only within the renewal split |
| 6 | New/renewal | The book split |
| 7 | Seasoned book | Baseline |

**Primary cohort** = the first tier whose seasoned count ≥ `MinDeals` (10) **and** seasoned funded $ ≥ `MinFunded` ($150,000). Every tier is computed and displayed regardless, so the underwriter can see the closer-but-thinner cohorts.

**State** is computed as a single-dimension cohort and displayed with a fixed warning; it never enters the hierarchy. Position, ISO, industry and new/renewal single-dimension cohorts are also shown for the "largest impact" test.

**Per-cohort statistics (all 12 cohorts):** deals (all), seasoned deals, % seasoned, total funded (seasoned), open n / %, defaulted n / %, principal lost rate, collected on defaults, return on funded capital, avg advance/revenue (where revenue exists), avg term (business days), avg factor, avg deal size, avg days on book, renewal share, seasoned-enough flag, credibility Z, blended lost rate, relative index.

**Credibility (Bühlmann-style, transparent).**
`Z = n_seasoned ÷ (n_seasoned + K)`, K = 30 seasoned deals (Controls).
`Blended lost rate = Z × cohort lost rate + (1 − Z) × seasoned-book lost rate`.
`Relative index = blended ÷ seasoned-book lost rate` (1.00 = book average).
A 5-deal ISO gets Z = 0.14 and moves the index by at most a seventh of its raw gap; a 100-deal cohort gets Z = 0.77.

**Confidence level.**
- High: Z ≥ 0.60 and cohort open share ≤ 15%.
- Medium: Z ≥ 0.30.
- Low: everything else, or primary tier = 7 (no specific match).
Low confidence caps the modifier at one step from neutral in either direction and adds a manual-review-of-history stipulation. Adjustments below Medium confidence are never called "precise": the explanation says the evidence is thin.

**Largest-impact variable** = the single-dimension cohort (position, new/renewal, industry, ISO) with the largest |blended lost − book lost|, reported with its numbers. State is excluded by design.

---

## 5. Offer-adjustment methodology (Layer 3)

Modifier table (Controls, editable):

| Relative index | Level | Fund multiplier | Term step | Factor step | Frequency | Manual review |
|---|---|---|---|---|---|---|
| ≤ 0.60 | Strong positive | 1.00 (full cash-flow max) | 0 | 0 | as proposed | no |
| ≤ 0.85 | Mild positive | 0.95 | 0 | 0 | as proposed | no |
| ≤ 1.15 | Neutral | 0.90 | 0 | 0 | as proposed | no |
| ≤ 1.40 | Mild negative | 0.80 | −1 band | +1 tier | as proposed | no |
| ≤ 1.75 | Negative | 0.70 | −1 band | +1 tier | Weekly | no |
| > 1.75 | Severe | 0.55 | −2 bands | +1 tier | Weekly | **yes** |

Mechanics:
- Adjusted factor = next tier up in {1.40, 1.45, 1.49}, never above 1.49.
- Adjusted term = the band table row `step` places lower (195 → 150 → 130 → 100), floor 20 payments.
- Adjusted fund = MIN(base fund × multiplier, FLOOR(cash-flow max daily × adjusted term ÷ adjusted factor, 500)). The second term guarantees the daily payment never exceeds the cash-flow maximum even after the term is shortened.
- Payment = fund × factor ÷ term (daily); weekly = × 5.
- All four offer-side metrics are recomputed at the adjusted offer; if any is RED or KNOCKOUT the fund is reduced further until green (this cannot happen by construction, but the check is displayed).

Final decision:
- DECLINE if Layer 1 declines (knockout, red-flag limit, below minimum).
- MANUAL REVIEW if any trigger on Controls fires: severe modifier; low confidence with a red metric; credit score red plus any other red; positions ≥ 4; new 3rd-position deal in a construction industry; new ISO/industry never seen in the book; adjusted fund below minimum.
- CONDITIONAL if Layer 1 is CONDITIONAL, or the modifier is negative, or any conditional stipulation fires.
- APPROVE otherwise.

Stipulations: standard set always (DataMerch, UCC search, positions/payoff letters). Conditional rules on Controls: no-stacking covenant with default-to-daily (negative modifiers), bank login verification (neg-days/NSF not green), payoff letter per existing position (positions ≥ 2), landlord/lease letter (credit not green), manual review of history (low confidence), weekly remittance (frequency switched), 12-month statements (renewal claimed but no house-book renewal history).

---

## 6. Worked example (actual model output, scenario A in `build/run_scenarios.py`)

New deal, General Building Contractor, TX, ISO United Secured Capital 8, FICO 640, requested $75,000 @ 1.49 / 30 weeks. Four months of true revenue averaging $119,500, average balance $9,000, one negative day, one NSF, one existing weekly position (Fox Funding, $60,000 funded, $2,500/week). Positions incl. new = 2.

| Step | Result |
|---|---|
| Layer 1 metrics | Leverage 22.6% G, holdback 13.5% G, daily debt/balance 13.8% G, fund/rev 0.63 Y, positions 2 G, trend +2.6% G, neg days+NSF 2 Y, credit 640 Y |
| Layer 1 decision (requested offer) | CONDITIONAL, score 6.5 of 8 (scaled 5.7/7 = Solid) |
| Cash-flow maximum | $993/day, limited by holdback; term 100 daily payments (fund/rev green line caps it); factor 1.45; **max fund $68,500** |
| Primary cohort | Tier 1: new + General Building Contractors + position 2 — 27 matches, 23 seasoned, $414k funded, 4% open, 48% defaulted, **principal lost 34.0%** vs 21.1% seasoned book |
| Credibility | Z = 23/(23+30) = 0.43 → blended 26.7% → index 1.27 → Mild negative; confidence Medium |
| Adjustment | fund × 0.80, term −1 band (already at the 100 floor), factor +1 tier (1.45 → 1.49) |
| Final offer | **$54,500 @ 1.49, 20 weeks, $4,060/week** (daily-equivalent $812 ≤ $993 max); holdback 14.7% G, leverage 24.9% G |
| Decision | CONDITIONAL. Stips: standard + no-stacking covenant with default-to-daily, bank login verification (neg days/NSF yellow), landlord letter (credit yellow) |
| Largest historical impact | Industry (82 seasoned GBC deals, 29.4% blended vs 21.1%) |
| Evidence notes | ISO USC8 109 seasoned deals lost 26.5% (weaker); TX 118 deals lost 13.0% — informational only |

Written conclusion produced by the sheet: "Cash flow supports up to $68,500 (SP% holdback is the limit, 100 daily pmts @ 1.45). The recommendation is $54,500 (80% of the maximum) at 1.49 over 20 weeks, $4,060/week, because the primary comparable cohort (T1: new/renewal + industry + position) holds 23 / 27 seasoned deals with a principal lost rate of 34.0% against 21.1% for the seasoned book; at credibility 0.43 that blends to 26.7% (index 1.27), which is mild negative. Confidence is medium. ISO United Secured Capital 8: 109 seasoned deals, lost 26.5% vs 21.1% book, Z 0.78 -> weaker than book. Industry General Building Contractors: 82 seasoned deals, lost 32.5% vs 21.1% book, Z 0.73 -> weaker than book. State TX: 118 seasoned deals, lost 13.0% vs 21.1% book - informational only, not used in the offer."
