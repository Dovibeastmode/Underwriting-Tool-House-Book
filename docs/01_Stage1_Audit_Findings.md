# Stage 1 — Audit and Findings

Scope: `Updated Underwriting Tool (1).xlsx`, `House Book Analytics (2).xlsx`, `Aspire Book Analysis_V2.pptx`, `session_log.md`, `House_Book_Session_Log.md`. Untouched copies of all five are in `originals/` (MD5s recorded in `originals/CHECKSUMS.md`).

All figures below were recomputed from the raw `Data` sheet, not read off the deck.

---

## 1. How the Underwriting Tool works today

**Shape.** 78 sheets (the session log says 68 + 1). Live template = `New Template`. 62 `Copy of New Template n` tabs, 7 contract-number tabs (`440559`, `438249`, `435983`, `431725`, `420320`, `4226498`, `426120`, `425209`), `Hypothetical`, `test`, and 4 `Copy of Template n` tabs in an older 51-row format that is a different tool (average-balance-over-revenue box, no offer engine).

**Data flow of `New Template` (one deal per tab).**

| Block | Cells | What it does |
|---|---|---|
| A. Offer | B6 fund, C6 factor, D6 term, E6 freq (Daily/Weekly), F6 fee 5% | B8 payback = fund × factor. C8 daily $ = payback ÷ term (÷5 if weekly). D8 weekly = daily × 5. E8 monthly = daily × 21.655. F8 term months = weeks ÷ 4.331. G8 net = fund × (1 − fee). |
| MoneyBadger cross-check | J4:M9 | L7 = ROUND(weekly×4 or daily×21 ÷ avg rev, 2 dp). L8 = 4, L9 = 21 (MB constants), K8 = 4.331, K9 = 21.655 (ours). |
| B. Bank data | B12:F23, newest month first | C24 avg true revenue, D24 avg daily balance, E24 total neg days, F24 total NSFs, G24 average month-over-month change. |
| Pricing bands | J12:L16, K17 | Score ≥0 → 100 pmts @1.49; ≥5 → 130 @1.45; ≥6 → 150 @1.45; ≥7 → 195 @1.40. K17 max term 9 months. |
| Existing positions | B28:J33 | F = daily equivalent (weekly ÷ 5, monthly ÷ 21.655; funded-only rows spread at 1.45 over 9 months). H est. term in payments = funded × 1.45 ÷ payment. I est. term months. J months remaining. C34 position count, E34 total existing daily, F34 monthly. |
| C. Metrics | B38:G44 | 7 metrics, each Green ≤ / Red ≥ / Knockout ≥ (trend is inverted). Status text drives conditional colour. |
| D. Decision | C49:E51 | Red flags = RED + KNOCKOUT count. DECLINE if any knockout, or recommended fund < E50 minimum ($20k), or red flags ≥ C50 (2). APPROVE only if all 7 green. Else CONDITIONAL. |
| E. Offer engine | C54:D62 | Score = greens + ½ yellows (of 7). C56–C58 = max daily payment that keeps holdback, total leverage and daily-debt/balance at their **green** lines. C59 = MIN of the three (the limiting constraint is named in D59). C61 term = MAX(20, MIN(score band cap, 9 mo × 21.655, term that keeps fund/rev ≤ green)). D61 factor from band. C62 fund = FLOOR(C59 × C61 ÷ D61, 500). |

**Threshold set in force on the live template**

| Metric | Green | Red | Knockout |
|---|---|---|---|
| Total leverage (all debt ÷ revenue) | ≤ 30% | ≥ 45% | ≥ 55% |
| New holdback % | ≤ 18% | ≥ 25% | ≥ 35% |
| Daily debt ÷ avg balance | ≤ 18% | ≥ 30% | ≥ 35% |
| Fund ÷ monthly revenue | ≤ 0.60x | ≥ 0.90x | ≥ 1.25x |
| Positions incl. new | ≤ 2 | ≥ 5 | ≥ 6 |
| Revenue trend (avg MoM) | ≥ −5% | ≤ −15% | ≤ −30% |
| Neg days + NSFs (all months entered) | ≤ 1 | ≥ 6 | ≥ 10 |

**Rules that are intentional (confirmed by log + formulas)**
- Offer-side knockouts (leverage, holdback, fund/rev) do **not** block a recommended fund; C62 only blanks on merchant-side knockouts (positions, trend, neg days/NSF). The offer engine resizes the deal to green; the decision cell still says DECLINE for the *requested* offer. Keep.
- Money converts with 5 and 21.655; time with 4.331. Consistent everywhere in the live template.
- Monthly-frequency debt counts toward leverage but not toward position count (Intuit/QuickBooks rule from the log).
- Only 1.40 / 1.45 / 1.49 factors are issued; 1.35 is prepay. Only a perfect 7.0 score reaches 1.40/195 — a 6.5 "Strong" gets 1.45/150. That is how the band table is built; treated as intentional.

**Workbook accidents / stale items**
1. `E35 = E34 + C8` — orphan, referenced by nothing (log also notes it). Removed in the new build.
2. `F34 = E34*21.655` hard-codes the constant the log says was moved to `$K$9`.
3. `C34` counts positions as `(D28:D33<>"")` — a funded-only row (no payment) is **not** counted as a position, contrary to the log's shipped formula `((C<>"")+(D<>""))>0`. New build uses the log's version.
4. `J6`/`K6`/`L6` are the **old** MoneyBadger formulas (label fixed to "Weekly remittance", K6 = D8 always, L6 = L7×C24÷4). Only `L7`, `L8`, `L9` from "Fix 1" shipped. On a daily deal the block compares a weekly remittance against 5.25 × daily and reports a bogus diff. New build uses the log's four formulas.
5. `E50` minimum fund drifted across tabs ($10k in tabs 13–21, $15k in tab 52, $20k live).
6. Threshold drift: tabs 10 and 13–30 use Daily/Balance 0.45/0.75/1.00 and Fund/rev 0.75/1.25/1.50; tab 14 uses NSF knockout 15. Tabs 1–12 have no band table (term/factor typed as 130/1.45). Tabs ≤ 54 carry the pre-fix MoneyBadger and daily-equivalent formulas. All of these are closed deals; left alone.
7. `Neg days + NSFs` is a **total across however many months were entered**, so a 12-month file is judged 3× harder than a 4-month file. Kept as-is (it is your rule) but noted; a per-month variant is offered as an optional control.
8. Term band lookup uses `MATCH(score, {0,5,6,7}, 1)`; a score of 4.5 lands on the 100-payment band, and blank score → "". Fine.
9. The `test` and `Copy of Template n` tabs are a different, older layout and should not be copied.

**Session log vs workbook (reconciliation)**

| Log says | Workbook | Verdict |
|---|---|---|
| 68 sheets | 78 | Log stale |
| Fix 1 shipped J6/K6/L6/L7 | Only L7 + constants | Partially shipped |
| Fix 2 position count counts funded-only rows | Payment-only | Not shipped |
| F34 uses $K$9 | Hard-coded 21.655 | Not shipped |
| E35 orphan noted | Still present | Not removed |
| "Zero #DIV/0! on blank" | Confirmed (LibreOffice recalc of live template = 0 errors) | OK |
| MB verification table (8 tabs) | Tabs 46–54 still carry the old L7 formula, so the table in the log was produced outside those tabs | Informational |
| 1.35 not a tier; 1.40/1.45/1.49 only | Band table matches | OK |

---

## 2. How the House Book workbook works today

**Sheets:** `Verdict` (headline tiles + narrative), `Controls` (date window B5/B6, conversions, category selection, status mapping A25:D35, include lists), `Visual Metrics` (charts), `Data` (1,067 deals + a column dictionary below row 1071), `Calc` (one master block per dimension; columns A–W as in the log; book-level figures in BS1:BT6), `Ref` (SIC-2 → industry, Federal Reserve holidays).

**Raw columns used (Data A–AA):** Contract ID, DBA, legal name, Status, Start date (funding), Last active (last cleared payment), Position, SIC/SIC2/MCC, State, ISO, New or renewal, Factor, Advance, Commission, Payback, Collected, Balance, % paid, Exp dur (raw), Days since pmt, NSF count, Avg rev/mo, Holdback %, Biz start.

**Derived (AB–BB):** duration unit (raw < 12 = months = daily deal; ≥ 12 = weeks = weekly deal), pay frequency, expected term in payments and business days, days on book (calendar and NETWORKDAYS), Default / Resolved / Closed flags from the Controls status table, in-scope flag, net cash, principal lost = MAX(advance − collected, 0) on defaulted deals, advance/revenue, has-revenue, never-paid, size/factor/term/adv-rev buckets, industry, quarter/month/year.

**Status mapping in force**

| Status | Deals | Default | Resolved | Closed |
|---|---|---|---|---|
| Closed | 378 | No | Yes | Yes |
| Open | 300 | No | No | No |
| Legal | 248 | Yes | No | No |
| In House Collection | 73 | Yes | No | No |
| Lowered Payments | 22 | **No** | No | No |
| Written-Off | 15 | Yes | Yes | No |
| Bankruptcy | 12 | Yes | Yes | No |
| Settlement Agreement | 7 | Yes | Yes | No |
| Closed/NSF | 6 | No | Yes | Yes |
| Closed/NSF Unpaid | 5 | No | Yes | Yes |
| At-Risk | 1 | Yes | No | No |

**Date window actually in the uploaded workbook: 11/01/2024 – 06/30/2026 (994 deals, $30,540,626).** The session log says the workbook runs to 08/25/2026; the uploaded file is set to the deck window. The 1,067-deal / $34,481,626 "full book" figures in the log are reproduced from raw data in the new build's reconciliation block.

**Recomputed full-file figures (all 1,067 deals, 11/20/2024 – 08/07/2026 funding dates):** funded $34,481,626; average advance $32,316; median $20,000; 28% open; 33% in a default status.

**Log / deck / workbook conflicts**
- Deck is `Aspire Book Analysis_V2.pptx` (13 slides). The log describes `Aspire_House_Book_Dimensions_v11.pptx` (14 slides). Figures on the uploaded deck match the uploaded workbook's window, so they are consistent with each other, but the deck version in hand is not the one the log narrates.
- Slide 10 note: "152 deals have no revenue data (79 closed, 69 defaulted, 4 open)". Log §9 lists 152 with a different status split; raw file has **188** with no revenue across all 1,067 deals (152 inside the deck window). The slide note is the one the log itself flags as still wrong.
- Log §2 says Written-Off = 12 deals; mapping table shows 15 (window difference).
- `Data` dictionary rows 1101–1105 document columns BC/BD/BJ/BK/BM (scenario term/factor buckets, conservative/optimistic collections, "meets suggested box"). Those columns no longer exist. Stale documentation.
- Verdict K8 "Return on funded capital" = −13.7% (all in-window deals incl. open) while the deck title tile says −4.2% "excl. open". Both are correct for what they measure; the labels differ.
- Log §11 recommends "no 3rd position to new merchants". See §4 below — the raw effect mostly dissolves after controls.

**Data-quality issues (raw file)**

| Issue | Count | Handling in the new build |
|---|---|---|
| No revenue figure | 154 (14.4%) | Excluded from advance/revenue stats; counted and shown |
| No ISO / "No state" / no SIC-2 | 0 / 37 / 6 | Never match on a blank; blanks fall to the next tier. One deal has no position. |
| Position field: 49% of deals are position 3 (61% of renewals) | 519 | Used, but flagged as a CRM-field reliability question |
| Never paid (collected ≤ 0) | 15 | Kept; counted as defaults per status |
| Status Open but ≥ 99.9% paid | 8 | Status stale; flagged in Checks |
| Open deals more than 1.5× contracted term old | 20 | Flagged in Checks |
| Weekly durations above 39 weeks (9-month box) | 23 (incl. 56, 90, 100 weeks) | Kept in cohorts; flagged as implausible/out-of-box |
| Lowered Payments treated as performing | 22 (avg 52% paid, past term) | Mapping kept as the house book defines it, but flagged; one switch on Controls changes it |
| Commission zero | 108 | Informational |
| **NSF count is post-funding** (Open avg 22.9 NSFs, Closed avg 2.8) | all | **Must not be used as a matching input** — it is an outcome, not underwriting-time data |
| Written-Off deals collected 64% of payback on average (up to 99%) | 15 | Status marks "stopped pursuing", not "lost" |

---

## 3. Seasoning and the four bases

Seasoned = business-day age at the as-of date (08/25/2026) ≥ 1.25 × contracted term in business days (the v13 rule from the log, now a Controls cell). The table below comes from the analysis script (numpy business-day count, end-exclusive: 629 deals). The workbook uses Excel `NETWORKDAYS`, which counts both end days, and lands on 637 seasoned deals, $12,563,126 funded, 21.1% principal lost; both were verified independently and the workbook figure is the one the model uses.

| Basis | Deals | Funded | Open | Default rate | Principal lost | Collected on defaults | Return on funded |
|---|---|---|---|---|---|---|---|
| Full book | 1,067 | $34.48M | 28.1% | 33.4% | 15.3% | 40.9% | −20.4% |
| Decided (ex-open) | 767 | $19.88M | 0% | 46.4% | 26.5% | 40.9% | −4.1% |
| Resolved | 423 | $11.29M | 0% | 8.0% | 2.6% | 67.8% | +37.6% |
| **Seasoned** | **629** | **$12.29M** | 6.8% | 42.8% | 21.4% | 47.5% | +1.7% |

Seasoned is the basis for every offer adjustment. "Decided" over-counts losses (Legal and In-House are mid-collection); "Resolved" is survivorship. 43 seasoned deals are still Open (median 72% paid) — they stay in the cohort and are shown as open.

---

## 4. Which historical signals survive basic controls

Logistic regression on the 629 seasoned deals (default = 1) with position, renewal, construction, health, log advance, factor ≥ 1.49, weekly, adv/rev < 0.15, missing revenue, and the five largest ISOs. Odds ratios with |z| ≥ 2 are treated as surviving.

| Signal | Seasoned raw | After controls | Verdict |
|---|---|---|---|
| Renewal vs new | lost 12.3% vs 23.6% | OR 0.43, z −3.3 | **Robust. Safe for offer logic.** Confirms the log's "renewals first" (the sign flip vs the full book is age). |
| Construction (SIC 15/16/17) | GBC 32.0%, Special Trade 24.5% | OR 1.84, z 2.7 | **Robust.** Survives new/renewal split (new 29%, renewal 35%). |
| Health Services | 30.4% (n 34) | OR 2.17, z 2.0 | Moderate; thin sample. |
| Westwood ISO | 4.0% (n 36) | OR 0.31, z −2.6 | **Robust positive**, holds in every position. |
| PMF ISO | 42.1% (n 25) | OR 1.99, z 1.6 | Suggestive, not significant; half its deals are 3 months old. Shrinkage handles it. |
| United Secured Capital 8 | 26.5% (n 106) | OR 1.09, z 0.4 | **Composition, not the ISO**: its book is 45% 3rd-position, 87% new, 1.49+. Raw underperformance dissolves. |
| Position 3 (new deals) | 29.4% vs 20.6% | OR 1.25, z 1.0 | Weak after renewal/industry/factor are held. Non-construction new 3rd: 29.3% vs 19.0%, so a raw gap remains. **Use position inside the matching hierarchy, not as a standalone penalty.** The deck's "stop 3rd position" is mostly composition. |
| Factor ≥ 1.49 | 24–30% vs 2–23% | OR 2.67, z 3.9 | Real but it is **selection**: price was set high on deals judged risky. Informational; confirms pricing is not a cure. Never a lever the history can "raise" on its own. |
| Term, deal size | no gradient on seasoned | n.s. | Not usable (log agrees). 150+ business-day seasoned cohort is 20 deals. |
| Weekly vs daily | 17.1% vs 23.9% | OR 0.75, z −1.4 | Not significant once size is held. Informational. |
| Advance/revenue < 0.15x | 30.0% (n 80) | OR 1.25, z 0.9 | Inverted gradient persists on seasoned data; likely the revenue field. Informational until revenue data is cleaned. |
| State | NC 51% (n 22) | none significant | NC is 32% one ISO. **Informational only, always with a warning.** |
| Vintage | 2026-Q1 28% vs 2025-Q4 16% | n.s. after seasoning | Used only as the seasoning gate, never as a match key. |
| Holdback % | no monotonic pattern | — | Not used. |
| NSF count | dominant predictor | z 11.5 | **Leakage** — post-funding. Excluded. |

---

## 5. Questions and assumptions that shape the design

1. **Position field** — assumed to be the position the house deal took at funding. If it is a CRM default, position tiers are wrong; the Checks sheet reports the position distribution so this is visible.
2. **Lowered Payments** — kept as "not default" to reconcile with the house book. One cell on Controls (status mapping) flips it.
3. **As-of date** for seasoning = 08/25/2026 (Controls). Move it when the data refreshes.
4. **Credit score bands** (your rule): < 630 red, 630–649 yellow, ≥ 650 green. "580 and below" and "580–620" are both inside the red band, so the single red line is 630. Knockout line is a Controls cell left at 0 (off).
5. **Match on ISO name and industry name exactly** as spelled in the house book; the Deal sheet uses dropdowns fed from the data so this cannot drift.
6. **The 9-month term cap, 1.40/1.45/1.49 factor set, 5% fee, and all seven thresholds are unchanged** and now live on Controls.
7. **Neutral historical evidence sizes the offer at 90% of the cash-flow maximum** (Controls "Base utilisation"). The seasoned book lost 21% of principal while funding at those green lines; sitting at the maximum is reserved for strong, credible evidence. Set to 100% to reproduce the old behaviour.
