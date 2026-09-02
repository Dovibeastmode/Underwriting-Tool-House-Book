# Stage 3 — Implementation Notes and Test Record

Deliverable: `Aspire_Combined_Underwriting_Model.xlsx` (blank, ready for a deal). Built by `build/build_model.py` from the two untouched originals in `originals/`. The worked example (scenario A) is reproducible with `build/run_scenarios.py`.

## What was preserved

- The `Deal` sheet **is** the original `New Template` sheet: same rows, columns, fonts, fills, borders, dropdowns and colour rules. Cells that changed: threshold columns are now green links to `Controls`; the four MoneyBadger cells and the position-count / monthly-total cells carry the session-log formulas that had not shipped; the orphan `E35` is gone; a credit-score metric row was inserted at row 45 (everything below moved down one row); the old pricing-band block became the Deal Profile inputs; Section F was added at the bottom.
- All seven threshold sets, the 2-red-flag limit, the $20k minimum, the 9-month cap, the 100/130/150/195 bands, 1.40/1.45/1.49 factors and the 5% fee are the same numbers, now on `Controls`.
- Every raw house-book field (A–X) is copied as values into `Historical Data`; the SIC map and holiday calendar came from the original `Ref` sheet; the status mapping is the original table plus an "Open?" column.

## Sheet-by-sheet formula map

| Sheet | Key cells |
|---|---|
| Deal | C24/D24 averages; C34 positions; E34/F34 existing daily/monthly; C38:C45 metric values; F38:F45 statuses; C50/E50 red flags/knockouts; C52 Layer-1 decision; C55/D55 score and metrics scored; D57 score scaled to 7; C57:C60 max daily payments; C62/D62 term/factor; **C63 cash-flow max fund**; K12:K18 profile; C67:C79 Section F read-outs |
| Historical Data | Y industry; Z frequency; AA term (bd); AB age (bd, NETWORKDAYS to as-of); AC seasoned; AD in window; AE–AH default/resolved/closed/open flags via Controls; AK principal lost; AL adv/rev (0 when missing) + AM has-revenue; AN days on book; AP eligible; AQ–AU match flags; AV–BB tier flags; BC best tier; BD in primary cohort; BE sequence for the comparable list |
| Historical Score | C5:C9 deal keys; rows 14–20 tiers 1–7, rows 22–26 single dimensions (D–AA: counts, funded, open, default, lost, collected on defaults, return, averages, seasoned-enough, Z, blended, index, meets-min, impact); C29 primary tier; C33 cohort lost; C37 Z; C38 blended; C39 index; C40 level; C41 confidence; C42 capped level; C44/C45 largest impact; C46:C49 evidence sentences; C50 conflict; C51 warnings |
| Offer Engine | C5:C9 base; D6 adjusted term; D7 adjusted factor; **D8 adjusted fund**; D9 daily; D16 frequency; D17/D18 term and payment in final units; E26:E33 statuses at adjusted offer; D36:D39 safety proofs; D40 red flags at adjusted offer |
| Decision Summary | C5 final decision; C6:C29 the 20 output items; C26 stipulations; C28 manual-review triggers; C30 written conclusion; C31 risk-stip count |
| Checks | 42 checks: missing inputs, formula errors, safety proofs, cohort quality, data quality, reconciliation |

## Reconciliation to the original house-book workbook

| Figure | Original (full file) | Rebuilt | Check |
|---|---|---|---|
| Deals | 1,067 | 1,067 | OK |
| Funded | $34,481,626 | $34,481,626 | OK |
| Collected | $27,438,510 | $27,438,510 | OK |
| Principal lost | $5,270,163 | $5,270,163 | OK |
| Principal lost rate | 15.3% | 15.3% | OK |
| Payback = advance × factor | all | 0 mismatches | OK |

## Manual verification of historical deals (against the raw `Data` sheet, outside Excel)

| Contract | Raw status / advance / collected | Model row | Seasoned (age vs 1.25 × term) | Match |
|---|---|---|---|---|
| 5674154422 Bushco Contracting | Legal, $14,000 / $2,318.06 | net −$11,681.94, lost $11,681.94, term 90 bd, age 349 bd, days on book 25 | yes | exact |
| 5674160797 Macs Custom Installs | At-Risk, $10,000 / $8,473.92 | net −$1,526.08, lost $1,526.08, term 56.8 bd, age 291 | yes | exact |
| 5674162609 Ribeiro General Construction | Closed, $10,000 / $13,500 | net +$3,500, lost 0, term 42.6 bd, age 273, days on book 72 | yes | exact |
| Tier 1 cohort (new + GBC + position 2) | — | 27 matches / 23 seasoned / $414,000 / 1 open / 11 defaulted / 34.0% lost | — | exact |
| Seasoned book | — | 637 deals / $12,563,126 / 21.1% lost | — | exact |

## Scenario test record (`build/run_scenarios.py`, LibreOffice recalculation, 45,098 formulas)

| Scenario | Layer 1 | Cash-flow max | Primary cohort | Modifier / confidence | Final | Errors / check fails |
|---|---|---|---|---|---|---|
| A new, construction, USC8, FICO 640, 2nd position, weekly | CONDITIONAL | $68,500 | T1, 23 seasoned, lost 34.0% | Mild negative / Medium | CONDITIONAL $54,500 @1.49, 20 wks | 0 / 0 |
| B renewal restaurant, Westwood, FICO 705, all green | APPROVE | $85,500 | T3 (T4 Westwood-renewal had only 9 seasoned → skipped), lost 2.5% | Mild positive / Medium | APPROVE $81,000 @1.40 | 0 / 0 |
| C 11 negative days + 6 NSFs, daily | DECLINE | knockout | T2 | Mild positive (ignored) | DECLINE, "history cannot override a knockout" | 0 / 0 |
| D zero revenue, no bank data | blank | blank | T3 | Neutral | blank, "Enter bank data" | 0 / 0 |
| E ISO not in book, no industry, monthly-only debt | APPROVE | $47,500 | T5 | Neutral / High | MANUAL REVIEW (unknown ISO) | 0 / 0 |
| F renewal, ISO with 8 seasoned deals, 5th position, daily | CONDITIONAL | $30,000 | T6 (tiers 3–5 too small) | Mild positive / High | MANUAL REVIEW (positions ≥ 4) | 0 / 0 |
| G daily health-care deal, PMF, 6 months data, Kapitus monthly + Mulligan weekly | CONDITIONAL | $49,000 | T3 | Neutral / Medium | APPROVE $44,000 @1.45, 150 pmts | 0 / 0 |
| H five existing daily positions | DECLINE | none | T3 | Neutral | DECLINE (positions knockout) | 0 / 0 |
| I FICO 590 + revenue trend −16% | DECLINE | $48,500 | T2 | Neutral | DECLINE (two reds at adjusted offer) | 0 / 0 |
| J strong file, position override 3, Premium Capital, Legal Services (4 seasoned → too few) | APPROVE | $238,000 | T4 | Mild positive / Medium | APPROVE $226,000 @1.40 | 0 / 0 |

In every scenario the safety proofs held: adjusted fund ≤ cash-flow max fund, adjusted daily ≤ cash-flow max daily, no offer-side metric red at the adjusted offer, merchant knockouts untouched.

## Known limitations

1. Section F on a copied `Deal` tab still points to the shared historical sheets (they compute for the sheet named `Deal`). Paste values before archiving.
2. Excel opens the blank model and recalculates on load (`fullCalcOnLoad`); the first open takes a few seconds because of the 45k formulas.
3. Only the first 400 comparable deals are listed; the cohort statistics use all of them.
4. The `Position` field is used as delivered by the CRM. If it is not the position at funding, tiers 1, 2 and 5 are mis-keyed.
5. The historical modifier is a hierarchical, credibility-weighted lost-rate comparison, not a multivariable model. The Stage 1 regression is what justifies the tier order; the sheet itself does not run a regression.
