# v2 changes (built by `build/build_model_v2.py`; `demo` argument = 8 sample deals + populated Deal sheet)

**History can now loosen the cash-flow sizing.** The modifier table on Controls (section 5) has a new column, *Sizing flex*: 0 = size at the green lines (the tool exactly as built), 1 = size at the red lines. Recommended values: Strong positive 90%, Mild positive 45%, everything else 0. Section C on the Deal sheet shows the sizing line in use in column H (purple); section E sizes the max offer at those lines, so a supportive book produces a larger offer whose metrics may show YELLOW. Flex is capped below 100%, so RED is impossible. Negative evidence keeps the green lines and applies the fund multiplier, band step, factor step and weekly rule as before.

Full-book test (renewal restaurant via Westwood): green-line max $85,500 → sized max $105,500 at 45% flex; holdback 21.1% and fund/revenue 0.73 turn YELLOW, nothing red; every safety proof OK.

**Section C extended (C2, rows 49–57).** Same Green / Red / Status layout: primary cohort, ISO, industry, new/renewal, position, state (informational), and the seasoned-deal count behind the primary cohort. Green/red lines are the modifier table's 0.85 / 1.15 index bounds expressed as lost rates. These rows are not counted in the cash-flow score or red flags.

**Removed:** confidence level and its cap, stipulations (Decision Summary and Controls section 8), the written conclusion paragraph, the proposed-position input and override (matching reads C42). State is optional; a blank state never blocks an offer.

**Decision Summary** is now a bulleted read-out: decision, offer, sizing line used, the four offer-side metrics with status, modifier, cohort numbers, largest impact, limiting metric, triggers, one evidence line per dimension, warnings, reds and greens at the requested offer.

**Demo file minimums** are lowered (3 seasoned deals / $30,000 / rate shown at 2 / indicator green at 5) so eight sample deals produce cohorts. The production values (10 / $150,000 / 5 / 30) are noted beside each cell on Controls.
