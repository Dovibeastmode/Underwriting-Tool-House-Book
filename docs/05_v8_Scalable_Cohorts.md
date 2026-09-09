# v8 — scalable cohort engine

Built on the v7 file you uploaded, so your edits (Section E moved to C87:D90,
L14, "NO DEALS", widened COUNTIF ranges) are preserved. C36's XLOOKUP is
untouched.

Verified against v7 across 10 deal profiles (7 tier hits, whole-book fallback,
blank deal, unmatchable). Primary cohort, seasoned count, PLR, ROF, index,
outcome, flex, Deal F51/C56, the full 8x5 tier table and the BA column sum are
identical in every case.

## What changed, by cell

### Historical Score — NEW
| Cell | Content |
|---|---|
| K13:N13 | headers `req NR` / `req ISO` / `req Ind` / `req Pos` |
| K14:N21 | the tier definition flags (1 = must match, 0 = ignore) |
| K29:N29 | `=IF($D$29="","",INDEX(K$14:K$21,$D$29))` — selected tier's flags |
| D30 | `=COUNTA($B$14:$B$21)` — fallback row # |
| B37/C37 | `HAS EVIDENCE?` / `=IF(OR($D$29="",$D$29=$D$30),"NO","YES")` |

### Historical Score — REPLACED
- C14:D21, F14:G21 — one generic mask formula, identical in all 8 rows
- D29 — IFERROR dropped; row 21 always qualifies so MATCH cannot fail
- C30:C34 — guards now read `$C$37`, not the hard-coded 8

### Historical Data
- AU:AZ cleared (combos are data now, not columns)
- BA2:BA1068 — mask against K29:N29; no CHOOSE

### Deal
- F51, C56 — guards read `'Historical Score'!$C$37`

## Adding a dimension later
1. New match column on Historical Data (same shape as AQ-AT)
2. New flag column at O on Historical Score + one factor in the mask formula, filled down
3. One more factor in BA, filled down
4. New INDEX cell at O29

No magic numbers, no downstream edits.
