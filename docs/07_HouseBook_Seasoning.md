# House Book Analytics v4 - seasoning

## Why
Every metric in the workbook filtered on `Data!AK` ("In scope"), which tested the
funding-date window only. Deals funded last month sat in the same loss rates as
deals that have run their full term, pulling PLR down and inflating ROF.

## Where it was implemented
`AK` is referenced 9,116 times in Calc and 27 in Verdict, so seasoning was folded
INTO that flag. No formula in Calc, Verdict or Visual Metrics had to change.

## Controls - new block (rows 8-10)
| Cell | Meaning |
|---|---|
| B9  | As-of date - when the data was pulled |
| B10 | Seasoning multiple (age >= multiple x expected term) |
| E9  | Require seasoning?  Y = seasoned only, anything else = all deals |
| E10 | Live count of seasoned deals in the window |

## Data - new columns
| Col | Formula |
|---|---|
| BC | `NETWORKDAYS(E, Controls!$B$9, Ref!$F$4:$F$47)` - age at the as-of date |
| BD | `IF(OR(AH=1,AJ=1),1,0)` - reached a terminal outcome (defaulted or closed) |
| BE | `IF(OR(BD=1, AND(AE<>"", BC>=Controls!$B$10*AE)),1,0)` - Seasoned |

A deal is seasoned if it has had time to fail OR has already reached a final
outcome. The terminal-status arm matters: a deal that closed early and paid in
full is evidence even if it is younger than the age bar.

## AK rewritten
`=IF(AND(E>=Controls!$B$5, E<=Controls!$B$6, OR(Controls!$E$9<>"Y", BE=1)),1,0)`

## Separate fix - stale row bound
Data holds 1,104 deals (rows 2-1105) but every SUMIFS/COUNTIFS stopped at row
1068, silently excluding 34 deals from every metric in the workbook.
5,682 formulas had their Data ranges extended to 1105.
