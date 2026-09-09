# v9 - scalable cohort engine applied to the combined main workbook

Source: Underwriting_Tool_Combined_with_House_book_analytics.xlsx (pre-v8 lineage:
AU:AZ combo columns, CHOOSE in BA, no flag table, no guard cell).
Every changed or new cell is filled light blue in the delivered file.

## Reconciliation
6 deal profiles recalculated on the original and the rebuild.
5 identical across the result block, the 8x5 tier table, Deal F51/C56/C78 and the
BA column sum. The 6th (whole-book fallback) differs deliberately: C30:C33 now
blank instead of reporting the whole book as if it were a matched cohort. F51,
C56 and C78 are unchanged even there.

## Changed cells
- Historical Score: 87 (flag table K13:N21 + O13 note, tier formulas C/D/F/G 14:21,
  D29 E29, K29:N29 + O29 note, D30 E30, B37 C37 D37, C29:C34)
- Historical Data: AU:AZ cleared rows 1-1068; BA2:BA1068 rewritten
- Deal: F51, C56

## Not changed
Controls, Ref, README, Comparable Deals, the cash-flow engine, C35 (the symmetric
sizing flex already applied), C36 (XLOOKUP left as-is).
