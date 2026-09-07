# Underwriting Tool + House Book — combined MCA underwriting model

| Path | What it is |
|---|---|
| `Aspire_Combined_Underwriting_Model_OUTLINE.xlsx` | **Brainstorm-phase outline.** Deal sheet = your live template (v2 upload, untouched above row 64); every other tab shows the planned logic in words, with Controls values real and 8 sample house-book deals. Built by `build/build_outline.py`. |
| `Aspire_Combined_Underwriting_Model.xlsx` | The fully-formulated version (reference for the build phase). Blank combined model: `Deal` (the original template + profile inputs + Section F), `Decision Summary`, `Offer Engine`, `Historical Score`, `Comparable Deals`, `Historical Data` (1,067 deals), `Controls`, `Checks`, `Ref`, `README`. |
| `docs/01_Stage1_Audit_Findings.md` | Audit of both workbooks, the deck and both session logs; what survives controls. |
| `docs/02_Stage2_Combined_Architecture.md` | Sheet structure, input dictionary, hard rules, hierarchy, credibility, modifier table, worked example. |
| `docs/03_Stage3_Implementation_Notes.md` | Formula map, reconciliation, manual verification, ten-scenario test record, limitations. |
| `originals/` | Untouched uploads with MD5 checksums. |
| `build/build_model.py` | Rebuilds the model from the originals. `build/run_scenarios.py` runs the test scenarios through LibreOffice. |

## Google Sheets

1. Google Drive → **New → File upload** → `Aspire_Combined_Underwriting_Model.xlsx`.
2. Open it → **File → Save as Google Sheets**. (Or in a blank Sheet: File → Import → Upload → Replace spreadsheet.)
3. Give the first recalculation a few seconds. Check the `Checks` sheet reads 0 FAIL, then use the `Deal` sheet.

All functions in the model are native to Sheets (SUMPRODUCT, NETWORKDAYS, INDEX/MATCH, COUNTIF, CHOOSE, FLOOR, TEXT, IFERROR, VLOOKUP). Dropdowns, colour rules, merged cells and number formats survive the import.
