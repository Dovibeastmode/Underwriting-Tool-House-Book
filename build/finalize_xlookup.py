"""Post-recalc step: swap Historical Score C36 to XLOOKUP (Google Sheets / Excel 365).
Run AFTER recalc.py so LibreOffice never rewrites the cell."""
import openpyxl, sys, shutil
f=sys.argv[1]
wb=openpyxl.load_workbook(f)
wb['Historical Score']['C36']='=IF(MAX($I$23:$I$26)=0,"none rated",_xlfn._xlws.XLOOKUP(MAX($I$23:$I$26),$I$23:$I$26,$B$23:$B$26))'
from openpyxl.workbook.properties import CalcProperties
wb.calculation=CalcProperties(fullCalcOnLoad=True)
wb.save(f); print('XLOOKUP written into',f)
