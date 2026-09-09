"""v7: blank-deal fixes. Whole-book row can no longer present itself as cohort evidence;
Section F stays silent until a deal profile is entered. Applied in place on a v6 file."""
import openpyxl, sys
from openpyxl.workbook.properties import CalcProperties
f_in, f_out = sys.argv[1], sys.argv[2]
wb=openpyxl.load_workbook(f_in)
H=wb['Historical Score']; D=wb['Deal']; HD=wb['Historical Data']; NR=HD.max_row
# --- Historical Score: no cohort until new/renewal entered; book row never shows as evidence
H['D29']='=IF($C$5="","",IFERROR(MATCH("Yes",$E$14:$E$21,0),8))'
H['C29']='=IF($D$29="","- no deal entered -",INDEX($B$14:$B$21,$D$29))'
for cell,rng in (('C30','$C$14:$C$21'),('C31','$F$14:$F$21'),('C32','$G$14:$G$21'),('C33','$H$14:$H$21')):
    H[cell]=f'=IF(OR($D$29="",$D$29=8),"",INDEX({rng},$D$29))'
H['C34']='=IF(OR($D$29="",$D$29=8,$C$33=""),"Neutral",IF($C$33<=Controls!$B$29,"Supportive",IF($C$33>Controls!$B$30,"Negative","Neutral")))'
# --- Historical Data: CHOOSE must not receive a blank index
for r in range(2,NR+1):
    HD[f'BA{r}']=(f'=IF(OR(AD{r}=0,\'Historical Score\'!$D$29=""),0,'
                  f'CHOOSE(\'Historical Score\'!$D$29,AU{r},AV{r},AW{r},AX{r},AY{r},AZ{r},AR{r},1))')
# --- Deal C2 block
D['F51']=('=IF(OR(\'Historical Score\'!$D$29="",\'Historical Score\'!$D$29=8),"NO EVIDENCE",'
          'IF(\'Historical Score\'!$C$33<=Controls!$B$29,"GREEN",'
          'IF(\'Historical Score\'!$C$33>Controls!$B$30,"RED","YELLOW")))')
D['G51']="='Historical Score'!$C$30"
D['C56']='=IF(OR(\'Historical Score\'!$D$29="",\'Historical Score\'!$D$29=8),"",\'Historical Score\'!$C$30)'
D['F56']='=IF(C56="","",IF(C56>=D56,"GREEN",IF(C56>=E56,"YELLOW","RED")))'
# --- Deal Section F: silent until a deal profile exists
D['C76']=('=IF(OR(C24=0,$K$14=""),"",IF(OR(C73="-",C73="",C78="",C78=0,C78<E61,'
          'COUNTIF(F42:F45,"RED")+COUNTIF(F42:F45,"KNOCKOUT")+COUNTIF($J$76:$J$79,"RED")>=C61),"DECLINE",'
          'IF(OR(COUNTIF($J$76:$J$79,"YELLOW")>0,COUNTIF(F42:F45,"YELLOW")>0,C84="Negative"),"CONDITIONAL","APPROVE")))')
D['C77']='=IF($K$14="","",F73)'
D['C78']=('=IF($K$14="","",IF(OR(C73="",C73="-"),C73,'
          'IF(\'Historical Score\'!$C$34="Negative",FLOOR(C73*(1-Controls!$B$32),500),C73)))')
D['C84']='=IF($K$14="","",\'Historical Score\'!$C$34)'
wb.calculation=CalcProperties(fullCalcOnLoad=True)
wb.save(f_out); print('v7 fixes applied ->',f_out)
