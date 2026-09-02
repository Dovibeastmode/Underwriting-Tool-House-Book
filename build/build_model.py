"""Build Aspire_Combined_Underwriting_Model.xlsx from the two original workbooks.
Layer 1 (Deal) = the original New Template, restyled in place. Layers 2-3 generated."""
import openpyxl, datetime as dt, copy, sys
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.utils import get_column_letter as L
from openpyxl.comments import Comment

ROOT='/home/user/Underwriting-Tool-House-Book'
UW=f'{ROOT}/originals/Updated Underwriting Tool (1).xlsx'
HB=f'{ROOT}/originals/House Book Analytics (2).xlsx'
OUT=sys.argv[1] if len(sys.argv)>1 else f'{ROOT}/Aspire_Combined_Underwriting_Model.xlsx'

# ---------- styles ----------
DARK=PatternFill('solid',fgColor='1F2937'); MID=PatternFill('solid',fgColor='4B5563'); GREY=PatternFill('solid',fgColor='E5E7EB'); INPUT=PatternFill('solid',fgColor='FFF9C4')
WHITE=Font(name='Calibri',size=11,bold=True,color='FFFFFF'); SUB=Font(name='Calibri',size=9,bold=True,color='FFFFFF')
B=Font(name='Calibri',size=11,bold=True); N=Font(name='Calibri',size=11); NOTE=Font(name='Calibri',size=9,color='374151'); LINK=Font(name='Calibri',size=11,color='008000'); INP=Font(name='Calibri',size=11,color='0000FF')
thin=Side(style='thin',color='9CA3AF'); med=Side(style='medium',color='374151')
BOX=Border(left=thin,right=thin,top=thin,bottom=thin); IBOX=Border(left=med,right=med,top=med,bottom=med)
CEN=Alignment(horizontal='center',vertical='center'); LEFT=Alignment(horizontal='left',vertical='center'); WRAP=Alignment(horizontal='left',vertical='top',wrap_text=True)
G_FILL=PatternFill('solid',fgColor='00B050'); Y_FILL=PatternFill('solid',fgColor='FFFF00'); R_FILL=PatternFill('solid',fgColor='FF0000'); GREY2=PatternFill('solid',fgColor='9CA3AF')
def rag_rules(ws,rng):
    ws.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"GREEN"'],fill=G_FILL,font=Font(color='FFFFFF',bold=True)))
    ws.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"YELLOW"'],fill=Y_FILL,font=Font(color='000000',bold=True)))
    ws.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"RED"'],fill=R_FILL,font=Font(color='FFFFFF',bold=True)))
    ws.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"KNOCKOUT"'],fill=R_FILL,font=Font(color='FFFFFF',bold=True)))
def dec_rules(ws,rng):
    ws.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"APPROVE"'],fill=G_FILL,font=Font(color='FFFFFF',bold=True)))
    ws.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"CONDITIONAL"'],fill=Y_FILL,font=Font(color='000000',bold=True)))
    ws.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"DECLINE"'],fill=R_FILL,font=Font(color='FFFFFF',bold=True)))
    ws.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"MANUAL REVIEW"'],fill=PatternFill('solid',fgColor='F97316'),font=Font(color='FFFFFF',bold=True)))
def chk_rules(ws,rng):
    ws.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"OK"'],fill=G_FILL,font=Font(color='FFFFFF',bold=True)))
    ws.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"WARN"'],fill=Y_FILL,font=Font(color='000000',bold=True)))
    ws.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"FAIL"'],fill=R_FILL,font=Font(color='FFFFFF',bold=True)))
def hdr(ws,cell,text,span=None):
    ws[cell]=text; ws[cell].font=WHITE; ws[cell].fill=DARK; ws[cell].alignment=LEFT
    if span: ws.merge_cells(span)
def sub(ws,row,cols,labels,start=1):
    for i,t in enumerate(labels):
        c=ws.cell(row,start+i,t); c.font=SUB; c.fill=MID; c.alignment=CEN; c.border=BOX
def put(ws,addr,val,font=N,fmt=None,align=CEN,border=BOX,fill=None):
    c=ws[addr]; c.value=val; c.font=font; c.alignment=align; c.border=border
    if fmt: c.number_format=fmt
    if fill: c.fill=fill
    return c
def label(ws,addr,text,font=N): 
    c=ws[addr]; c.value=text; c.font=font; c.alignment=LEFT; c.border=BOX; return c
def note(ws,addr,text): 
    c=ws[addr]; c.value=text; c.font=NOTE; c.alignment=LEFT; return c
def widths(ws,d):
    for k,v in d.items(): ws.column_dimensions[k].width=v

# ---------- read house book raw data ----------
hv=openpyxl.load_workbook(HB,data_only=True)
hd=hv['Data']
raw=[]
for r in hd.iter_rows(min_row=2,max_row=1068,max_col=27,values_only=True):
    if r[0] is None: continue
    raw.append(list(r))
assert len(raw)==1067, len(raw)
ref=hv['Ref']
sic=[(ref.cell(r,1).value,ref.cell(r,2).value,ref.cell(r,3).value) for r in range(4,68)]
hols=[ref.cell(r,6).value for r in range(4,45) if ref.cell(r,6).value]
isos=sorted({r[12] for r in raw if r[12]},key=str.lower)
states=sorted({r[11] for r in raw if r[11]})
statuses=[('Closed','No','Yes','Yes','No'),('Open','No','No','No','Yes'),('Legal','Yes','No','No','No'),('In House Collection','Yes','No','No','No'),('Lowered Payments','No','No','No','No'),('Written-Off','Yes','Yes','No','No'),('Bankruptcy','Yes','Yes','No','No'),('Settlement Agreement','Yes','Yes','No','No'),('Closed/NSF','No','Yes','Yes','No'),('Closed/NSF Unpaid','No','Yes','Yes','No'),('At-Risk','Yes','No','No','No')]
industries=sorted({s[1] for s in sic})+['No SIC','Unclassified']

# ---------- start from the original underwriting template ----------
wb=openpyxl.load_workbook(UW)
for n in wb.sheetnames:
    if n!='New Template': del wb[n]
ws=wb['New Template']; ws.title='Deal'
NR=1068  # last data row in Historical Data
HDR="'Historical Data'!"
def hr(col): return f"{HDR}${col}$2:${col}${NR}"

# shift rows 45..63 down one (styles, heights, merges) to make room for the credit-score metric
merges=[m for m in list(ws.merged_cells.ranges) if m.min_row>=45]
for m in merges: ws.unmerge_cells(str(m))
for r in range(63,44,-1):
    for c in range(2,14):
        src=ws.cell(r,c); dst=ws.cell(r+1,c); dst._style=copy.copy(src._style); dst.value=None
    if ws.row_dimensions[r].height: ws.row_dimensions[r+1].height=ws.row_dimensions[r].height
for m in merges:
    ws.merge_cells(start_row=m.min_row+1,start_column=m.min_col,end_row=m.max_row+1,end_column=m.max_col)
for c in range(2,14): ws.cell(45,c)._style=copy.copy(ws.cell(44,c)._style)
# unmerge J11:M11 & J18:M18 remnants, then clear old pricing band block J11:M18 and old values rows 45-64
for m in list(ws.merged_cells.ranges):
    if m.min_row in (11,18) and m.min_col>=10: ws.unmerge_cells(str(m))
from openpyxl.cell.cell import MergedCell
for r in range(11,19):
    for c in range(10,14):
        if not isinstance(ws.cell(r,c),MergedCell): ws.cell(r,c).value=None
for r in range(45,65):
    for c in range(2,14):
        if not isinstance(ws.cell(r,c),MergedCell): ws.cell(r,c).value=None
ws.conditional_formatting=ConditionalFormattingList()
ws.data_validations.dataValidation=[dv for dv in ws.data_validations.dataValidation if str(dv.sqref) in ('E28:E33','B12:B23','E6')]

C='Controls!'
# ---- header / decision line
ws['C2']='=C52'
ws['D2']='=IF(C52="","",C50&" red flag(s) vs limit "&C51&"   |   "&COUNTIF(F38:F45,"GREEN")&"G / "&COUNTIF(F38:F45,"YELLOW")&"Y / "&COUNTIF(F38:F45,"RED")&"R   |   score "&TEXT(C55,"0.0")&"/"&D55&" = "&C56&"   |   FINAL (with house book): "&IF(C67="","-",C67))'
# ---- A. offer (unchanged) with fee & MB constants linked to Controls
ws['F6']=f'={C}$B$15'; ws['F6'].font=Font(name='Calibri',size=12,bold=True,color='008000')
ws['J6']='=IF($E$6="Weekly","Weekly remittance","Daily remittance")'
ws['K6']='=IF(OR($E$6="",C8="",C24=0),"",IF($E$6="Weekly",D8,C8))'
ws['L6']='=IF(OR(L7="",C24=0),"",L7*C24/IF($E$6="Weekly",$L$8,$L$9))'
ws['L7']='=IF(OR($E$6="",C8="",C24=0),"",ROUND(IF($E$6="Weekly",D8*$L$8,C8*$L$9)/C24*100,2)/100)'
ws['K8']=f'={C}$B$28'; ws['K9']=f'={C}$B$29'; ws['L8']=f'={C}$B$30'; ws['L9']=f'={C}$B$31'
for a in ['K8','K9','L8','L9']: ws[a].font=LINK
ws['E8']='=IF(C8="","",C8*$K$9)'; ws['F8']='=IF(OR(D6="",D6=0),"",IF(E6="Weekly",D6,D6/5)/($K$9/5))'
# ---- Deal profile block J11:M18
hdr(ws,'J11','DEAL PROFILE  (house-book match inputs)','J11:M11')
prof=[('ISO','K12','dropdown = house-book ISO names'),('Industry (SIC major group)','K13','dropdown = Ref list'),('State','K14','2-letter'),('New deal or renewal','K15','New Deal / Renewal'),('Credit score (FICO)','K16','number'),('Proposed position (auto)','K17','auto: positions incl. new'),('Position override (optional)','K18','type to override')]
for i,(t,a,n_) in enumerate(prof):
    r=12+i; c=ws.cell(r,10,t); c.font=Font(name='Calibri',size=9,bold=True,color='FFFFFF'); c.fill=MID; c.alignment=LEFT; c.border=BOX
    cell=ws[a]; cell.border=IBOX; cell.alignment=CEN; cell.font=Font(name='Calibri',size=11,bold=True)
    nc=ws.cell(r,12,n_); nc.font=NOTE; nc.alignment=LEFT
ws['K17']='=IF(C42="","",C42)'
ws['E35']=None; ws['K17'].font=Font(name='Calibri',size=11,bold=True,color='000000')
ws['L12']=f'=IF(K12="","",COUNTIF({hr("K")},K12)&" deals in book")'
ws['L13']=f'=IF(K13="","",COUNTIF({hr("Y")},K13)&" deals in book")'
ws['L14']=f'=IF(K14="","",COUNTIF({hr("J")},K14)&" deals in book")'
ws['L15']='=IF(K15="","REQUIRED for history match","")'
ws['L16']=f'=IF(K16="","not scored",IF(K16>={C}$B$12,"GREEN",IF(K16>={C}$C$12,"YELLOW","RED")))'
ws['L18']='=IF(K18="","",IF(K18<>K17,"override in use",""))'
ws.merge_cells('L12:M12'); 
for r in range(13,19): ws.merge_cells(f'L{r}:M{r}')
ws['K12'].fill=INPUT; ws['K13'].fill=INPUT; ws['K14'].fill=INPUT; ws['K15'].fill=INPUT; ws['K16'].fill=INPUT
for a,src in [('K12','=Ref!$E$2:$E$'+str(1+len(isos))),('K13','=Ref!$G$2:$G$'+str(1+len(industries))),('K14','=Ref!$I$2:$I$'+str(1+len(states))),('K15','"New Deal,Renewal"')]:
    dv=DataValidation(type='list',formula1=src,allow_blank=True); dv.add(a); ws.add_data_validation(dv)
# ---- positions block fixes
for r in range(28,34):
    ws[f'F{r}']=f'=IF(D{r}<>"",IF(E{r}="Daily",D{r},IF(E{r}="Weekly",D{r}/5,D{r}/$K$9)),IF(C{r}="","",C{r}*{C}$B$33/({C}$B$23*$K$9)))'
    ws[f'H{r}']=f'=IF(OR(C{r}="",D{r}=""),"",ROUND(C{r}*{C}$B$33/D{r},0))'
    ws[f'I{r}']=f'=IF(H{r}="","",H{r}/IF(E{r}="Daily",$K$9,IF(E{r}="Weekly",$K$9/5,1)))'
    ws[f'J{r}']=f'=IF(OR(I{r}="",G{r}=""),"",MAX(0,I{r}-(TODAY()-G{r})/{C}$B$32))'
ws['C34']='=SUMPRODUCT(((C28:C33<>"")+(D28:D33<>"")>0)*(E28:E33<>"Monthly"))'
ws['F34']='=E34*$K$9'
# ---- metrics: thresholds linked to Controls
metrics=[(38,'Total LVG (all debt / rev)',5,'=IF(C24=0,"",(E34+IF(C8="",0,C8))*$K$9/C24)','0.00%'),
         (39,'New holdback SP%',6,'=IF(OR(C24=0,E8=""),"",E8/C24)','0.00%'),
         (40,'Daily debt / balance',7,'=IF(D24=0,"",(E34+IF(C8="",0,C8))/D24)','0.00%'),
         (41,'Fund / revenue',8,'=IF(OR(C24=0,B6=""),"",B6/C24)','0.00'),
         (42,'# positions (incl. new)',9,'=IF(C24=0,"",C34+1)','0'),
         (43,'Revenue trend (avg MoM)',10,'=G24','0.0%'),
         (44,'Neg days + NSFs (total)',11,'=E24+F24','0'),
         (45,'Credit score (FICO)',12,'=IF(K16="","",K16)','0')]
for r,t,cr,f,fmt in metrics:
    ws[f'B{r}']=t; ws[f'C{r}']=f; ws[f'C{r}'].number_format=fmt if r not in (38,39,40) else '0.00%'
    for col,cc in (('D','B'),('E','C'),('G','D')):
        ws[f'{col}{r}']=f'={C}${cc}${cr}'; ws[f'{col}{r}'].font=LINK
        ws[f'{col}{r}'].number_format = '0.0%' if r in (38,39,40,43) else ('0.00' if r==41 else '0')
    if r in (43,45):  # lower is worse
        ws[f'F{r}']=f'=IF(C{r}="","",IF(AND(G{r}<>0,C{r}<=G{r}),"KNOCKOUT",IF(C{r}>=D{r},"GREEN",IF(C{r}<E{r},"RED","YELLOW"))))' if r==45 else '=IF(C24=0,"",IF(NOT(ISNUMBER(C43)),"YELLOW",IF(C43<=G43,"KNOCKOUT",IF(C43>=D43,"GREEN",IF(C43<=E43,"RED","YELLOW")))))'
    else:
        ws[f'F{r}']=f'=IF(C{r}="","",IF(C{r}>=G{r},"KNOCKOUT",IF(C{r}<=D{r},"GREEN",IF(C{r}>=E{r},"RED","YELLOW"))))'
ws['C45'].number_format='0'
ws['B46']='Existing leverage (before new)'; ws['C46']='=IF(C24=0,"",F34/C24)'; ws['D46']='-'; ws['E46']='-'; ws['F46']='INFO'
ws['B47']='Balance / revenue'; ws['C47']='=IF(C24=0,"",D24/C24)'; ws['D47']='-'; ws['E47']='-'; ws['F47']='INFO'
ws['C46'].number_format='0.0%'; ws['C47'].number_format='0.0%'
rag_rules(ws,'F38:F45'); dec_rules(ws,'C2'); dec_rules(ws,'C52'); dec_rules(ws,'C67')
# ---- D. decision
ws['B49']='D.  DECISION  (cash-flow layer, on the requested offer)'
ws['B50']='Red flags'; ws['C50']='=COUNTIF(F38:F45,"RED")+COUNTIF(F38:F45,"KNOCKOUT")'; ws['D50']='Knockouts'; ws['E50']='=COUNTIF(F38:F45,"KNOCKOUT")'
ws['B51']='Decline if red flags >=  (Controls)'; ws['C51']=f'={C}$B$13'; ws['D51']='Minimum fund $'; ws['E51']=f'={C}$B$14'
ws['C51'].font=Font(name='Calibri',size=12,bold=True,color='008000'); ws['E51'].font=Font(name='Calibri',size=12,bold=True,color='008000')
ws['B52']='DECISION'
ws['C52']='=IF(C24=0,"",IF(OR(E50>0,IF(ISNUMBER(C63),C63,0)<E51),"DECLINE",IF(OR(B6="",B6=0),"",IF(C50>=C51,"DECLINE",IF(COUNTIF(F38:F45,"GREEN")=D55,"APPROVE","CONDITIONAL")))))'
ws['D52']='=IF(C52="DECLINE",IF(E50>0,"KNOCKOUT: "&INDEX(B38:B45,MATCH("KNOCKOUT",F38:F45,0))&" is past its knockout line.",IF(C50>=C51,"Red flags reached the threshold.","Recommended fund "&IF(ISNUMBER(C63),TEXT(C63,"$#,##0"),"$0")&" is below your minimum "&TEXT(E51,"$#,##0")&".")),IF(C52="APPROVE","All "&D55&" decision metrics green.",IF(C52="","Enter offer + bank data.","Fundable but not clean - review yellows/reds above.")))'
# ---- E. offer engine (cash-flow max)
ws['B54']='E.  DEAL QUALITY  &  CASH-FLOW MAXIMUM OFFER  (bank data alone)'
ws['B55']='Score (G=1, Y=0.5, R=0)'; ws['C55']='=COUNTIF(F38:F45,"GREEN")+0.5*COUNTIF(F38:F45,"YELLOW")'; ws['D55']='=SUMPRODUCT(--(F38:F45<>""))'; ws['E55']='scored metrics'
ws['B56']='Tier  (score scaled to 7 for the bands)'; ws['C56']=f'=IF(C24=0,"",IF(D57>={C}$B$25,"Strong",IF(D57>={C}$C$25,"Solid",IF(D57>={C}$D$25,"Moderate","Weak"))))'
ws['B57']='Max daily by SP% green'; ws['C57']='=IF(C24=0,"",D39*C24/$K$9)'; ws['D57']='=IF(D55=0,0,C55*7/D55)'; ws['E57']='scaled score /7'
ws['B58']='Max daily by Total LVG green'; ws['C58']='=IF(C24=0,"",(D38*C24-F34)/$K$9)'
ws['B59']='Max daily by Daily/Balance green'; ws['C59']='=IF(OR(C24=0,D24=0),"",D40*D24-E34)'
ws['B60']='CASH-FLOW MAX DAILY  (weakest link)'; ws['C60']='=IF(C24=0,"",MAX(0,MIN(C57:C59)))'
ws['D60']='=IF(C60="","",IF(C60=0,"no clean room - existing debt is over a green line",INDEX({"SP% holdback";"Total LVG";"Daily/Balance"},MATCH(C60,C57:C59,0))&" is the limit"))'
ws['B61']='Weekly / monthly equivalent'; ws['C61']='=IF(OR(C60="",C60=0),"",C60*5)'; ws['D61']='=IF(OR(C60="",C60=0),"",C60*$K$9)'
ws['B62']='Term (daily pmts) / factor from bands'
ws['C62']=f'=IF(OR(C60="",C60=0,C55=""),"",MAX(20,MIN(INDEX({C}$B$19:$B$22,MATCH(D57,{C}$A$19:$A$22,1)),FLOOR({C}$B$23*$K$9,5),FLOOR(D41*C24*D62/C60,5))))'
ws['D62']=f'=IF(OR(C55="",C24=0),"",INDEX({C}$C$19:$C$22,MATCH(D57,{C}$A$19:$A$22,1)))'
ws['B63']='CASH-FLOW MAX FUND'
ws['C63']='=IF(COUNTIF(F42:F45,"KNOCKOUT")>0,"-",IF(OR(C60="",C60=0,C62="",D62=""),"",FLOOR(C60*C62/D62,500)))'
ws['D63']='=IF(C63="-","KNOCKOUT on the merchant - do not fund.",IF(C63="","","payback "&TEXT(C63*D62,"$#,##0")&" over "&C62&" daily pmts ("&TEXT(C62/5,"0")&" wks) - this is the ceiling; see F below"))'
ws['C57'].number_format=ws['C58'].number_format=ws['C59'].number_format=ws['C60'].number_format='"$"#,##0.00'; ws['D57'].number_format='0.0'
ws['C61'].number_format=ws['D61'].number_format='"$"#,##0'; ws['C62'].number_format='0'; ws['D62'].number_format='0.00'; ws['C63'].number_format='"$"#,##0'
# ---- F. final offer (house book applied)
DS="'Decision Summary'!"; OE="'Offer Engine'!"; HS="'Historical Score'!"
r0=65
hdr(ws,f'B{r0}','F.  FINAL OFFER  =  cash-flow maximum  x  house-book evidence   (full detail on Decision Summary)',f'B{r0}:H{r0}')
ws.row_dimensions[r0].height=18
sub(ws,r0+1,None,['Item','Value','Detail'],start=2); ws.merge_cells(f'D{r0+1}:H{r0+1}')
rows=[('FINAL DECISION',f'={DS}$C$5',f'={DS}$D$5'),
      ('Recommended fund',f'={DS}$C$6',f'={DS}$D$6'),
      ('Factor',f'={DS}$C$7',f'={DS}$D$7'),
      ('Term',f'={DS}$C$8',f'={DS}$D$8'),
      ('Payment',f'={DS}$C$9',f'={DS}$D$9'),
      ('Frequency',f'={DS}$C$10',f'={DS}$D$10'),
      ('Holdback % at final offer',f'={DS}$C$11',f'={DS}$D$11'),
      ('Historical modifier',f'={DS}$C$16',f'={DS}$D$16'),
      ('Comparable deals (seasoned / all)',f'={DS}$C$19',f'={DS}$D$19'),
      ('Cohort principal lost rate',f'={DS}$C$21',f'={DS}$D$21'),
      ('Confidence',f'={DS}$C$18',f'={DS}$D$18'),
      ('Stipulations',f'={DS}$C$26',f'={DS}$D$26')]
for i,(t,v,d) in enumerate(rows):
    r=r0+2+i
    c=ws.cell(r,2,t); c.font=B; c.alignment=Alignment(horizontal='right',vertical='center'); c.border=BOX
    c=ws.cell(r,3,v); c.font=Font(name='Calibri',size=12 if i in (0,1) else 11,bold=True); c.alignment=CEN; c.border=BOX
    c=ws.cell(r,4,d); c.font=NOTE; c.alignment=LEFT; ws.merge_cells(f'D{r}:H{r}')
ws['C68'].number_format='"$"#,##0'; ws['C69'].number_format='0.00'; ws['C71'].number_format='"$"#,##0.00'; ws['C73'].number_format='0.0%'; ws['C76'].number_format='0.0%'
ws['C78'].alignment=Alignment(horizontal='left',vertical='top',wrap_text=True); ws.merge_cells('C78:H78'); ws.row_dimensions[78].height=42
rx=r0+2+len(rows)
c=ws.cell(rx,2,"Underwriter's conclusion"); c.font=B; c.alignment=Alignment(horizontal='right',vertical='top'); c.border=BOX
c=ws.cell(rx,3,f'={DS}$C$30'); c.font=N; c.alignment=WRAP; ws.merge_cells(f'C{rx}:H{rx}'); ws.row_dimensions[rx].height=105
note(ws,f'B{rx+1}','Layer 1 (sections A-E) is the safety ceiling. Section F never exceeds C63, never removes a knockout and never turns a green metric red - see Checks. Thresholds live on Controls.')
ws.column_dimensions['J'].width=27; ws.column_dimensions['K'].width=26; ws.column_dimensions['L'].width=16; ws.column_dimensions['M'].width=12
ws.sheet_view.showGridLines=False

# ======================= Controls =======================
cs=wb.create_sheet('Controls'); widths(cs,{'A':52,'B':16,'C':16,'D':16,'E':16,'F':14,'G':14,'H':60})
cs['A1']='CONTROLS  -  every editable underwriting threshold, pricing tier, historical weight and rule'; cs['A1'].font=Font(name='Calibri',size=14,bold=True)
note(cs,'A2','Blue = editable input. Change values here only; the Deal, Historical Score, Offer Engine and Decision Summary sheets read them. Cell addresses are referenced by formulas - do not insert or delete rows.')
hdr(cs,'A3','1. CASH-FLOW THRESHOLDS  (Deal sheet metrics)','A3:H3'); sub(cs,4,None,['Metric','Green (at or better)','Red (at or worse)','Knockout','Direction','','','Note'])
thr=[('Total LVG (all debt / rev)',0.30,0.45,0.55,'higher is worse','0.0%','Monthly debt service incl. new deal / avg true revenue'),
     ('New holdback SP%',0.18,0.25,0.35,'higher is worse','0.0%','New deal monthly payment / avg true revenue'),
     ('Daily debt / balance',0.18,0.30,0.35,'higher is worse','0.0%','All daily debt incl. new / avg daily balance'),
     ('Fund / revenue',0.60,0.90,1.25,'higher is worse','0.00','Fund $ / avg monthly true revenue'),
     ('# positions (incl. new)',2,5,6,'higher is worse','0','Existing non-monthly positions + 1'),
     ('Revenue trend (avg MoM)',-0.05,-0.15,-0.30,'lower is worse','0.0%','Average month-over-month change of entered months'),
     ('Neg days + NSFs (total)',1,6,10,'higher is worse','0','Total across ALL months entered (not per month)'),
     ('Credit score (FICO)',650,630,0,'lower is worse','0','Your bands: 650+ green, 630-649 yellow, below 630 red. Knockout: score at or below this = knockout; 0 = off')]
for i,(t,g,r_,k,d,fmt,n_) in enumerate(thr):
    r=5+i; label(cs,f'A{r}',t)
    for col,v in (('B',g),('C',r_),('D',k)): put(cs,f'{col}{r}',v,INP,fmt,border=IBOX)
    put(cs,f'E{r}',d,N,align=LEFT); note(cs,f'H{r}',n_)
label(cs,'A13','Decline if red flags (RED + KNOCKOUT) >='); put(cs,'B13',2,INP,'0',border=IBOX)
label(cs,'A14','Minimum fund $ (below this = decline)'); put(cs,'B14',20000,INP,'"$"#,##0',border=IBOX)
label(cs,'A15','Underwriting fee % (deducted from disbursement)'); put(cs,'B15',0.05,INP,'0%',border=IBOX)
hdr(cs,'A17','2. PRICING BANDS  (score on the 7-point scale -> term cap and factor)','A17:H17'); sub(cs,18,None,['Score >=','Term cap (daily pmts)','Factor','','','','','Note'])
for i,(s,t,f) in enumerate([(0,100,1.49),(5,130,1.45),(6,150,1.45),(7,195,1.40)]):
    r=19+i; put(cs,f'A{r}',s,INP,'0.0',border=IBOX); put(cs,f'B{r}',t,INP,'0',border=IBOX); put(cs,f'C{r}',f,INP,'0.00',border=IBOX)
note(cs,'H19','Bands are unchanged from the original tool. Term step -1 moves one row up this table.')
label(cs,'A23','Max term (months)'); put(cs,'B23',9,INP,'0',border=IBOX); note(cs,'H23','195 collection days at 21.655/month')
label(cs,'A24','Approved factor tiers (low -> high)'); put(cs,'B24',1.40,INP,'0.00',border=IBOX); put(cs,'C24',1.45,INP,'0.00',border=IBOX); put(cs,'D24',1.49,INP,'0.00',border=IBOX); note(cs,'H24','1.35 is the prepay rate, not a tier. 1.50+ is a hard no.')
label(cs,'A25','Tier cut-offs on 7-point scale: Strong / Solid / Moderate'); put(cs,'B25',6.5,INP,'0.0',border=IBOX); put(cs,'C25',5,INP,'0.0',border=IBOX); put(cs,'D25',3.5,INP,'0.0',border=IBOX)
hdr(cs,'A27','3. CONVERSION CONSTANTS','A27:H27')
for r,(t,v,fmt,n_) in enumerate([('Weeks per month (ours)',4.331,'0.000','Time converts with 4.331'),('Business days per month (ours)',21.655,'0.000','Money converts with 5 and 21.655'),('MoneyBadger weeks per month',4,'0','MB uses 4 on weekly deals'),('MoneyBadger business days per month',21,'0','MB uses 21 on daily deals (one confirmed data point)'),('Calendar days per month (months-remaining calc)',30.4375,'0.0000',''),('Assumed factor on existing positions (est. term)',1.45,'0.00','Used when only funded $ / payment is known')],start=28):
    label(cs,f'A{r}',t); put(cs,f'B{r}',v,INP,fmt,border=IBOX); note(cs,f'H{r}',n_)
hdr(cs,'A35','4. HISTORICAL EVIDENCE SETTINGS','A35:H35')
hist=[('As-of date for seasoning (date of the CRM export)',dt.datetime(2026,8,25),'mm/dd/yyyy','Age = business days from funding to this date'),
      ('Window: include deals funded on or after',dt.datetime(2024,11,1),'mm/dd/yyyy',''),
      ('Window: include deals funded on or before',dt.datetime(2026,8,25),'mm/dd/yyyy','Full file. Set to 12/31/2025 for the settled book'),
      ('Seasoning multiple (age >= multiple x contracted term)',1.25,'0.00','v13 rule from the house-book session'),
      ('Minimum seasoned deals for a cohort to be PRIMARY',10,'0','Below this the engine drops to the next tier'),
      ('Minimum seasoned funded $ for a cohort to be PRIMARY',150000,'"$"#,##0',''),
      ('Credibility constant K (seasoned deals for Z = 0.5)',30,'0','Z = n / (n + K). 30 deals -> half weight; 90 -> 75%'),
      ('Max open share for a cohort to count as seasoned-enough',0.15,'0%','Above this the cohort is flagged and confidence cannot be High'),
      ('High confidence: minimum Z',0.60,'0.00','and open share within the line above'),
      ('Medium confidence: minimum Z',0.30,'0.00','Below = Low'),
      ('Low confidence: max steps from Neutral the modifier may move',1,'0','Weak evidence can never drive a severe adjustment'),
      ('Minimum seasoned deals before a RATE is displayed',5,'0','Counts and dollars always show')]
for i,(t,v,fmt,n_) in enumerate(hist):
    r=36+i; label(cs,f'A{r}',t); put(cs,f'B{r}',v,INP,fmt,border=IBOX); note(cs,f'H{r}',n_)
hdr(cs,'A52','5. HISTORICAL MODIFIER TABLE  (relative index = blended cohort lost rate / seasoned-book lost rate)','A52:H52')
sub(cs,53,None,['Index at or below','Level','Fund multiplier (x cash-flow max)','Term step (bands)','Factor step (tiers)','Frequency','Manual review?','Note'])
mods=[(0.60,'Strong positive',1.00,0,0,'As proposed','N','Full cash-flow maximum only with strong, credible evidence'),
      (0.85,'Mild positive',0.95,0,0,'As proposed','N',''),
      (1.15,'Neutral',0.90,0,0,'As proposed','N','Base utilisation. Set 1.00 to reproduce the old tool'),
      (1.40,'Mild negative',0.80,-1,1,'As proposed','N','Shorter band, next factor tier'),
      (1.75,'Negative',0.70,-1,1,'Weekly','N','Adds weekly remittance + no-stacking covenant'),
      (999,'Severe',0.55,-2,1,'Weekly','Y','Sent to manual review')]
for i,m in enumerate(mods):
    r=54+i
    for j,v in enumerate(m[:7]):
        fmt={0:'0.00',2:'0.00',3:'0',4:'0'}.get(j,None); put(cs,f'{L(1+j)}{r}',v,INP if j!=1 else B,fmt,border=IBOX if j!=1 else BOX)
    note(cs,f'H{r}',m[7])
hdr(cs,'A62','6. STATUS MAPPING  (the only place default / resolved / closed / open are defined)','A62:H62')
sub(cs,63,None,['Status','Counts as default?','Counts as resolved?','Counts as Closed?','Counts as Open?','Deals in file','',''])
for i,s in enumerate(statuses):
    r=64+i; label(cs,f'A{r}',s[0])
    for j in range(4): put(cs,f'{L(2+j)}{r}',s[j+1],INP,border=IBOX)
    put(cs,f'F{r}',f'=COUNTIF({hr("D")},A{r})',N,'0')
note(cs,'H68','Lowered Payments = restructured, still collecting. House book treats it as performing; switch B68 to Yes to treat as impaired.')
note(cs,'H74','At-Risk: 1 deal, defaulted per house book.')
hdr(cs,'A77','7. MANUAL-REVIEW TRIGGERS  (Y = active)','A77:H77'); sub(cs,78,None,['Trigger','Active?','Parameter','','','','','Note'])
trig=[('Historical modifier is Severe','Y',None,''),('Confidence Low AND any metric RED','Y',None,''),('Credit score RED AND another metric RED','Y',None,''),('Positions incl. new at or above',
'Y',4,'Log rule: fund up to a 4th; parameter = position count that triggers review'),('New deal, 3rd position or higher, construction industry (SIC 15/16/17)','Y',3,'Parameter = position at or above. The one raw house-book effect that stayed visible after controls'),('ISO or industry has no deals in the house book','Y',None,'Unknown relationship'),('Adjusted fund below minimum fund $','Y',None,''),('Primary cohort open share above the max open share','N',None,'Off by default; the flag is shown anyway')]
for i,(t,a,p,n_) in enumerate(trig):
    r=79+i; label(cs,f'A{r}',t); put(cs,f'B{r}',a,INP,border=IBOX); 
    if p is not None: put(cs,f'C{r}',p,INP,'0',border=IBOX)
    note(cs,f'H{r}',n_)
hdr(cs,'A88','8. STIPULATION RULES  (Y = active)','A88:H88'); sub(cs,89,None,['Stipulation','Active?','Condition','','','','','Note'])
stips=[('DataMerch, UCC search, positions / payoff letters','Y','Always','Standing rule from the underwriting session'),
       ('No-stacking covenant with default-to-daily','Y','Modifier Mild negative or worse',''),
       ('Bank login verification (Plaid / DecisionLogic)','Y','Neg days + NSFs not GREEN',''),
       ('Payoff / balance letter for every existing position','Y','Existing positions >= 2',''),
       ('Landlord / lease letter','Y','Credit score not GREEN',''),
       ('Underwriter review of the historical evidence','Y','Confidence Low',''),
       ('Weekly remittance (no daily option)','Y','Frequency switched to Weekly by the modifier',''),
       ('Prior-contract payoff and renewal history','Y','Renewal selected','')]
for i,(t,a,cnd,n_) in enumerate(stips):
    r=90+i; label(cs,f'A{r}',t); put(cs,f'B{r}',a,INP,border=IBOX); put(cs,f'C{r}',cnd,N,align=LEFT); cs.merge_cells(f'C{r}:G{r}'); note(cs,f'H{r}',n_)
cs.sheet_view.showGridLines=False

# ======================= Ref =======================
rf=wb.create_sheet('Ref'); widths(rf,{'A':8,'B':40,'C':32,'E':44,'G':40,'I':10,'K':14,'M':22,'O':12})
for col,t in zip('ABCEGIKM',['SIC2','Industry (SIC major group)','Division','ISO (house-book spelling)','Industry list (dropdown)','State','Fed holiday','Status list']):
    c=rf[f'{col}1']; c.value=t; c.font=SUB; c.fill=MID; c.alignment=CEN; c.border=BOX
for i,(a,b,c_) in enumerate(sic): rf.cell(2+i,1,a); rf.cell(2+i,2,b); rf.cell(2+i,3,c_)
for i,v in enumerate(isos): rf.cell(2+i,5,v)
for i,v in enumerate(industries): rf.cell(2+i,7,v)
for i,v in enumerate(states): rf.cell(2+i,9,v)
for i,v in enumerate(hols): rf.cell(2+i,11,v).number_format='mm/dd/yyyy'
for i,s in enumerate(statuses): rf.cell(2+i,13,s[0])
note(rf,'O1','Dropdown sources. ISO and state lists come from the house-book data so a Deal entry always matches the history spelling.')
HOL=f'Ref!$K$2:$K${1+len(hols)}'

# ======================= Historical Data =======================
hs=wb.create_sheet('Historical Data')
cols=['Contract ID','DBA','Legal name','Status','Start date','Last active','Position','SIC','SIC2','State','ISO','New or renewal','Factor','Advance','Commission','Payback','Collected','Balance','% paid','Exp dur (raw)','NSF count','Avg rev/mo','Holdback %','Biz start',
      'Industry','Pay frequency','Term (bus. days)','Age at as-of (bus. days)','Seasoned','In window','Default','Resolved','Closed family','Open','Status valid','Net cash','Principal lost','Adv / revenue','Has revenue','Days on book (bus.)','Renewal',
      'Eligible (seasoned & in window)','m: position','m: new/renewal','m: industry','m: ISO','m: state','T1 NR+Ind+Pos','T2 NR+ISO+Pos','T3 NR+Ind','T4 NR+ISO','T5 NR+Pos','T6 NR','T7 Book','Best tier','In primary cohort','Seq']
for i,t in enumerate(cols):
    c=hs.cell(1,1+i,t); c.font=SUB; c.fill=MID if i<24 else DARK; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=BOX
hs.row_dimensions[1].height=42
# raw: A..AA in source; we keep A..H, J(SIC2), L..AA -> skip I (SIC description) and K (MCC)
keep=[0,1,2,3,4,5,6,7,9,11,12,13,14,15,16,17,18,19,20,21,23,24,25,26]
for i,r in enumerate(raw):
    rr=2+i
    for j,k in enumerate(keep):
        v=r[k]
        if k in (7,9) and v is not None: v=str(v)
        hs.cell(rr,1+j,v)
    st=f"Controls!$A$64:$E$74"
    f={ 'Y':f'=IF(I{rr}="","No SIC",IFERROR(VLOOKUP(I{rr},Ref!$A$2:$B$65,2,FALSE),"Unclassified"))',
        'Z':f'=IF(T{rr}="","",IF(T{rr}<12,"Daily","Weekly"))',
        'AA':f'=IF(T{rr}="",0,IF(T{rr}<12,T{rr}*Controls!$B$29,T{rr}*5))',
        'AB':f'=IF(E{rr}="",0,NETWORKDAYS(E{rr},Controls!$B$36,{HOL}))',
        'AC':f'=IF(AND(AA{rr}>0,AB{rr}>=Controls!$B$39*AA{rr}),1,0)',
        'AD':f'=IF(AND(E{rr}<>"",E{rr}>=Controls!$B$37,E{rr}<=Controls!$B$38),1,0)',
        'AE':f'=IF(IFERROR(VLOOKUP(D{rr},{st},2,FALSE),"No")="Yes",1,0)',
        'AF':f'=IF(IFERROR(VLOOKUP(D{rr},{st},3,FALSE),"No")="Yes",1,0)',
        'AG':f'=IF(IFERROR(VLOOKUP(D{rr},{st},4,FALSE),"No")="Yes",1,0)',
        'AH':f'=IF(IFERROR(VLOOKUP(D{rr},{st},5,FALSE),"No")="Yes",1,0)',
        'AI':f'=IF(ISNUMBER(MATCH(D{rr},Controls!$A$64:$A$74,0)),1,0)',
        'AJ':f'=Q{rr}-N{rr}',
        'AK':f'=IF(AE{rr}=1,MAX(0,N{rr}-Q{rr}),0)',
        'AL':f'=IF(OR(V{rr}="",V{rr}=0),0,N{rr}/V{rr})',
        'AM':f'=IF(OR(V{rr}="",V{rr}=0),0,1)',
        'AN':f'=IF(OR(E{rr}="",F{rr}=""),0,NETWORKDAYS(E{rr},F{rr},{HOL}))',
        'AO':f'=IF(L{rr}="Renewal",1,0)',
        'AP':f'=AC{rr}*AD{rr}',
        'AQ':f"=IF(AND(G{rr}<>\"\",'Historical Score'!$C$6<>\"\"),IF(G{rr}='Historical Score'!$C$6,1,0),0)",
        'AR':f"=IF(AND(L{rr}<>\"\",'Historical Score'!$C$5<>\"\"),IF(L{rr}='Historical Score'!$C$5,1,0),0)",
        'AS':f"=IF(AND(Y{rr}<>\"\",'Historical Score'!$C$7<>\"\"),IF(Y{rr}='Historical Score'!$C$7,1,0),0)",
        'AT':f"=IF(AND(K{rr}<>\"\",'Historical Score'!$C$8<>\"\"),IF(K{rr}='Historical Score'!$C$8,1,0),0)",
        'AU':f"=IF(AND(J{rr}<>\"\",'Historical Score'!$C$9<>\"\"),IF(J{rr}='Historical Score'!$C$9,1,0),0)",
        'AV':f'=AR{rr}*AS{rr}*AQ{rr}', 'AW':f'=AR{rr}*AT{rr}*AQ{rr}', 'AX':f'=AR{rr}*AS{rr}', 'AY':f'=AR{rr}*AT{rr}', 'AZ':f'=AR{rr}*AQ{rr}', 'BA':f'=AR{rr}', 'BB':'=1',
        'BC':f'=IF(AV{rr}=1,1,IF(AW{rr}=1,2,IF(AX{rr}=1,3,IF(AY{rr}=1,4,IF(AZ{rr}=1,5,IF(BA{rr}=1,6,7))))))',
        'BD':f"=IF(AD{rr}=1,CHOOSE('Historical Score'!$C$29,AV{rr},AW{rr},AX{rr},AY{rr},AZ{rr},BA{rr},BB{rr}),0)",
        'BE':f'=IF(BD{rr}=1,COUNTIF($BD$2:BD{rr},1),"")'}
    for col,fx in f.items(): hs[f'{col}{rr}']=fx
    for col in ('E','F','X'): hs[f'{col}{rr}'].number_format='mm/dd/yyyy'
    for col in ('N','O','P','Q','R'): hs[f'{col}{rr}'].number_format='"$"#,##0'
    hs[f'S{rr}'].number_format='0.0%'; hs[f'W{rr}'].number_format='0.0%'; hs[f'M{rr}'].number_format='0.000'
hs.freeze_panes='E2'; widths(hs,{'A':12,'B':28,'C':28,'D':18,'K':24,'Y':26})
assert L(len(cols))=='BE'

# ======================= Historical Score =======================
sc=wb.create_sheet('Historical Score'); widths(sc,{'A':7,'B':34,'C':30,'D':9,'E':9,'F':9,'G':13,'H':8,'I':8,'J':8,'K':8,'L':10,'M':10,'N':10,'O':9,'P':9,'Q':8,'R':11,'S':9,'T':9,'U':10,'V':9,'W':10,'X':9,'Y':8,'Z':12,'AA':12})
sc['A1']='HISTORICAL SCORE  -  comparable cohorts, credibility and the risk modifier'; sc['A1'].font=Font(name='Calibri',size=14,bold=True)
note(sc,'A2','Rates use SEASONED deals only (age >= seasoning multiple x contracted term, as of the Controls date). Unseasoned deals are counted, never rated. Blended lost rate = Z x cohort + (1-Z) x seasoned book, Z = n/(n+K).')
hdr(sc,'A4','DEAL BEING MATCHED  (from the Deal sheet)','A4:C4')
for r,(t,f) in enumerate([('New deal or renewal','=IF(Deal!$K$15="","",Deal!$K$15)'),('Proposed position (effective)','=IF(Deal!$K$18<>"",Deal!$K$18,Deal!$K$17)'),('Industry','=IF(Deal!$K$13="","",Deal!$K$13)'),('ISO','=IF(Deal!$K$12="","",Deal!$K$12)'),('State','=IF(Deal!$K$14="","",Deal!$K$14)')],start=5):
    label(sc,f'B{r}',t); put(sc,f'C{r}',f,B)
label(sc,'B10','Seasoned-book principal lost rate (baseline)'); put(sc,'C10','=$L$20',B,'0.0%')
label(sc,'B11','Seasoned-book deals / funded'); put(sc,'C11','=TEXT($E$20,"#,##0")&" / "&TEXT($G$20,"$#,##0")',B)
heads=['Tier','Cohort','Keys','Deals (all)','Seasoned','% seasoned','Funded $ (seasoned)','Open n','Open %','Default n','Default %','Principal lost rate','Collected on defaults','Return on funded','Avg adv/rev','Avg term (bd)','Avg factor','Avg deal size','Avg days on book','Renewal share','Seasoned enough?','Credibility Z','Blended lost rate','Relative index','Meets min','Impact vs book','Lost rate shown?']
sub(sc,13,None,heads); sc.row_dimensions[13].height=40
for c in range(1,28): sc.cell(13,c).alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
tiers=[(1,'T1: new/renewal + industry + position','AV'),(2,'T2: new/renewal + ISO + position','AW'),(3,'T3: new/renewal + industry','AX'),(4,'T4: new/renewal + ISO','AY'),(5,'T5: new/renewal + position','AZ'),(6,'T6: new/renewal only','BA'),(7,'T7: seasoned book (baseline)','BB')]
single=[('Position only','AQ'),('New/renewal only','AR'),('Industry only','AS'),('ISO only','AT'),('State only (INFORMATIONAL - never in hierarchy)','AU')]
def cohort_row(r,tier,name,flag,keys):
    F=hr(flag); E_=hr('AP')
    put(sc,f'A{r}',tier,B); label(sc,f'B{r}',name); label(sc,f'C{r}',keys)
    fx={'D':f'=SUMPRODUCT({F}*{hr("AD")})','E':f'=SUMPRODUCT({F}*{E_})','F':f'=IF(D{r}=0,"",E{r}/D{r})','G':f'=SUMPRODUCT({F}*{E_}*{hr("N")})',
        'H':f'=SUMPRODUCT({F}*{E_}*{hr("AH")})','I':f'=IF(E{r}=0,"",H{r}/E{r})','J':f'=SUMPRODUCT({F}*{E_}*{hr("AE")})','K':f'=IF(E{r}=0,"",J{r}/E{r})',
        'L':f'=IF(OR(E{r}<Controls!$B$47,G{r}=0),"",SUMPRODUCT({F}*{E_}*{hr("AK")})/G{r})',
        'M':f'=IF(OR(E{r}<Controls!$B$47,SUMPRODUCT({F}*{E_}*{hr("AE")}*{hr("N")})=0),"",SUMPRODUCT({F}*{E_}*{hr("AE")}*{hr("Q")})/SUMPRODUCT({F}*{E_}*{hr("AE")}*{hr("N")}))',
        'N':f'=IF(OR(E{r}<Controls!$B$47,G{r}=0),"",SUMPRODUCT({F}*{E_}*{hr("AJ")})/G{r})',
        'O':f'=IF(SUMPRODUCT({F}*{E_}*{hr("AM")})=0,"",SUMPRODUCT({F}*{E_}*{hr("AM")}*{hr("AL")})/SUMPRODUCT({F}*{E_}*{hr("AM")}))',
        'P':f'=IF(E{r}=0,"",SUMPRODUCT({F}*{E_}*{hr("AA")})/E{r})','Q':f'=IF(E{r}=0,"",SUMPRODUCT({F}*{E_}*{hr("M")})/E{r})','R':f'=IF(E{r}=0,"",G{r}/E{r})',
        'S':f'=IF(E{r}=0,"",SUMPRODUCT({F}*{E_}*{hr("AN")})/E{r})','T':f'=IF(E{r}=0,"",SUMPRODUCT({F}*{E_}*{hr("AO")})/E{r})',
        'U':f'=IF(AND(E{r}>=Controls!$B$40,G{r}>=Controls!$B$41,IF(I{r}="",1,I{r})<=Controls!$B$43),"Yes","No")',
        'V':f'=IF(E{r}+Controls!$B$42=0,0,E{r}/(E{r}+Controls!$B$42))',
        'W':f'=IF(L{r}="",$C$10,V{r}*L{r}+(1-V{r})*$C$10)','X':f'=IF(OR($C$10="",$C$10=0),"",W{r}/$C$10)',
        'Y':f'=IF(AND(E{r}>=Controls!$B$40,G{r}>=Controls!$B$41),1,0)','Z':f'=IF(L{r}="",0,ABS(W{r}-$C$10))','AA':f'=IF(L{r}="","No","Yes")'}
    for col,f in fx.items(): put(sc,f'{col}{r}',f,N)
    for col in 'FIKLMNT': sc[f'{col}{r}'].number_format='0.0%'
    for col in 'GR': sc[f'{col}{r}'].number_format='"$"#,##0'
    for col in 'OQVX': sc[f'{col}{r}'].number_format='0.00'
    sc[f'W{r}'].number_format='0.0%'; sc[f'Z{r}'].number_format='0.0%'; sc[f'P{r}'].number_format='0'; sc[f'S{r}'].number_format='0'
keysd={1:'NR + industry + position',2:'NR + ISO + position',3:'NR + industry',4:'NR + ISO',5:'NR + position',6:'NR',7:'all seasoned deals in window'}
for t,nme,flag in tiers: cohort_row(13+t,t,nme,flag,keysd[t])
sub(sc,21,None,['','SINGLE-DIMENSION EVIDENCE (informational + largest-impact test)']+['']*25)
for i,(nme,flag) in enumerate(single): cohort_row(22+i,'S'+str(i+1),nme,flag,nme.split(' (')[0])
sc.conditional_formatting.add('U14:U26',CellIsRule(operator='equal',formula=['"Yes"'],fill=G_FILL,font=Font(color='FFFFFF',bold=True)))
sc.conditional_formatting.add('U14:U26',CellIsRule(operator='equal',formula=['"No"'],fill=Y_FILL))
hdr(sc,'A28','RESULT','A28:C28')
res=[('Primary tier (first tier meeting the minimums)','=IFERROR(MATCH(1,$Y$14:$Y$20,0),7)','0'),
     ('Primary cohort','=INDEX($B$14:$B$20,$C$29)',None),
     ('Seasoned deals in primary cohort / all matches','=TEXT(INDEX($E$14:$E$20,$C$29),"#,##0")&" / "&TEXT(INDEX($D$14:$D$20,$C$29),"#,##0")',None),
     ('% of comparable deals seasoned','=INDEX($F$14:$F$20,$C$29)','0.0%'),
     ('Cohort principal lost rate (seasoned)','=INDEX($L$14:$L$20,$C$29)','0.0%'),
     ('Cohort open share','=INDEX($I$14:$I$20,$C$29)','0.0%'),
     ('Cohort return on funded capital','=INDEX($N$14:$N$20,$C$29)','0.0%'),
     ('Cohort collected on defaults','=INDEX($M$14:$M$20,$C$29)','0.0%'),
     ('Credibility Z','=INDEX($V$14:$V$20,$C$29)','0.00'),
     ('Blended lost rate','=INDEX($W$14:$W$20,$C$29)','0.0%'),
     ('Relative index (blended / book)','=INDEX($X$14:$X$20,$C$29)','0.00'),
     ('Modifier level from index (1 strong positive .. 6 severe)','=IF(C39="",3,COUNTIF(Controls!$A$54:$A$59,"<"&C39)+1)','0'),
     ('Confidence','=IF(OR($C$29=7,C37<Controls!$B$45),"Low",IF(AND(C37>=Controls!$B$44,IF(C34="",1,C34)<=Controls!$B$43),"High","Medium"))',None),
     ('Modifier level after confidence cap','=IF(C41="Low",MIN(MAX(C40,3-Controls!$B$46),3+Controls!$B$46),C40)','0'),
     ('Modifier level name','=INDEX(Controls!$B$54:$B$59,C42)',None),
     ('Largest-impact variable (position / NR / industry / ISO)','=IF(MAX($Z$22:$Z$25)=0,"none - no single-dimension cohort has a rate",INDEX($B$22:$B$25,MATCH(MAX($Z$22:$Z$25),$Z$22:$Z$25,0)))',None),
     ('   its blended lost rate vs book','=IF(MAX($Z$22:$Z$25)=0,"",TEXT(INDEX($W$22:$W$25,MATCH(MAX($Z$22:$Z$25),$Z$22:$Z$25,0)),"0.0%")&" vs "&TEXT($C$10,"0.0%")&" book ("&TEXT(INDEX($E$22:$E$25,MATCH(MAX($Z$22:$Z$25),$Z$22:$Z$25,0)),"0")&" seasoned deals)")',None),
     ('ISO evidence','=IF($C$8="","ISO not entered",IF($E$25=0,"ISO "&$C$8&" has no seasoned deals in the book",IF($L$25="","ISO "&$C$8&": "&$E$25&" seasoned deals - too few to rate","ISO "&$C$8&": "&$E$25&" seasoned deals, lost "&TEXT($L$25,"0.0%")&" vs "&TEXT($C$10,"0.0%")&" book, Z "&TEXT($V$25,"0.00")&" -> "&IF($X$25<=Controls!$A$55,"stronger than book",IF($X$25<=Controls!$A$56,"neutral","weaker than book")))))',None),
     ('Industry evidence','=IF($C$7="","Industry not entered",IF($E$24=0,"Industry has no seasoned deals in the book",IF($L$24="","Industry "&$C$7&": "&$E$24&" seasoned deals - too few to rate","Industry "&$C$7&": "&$E$24&" seasoned deals, lost "&TEXT($L$24,"0.0%")&" vs "&TEXT($C$10,"0.0%")&" book, Z "&TEXT($V$24,"0.00")&" -> "&IF($X$24<=Controls!$A$55,"stronger than book",IF($X$24<=Controls!$A$56,"neutral","weaker than book")))))',None),
     ('State evidence (informational only)','=IF($C$9="","State not entered",IF($E$26=0,"State "&$C$9&" has no seasoned deals",IF($L$26="","State "&$C$9&": "&$E$26&" seasoned deals - too few to rate; not used","State "&$C$9&": "&$E$26&" seasoned deals, lost "&TEXT($L$26,"0.0%")&" vs "&TEXT($C$10,"0.0%")&" book - informational only, not used in the offer")))',None),
     ('Position evidence','=IF($C$6="","Position not entered",IF($L$22="","Position "&$C$6&": "&$E$22&" seasoned deals - too few to rate","Position "&$C$6&": "&$E$22&" seasoned deals, lost "&TEXT($L$22,"0.0%")&" vs "&TEXT($C$10,"0.0%")&" book"))',None),
     ('Conflicting signals? (ISO and industry on opposite sides of neutral)','=IF(OR($L$24="",$L$25=""),"No",IF(OR(AND($X$24<Controls!$A$55,$X$25>Controls!$A$56),AND($X$25<Controls!$A$55,$X$24>Controls!$A$56)),"YES - review both cohorts","No"))',None),
     ('Warnings','=IF($C$29=7,"No cohort met the minimums - book baseline used; ","")&IF(INDEX($U$14:$U$20,$C$29)="No","Primary cohort not seasoned-enough (count, dollars or open share); ","")&IF(C41="Low","Low confidence - modifier capped; ","")&IF(C33<Controls!$B$43,"","Cohort open share high; ")',None)]
for i,(t,f,fmt) in enumerate(res):
    r=29+i; label(sc,f'B{r}',t); put(sc,f'C{r}',f,B,fmt,align=LEFT)
sc.sheet_view.showGridLines=False; sc.freeze_panes='D14'

# ======================= Offer Engine =======================
oe=wb.create_sheet('Offer Engine'); widths(oe,{'A':3,'B':46,'C':20,'D':20,'E':14,'F':60})
oe['B1']='OFFER ENGINE  -  base cash-flow offer, historical adjustment, adjusted offer'; oe['B1'].font=Font(name='Calibri',size=14,bold=True)
note(oe,'B2','Adjusted fund = MIN(base fund x multiplier, cash-flow max daily x adjusted term / adjusted factor). The second term keeps the daily payment at or below the cash-flow maximum even when the term is shortened.')
sub(oe,4,None,['','Item','Base (cash flow only)','Adjusted (with history)','Status','How'],start=1)
D='Deal!'
rows=[(5,'Cash-flow max daily payment',f'={D}$C$60',f'={D}$C$60','"$"#,##0.00','Deal C60 - weakest of holdback / total LVG / daily-debt-to-balance green lines'),
      (6,'Term (daily payments)',f'={D}$C$62','=IF(OR(C6="",C8="-"),"",MAX(20,MIN(C6,INDEX(Controls!$B$19:$B$22,MAX(1,MATCH(Deal!$D$57,Controls!$A$19:$A$22,1)+C12)))))','0','Base band term; step moves down the Controls band table'),
      (7,'Factor',f'={D}$D$62','=IF(C7="","",INDEX(Controls!$B$24:$D$24,MIN(3,MAX(1,MATCH(C7,Controls!$B$24:$D$24,0)+C13))))','0.00','Next approved tier, never above 1.49'),
      (8,'Fund $',f'={D}$C$63','=IF(C8="-","-",IF(OR(C8="",D6="",D7=""),"",MAX(0,FLOOR(MIN(C8*C11,C5*D6/D7),500))))','"$"#,##0','Never above the base fund'),
      (9,'Daily payment','=IF(OR(C8="",C8="-"),"",C8*C7/C6)','=IF(OR(D8="",D8="-",D8=0),"",D8*D7/D6)','"$"#,##0.00','fund x factor / term'),
      (10,'Historical modifier level (after confidence cap)',"='Historical Score'!$C$42","='Historical Score'!$C$42",'0',"1 strong positive .. 3 neutral .. 6 severe"),
      (11,'Fund multiplier','=INDEX(Controls!$C$54:$C$59,C10)','=INDEX(Controls!$C$54:$C$59,C10)','0.00','Controls modifier table'),
      (12,'Term step (bands)','=INDEX(Controls!$D$54:$D$59,C10)','=INDEX(Controls!$D$54:$D$59,C10)','0',''),
      (13,'Factor step (tiers)','=INDEX(Controls!$E$54:$E$59,C10)','=INDEX(Controls!$E$54:$E$59,C10)','0',''),
      (14,'Frequency rule','=INDEX(Controls!$F$54:$F$59,C10)','=INDEX(Controls!$F$54:$F$59,C10)',None,''),
      (15,'Manual review flag from modifier','=INDEX(Controls!$G$54:$G$59,C10)','=INDEX(Controls!$G$54:$G$59,C10)',None,''),
      (16,'Frequency (final)',f'=IF({D}$E$6="","Daily",{D}$E$6)',f'=IF(C14="Weekly","Weekly",IF({D}$E$6="","Daily",{D}$E$6))',None,'Proposed frequency unless the modifier requires weekly'),
      (17,'Term in final units (weeks or daily pmts)','=IF(OR(C6="",C8="-"),"",IF(C16="Weekly",C6/5,C6))','=IF(OR(D6="",D8="-"),"",IF(D16="Weekly",D6/5,D6))','0',''),
      (18,'Payment in final units','=IF(C9="","",IF(C16="Weekly",C9*5,C9))','=IF(D9="","",IF(D16="Weekly",D9*5,D9))','"$"#,##0.00',''),
      (19,'Monthly payment','=IF(C9="","",C9*Controls!$B$29)','=IF(D9="","",D9*Controls!$B$29)','"$"#,##0',''),
      (20,'Payback','=IF(OR(C8="",C8="-"),"",C8*C7)','=IF(OR(D8="",D8="-"),"",D8*D7)','"$"#,##0',''),
      (21,'Net disbursed (after fee)','=IF(OR(C8="",C8="-"),"",C8*(1-Controls!$B$15))','=IF(OR(D8="",D8="-"),"",D8*(1-Controls!$B$15))','"$"#,##0',''),
      (22,'Term (months)','=IF(C6="","",C6/Controls!$B$29)','=IF(D6="","",D6/Controls!$B$29)','0.0','')]
for r,t,cb,ca,fmt,how in rows:
    label(oe,f'B{r}',t); put(oe,f'C{r}',cb,N,fmt); put(oe,f'D{r}',ca,B,fmt); note(oe,f'F{r}',how)
hdr(oe,'B24','METRICS RE-CHECKED AT EACH OFFER  (must stay off RED / KNOCKOUT)','B24:F24')
sub(oe,25,None,['','Metric','Base offer','Adjusted offer','Status (adjusted)','Green / Red / Knockout'],start=1)
mt=[(26,'Total LVG (all debt / rev)',f'=IF(OR({D}$C$24=0,C9=""),"",({D}$E$34+C9)*Controls!$B$29/{D}$C$24)',f'=IF(OR({D}$C$24=0,D9=""),"",({D}$E$34+D9)*Controls!$B$29/{D}$C$24)',5,'0.0%'),
    (27,'New holdback SP%',f'=IF(OR({D}$C$24=0,C9=""),"",C9*Controls!$B$29/{D}$C$24)',f'=IF(OR({D}$C$24=0,D9=""),"",D9*Controls!$B$29/{D}$C$24)',6,'0.0%'),
    (28,'Daily debt / balance',f'=IF(OR({D}$D$24=0,C9=""),"",({D}$E$34+C9)/{D}$D$24)',f'=IF(OR({D}$D$24=0,D9=""),"",({D}$E$34+D9)/{D}$D$24)',7,'0.0%'),
    (29,'Fund / revenue',f'=IF(OR({D}$C$24=0,C8="",C8="-"),"",C8/{D}$C$24)',f'=IF(OR({D}$C$24=0,D8="",D8="-"),"",D8/{D}$C$24)',8,'0.00')]
for r,t,fb,fa,cr,fmt in mt:
    label(oe,f'B{r}',t); put(oe,f'C{r}',fb,N,fmt); put(oe,f'D{r}',fa,B,fmt)
    put(oe,f'E{r}',f'=IF(D{r}="","",IF(D{r}>=Controls!$D${cr},"KNOCKOUT",IF(D{r}<=Controls!$B${cr},"GREEN",IF(D{r}>=Controls!$C${cr},"RED","YELLOW"))))',B)
    put(oe,f'F{r}',f'=TEXT(Controls!$B${cr},"{fmt}")&" / "&TEXT(Controls!$C${cr},"{fmt}")&" / "&TEXT(Controls!$D${cr},"{fmt}")',NOTE,align=LEFT)
for r,t,cr in [(30,'# positions (incl. new)',42),(31,'Revenue trend (avg MoM)',43),(32,'Neg days + NSFs (total)',44),(33,'Credit score (FICO)',45)]:
    label(oe,f'B{r}',t); put(oe,f'C{r}',f'={D}$C${cr}',N); put(oe,f'D{r}',f'={D}$C${cr}',N); put(oe,f'E{r}',f'={D}$F${cr}',B); note(oe,f'F{r}','merchant-side metric - unchanged by the offer')
oe['C31'].number_format=oe['D31'].number_format='0.0%'
rag_rules(oe,'E26:E33')
hdr(oe,'B35','SAFETY PROOF','B35:F35')
prf=[(36,'Adjusted fund <= cash-flow max fund','=IF(OR(D8="",D8="-",C8="",C8="-"),"OK",IF(D8<=C8,"OK","FAIL"))'),
     (37,'Adjusted daily payment <= cash-flow max daily','=IF(OR(D9="",C5=""),"OK",IF(D9<=C5+0.005,"OK","FAIL"))'),
     (38,'No offer-side metric RED or KNOCKOUT at adjusted offer','=IF(COUNTIF(E26:E29,"RED")+COUNTIF(E26:E29,"KNOCKOUT")=0,"OK","FAIL")'),
     (39,'Merchant knockouts untouched (history cannot remove them)',f'=IF(COUNTIF({D}$F$42:$F$45,"KNOCKOUT")>0,IF(D8="-","OK","FAIL"),"OK")'),
     (40,'Red flags at adjusted offer (offer-side + merchant-side)','=COUNTIF(E26:E33,"RED")+COUNTIF(E26:E33,"KNOCKOUT")')]
for r,t,f in prf: label(oe,f'B{r}',t); put(oe,f'D{r}',f,B)
chk_rules(oe,'D36:D39')
oe.sheet_view.showGridLines=False

# ======================= Decision Summary =======================
ds=wb.create_sheet('Decision Summary'); widths(ds,{'A':3,'B':44,'C':26,'D':95})
ds['B1']='DECISION SUMMARY'; ds['B1'].font=Font(name='Calibri',size=14,bold=True)
note(ds,'B2','Everything an underwriter needs to sign or override. Cash flow sets the ceiling (Layer 1); house-book evidence sets how much of it to use (Layers 2-3).')
sub(ds,4,None,['','Item','Value','Detail'],start=1)
HSr="'Historical Score'!"
S=[(5,'DECISION','=IF(Deal!$C$24=0,"",IF(OR(Deal!$C$63="-",Deal!$C$63="",'+OE+'$D$8="",'+OE+'$D$8=0,'+OE+'$D$8<Controls!$B$14,'+OE+'$D$40>=Controls!$B$13),"DECLINE",IF(C28<>"","MANUAL REVIEW",IF(OR(COUNTIF('+OE+'$E$26:$E$33,"YELLOW")+COUNTIF('+OE+'$E$26:$E$33,"RED")>0,'+OE+'$C$10>=4,C31>0),"CONDITIONAL","APPROVE"))))',
   '=IF(C5="","Enter bank data + deal profile.",IF(C5="DECLINE",IF(COUNTIF(Deal!$F$42:$F$45,"KNOCKOUT")>0,"Merchant knockout: "&INDEX(Deal!$B$42:$B$45,MATCH("KNOCKOUT",Deal!$F$42:$F$45,0))&".",IF(OR('+OE+'$D$8="",'+OE+'$D$8=0),"No clean room under the green lines.",IF('+OE+'$D$8<Controls!$B$14,"Adjusted fund "&TEXT('+OE+'$D$8,"$#,##0")&" is below the minimum "&TEXT(Controls!$B$14,"$#,##0")&".","Red flags at the adjusted offer reached the limit."))),IF(C5="MANUAL REVIEW","Triggers: "&C28,IF(C5="APPROVE","All metrics green at the adjusted offer; evidence neutral or better; standard stips only.","Yellows/reds remain, evidence is negative, or conditional stips apply - see below."))))',None),
   (6,'Recommended funding amount',"='Offer Engine'!$D$8",'=IF(OR(C6="",C6="-"),"",IF(ISNUMBER(Deal!$C$63),"cash-flow max "&TEXT(Deal!$C$63,"$#,##0")&"  ->  "&TEXT(C6/Deal!$C$63,"0%")&" used  (requested "&IF(Deal!$B$6="","n/a",TEXT(Deal!$B$6,"$#,##0"))&")",""))','"$"#,##0'),
   (7,'Recommended factor rate',"='Offer Engine'!$D$7",'=IF(C7="","","base "&TEXT('+OE+'$C$7,"0.00")&IF('+OE+'$D$7>'+OE+'$C$7," -> raised one approved tier by history",""))','0.00'),
   (8,'Recommended term',"='Offer Engine'!$D$17",'=IF(C8="","",IF('+OE+'$D$16="Weekly","weeks","daily payments")&"  ("&TEXT('+OE+'$D$22,"0.0")&" months; base "&'+OE+'$C$6&" daily pmts"&IF('+OE+'$D$6<'+OE+'$C$6,", shortened by history","")&")")','0'),
   (9,'Payment',"='Offer Engine'!$D$18",'=IF(C9="","","per "&IF('+OE+'$D$16="Weekly","week","business day")&";  daily-equivalent "&TEXT('+OE+'$D$9,"$#,##0.00")&", monthly "&TEXT('+OE+'$D$19,"$#,##0")&";  payback "&TEXT('+OE+'$D$20,"$#,##0")&", net "&TEXT('+OE+'$D$21,"$#,##0"))','"$"#,##0.00'),
   (10,'Daily or weekly',"='Offer Engine'!$D$16",'=IF('+OE+'$C$14="Weekly","Weekly required by the historical modifier","As proposed on the Deal sheet")',None),
   (11,'Holdback %',"='Offer Engine'!$D$27",'=IF(C11="","",'+OE+'$E$27&"  (green line "&TEXT(Controls!$B$6,"0%")&")")','0.0%'),
   (12,'Total leverage (all debt / revenue)',"='Offer Engine'!$D$26",'=IF(C12="","",'+OE+'$E$26&"  (green line "&TEXT(Controls!$B$5,"0%")&")")','0.0%'),
   (13,'Funding-to-revenue',"='Offer Engine'!$D$29",'=IF(C13="","",'+OE+'$E$29&"  (green line "&TEXT(Controls!$B$8,"0.00")&"x)")','0.00'),
   (14,'Daily debt-to-balance',"='Offer Engine'!$D$28",'=IF(C14="","",'+OE+'$E$28&"  (green line "&TEXT(Controls!$B$7,"0%")&")")','0.0%'),
   (15,'Number of positions (incl. new)','=Deal!$C$42','=IF(C15="","",Deal!$F$42&IF(Deal!$K$18<>"","  (position override "&Deal!$K$18&" used for matching)",""))','0'),
   (16,'Historical risk modifier',"='Historical Score'!$C$43",'=IF(C16="","","relative index "&TEXT('+HSr+'$C$39,"0.00")&" (1.00 = seasoned book); fund x "&TEXT('+OE+'$D$11,"0.00")&", term step "&'+OE+'$D$12&", factor step "&'+OE+'$D$13&IF('+HSr+'$C$42<>'+HSr+'$C$40,"; capped by Low confidence",""))',None),
   (17,'Cash-flow risk score','=Deal!$C$55','=IF(C17="","","of "&Deal!$D$55&" scored metrics = "&Deal!$C$56&"  (scaled "&TEXT(Deal!$D$57,"0.0")&"/7 for pricing bands)")','0.0'),
   (18,'Overall confidence in the historical evidence',"='Historical Score'!$C$41",'="Z = "&TEXT('+HSr+'$C$37,"0.00")&"; open share "&IF('+HSr+'$C$34="","n/a",TEXT('+HSr+'$C$34,"0%"))&"; "&IF('+HSr+'$C$29=7,"no specific cohort met the minimums","primary tier "&'+HSr+'$C$29)',None),
   (19,'Number of comparable deals (seasoned / all matches)',"='Historical Score'!$C$31","='Historical Score'!$C$30",None),
   (20,'Percentage of comparable deals seasoned',"='Historical Score'!$C$32",'="unseasoned deals are listed on Comparable Deals but never rated"','0.0%'),
   (21,'Comparable-cohort principal lost rate',"='Historical Score'!$C$33",'=IF(C21="","too few seasoned deals to show a rate","vs seasoned book "&TEXT('+HSr+'$C$10,"0.0%")&";  blended (credibility-weighted) "&TEXT('+HSr+'$C$38,"0.0%")&";  collected on defaults "&IF('+HSr+'$C$36="","n/a",TEXT('+HSr+'$C$36,"0%")))','0.0%'),
   (22,'Comparable-cohort return on funded capital',"='Historical Score'!$C$35",'=IF(C22="","","seasoned deals only; open deals in the cohort depress this figure")','0.0%'),
   (23,'Largest historical impact',"='Historical Score'!$C$44","='Historical Score'!$C$45",None),
   (24,'Metric that limited the offer','=IF(Deal!$D$60="","",Deal!$D$60)','=IF(Deal!$C$60="","",IF(Deal!$C$60=0,"existing debt already exceeds a green line",IF(Deal!$C$62=FLOOR(Deal!$D$41*Deal!$C$24*Deal!$D$62/Deal!$C$60,5),"term also capped by the fund/revenue green line","term set by the score band")))',None),
   (25,'ISO / industry / state / position evidence',"='Historical Score'!$C$46",'='+HSr+'$C$47&" | "&'+HSr+'$C$48&" | "&'+HSr+'$C$49&IF('+HSr+'$C$50="No",""," | CONFLICT: "&'+HSr+'$C$50)',None),
   (26,'Required stipulations','=C27&IF(D26="","","; "&D26)','=IF(AND(Controls!$B$91="Y",'+OE+'$C$10>=4),"No-stacking covenant with default-to-daily; ","")&IF(AND(Controls!$B$92="Y",Deal!$F$44<>"GREEN",Deal!$F$44<>""),"Bank login verification (Plaid/DecisionLogic); ","")&IF(AND(Controls!$B$93="Y",Deal!$C$34>=2),"Payoff/balance letter for each existing position; ","")&IF(AND(Controls!$B$94="Y",Deal!$F$45<>"GREEN",Deal!$F$45<>""),"Landlord/lease letter; ","")&IF(AND(Controls!$B$95="Y",'+HSr+'$C$41="Low"),"Underwriter review of historical evidence; ","")&IF(AND(Controls!$B$96="Y",'+OE+'$C$14="Weekly"),"Weekly remittance only; ","")&IF(AND(Controls!$B$97="Y",Deal!$K$15="Renewal"),"Prior-contract payoff and renewal history; ","")',None),
   (27,'Standard stipulations','=IF(Controls!$B$90="Y",Controls!$A$90,"")','="always required"',None),
   (28,'Manual-review triggers fired','=IF(AND(Controls!$B$79="Y",'+OE+'$C$15="Y"),"Severe historical modifier; ","")&IF(AND(Controls!$B$80="Y",'+HSr+'$C$41="Low",COUNTIF('+OE+'$E$26:$E$33,"RED")>0),"Low confidence with a RED metric; ","")&IF(AND(Controls!$B$81="Y",Deal!$F$45="RED",COUNTIF('+OE+'$E$26:$E$32,"RED")>0),"Credit RED plus another RED; ","")&IF(AND(Controls!$B$82="Y",ISNUMBER(Deal!$C$42),IF(ISNUMBER(Deal!$C$42),Deal!$C$42,0)>=Controls!$C$82),"Positions incl. new at/above "&Controls!$C$82&"; ","")&IF(AND(Controls!$B$83="Y",Deal!$K$15="New Deal",ISNUMBER('+HSr+'$C$6),IF(ISNUMBER('+HSr+'$C$6),'+HSr+'$C$6,0)>=Controls!$C$83,OR(Deal!$K$13="General Building Contractors",Deal!$K$13="Special Trade Contractors",Deal!$K$13="Heavy Construction")),"New 3rd+ position construction deal; ","")&IF(AND(Controls!$B$84="Y",OR(AND(Deal!$K$12<>"",COUNTIF('+hr("K")+',Deal!$K$12)=0),AND(Deal!$K$13<>"",COUNTIF('+hr("Y")+',Deal!$K$13)=0))),"ISO or industry not in house book; ","")&IF(AND(Controls!$B$85="Y",ISNUMBER('+OE+'$D$8),IF(ISNUMBER('+OE+'$D$8),'+OE+'$D$8,0)<Controls!$B$14,IF(ISNUMBER('+OE+'$D$8),'+OE+'$D$8,0)>0),"Adjusted fund below minimum; ","")&IF(AND(Controls!$B$86="Y",ISNUMBER('+HSr+'$C$34),IF(ISNUMBER('+HSr+'$C$34),'+HSr+'$C$34,0)>Controls!$B$43),"Cohort open share high; ","")','="blank = none"',None),
   (29,'Main risks / positive factors','=IF(COUNTIF(Deal!$F$38:$F$45,"RED")+COUNTIF(Deal!$F$38:$F$45,"KNOCKOUT")=0,"Risks: none red on the requested offer","Risks: "&IF(Deal!$F$38="RED","total leverage; ","")&IF(Deal!$F$39="RED","holdback; ","")&IF(Deal!$F$40="RED","daily debt/balance; ","")&IF(Deal!$F$41="RED","fund/revenue; ","")&IF(Deal!$F$42="RED","positions; ","")&IF(Deal!$F$43="RED","revenue trend; ","")&IF(Deal!$F$44="RED","neg days/NSFs; ","")&IF(Deal!$F$45="RED","credit; ","")&IF(COUNTIF(Deal!$F$38:$F$45,"KNOCKOUT")>0,"KNOCKOUT present",""))&IF('+OE+'$C$10>=4," | history: "&'+HSr+'$C$43,"")','="Positives: "&IF(COUNTIF(Deal!$F$38:$F$45,"GREEN")=0,"none green",IF(Deal!$F$38="GREEN","leverage; ","")&IF(Deal!$F$39="GREEN","holdback; ","")&IF(Deal!$F$40="GREEN","daily debt/balance; ","")&IF(Deal!$F$41="GREEN","fund/revenue; ","")&IF(Deal!$F$42="GREEN","positions; ","")&IF(Deal!$F$43="GREEN","revenue trend; ","")&IF(Deal!$F$44="GREEN","clean neg days/NSFs; ","")&IF(Deal!$F$45="GREEN","credit; ",""))&IF('+OE+'$C$10<=2,"history: "&'+HSr+'$C$43,"")&IF(Deal!$K$15="Renewal"," | renewal (strongest surviving positive in the book)","")',None),
   (31,'Risk-based conditional stips (drive CONDITIONAL)','=IF(AND(Controls!$B$91="Y",'+OE+'$C$10>=4),1,0)+IF(AND(Controls!$B$92="Y",Deal!$F$44<>"GREEN",Deal!$F$44<>""),1,0)+IF(AND(Controls!$B$94="Y",Deal!$F$45<>"GREEN",Deal!$F$45<>""),1,0)+IF(AND(Controls!$B$95="Y",'+HSr+'$C$41="Low"),1,0)+IF(AND(Controls!$B$96="Y",'+OE+'$C$14="Weekly"),1,0)','="documentary stips (payoff letters, renewal history) do not by themselves make a deal CONDITIONAL"','0'),
   (30,"Underwriter's conclusion",'=IF(Deal!$C$24=0,"Enter bank data.",IF(Deal!$C$63="-","Do not fund: "&INDEX(Deal!$B$42:$B$45,MATCH("KNOCKOUT",Deal!$F$42:$F$45,0))&" is past its knockout line; history cannot override a knockout.",IF(OR(Deal!$C$63="",Deal!$C$63=0),"Cash flow leaves no clean room under the green lines - "&Deal!$D$60&".","Cash flow supports up to "&TEXT(Deal!$C$63,"$#,##0")&" ("&Deal!$D$60&", "&Deal!$C$62&" daily pmts @ "&TEXT(Deal!$D$62,"0.00")&"). "&IF(C5="DECLINE","The deal is declined: "&D5&" For reference, ","The recommendation is "&TEXT(C6,"$#,##0")&" ("&TEXT(C6/Deal!$C$63,"0%")&" of the maximum) at "&TEXT(C7,"0.00")&" over "&TEXT(C8,"0")&IF(C10="Weekly"," weeks"," daily payments")&", "&TEXT(C9,"$#,##0")&IF(C10="Weekly","/week","/day")&", because ")&"the primary comparable cohort ("&'+HSr+'$C$30&") holds "&'+HSr+'$C$31&" seasoned deals with a principal lost rate of "&IF('+HSr+'$C$33="","n/a",TEXT('+HSr+'$C$33,"0.0%"))&" against "&TEXT('+HSr+'$C$10,"0.0%")&" for the seasoned book; at credibility "&TEXT('+HSr+'$C$37,"0.00")&" that blends to "&TEXT('+HSr+'$C$38,"0.0%")&" (index "&TEXT('+HSr+'$C$39,"0.00")&"), which is "&LOWER('+HSr+'$C$43)&IF('+HSr+'$C$42<>'+HSr+'$C$40," after the low-confidence cap","")&". Confidence is "&LOWER('+HSr+'$C$41)&". "&IF(Deal!$K$15="Renewal","Renewal status is the strongest surviving positive in the book. ","")&'+HSr+'$C$46&". "&'+HSr+'$C$47&". "&'+HSr+'$C$48&". "&IF(C28<>"","Sent to manual review: "&C28,"")&IF(D26<>""," Conditional stips: "&D26,""))))',None,None)]
for item in S:
    r,t,v,d=item[0],item[1],item[2],item[3]; fmt=item[4] if len(item)>4 else None
    label(ds,f'B{r}',t,B); put(ds,f'C{r}',v,B if r in (5,6) else N,fmt,align=LEFT if r>=23 else CEN)
    if d is not None: c=ds[f'D{r}']; c.value=d; c.font=NOTE; c.alignment=WRAP; c.border=BOX
ds['C5'].font=Font(name='Calibri',size=13,bold=True); ds['C6'].font=Font(name='Calibri',size=13,bold=True)
ds['C30'].alignment=WRAP; ds.merge_cells('C30:D30'); ds.row_dimensions[30].height=150; ds['C31'].alignment=CEN
for r in (25,26,28,29): ds.row_dimensions[r].height=45; ds[f'C{r}'].alignment=WRAP
ds.row_dimensions[5].height=32
dec_rules(ds,'C5'); ds.sheet_view.showGridLines=False

# ======================= Comparable Deals =======================
cd=wb.create_sheet('Comparable Deals'); 
cd['A1']='COMPARABLE DEALS  -  every house-book deal in the primary cohort (seasoned and unseasoned), inspect by Contract ID'; cd['A1'].font=Font(name='Calibri',size=14,bold=True)
cd['A2']="=\"Primary cohort: \"&'Historical Score'!$C$30&\"   |   \"&'Historical Score'!$C$31&\" seasoned / all   |   only the first 400 are listed\""; cd['A2'].font=B
note(cd,'A3','Similarity tier 1 = closest (matches NR + industry + position). Deals shown with Seasoned = 0 are counted but excluded from every rate.')
ch=['#','Contract ID','DBA','Status','Funded','Position','Industry','ISO','State','New/renewal','Factor','Advance','Collected','% paid','Net cash','Principal lost','Term (bd)','Age (bd)','Days on book','Seasoned','Default','Open','Similarity tier','Why included']
sub(cd,4,None,ch); cd.row_dimensions[4].height=30
src={'Contract ID':'A','DBA':'B','Status':'D','Funded':'E','Position':'G','Industry':'Y','ISO':'K','State':'J','New/renewal':'L','Factor':'M','Advance':'N','Collected':'Q','% paid':'S','Net cash':'AJ','Principal lost':'AK','Term (bd)':'AA','Age (bd)':'AB','Days on book':'AN','Seasoned':'AC','Default':'AE','Open':'AH','Similarity tier':'BC'}
for k in range(1,401):
    r=4+k; cd.cell(r,1,k).font=NOTE
    for j,h in enumerate(ch[1:-1]):
        col=src[h]; c=cd.cell(r,2+j,f"=IFERROR(INDEX({hr(col)},MATCH($A{r},{hr('BE')},0)),\"\")"); c.font=N
        if h in ('Advance','Collected','Net cash','Principal lost'): c.number_format='"$"#,##0'
        if h=='Funded': c.number_format='mm/dd/yyyy'
        if h=='% paid': c.number_format='0%'
        if h=='Factor': c.number_format='0.00'
    c=cd.cell(r,len(ch),f'=IF(B{r}="","",IF(W{r}=1,"NR + industry + position",IF(W{r}=2,"NR + ISO + position",IF(W{r}=3,"NR + industry",IF(W{r}=4,"NR + ISO",IF(W{r}=5,"NR + position",IF(W{r}=6,"NR only","book baseline"))))))&"  ["&IF(T{r}=1,"seasoned","unseasoned")&IF(U{r}=1,", defaulted","")&IF(V{r}=1,", open","")&"]")'); c.font=NOTE
widths(cd,{'A':5,'B':12,'C':30,'D':18,'E':11,'F':8,'G':26,'H':24,'I':6,'J':10,'K':7,'L':11,'M':11,'N':7,'O':11,'P':11,'Q':8,'R':8,'S':8,'T':8,'U':8,'V':7,'W':9,'X':40})
cd.freeze_panes='C5'

# ======================= Checks =======================
ck=wb.create_sheet('Checks'); widths(ck,{'A':3,'B':64,'C':16,'D':70})
ck['B1']='CHECKS / AUDIT'; ck['B1'].font=Font(name='Calibri',size=14,bold=True)
sub(ck,3,None,['','Check','Result','Detail'],start=1)
HDr=hr
checks=[('MISSING INPUTS',None,None),
 ('Bank data entered (at least one month of true revenue)','=IF(Deal!$C$24=0,"WARN","OK")','="months entered: "&COUNT(Deal!$C$12:$C$23)'),
 ('Average daily balance entered','=IF(Deal!$D$24=0,"WARN","OK")',None),
 ('New deal or renewal entered (required for any cohort tighter than the book)','=IF(Deal!$K$15="","WARN","OK")',None),
 ('ISO entered and found in the house book','=IF(Deal!$K$12="","WARN",IF(COUNTIF('+HDr("K")+',Deal!$K$12)=0,"WARN","OK"))','=IF(Deal!$K$12="","",COUNTIF('+HDr("K")+',Deal!$K$12)&" deals")'),
 ('Industry entered and found in the house book','=IF(Deal!$K$13="","WARN",IF(COUNTIF('+HDr("Y")+',Deal!$K$13)=0,"WARN","OK"))','=IF(Deal!$K$13="","",COUNTIF('+HDr("Y")+',Deal!$K$13)&" deals")'),
 ('State entered','=IF(Deal!$K$14="","WARN","OK")','="informational only"'),
 ('Credit score entered','=IF(Deal!$K$16="","WARN","OK")','="blank = metric not scored (score out of 7)"'),
 ('Requested offer entered (fund, factor, term, frequency)','=IF(OR(Deal!$B$6="",Deal!$C$6="",Deal!$D$6="",Deal!$E$6=""),"WARN","OK")','="optional - the engine still sizes an offer from bank data"'),
 ('Requested factor is an approved tier','=IF(Deal!$C$6="","OK",IF(ISNUMBER(MATCH(Deal!$C$6,Controls!$B$24:$D$24,0)),"OK","WARN"))','="1.40 / 1.45 / 1.49 only"'),
 ('FORMULA ERRORS',None,None),
 ('Deal sheet','=IF(SUMPRODUCT(--ISERROR(Deal!$B$1:$M$82))=0,"OK","FAIL")','=SUMPRODUCT(--ISERROR(Deal!$B$1:$M$82))&" error cells"'),
 ('Historical Data','=IF(SUMPRODUCT(--ISERROR('+HDr("A")+'))+SUMPRODUCT(--ISERROR('+HDr("AA")+'))+SUMPRODUCT(--ISERROR('+HDr("BC")+'))+SUMPRODUCT(--ISERROR('+HDr("BE")+'))=0,"OK","FAIL")',None),
 ('Historical Score','=IF(SUMPRODUCT(--ISERROR(\'Historical Score\'!$A$1:$AA$52))=0,"OK","FAIL")',None),
 ('Offer Engine / Decision Summary','=IF(SUMPRODUCT(--ISERROR(\'Offer Engine\'!$B$1:$F$40))+SUMPRODUCT(--ISERROR(\'Decision Summary\'!$B$1:$D$31))=0,"OK","FAIL")',None),
 ('SAFETY (history may never override cash flow)',None,None),
 ('Adjusted fund <= cash-flow max fund',"='Offer Engine'!$D$36",None),
 ('Adjusted daily payment <= cash-flow max daily',"='Offer Engine'!$D$37",None),
 ('No offer-side metric red/knockout at the adjusted offer',"='Offer Engine'!$D$38",None),
 ('Merchant knockouts untouched',"='Offer Engine'!$D$39",None),
 ('COHORT QUALITY',None,None),
 ('Primary cohort met the minimum seasoned count and dollars',"=IF('Historical Score'!$C$29=7,\"WARN\",\"OK\")","=IF('Historical Score'!$C$29=7,\"book baseline in use - no specific evidence\",'Historical Score'!$C$30)"),
 ('Primary cohort seasoned-enough (open share within limit)',"=IF(INDEX('Historical Score'!$U$14:$U$20,'Historical Score'!$C$29)=\"Yes\",\"OK\",\"WARN\")","=\"open share \"&IF('Historical Score'!$C$34=\"\",\"n/a\",TEXT('Historical Score'!$C$34,\"0%\"))"),
 ('Credibility Z at least the Medium line',"=IF('Historical Score'!$C$37>=Controls!$B$45,\"OK\",\"WARN\")","=\"Z = \"&TEXT('Historical Score'!$C$37,\"0.00\")"),
 ('Conflicting ISO vs industry signals',"=IF('Historical Score'!$C$50=\"No\",\"OK\",\"WARN\")","='Historical Score'!$C$50"),
 ('Tighter cohorts skipped for being too small (tiers above the primary)',"=IF('Historical Score'!$C$29=1,\"OK\",\"WARN\")","=IF('Historical Score'!$C$29=1,\"\",\"tiers 1 to \"&('Historical Score'!$C$29-1)&\" had fewer than \"&Controls!$B$40&\" seasoned deals or under \"&TEXT(Controls!$B$41,\"$#,##0\"))"),
 ('State cohort is informational only','="OK"',"='Historical Score'!$C$48"),
 ('DATA QUALITY (house book)',None,None),
 ('Deals loaded','=IF(COUNTA('+HDr("A")+')=1067,"OK","WARN")','=COUNTA('+HDr("A")+')&" deals (expected 1,067 from the 08/2026 export)"'),
 ('Statuses not in the Controls mapping','=IF(SUMPRODUCT(1-'+HDr("AI")+')=0,"OK","FAIL")','=SUMPRODUCT(1-'+HDr("AI")+')&" deals with an unmapped status"'),
 ('Deals with no revenue figure','=IF(SUMPRODUCT(1-'+HDr("AM")+')>0,"WARN","OK")','=SUMPRODUCT(1-'+HDr("AM")+')&" deals - excluded from advance/revenue averages"'),
 ('Deals with no ISO / no state / no SIC','="WARN"','=COUNTBLANK('+HDr("K")+')&" no ISO, "&(COUNTBLANK('+HDr("J")+')+COUNTIF('+HDr("J")+',"No state"))&" no state, "&COUNTIF('+HDr("Y")+',"No SIC")&" no SIC - never matched"'),
 ('Never-paid deals (collected <= 0)','="WARN"','=COUNTIF('+HDr("Q")+',"<=0")&" deals"'),
 ('Open status but 99.9%+ paid (stale status)','=IF(SUMPRODUCT('+HDr("AH")+'*('+HDr("S")+'>=0.999))=0,"OK","WARN")','=SUMPRODUCT('+HDr("AH")+'*('+HDr("S")+'>=0.999))&" deals"'),
 ('Open deals older than 1.5x contracted term','="WARN"','=SUMPRODUCT('+HDr("AH")+'*('+HDr("AB")+'>1.5*'+HDr("AA")+'))&" deals"'),
 ('Implausible durations: weekly term above 39 weeks (9-month box)','="WARN"','=SUMPRODUCT(('+HDr("T")+'>=12)*('+HDr("T")+'>39))&" deals, incl. 56 / 90 / 100 weeks"'),
 ('Position field concentration (49% of the book is position 3)','="WARN"','=COUNTIF('+HDr("G")+',3)&" deals at position 3 of "&COUNT('+HDr("G")+')&" - confirm the CRM field is position at funding"'),
 ('Lowered Payments treated as performing (Controls B68)','=IF(Controls!$B$68="No","WARN","OK")','=COUNTIF('+HDr("D")+',"Lowered Payments")&" restructured deals"'),
 ('Written-Off deals collected most of their payback','="WARN"','="avg % paid "&TEXT(SUMPRODUCT(('+HDr("D")+'="Written-Off")*'+HDr("S")+')/MAX(1,COUNTIF('+HDr("D")+',"Written-Off")),"0%")&" - status means stopped pursuing, not lost"'),
 ('RECONCILIATION to the original House Book Analytics workbook (full file)',None,None),
 ('Total funded = $34,481,626','=IF(ABS(SUM('+HDr("N")+')-34481626)<1,"OK","FAIL")','=TEXT(SUM('+HDr("N")+'),"$#,##0")'),
 ('Total collected = $27,438,510','=IF(ABS(SUM('+HDr("Q")+')-27438510)<1,"OK","FAIL")','=TEXT(SUM('+HDr("Q")+'),"$#,##0")'),
 ('Principal lost (house-book status mapping) = $5,270,163','=IF(ABS(SUMPRODUCT('+HDr("AK")+')-5270163)<1,"OK","WARN")','=TEXT(SUMPRODUCT('+HDr("AK")+'),"$#,##0")&"  (differs if the status mapping was changed)"'),
 ('Principal lost rate (full file) = 15.3%','=IF(ABS(SUMPRODUCT('+HDr("AK")+')/SUM('+HDr("N")+')-0.1528)<0.001,"OK","WARN")','=TEXT(SUMPRODUCT('+HDr("AK")+')/SUM('+HDr("N")+'),"0.0%")'),
 ('Payback = advance x factor on every deal','=IF(SUMPRODUCT(--(ABS('+HDr("P")+'-'+HDr("N")+'*'+HDr("M")+')>1))=0,"OK","WARN")','=SUMPRODUCT(--(ABS('+HDr("P")+'-'+HDr("N")+'*'+HDr("M")+')>1))&" mismatches"'),
 ('Seasoned book size (637 deals at as-of 08/25/2026, multiple 1.25, NETWORKDAYS inclusive)',"=IF('Historical Score'!$E$20=637,\"OK\",\"WARN\")","='Historical Score'!$E$20&\" seasoned deals; changes with the as-of date, window or multiple\""),
 ('Sum of tier-7 cohort funded equals in-window seasoned funded',"=IF(ABS('Historical Score'!$G$20-SUMPRODUCT("+HDr("AP")+"*"+HDr("N")+"))<1,\"OK\",\"FAIL\")",None),
 ('Comparable list count equals primary cohort matches (first 400)',"=IF(MIN(400,INDEX('Historical Score'!$D$14:$D$20,'Historical Score'!$C$29))=SUMPRODUCT(--('Comparable Deals'!$B$5:$B$404<>\"\")),\"OK\",\"FAIL\")","=SUMPRODUCT(--('Comparable Deals'!$B$5:$B$404<>\"\"))&\" listed\"")]
r=4
for t,f,d in checks:
    if f is None: hdr(ck,f'B{r}',t,f'B{r}:D{r}')
    else:
        label(ck,f'B{r}',t); put(ck,f'C{r}',f,B)
        if d: c=ck[f'D{r}']; c.value=d; c.font=NOTE; c.alignment=LEFT; c.border=BOX
    r+=1
chk_rules(ck,f'C4:C{r}')
ck['B2']='=COUNTIF(C4:C'+str(r)+',"FAIL")&" FAIL  |  "&COUNTIF(C4:C'+str(r)+',"WARN")&" WARN  |  "&COUNTIF(C4:C'+str(r)+',"OK")&" OK"'; ck['B2'].font=B
ck.sheet_view.showGridLines=False

# ======================= README =======================
rd=wb.create_sheet('README',0); widths(rd,{'A':3,'B':120})
txt=["ASPIRE COMBINED UNDERWRITING MODEL  -  cash-flow engine + house-book evidence  (v1)",
"",
"HOW TO RUN A DEAL (everything happens on the Deal sheet)",
"1. Section A: requested fund, factor, term, frequency (optional - the engine sizes an offer from bank data even if blank).",
"2. Section B: up to 12 months of true revenue, average daily balance, negative days, NSFs - newest month first.",
"3. Existing positions: funder, funded $, payment, frequency, start date. Monthly-frequency debt = leverage only, not a position.",
"4. DEAL PROFILE (right of the bank data): ISO, industry (SIC major group), state, new deal or renewal, credit score. Position is automatic; override only if you know better.",
"5. Read Section F at the bottom: final decision, fund, factor, term, payment, frequency, modifier, comparables, confidence, stipulations and the written conclusion.",
"6. Decision Summary has the full 20-item output; Comparable Deals lists every house-book deal behind it by Contract ID; Historical Score shows all seven tiers and the credibility math.",
"",
"THE THREE LAYERS",
"Layer 1 - Deal sheet sections A-E: the original cash-flow tool. Seven metrics plus credit score, green/yellow/red/knockout, score, cash-flow maximum daily payment (weakest of holdback, total leverage, daily-debt-to-balance green lines), term band, factor band, cash-flow max fund. Thresholds live on Controls.",
"Layer 2 - Historical Data + Historical Score: 1,067 house-book deals, seasoned = age >= 1.25 x contracted term at the as-of date. Seven matching tiers (new/renewal + industry + position ... down to the seasoned book). Primary cohort = first tier with >= 10 seasoned deals and >= $150k seasoned funded. Credibility Z = n/(n+30). Blended lost rate = Z x cohort + (1-Z) x book. Relative index = blended / book -> modifier level (Controls table). Confidence High / Medium / Low; Low caps the modifier at one step from neutral.",
"Layer 3 - Offer Engine + Decision Summary: adjusted fund = MIN(cash-flow max x multiplier, cash-flow max daily x adjusted term / adjusted factor). Term shortens by bands, factor rises by approved tiers, frequency may switch to weekly. All four offer-side metrics are recomputed and must stay off red. Decision = DECLINE / MANUAL REVIEW / CONDITIONAL / APPROVE.",
"",
"WHAT HISTORY CAN AND CANNOT DO",
"Can: reduce funding, shorten term, raise factor within 1.40/1.45/1.49, require weekly, add stipulations, send to manual review, or allow the full cash-flow maximum when the evidence is strong and credible.",
"Cannot: remove a knockout, push a green metric to red, exceed the cash-flow maximum, or make a precise adjustment on a tiny or unseasoned cohort (Low confidence caps it and adds a stip). Checks proves each of these on every deal.",
"",
"WHAT THE HOUSE BOOK ACTUALLY SUPPORTS (seasoned deals, after basic controls - see docs/01_Stage1_Audit_Findings.md)",
"Robust: renewal vs new (odds of default x0.43), construction trades (x1.8), Westwood ISO (x0.3). Composition, not the variable: United Secured Capital 8, position 3 (mostly new + 1.49 + construction mix). Informational only: state, factor, term, deal size, weekly vs daily, advance/revenue, holdback. Leakage - never used: NSF count (post-funding).",
"",
"MAINTENANCE",
"- Refresh: paste the CRM export into Historical Data columns A-X (values only, same column order), update the as-of date on Controls, check Checks.",
"- Archiving a deal: copy the Deal sheet as before. Sections A-E stay self-contained; Section F points at the shared sheets and always shows the deal currently on the sheet named Deal - paste values before filing.",
"- Never insert or delete rows on Controls; formulas reference fixed cells.",
"- Score scaling: with a credit score entered the scorecard has 8 metrics; the score is scaled to the 7-point bands (score x 7 / metrics scored) so the pricing bands and tier names keep their meaning. APPROVE still requires every scored metric green.",
"",
"GOOGLE SHEETS",
"Import: Google Drive > New > File upload > this .xlsx > open > File > Save as Google Sheets (or Sheets > File > Import > Upload > Replace spreadsheet). Every function used (SUMPRODUCT, NETWORKDAYS, INDEX/MATCH, COUNTIF, CHOOSE, FLOOR, TEXT, IFERROR, VLOOKUP) is native to Sheets. The first recalculation of the 45,000 formulas takes a few seconds; after that only the changed deal recalculates.",
"Dropdowns, colour rules, merged cells and number formats survive the import. Page setup does not matter in Sheets. If a dropdown shows no choices after import, re-point it at the Ref sheet lists (Data > Data validation).",
"",
"CONVERSIONS: money with 5 and 21.655 (weekly / 5 = daily; monthly / 21.655 = daily); time with 4.331 (weeks / 4.331 = months). MoneyBadger uses 4 and 21."]
for i,t in enumerate(txt):
    c=rd.cell(2+i,2,t); c.font=B if (t.isupper() or i==0) else N; c.alignment=Alignment(wrap_text=True,vertical='top')
rd.sheet_view.showGridLines=False

for shn in wb.sheetnames:
    sh=wb[shn]; sh.sheet_properties.pageSetUpPr.fitToPage=True; sh.page_setup.fitToWidth=1; sh.page_setup.fitToHeight=0
    sh.page_setup.orientation='landscape' if shn in ('Deal','Historical Score','Comparable Deals','Historical Data','Offer Engine') else 'portrait'
ws.print_area='B1:M82'; ws.page_setup.fitToHeight=1
wb.move_sheet('Deal',offset=-(wb.sheetnames.index('Deal')-1))
order=['README','Deal','Decision Summary','Offer Engine','Historical Score','Comparable Deals','Historical Data','Controls','Checks','Ref']
wb._sheets=[wb[n] for n in order]
wb.active=1
from openpyxl.workbook.properties import CalcProperties
wb.calculation=CalcProperties(fullCalcOnLoad=True)
wb.save(OUT); print('saved',OUT)
