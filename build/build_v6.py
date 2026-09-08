"""v6: seasoning = age OR terminal; ISO+Ind tier 1; dead columns cut; XLOOKUP impact; guards stripped; green-line side-by-side; Section E hidden; Section F restructured."""
import openpyxl, datetime as dt, copy, sys
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.properties import CalcProperties
ROOT='/home/user/Underwriting-Tool-House-Book'
BASE=f'{ROOT}/originals/v4_DEMO_user_edited.xlsx'; FULLBOOK=f'{ROOT}/originals/House Book Analytics (2).xlsx'
OUT=sys.argv[1]; FULL=len(sys.argv)>2 and sys.argv[2]=='full'
wb=openpyxl.load_workbook(BASE)
ws=wb['Deal']; cs=wb['Controls']; sc=wb['Historical Score']; hs=wb['Historical Data']; cd=wb['Comparable Deals']
DARK=PatternFill('solid',fgColor='1F2937'); MID=PatternFill('solid',fgColor='4B5563'); INPUT=PatternFill('solid',fgColor='E5E7EB')
WHITE=Font(name='Calibri',size=11,bold=True,color='FFFFFF'); SUB=Font(name='Calibri',size=9,bold=True,color='FFFFFF')
B=Font(name='Calibri',size=11,bold=True); N=Font(name='Calibri',size=11); NOTE=Font(name='Calibri',size=9,color='374151'); INP=Font(name='Calibri',size=11,bold=True,color='000000')
thin=Side(style='thin',color='9CA3AF'); med=Side(style='medium',color='374151'); BOX=Border(left=thin,right=thin,top=thin,bottom=thin); IBOX=Border(left=med,right=med,top=med,bottom=med)
CEN=Alignment(horizontal='center',vertical='center'); LEFT=Alignment(horizontal='left',vertical='center')
def hdr(w,cell,text,span=None):
    w[cell]=text; w[cell].font=WHITE; w[cell].fill=DARK; w[cell].alignment=LEFT
    if span: w.merge_cells(span)
def sub(w,row,labels,start=1):
    for i,t in enumerate(labels):
        c=w.cell(row,start+i,t); c.font=SUB; c.fill=MID; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=BOX
def put(w,addr,val,font=N,fmt=None,align=CEN,border=BOX,fill=None):
    c=w[addr]; c.value=val; c.font=font; c.alignment=align; c.border=border
    if fmt: c.number_format=fmt
    if font is INP and fill is None: fill=INPUT
    if fill: c.fill=fill
    return c
def note(w,addr,text):
    c=w[addr]; c.value=text; c.font=NOTE; c.alignment=LEFT; return c
def label(w,addr,text,font=N):
    c=w[addr]; c.value=text; c.font=font; c.alignment=LEFT; c.border=BOX; return c
def clear(w,rng):
    for row in w[rng]:
        for c in row:
            if type(c).__name__!='MergedCell': c.value=None

# ---------- data rows ----------
if FULL:
    hv=openpyxl.load_workbook(FULLBOOK,data_only=True)
    raw=[list(r) for r in hv['Data'].iter_rows(min_row=2,max_row=1068,max_col=27,values_only=True) if r[0] is not None]
    NR=1+len(raw)
else:
    NR=9
HDR="'Historical Data'!"; HS="'Historical Score'!"; C='Controls!'
def hr(col): return f"{HDR}${col}$2:${col}${NR}"
hols_n=len([1 for r in range(2,60) if wb['Ref'].cell(r,11).value]); HOL=f'Ref!$K$2:$K${1+hols_n}'

# ======================= Controls: rebuild content, keep the sheet's widths / title =======================
for m in list(cs.merged_cells.ranges): cs.unmerge_cells(str(m))
clear(cs,'A2:I100')
for r in range(2,101): cs.row_dimensions[r].height=None
hdr(cs,'A3','1. PRICING BANDS','A3:D3'); sub(cs,4,['Score >=','Term cap (daily pmts)','Factor'])
for i,(s_,t,f) in enumerate([(0,100,1.49),(5,130,1.45),(6,150,1.45),(7,195,1.40)]):
    r=5+i; put(cs,f'A{r}',s_,INP,'0.0',border=IBOX); put(cs,f'B{r}',t,INP,'0',border=IBOX); put(cs,f'C{r}',f,INP,'0.00',border=IBOX)
label(cs,'A9','Max term (months)'); put(cs,'B9',9,INP,'0',border=IBOX)
hdr(cs,'A11','2. CONVERSION CONSTANTS','A11:D11')
for r,(t,v,fmt) in enumerate([('Weeks per month (ours)',4.331,'0.000'),('Business days per month (ours)',21.655,'0.000'),('MoneyBadger weeks per month',4,'0'),('MoneyBadger business days per month',21,'0'),('Calendar days per month',30.4375,'0.0000'),('Assumed factor on existing positions',1.45,'0.00')],start=12):
    label(cs,f'A{r}',t); put(cs,f'B{r}',v,INP,fmt,border=IBOX)
hdr(cs,'A19','3. HISTORICAL EVIDENCE SETTINGS','A19:D19')
for r,(t,v,fmt) in enumerate([('As-of date (date of the CRM export)',dt.datetime(2026,8,25),'mm/dd/yyyy'),('Window: funded on or after',dt.datetime(2024,11,1),'mm/dd/yyyy'),('Window: funded on or before',dt.datetime(2026,8,25),'mm/dd/yyyy'),('Seasoning multiple (age >= multiple x term)',1.25,'0.00'),('Minimum seasoned deals for a cohort to be primary',10 if FULL else 3,'0'),('Minimum seasoned funded $ for a cohort to be primary',150000 if FULL else 30000,'"$"#,##0'),('Seasoned deals for the sample-size indicator to be GREEN',30 if FULL else 5,'0')],start=20):
    label(cs,f'A{r}',t); put(cs,f'B{r}',v,INP,fmt,border=IBOX)
hdr(cs,'A28','4. HOUSE-BOOK OUTCOME   (index = primary cohort PLR / seasoned-book PLR)','A28:D28')
for r,(t,v,fmt) in enumerate([('SUPPORTIVE when index at or below',0.85,'0.00'),('NEGATIVE when index above',1.15,'0.00'),('Supportive: sizing flex (0 = green line, 1 = red line)',0.50,'0%'),('Negative: fund cut % (Section F, planned)',0.20,'0%')],start=29):
    label(cs,f'A{r}',t); put(cs,f'B{r}',v,INP,fmt,border=IBOX)
hdr(cs,'A34','5. STATUS','A34:F34'); sub(cs,35,['Status','Default?','Resolved?','Closed?','Open?','Deals in file'])
statuses=[('Closed','No','Yes','Yes','No'),('Open','No','No','No','Yes'),('Legal','Yes','No','No','No'),('In House Collection','Yes','No','No','No'),('Lowered Payments','No','No','No','No'),('Written-Off','Yes','Yes','No','No'),('Bankruptcy','Yes','Yes','No','No'),('Settlement Agreement','Yes','Yes','No','No'),('Closed/NSF','No','Yes','Yes','No'),('Closed/NSF Unpaid','No','Yes','Yes','No'),('At-Risk','Yes','No','No','No')]
for i,s_ in enumerate(statuses):
    r=36+i; label(cs,f'A{r}',s_[0])
    for j in range(4): put(cs,f'{L(2+j)}{r}',s_[j+1],INP,border=IBOX)
    put(cs,f'F{r}',f'=COUNTIF({hr("D")},A{r})',N,'0')
cs.column_dimensions['A'].width=52; cs.column_dimensions['B'].width=15
ST='Controls!$A$36:$E$46'

# ======================= Deal =======================
# thresholds typed on the sheet (as in the original tool)
thr={38:(0.30,0.45,0.55),39:(0.18,0.25,0.35),40:(0.18,0.30,0.35),41:(0.60,0.90,1.25),42:(2,5,6),43:(-0.05,-0.15,-0.30),44:(1,6,10),45:(650,630,0)}
for r,(g,rd,k) in thr.items():
    for col,v in (('D',g),('E',rd),('G',k)):
        c=ws[f'{col}{r}']; c.value=v; c.font=Font(name='Calibri',size=11)
ws['F6']=0.05; ws['F6'].font=Font(name='Calibri',size=12,bold=True)
ws['C61']=2; ws['E61']=10000
for a,v in (('K8',4.331),('K9',21.655),('L8',4),('L9',21)): ws[a]=v; ws[a].font=Font(name='Calibri',size=11)
ws['L15']='=IF(K15="","",IF(K15>=D45,"GREEN",IF(K15>=E45,"YELLOW","RED")))'
for r in range(28,34):
    ws[f'F{r}']=f'=IF(D{r}<>"",IF(E{r}="Daily",D{r},IF(E{r}="Weekly",D{r}/5,D{r}/$K$9)),IF(C{r}="","",C{r}*{C}$B$17/({C}$B$9*$K$9)))'
    ws[f'H{r}']=f'=IF(OR(C{r}="",D{r}=""),"",ROUND(C{r}*{C}$B$17/D{r},0))'
    ws[f'J{r}']=f'=IF(OR(I{r}="",G{r}=""),"",MAX(0,I{r}-(TODAY()-G{r})/{C}$B$16))'
SCORED='SUMPRODUCT(--(F38:F45<>""))'; SCALED=f'IF({SCORED}=0,0,C65*7/{SCORED})'
ws['C62']=f'=IF(C24=0,"",IF(OR(E60>0,IF(ISNUMBER(C73),C73,0)<E61),"DECLINE",IF(OR(B6="",B6=0),"",IF(C60>=C61,"DECLINE",IF(COUNTIF(F38:F45,"GREEN")={SCORED},"APPROVE","CONDITIONAL")))))'
ws['C66']=f'=IF(C24=0,"",IF({SCALED}>=6.5,"Strong",IF({SCALED}>=5,"Solid",IF({SCALED}>=3.5,"Moderate","Weak"))))'
ws['C72']=f'=IF(OR(C70="",C70=0,C65=""),"",MAX(20,MIN(INDEX({C}$B$5:$B$8,MATCH({SCALED},{C}$A$5:$A$8,1)),FLOOR({C}$B$9*$K$9,5),FLOOR(H41*C24*D72/C70,5))))'
ws['D72']=f'=IF(OR(C65="",C24=0),"",INDEX({C}$C$5:$C$8,MATCH({SCALED},{C}$A$5:$A$8,1)))'
for r in range(38,42): ws[f'H{r}']=f'=D{r}+{HS}$C$35*(E{r}-D{r})'
book=f'{HS}$C$10'
for r in range(51,56):
    ws[f'D{r}']=f'={book}*{C}$B$29'; ws[f'E{r}']=f'={book}*{C}$B$30'
ws['B51']=f'="Primary cohort: "&{HS}$C$29'; ws['C51']=f'={HS}$C$31'; ws['H51']=f'={HS}$C$32'
ws['F51']=f'=IF({HS}$D$29=8,"NO EVIDENCE",IF({HS}$C$33<={C}$B$29,"GREEN",IF({HS}$C$33>{C}$B$30,"RED","YELLOW")))'; ws['G51']=f'={HS}$C$30'
for r,hsrow in ((52,26),(53,25),(54,24),(55,23)):
    ws[f'C{r}']=f'={HS}$F${hsrow}'; ws[f'H{r}']=f'={HS}$G${hsrow}'
    ws[f'F{r}']=f'=IF({HS}$C${hsrow}=0,"NO DEALS",IF({HS}$F${hsrow}="","NO RATE",IF({HS}$H${hsrow}<={C}$B$29,"GREEN",IF({HS}$H${hsrow}>{C}$B$30,"RED","YELLOW"))))'
    ws[f'G{r}']=f'={HS}$C${hsrow}'
ws['C56']=f'=IF({HS}$D$29=8,0,{HS}$C$30)'; ws['D56']=f'={C}$B$26'; ws['E56']=f'={C}$B$24'
for a in ('D56','E56'): ws[a].font=Font(name='Calibri',size=11,italic=True,color='374151')

# ======================= Section E green-line column + Section F =======================
from openpyxl.formatting.rule import CellIsRule
G_FILL=PatternFill('solid',fgColor='00B050'); Y_FILL=PatternFill('solid',fgColor='FFFF00'); R_FILL=PatternFill('solid',fgColor='FF0000')
SMALL=Font(name='Calibri',size=9,color='374151')
SCORED='SUMPRODUCT(--(F38:F45<>""))'; SCALED=f'IF({SCORED}=0,0,C65*7/{SCORED})'
for m in list(ws.merged_cells.ranges):
    if 64<=m.min_row<=73: ws.unmerge_cells(str(m))
put(ws,'F64','At green lines',Font(name='Calibri',size=9,bold=True,color='FFFFFF'),fill=MID)
for r,f in {67:'=IF(C24=0,"",D39*C24/$K$9)',68:'=IF(C24=0,"",(D38*C24-F34)/$K$9)',69:'=IF(OR(C24=0,D24=0),"",D40*D24-E34)',70:'=IF(C24=0,"",MAX(0,MIN(F67:F69)))'}.items():
    put(ws,f'F{r}',f,N,'"$"#,##0.00')
put(ws,'G70','=IF(F70="","",IF(F70=0,"no clean room",INDEX({"SP% holdback";"Total LVG";"Daily/Balance"},MATCH(F70,F67:F69,0))&" is the limit"))',SMALL,align=LEFT)
put(ws,'F72',f'=IF(OR(F70="",F70=0,C65=""),"",MAX(20,MIN(INDEX({C}$B$5:$B$8,MATCH({SCALED},{C}$A$5:$A$8,1)),FLOOR({C}$B$9*$K$9,5),FLOOR(D41*C24*D72/F70,5))))',N,'0')
put(ws,'F73','=IF(COUNTIF(F42:F45,"KNOCKOUT")>0,"-",IF(OR(F70="",F70=0,F72=""),"",FLOOR(F70*F72/D72,500)))',N,'"$"#,##0')
for m in list(ws.merged_cells.ranges):
    if m.min_row>=75: ws.unmerge_cells(str(m))
clear(ws,'B75:M95')
for r in range(75,96): ws.row_dimensions[r].height=None
hdr(ws,'B75','F.  FINAL OFFER','B75:D75')
FROWS=[(76,'FINAL DECISION','=IF(C24=0,"",IF(OR(C73="-",C73="",C78="",C78=0,C78<E61,COUNTIF(F42:F45,"RED")+COUNTIF(F42:F45,"KNOCKOUT")+COUNTIF($J$76:$J$79,"RED")>=C61),"DECLINE",IF(OR(COUNTIF($J$76:$J$79,"YELLOW")>0,COUNTIF(F42:F45,"YELLOW")>0,C84="Negative"),"CONDITIONAL","APPROVE")))',None,13),
 (77,'Cash flow only','=F73','"$"#,##0',9),
 (78,'RECOMMENDED FUND',f'=IF(OR(C73="",C73="-"),C73,IF({HS}$C$34="Negative",FLOOR(C73*(1-{C}$B$32),500),C73))','"$"#,##0',13),
 (79,'Factor','=IF(OR(C78="",C78="-"),"",D72)','0.00',11),
 (80,'Term (weeks)','=IF(OR(C78="",C78="-"),"",C72/5)','0',11),
 (81,'Term (daily pmts)','=IF(OR(C78="",C78="-"),"",C72)','0',11),
 (82,'Weekly payment','=IF(OR(C78="",C78="-"),"",C78*D72/C72*5)','"$"#,##0.00',11),
 (83,'Daily payment','=IF(OR(C78="",C78="-"),"",C78*D72/C72)','"$"#,##0.00',11),
 (84,'HOUSE-BOOK OUTCOME',f'={HS}$C$34',None,11)]
for r,t,f,fmt,sz in FROWS:
    c=ws.cell(r,2,t); c.font=Font(name='Calibri',size=9 if r==77 else 11,bold=True); c.alignment=Alignment(horizontal='right',vertical='center'); c.border=BOX
    c=ws.cell(r,3,f); c.font=Font(name='Calibri',size=sz,bold=True); c.alignment=CEN; c.border=BOX
    if fmt: c.number_format=fmt
put(ws,'I75','RE-CHECK AT FINAL OFFER',SUB,fill=MID,align=LEFT); ws.merge_cells('I75:K75')
for r,t,f,src,fmt in [(76,'Holdback','=IF($C$83="","",$C$83*$K$9/C24)',39,'0.0%'),(77,'Total LVG','=IF($C$83="","",(E34+$C$83)*$K$9/C24)',38,'0.0%'),(78,'Daily/balance','=IF(OR($C$83="",D24=0),"",(E34+$C$83)/D24)',40,'0.0%'),(79,'Fund/revenue','=IF(OR(C78="",C78="-"),"",C78/C24)',41,'0.00')]:
    put(ws,f'I{r}',t,SMALL,align=LEFT); put(ws,f'K{r}',f,SMALL,fmt)
    put(ws,f'J{r}',f'=IF(K{r}="","",IF(K{r}>=G{src},"KNOCKOUT",IF(K{r}<=D{src},"GREEN",IF(K{r}>=E{src},"RED","YELLOW"))))',SMALL)
for v,f,fc in (('GREEN',G_FILL,'FFFFFF'),('YELLOW',Y_FILL,'000000'),('RED',R_FILL,'FFFFFF'),('KNOCKOUT',R_FILL,'FFFFFF')):
    ws.conditional_formatting.add('J76:J79',CellIsRule(operator='equal',formula=[f'"{v}"'],fill=f,font=Font(color=fc,bold=True)))
for v,f,fc in (('APPROVE',G_FILL,'FFFFFF'),('CONDITIONAL',Y_FILL,'000000'),('DECLINE',R_FILL,'FFFFFF')):
    ws.conditional_formatting.add('C76',CellIsRule(operator='equal',formula=[f'"{v}"'],fill=f,font=Font(color=fc,bold=True)))
for r in range(64,75): ws.row_dimensions[r].hidden=True

# ======================= Historical Score =======================
for m in list(sc.merged_cells.ranges): sc.unmerge_cells(str(m))
clear(sc,'A4:I60')
hdr(sc,'A4','DEAL BEING MATCHED','A4:C4')
for r,(t,f) in enumerate([('New deal or renewal','=IF(Deal!$K$14="","",Deal!$K$14)'),('Position (Deal C42)','=IF(Deal!$C$42="","",Deal!$C$42)'),('Industry','=IF(Deal!$K$13="","",Deal!$K$13)'),('ISO','=IF(Deal!$K$12="","",Deal!$K$12)')],start=5):
    label(sc,f'B{r}',t); put(sc,f'C{r}',f,B)
label(sc,'B10','Seasoned-book principal lost rate'); put(sc,'C10','=IF($F$21="",0,$F$21)',B,'0.0%')
label(sc,'B11','Seasoned-book return on funded'); put(sc,'C11','=IF($G$21="",0,$G$21)',B,'0.0%')
note(sc,'A2','Seasoned = age >= multiple x contracted term OR the deal reached a terminal status (Closed family, Written-Off, Bankruptcy, Settlement).  SIDE NOTE: a blend rate and a credibility factor are under consideration; not built.')
sub(sc,13,['','Cohort','Seasoned deals','Funded $ (seasoned)','Meets min?','Principal lost rate','Return on funded','Index'])
def cohort_row(r,name,flag):
    E_=hr('AP'); F=hr(flag) if flag else None
    label(sc,f'B{r}',name)
    pre=f'{F}*' if F else ''
    put(sc,f'C{r}',f'=SUMPRODUCT({pre}{E_})',N,'0'); put(sc,f'D{r}',f'=SUMPRODUCT({pre}{E_}*{hr("N")})',N,'"$"#,##0')
    put(sc,f'E{r}',f'=IF(AND(C{r}>=Controls!$B$24,D{r}>=Controls!$B$25),"Yes","No")',B)
    put(sc,f'F{r}',f'=IF(D{r}=0,"",SUMPRODUCT({pre}{E_}*{hr("AK")})/D{r})',N,'0.0%')
    put(sc,f'G{r}',f'=IF(D{r}=0,"",SUMPRODUCT({pre}{E_}*{hr("AJ")})/D{r})',N,'0.0%')
    put(sc,f'H{r}',f'=IF(OR(F{r}="",$C$10=0),"",F{r}/$C$10)',N,'0.00')
names=[('New/renewal + ISO + industry','AU'),('New/renewal + industry + position','AV'),('New/renewal + ISO + position','AW'),('New/renewal + industry','AX'),('New/renewal + ISO','AY'),('New/renewal + position','AZ'),('New/renewal only','AR'),('Whole seasoned book',None)]
for i,(n_,flag) in enumerate(names): cohort_row(14+i,n_,flag)
sub(sc,22,['','SINGLE-DIMENSION EVIDENCE  (feeds C2 on the Deal sheet; does not pick the cohort)','','','','','','','|PLR - book|'])
for i,(n_,flag) in enumerate([('Position only','AQ'),('New/renewal only','AR'),('Industry only','AS'),('ISO only','AT')]):
    cohort_row(23+i,n_,flag); put(sc,f'I{23+i}',f'=IF(F{23+i}="",0,ABS(F{23+i}-$C$10))',NOTE,'0.0%')
hdr(sc,'A28','RESULT','A28:D28')
res=[(29,'Primary cohort  (lowest row that meets both minimums)','=INDEX($B$14:$B$21,$D$29)',None),(30,'Seasoned deals in it','=INDEX($C$14:$C$21,$D$29)','0'),(31,'Principal lost rate','=INDEX($F$14:$F$21,$D$29)','0.0%'),(32,'Return on funded','=INDEX($G$14:$G$21,$D$29)','0.0%'),(33,'Index (PLR / book PLR)','=INDEX($H$14:$H$21,$D$29)','0.00'),
     (34,'HOUSE-BOOK OUTCOME','=IF(OR(D29=8,C33=""),"Neutral",IF(C33<=Controls!$B$29,"Supportive",IF(C33>Controls!$B$30,"Negative","Neutral")))',None),(35,'Sizing flex applied (Supportive only)','=IF(C34="Supportive",Controls!$B$31,0)','0%'),
     (36,'Largest-impact variable','=IF(MAX($I$23:$I$26)=0,"none rated",INDEX($B$23:$B$26,MATCH(MAX($I$23:$I$26),$I$23:$I$26,0)))',None)]
for r,t,f,fmt in res: label(sc,f'B{r}',t); put(sc,f'C{r}',f,B,fmt,align=LEFT)
put(sc,'D29','=IFERROR(MATCH("Yes",$E$14:$E$21,0),8)',NOTE,'0'); sc['E29']='row # (read by Historical Data)'; sc['E29'].font=NOTE
sc['C34'].font=Font(name='Calibri',size=12,bold=True)

# ======================= Historical Data =======================
heads={'AU':'NR+ISO+Ind','AV':'NR+Ind+Pos','AW':'NR+ISO+Pos','AX':'NR+Ind','AY':'NR+ISO','AZ':'NR+Pos','BA':'In primary cohort'}
for col,t in heads.items(): hs[f'{col}1']=t
for col in ('BB','BC'):
    for r in range(1,NR+2): hs[f'{col}{r}'].value=None
def hd_formulas(rr):
    return {'Y':f'=IF(I{rr}="","No SIC",IFERROR(VLOOKUP(I{rr},Ref!$A$2:$B$65,2,FALSE),"Unclassified"))','Z':f'=IF(T{rr}="","",IF(T{rr}<12,"Daily","Weekly"))',
       'AA':f'=IF(T{rr}="",0,IF(T{rr}<12,T{rr}*Controls!$B$13,T{rr}*5))','AB':f'=IF(E{rr}="",0,NETWORKDAYS(E{rr},Controls!$B$20,{HOL}))','AC':f'=IF(OR(AF{rr}=1,AND(AA{rr}>0,AB{rr}>=Controls!$B$23*AA{rr})),1,0)',
       'AD':f'=IF(AND(E{rr}<>"",E{rr}>=Controls!$B$21,E{rr}<=Controls!$B$22),1,0)','AE':f'=IF(IFERROR(VLOOKUP(D{rr},{ST},2,FALSE),"No")="Yes",1,0)','AF':f'=IF(IFERROR(VLOOKUP(D{rr},{ST},3,FALSE),"No")="Yes",1,0)',
       'AG':f'=IF(IFERROR(VLOOKUP(D{rr},{ST},4,FALSE),"No")="Yes",1,0)','AH':f'=IF(IFERROR(VLOOKUP(D{rr},{ST},5,FALSE),"No")="Yes",1,0)','AI':f'=IF(ISNUMBER(MATCH(D{rr},Controls!$A$36:$A$46,0)),1,0)',
       'AJ':f'=Q{rr}-N{rr}','AK':f'=IF(AE{rr}=1,MAX(0,N{rr}-Q{rr}),0)','AL':f'=IF(OR(V{rr}="",V{rr}=0),0,N{rr}/V{rr})','AM':f'=IF(OR(V{rr}="",V{rr}=0),0,1)','AN':f'=IF(OR(E{rr}="",F{rr}=""),0,NETWORKDAYS(E{rr},F{rr},{HOL}))','AO':f'=IF(L{rr}="Renewal",1,0)','AP':f'=AC{rr}*AD{rr}',
       'AQ':f"=IF(AND(G{rr}<>\"\",{HS}$C$6<>\"\"),IF(G{rr}={HS}$C$6,1,0),0)",'AR':f"=IF(AND(L{rr}<>\"\",{HS}$C$5<>\"\"),IF(L{rr}={HS}$C$5,1,0),0)",'AS':f"=IF(AND(Y{rr}<>\"\",{HS}$C$7<>\"\"),IF(Y{rr}={HS}$C$7,1,0),0)",'AT':f"=IF(AND(K{rr}<>\"\",{HS}$C$8<>\"\"),IF(K{rr}={HS}$C$8,1,0),0)",
       'AU':f'=AR{rr}*AT{rr}*AS{rr}','AV':f'=AR{rr}*AS{rr}*AQ{rr}','AW':f'=AR{rr}*AT{rr}*AQ{rr}','AX':f'=AR{rr}*AS{rr}','AY':f'=AR{rr}*AT{rr}','AZ':f'=AR{rr}*AQ{rr}',
       'BA':f"=IF(AD{rr}=1,CHOOSE({HS}$D$29,AU{rr},AV{rr},AW{rr},AX{rr},AY{rr},AZ{rr},AR{rr},1),0)"}
keep=[0,1,2,3,4,5,6,7,9,11,12,13,14,15,16,17,18,19,20,21,23,24,25,26]
if FULL:
    styles=[copy.copy(hs.cell(2,c)._style) for c in range(1,54)]
    for i,r in enumerate(raw):
        rr=2+i
        for c in range(1,54): hs.cell(rr,c)._style=copy.copy(styles[c-1])
        for j,k in enumerate(keep):
            v=r[k]; v=str(v) if (k in (7,9) and v is not None) else v
            hs.cell(rr,1+j,v)
        for col,fx in hd_formulas(rr).items(): hs[f'{col}{rr}']=fx
    for rr in range(NR+1,NR+12):
        for c in range(1,54): hs.cell(rr,c).value=None
else:
    for rr in range(2,NR+1):
        for col,fx in hd_formulas(rr).items(): hs[f'{col}{rr}']=fx

# ======================= Comparable Deals =======================
cd['U4']='In primary cohort'
for r in range(4,NR+5): cd[f'V{r}'].value=None
src=['A','B','D','E','G','Y','K','L','M','N','Q','S','AJ','AK','AA','AB','AN','AC','AE','AH','BA']
if FULL:
    styles=[copy.copy(cd.cell(5,c)._style) for c in range(1,22)]
    for i in range(2,NR+1):
        r=3+i
        for j,col in enumerate(src):
            c=cd.cell(r,1+j,f"={HDR}{col}{i}"); c._style=copy.copy(styles[j])
    cd.auto_filter.ref=f'A4:U{NR+3}'
wb.calculation=CalcProperties(fullCalcOnLoad=True)
wb.save(OUT); print('saved',OUT,'deals',NR-1)
