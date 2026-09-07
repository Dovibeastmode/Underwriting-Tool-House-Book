"""Outline (brainstorm) version: Layer 1 Deal sheet kept live and identical; every other cell describes the planned logic in words."""
import openpyxl, datetime as dt, sys
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter as L
ROOT='/home/user/Underwriting-Tool-House-Book'
UW=f'{ROOT}/originals/Updated_Underwriting_Tool (v2 upload).xlsx'
SAMPLE=f'{ROOT}/originals/Copy_of_House_Book_Analytics (sample).xlsx'
FULL=f'{ROOT}/originals/House Book Analytics (2).xlsx'
OUT=sys.argv[1] if len(sys.argv)>1 else f'{ROOT}/Aspire_Combined_Underwriting_Model_OUTLINE.xlsx'
DARK=PatternFill('solid',fgColor='1F2937'); MID=PatternFill('solid',fgColor='4B5563'); INPUT=PatternFill('solid',fgColor='FFF9C4'); PLAN=PatternFill('solid',fgColor='EEF2FF')
WHITE=Font(name='Calibri',size=11,bold=True,color='FFFFFF'); SUB=Font(name='Calibri',size=9,bold=True,color='FFFFFF')
B=Font(name='Calibri',size=11,bold=True); N=Font(name='Calibri',size=11); NOTE=Font(name='Calibri',size=9,color='374151'); INP=Font(name='Calibri',size=11,color='0000FF')
LOGIC=Font(name='Calibri',size=10,italic=True,color='3730A3')
thin=Side(style='thin',color='9CA3AF'); med=Side(style='medium',color='374151')
BOX=Border(left=thin,right=thin,top=thin,bottom=thin); IBOX=Border(left=med,right=med,top=med,bottom=med)
CEN=Alignment(horizontal='center',vertical='center'); LEFT=Alignment(horizontal='left',vertical='center'); WRAP=Alignment(horizontal='left',vertical='top',wrap_text=True)
def hdr(ws,cell,text,span=None):
    ws[cell]=text; ws[cell].font=WHITE; ws[cell].fill=DARK; ws[cell].alignment=LEFT
    if span: ws.merge_cells(span)
def sub(ws,row,labels,start=1):
    for i,t in enumerate(labels):
        c=ws.cell(row,start+i,t); c.font=SUB; c.fill=MID; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=BOX
def label(ws,addr,text,font=N):
    c=ws[addr]; c.value=text; c.font=font; c.alignment=LEFT; c.border=BOX; return c
def logic(ws,addr,text):
    c=ws[addr]; c.value='→ '+text; c.font=LOGIC; c.fill=PLAN; c.alignment=WRAP; c.border=BOX; return c
def inp(ws,addr,text=None,fmt=None):
    c=ws[addr]; c.value=text; c.font=INP; c.fill=INPUT; c.alignment=CEN; c.border=IBOX
    if fmt: c.number_format=fmt
    return c
def note(ws,addr,text):
    c=ws[addr]; c.value=text; c.font=NOTE; c.alignment=WRAP; return c
def widths(ws,d):
    for k,v in d.items(): ws.column_dimensions[k].width=v

# ---------- reference data ----------
hv=openpyxl.load_workbook(FULL,data_only=True); ref=hv['Ref']
sic=[(ref.cell(r,1).value,ref.cell(r,2).value,ref.cell(r,3).value) for r in range(4,68)]
hols=[ref.cell(r,6).value for r in range(4,45) if ref.cell(r,6).value]
isos=sorted({r[12] for r in hv['Data'].iter_rows(min_row=2,max_row=1068,values_only=True) if r[12]},key=str.lower)
states=sorted({r[11] for r in hv['Data'].iter_rows(min_row=2,max_row=1068,values_only=True) if r[11]})
statuses=[('Closed','No','Yes','Yes','No'),('Open','No','No','No','Yes'),('Legal','Yes','No','No','No'),('In House Collection','Yes','No','No','No'),('Lowered Payments','No','No','No','No'),('Written-Off','Yes','Yes','No','No'),('Bankruptcy','Yes','Yes','No','No'),('Settlement Agreement','Yes','Yes','No','No'),('Closed/NSF','No','Yes','Yes','No'),('Closed/NSF Unpaid','No','Yes','Yes','No'),('At-Risk','Yes','No','No','No')]
industries=sorted({s[1] for s in sic})+['No SIC','Unclassified']
sv=openpyxl.load_workbook(SAMPLE,data_only=True)['Data']
sample=[list(r) for r in sv.iter_rows(min_row=2,max_row=9,max_col=27,values_only=True)]

# ---------- Deal sheet: the live template, untouched ----------
wb=openpyxl.load_workbook(UW)
for n in wb.sheetnames:
    if n!='New Template': del wb[n]
ws=wb['New Template']; ws.title='Deal'
# add the PLANNED house-book layer below the existing tool (rows 65+), nothing above row 64 is changed
r=66
hdr(ws,f'B{r}','F.  HOUSE-BOOK LAYER  (PLANNED - outline only, no formulas yet)',f'B{r}:M{r}')
note(ws,f'B{r+1}','Sections A-E above are the battle-tested cash-flow engine and stay exactly as built. Everything below is the plan for what gets added. Blue-on-yellow = new input cell; italic purple = logic to be built.')
r+=3
sub(ws,r,['New inputs (Deal Profile)','Cell','What it feeds'],start=2); ws.merge_cells(f'D{r}:H{r}')
prof=[('ISO','dropdown from Ref!ISO list','Historical tiers 2 and 4; ISO-only evidence'),
      ('Industry (SIC major group)','dropdown from Ref!Industry list','Tiers 1 and 3; industry-only evidence'),
      ('State','dropdown','State-only evidence (informational, never sizes the offer)'),
      ('New deal or renewal','New Deal / Renewal','Every tier - strongest surviving signal in the book'),
      ('Credit score (FICO)','number','8th metric: >= 650 GREEN, 630-649 YELLOW, < 630 RED (bands on Controls)'),
      ('Proposed position','auto = C42 (positions incl. new), manual override allowed','Tiers 1, 2, 5; position-only evidence')]
for t,cell,feed in prof:
    r+=1; label(ws,f'B{r}',t,B); inp(ws,f'C{r}',cell); logic(ws,f'D{r}',feed); ws.merge_cells(f'D{r}:H{r}')
r+=2
sub(ws,r,['Final offer read-out (from Decision Summary)','Value','How it will be computed'],start=2); ws.merge_cells(f'D{r}:H{r}')
outs=[('FINAL DECISION','APPROVE / CONDITIONAL / MANUAL REVIEW / DECLINE','DECLINE if Layer 1 declines (merchant knockout, red flags >= limit, below minimum). MANUAL REVIEW if any Controls trigger fires. CONDITIONAL if any yellow/red remains, the modifier is negative, or a risk stip applies. Otherwise APPROVE.'),
      ('Recommended fund','$','MIN(cash-flow max fund C62 x modifier multiplier, cash-flow max daily C59 x adjusted term / adjusted factor), floored to $500. Can never exceed C62.'),
      ('Factor','1.40 / 1.45 / 1.49','Band factor D61, stepped up one approved tier when the modifier says so; never above 1.49.'),
      ('Term','weeks or daily pmts','Band term C61, moved down one or two rows of the band table when the modifier says so; floor 20 payments.'),
      ('Payment','per week / per day','fund x factor / term (daily); x5 for weekly.'),
      ('Frequency','Daily / Weekly','As proposed in E6 unless the modifier requires Weekly.'),
      ('Historical modifier','Strong positive .. Severe','From Historical Score: relative index of the primary cohort, capped by confidence.'),
      ('Comparable deals','seasoned / all','Count of house-book deals in the primary cohort.'),
      ('Cohort principal lost rate','%','Primary cohort, seasoned deals only, vs seasoned-book baseline.'),
      ('Confidence','High / Medium / Low','Credibility Z and open share of the primary cohort.'),
      ('Stipulations','text','Standard set (DataMerch, UCC, payoff letters) + conditional rules on Controls.'),
      ("Underwriter's conclusion",'paragraph','Sentence built from the actual numbers: cash-flow max and its limiting metric, recommended offer and % of max, primary cohort name / size / lost rate vs book / credibility / index / level, confidence, ISO / industry / state sentences, triggers and stips.')]
for t,val,how in outs:
    r+=1; label(ws,f'B{r}',t,B); c=ws[f'C{r}']; c.value=val; c.font=NOTE; c.alignment=CEN; c.border=BOX; logic(ws,f'D{r}',how); ws.merge_cells(f'D{r}:H{r}'); ws.row_dimensions[r].height=30
ws.row_dimensions[r].height=75
r+=2
note(ws,f'B{r}','Rule of the house-book layer: it may reduce fund, shorten term, raise factor within tiers, require weekly, add stips or send to manual review. It may NOT remove a knockout, turn a green metric red, exceed C62, or make a precise adjustment on a tiny or unseasoned cohort.')
ws.sheet_view.showGridLines=False

# ---------- Controls ----------
cs=wb.create_sheet('Controls'); widths(cs,{'A':54,'B':16,'C':16,'D':16,'E':16,'F':14,'G':14,'H':62})
cs['A1']='CONTROLS  -  every threshold, tier and rule the model will read (values are real; nothing here is a formula)'; cs['A1'].font=Font(name='Calibri',size=14,bold=True)
note(cs,'A2','Blue = editable policy value. The Deal sheet currently holds its own copies of section 1 and 2 (D38:G44, J13:L17, C50, E50); in the build they become links to these cells.')
hdr(cs,'A3','1. CASH-FLOW THRESHOLDS  (unchanged from the Deal sheet)','A3:H3'); sub(cs,4,['Metric','Green (at or better)','Red (at or worse)','Knockout','Direction','','','Note'])
thr=[('Total LVG (all debt / rev)',0.30,0.45,0.55,'higher is worse','0.0%','Monthly debt service incl. new deal / avg true revenue'),
     ('New holdback SP%',0.18,0.25,0.35,'higher is worse','0.0%',''),('Daily debt / balance',0.18,0.30,0.35,'higher is worse','0.0%',''),
     ('Fund / revenue',0.60,0.90,1.25,'higher is worse','0.00',''),('# positions (incl. new)',2,5,6,'higher is worse','0',''),
     ('Revenue trend (avg MoM)',-0.05,-0.15,-0.30,'lower is worse','0.0%',''),('Neg days + NSFs (total)',1,6,10,'higher is worse','0','Total across all months entered'),
     ('Credit score (FICO)  - NEW',650,630,0,'lower is worse','0','650+ green, 630-649 yellow, below 630 red. Knockout 0 = off')]
for i,(t,g,r_,k,d,fmt,n_) in enumerate(thr):
    row=5+i; label(cs,f'A{row}',t); inp(cs,f'B{row}',g,fmt); inp(cs,f'C{row}',r_,fmt); inp(cs,f'D{row}',k,fmt); label(cs,f'E{row}',d); note(cs,f'H{row}',n_)
label(cs,'A13','Decline if red flags (RED + KNOCKOUT) >='); inp(cs,'B13',2,'0')
label(cs,'A14','Minimum fund $'); inp(cs,'B14',10000,'"$"#,##0'); note(cs,'H14','Deal sheet E50 currently 10,000')
label(cs,'A15','Underwriting fee %'); inp(cs,'B15',0.05,'0%')
hdr(cs,'A17','2. PRICING BANDS  (unchanged)','A17:H17'); sub(cs,18,['Score >=','Term cap (daily pmts)','Factor','','','','','Note'])
for i,(s,t,f) in enumerate([(0,100,1.49),(5,130,1.45),(6,150,1.45),(7,195,1.40)]):
    row=19+i; inp(cs,f'A{row}',s,'0.0'); inp(cs,f'B{row}',t,'0'); inp(cs,f'C{row}',f,'0.00')
note(cs,'H19','Term step -1 in the modifier table = one row up this table')
label(cs,'A23','Max term (months)'); inp(cs,'B23',9,'0')
label(cs,'A24','Approved factor tiers'); inp(cs,'B24',1.40,'0.00'); inp(cs,'C24',1.45,'0.00'); inp(cs,'D24',1.49,'0.00'); note(cs,'H24','1.35 = prepay, not a tier')
label(cs,'A25','Tier cut-offs (7-point scale): Strong / Solid / Moderate'); inp(cs,'B25',6.5,'0.0'); inp(cs,'C25',5,'0.0'); inp(cs,'D25',3.5,'0.0'); note(cs,'H25','With credit score the scorecard has 8 metrics; score x 7 / metrics-scored keeps the bands meaningful')
hdr(cs,'A27','3. CONVERSION CONSTANTS','A27:H27')
for row,(t,v,fmt,n_) in enumerate([('Weeks per month (ours)',4.331,'0.000','time converts with 4.331'),('Business days per month (ours)',21.655,'0.000','money converts with 5 and 21.655'),('MoneyBadger weeks / month',4,'0',''),('MoneyBadger business days / month',21,'0',''),('Calendar days / month',30.4375,'0.0000',''),('Assumed factor on existing positions',1.45,'0.00','')],start=28):
    label(cs,f'A{row}',t); inp(cs,f'B{row}',v,fmt); note(cs,f'H{row}',n_)
hdr(cs,'A35','4. HISTORICAL EVIDENCE SETTINGS  (new)','A35:H35')
hist=[('As-of date for seasoning (date of the CRM export)',dt.datetime(2026,8,25),'mm/dd/yyyy','age = business days from funding to this date'),
      ('Window: funded on or after',dt.datetime(2024,11,1),'mm/dd/yyyy',''),('Window: funded on or before',dt.datetime(2026,8,25),'mm/dd/yyyy','set 12/31/2025 to read the settled book'),
      ('Seasoning multiple (age >= multiple x contracted term)',1.25,'0.00','a deal is only rated once it has had 125% of its term to play out'),
      ('Minimum seasoned deals for a cohort to be PRIMARY',10,'0',''),('Minimum seasoned funded $ for a cohort to be PRIMARY',150000,'"$"#,##0',''),
      ('Credibility constant K',30,'0','Z = n / (n + K): 30 deals = half weight, 90 = 75%'),('Max open share for "seasoned-enough"',0.15,'0%',''),
      ('High confidence: min Z',0.60,'0.00',''),('Medium confidence: min Z',0.30,'0.00',''),('Low confidence: max steps from Neutral',1,'0','weak evidence can never drive a severe adjustment'),
      ('Minimum seasoned deals before a rate is shown',5,'0','')]
for i,(t,v,fmt,n_) in enumerate(hist):
    row=36+i; label(cs,f'A{row}',t); inp(cs,f'B{row}',v,fmt); note(cs,f'H{row}',n_)
hdr(cs,'A49','5. HISTORICAL MODIFIER TABLE  (relative index = blended cohort lost rate / seasoned-book lost rate)','A49:H49')
sub(cs,50,['Index at or below','Level','Fund multiplier (x cash-flow max)','Term step (bands)','Factor step (tiers)','Frequency','Manual review?','Note'])
for i,m in enumerate([(0.60,'Strong positive',1.00,0,0,'As proposed','N','full cash-flow max only with strong credible evidence'),(0.85,'Mild positive',0.95,0,0,'As proposed','N',''),(1.15,'Neutral',0.90,0,0,'As proposed','N','base utilisation; 1.00 reproduces the old tool'),(1.40,'Mild negative',0.80,-1,1,'As proposed','N',''),(1.75,'Negative',0.70,-1,1,'Weekly','N','adds weekly + no-stacking covenant'),(999,'Severe',0.55,-2,1,'Weekly','Y','')]):
    row=51+i
    for j,v in enumerate(m[:7]):
        if j==1: label(cs,f'B{row}',v,B)
        else: inp(cs,f'{L(1+j)}{row}',v,{0:'0.00',2:'0.00'}.get(j))
    note(cs,f'H{row}',m[7])
hdr(cs,'A59','6. STATUS MAPPING  (the only place default / resolved / closed / open are defined)','A59:H59'); sub(cs,60,['Status','Default?','Resolved?','Closed?','Open?','','',''])
for i,s in enumerate(statuses):
    row=61+i; label(cs,f'A{row}',s[0]); [inp(cs,f'{L(2+j)}{row}',s[j+1]) for j in range(4)]
note(cs,'H65','Lowered Payments = restructured, still collecting; house book treats it as performing')
hdr(cs,'A74','7. MANUAL-REVIEW TRIGGERS','A74:H74'); sub(cs,75,['Trigger','Active?','Parameter','','','','','Note'])
for i,(t,a,p,n_) in enumerate([('Historical modifier is Severe','Y',None,''),('Confidence Low AND any metric RED','Y',None,''),('Credit score RED AND another metric RED','Y',None,''),('Positions incl. new at or above','Y',4,'log rule: fund up to a 4th'),('New deal, 3rd+ position, construction industry','Y',3,'the one raw house-book effect still visible after controls'),('ISO or industry not in the house book','Y',None,''),('Adjusted fund below minimum','Y',None,''),('Primary cohort open share above max','N',None,'')]):
    row=76+i; label(cs,f'A{row}',t); inp(cs,f'B{row}',a)
    if p is not None: inp(cs,f'C{row}',p,'0')
    note(cs,f'H{row}',n_)
hdr(cs,'A85','8. STIPULATION RULES','A85:H85'); sub(cs,86,['Stipulation','Active?','Condition','','','','','Drives CONDITIONAL?'])
for i,(t,a,cnd,d) in enumerate([('DataMerch, UCC search, positions / payoff letters','Y','Always','No (standard)'),('No-stacking covenant with default-to-daily','Y','Modifier Mild negative or worse','Yes'),('Bank login verification (Plaid / DecisionLogic)','Y','Neg days + NSFs not GREEN','Yes'),('Payoff / balance letter for every existing position','Y','Existing positions >= 2','No (documentary)'),('Landlord / lease letter','Y','Credit score not GREEN','Yes'),('Underwriter review of the historical evidence','Y','Confidence Low','Yes'),('Weekly remittance only','Y','Frequency switched by the modifier','Yes'),('Prior-contract payoff and renewal history','Y','Renewal selected','No (documentary)')]):
    row=87+i; label(cs,f'A{row}',t); inp(cs,f'B{row}',a); label(cs,f'C{row}',cnd); cs.merge_cells(f'C{row}:G{row}'); note(cs,f'H{row}',d)
cs.sheet_view.showGridLines=False

# ---------- Ref ----------
rf=wb.create_sheet('Ref'); widths(rf,{'A':8,'B':40,'C':32,'E':44,'G':40,'I':10,'K':14,'M':22})
for col,t in zip('ABCEGIKM',['SIC2','Industry (SIC major group)','Division','ISO (house-book spelling, dropdown source)','Industry list (dropdown)','State','Fed holiday (ACH)','Status list']):
    c=rf[f'{col}1']; c.value=t; c.font=SUB; c.fill=MID; c.alignment=CEN; c.border=BOX
for i,(a,b,c_) in enumerate(sic): rf.cell(2+i,1,a); rf.cell(2+i,2,b); rf.cell(2+i,3,c_)
for i,v in enumerate(isos): rf.cell(2+i,5,v)
for i,v in enumerate(industries): rf.cell(2+i,7,v)
for i,v in enumerate(states): rf.cell(2+i,9,v)
for i,v in enumerate(hols): rf.cell(2+i,11,v).number_format='mm/dd/yyyy'
for i,s in enumerate(statuses): rf.cell(2+i,13,s[0])

# ---------- Historical Data (8 sample deals; house-book helper formulas LIVE, new columns outlined) ----------
hs=wb.create_sheet('Historical Data')
raw_cols=['Contract ID','DBA','Legal name','Status','Start date','Last active','Position','SIC','SIC2','State','ISO','New or renewal','Factor','Advance','Commission','Payback','Collected','Balance','% paid','Exp dur (raw)','NSF count','Avg rev/mo','Holdback %','Biz start']
ST='Controls!$A$61:$E$71'; HOL=f'Ref!$K$2:$K${1+len(hols)}'; BDM='Controls!$B$29'; BDW='5'; WFROM='Controls!$B$37'; WTO='Controls!$B$38'
# (header, description, formula-template with {r}) - these are the original House Book Analytics formulas, re-pointed to this workbook's Controls / Ref
live=[('Unit of raw duration','raw < 12 = Months (daily deal), >= 12 = Weeks (weekly deal)','=IF(T{r}="","",IF(T{r}<12,"Months","Weeks"))'),
 ('Pay frequency','Months -> Daily, Weeks -> Weekly','=IF(Y{r}="","",IF(Y{r}="Months","Daily","Weekly"))'),
 ('Expected term (payments)','months x 21.655, or weeks',f'=IF(T{{r}}="","",IF(Z{{r}}="Daily",T{{r}}*{BDM},T{{r}}))'),
 ('Expected term (business days)','months x 21.655, or weeks x 5',f'=IF(T{{r}}="","",IF(Z{{r}}="Daily",T{{r}}*{BDM},T{{r}}*{BDW}))'),
 ('Actual days on book (calendar)','last active - funding','=IF(OR(E{r}="",F{r}=""),"",F{r}-E{r})'),
 ('Actual days on book (business)','NETWORKDAYS(funding, last active, Fed holidays)',f'=IF(OR(E{{r}}="",F{{r}}=""),"",NETWORKDAYS(E{{r}},F{{r}},{HOL}))'),
 ('Default','status mapping, col 2',f'=IF(IFERROR(VLOOKUP(D{{r}},{ST},2,FALSE),"No")="Yes",1,0)'),
 ('Resolved','status mapping, col 3',f'=IF(IFERROR(VLOOKUP(D{{r}},{ST},3,FALSE),"No")="Yes",1,0)'),
 ('Closed family','status mapping, col 4',f'=IF(IFERROR(VLOOKUP(D{{r}},{ST},4,FALSE),"No")="Yes",1,0)'),
 ('In scope','funded inside the Controls window',f'=IF(AND(E{{r}}<>"",E{{r}}>={WFROM},E{{r}}<={WTO}),1,0)'),
 ('Net cash','collected - advance','=Q{r}-N{r}'),
 ('Principal lost','MAX(advance - collected, 0) on defaulted deals','=IF(AE{r}=1,MAX(0,N{r}-Q{r}),0)'),
 ('Collected %','collected / advance','=IFERROR(Q{r}/N{r},"")'),
 ('Advance / revenue','advance / avg monthly revenue; blank when missing','=IF(V{r}="","",N{r}/V{r})'),
 ('Has revenue','1 if revenue present','=IF(V{r}="",0,1)'),
 ('Never paid','1 if collected <= 0','=IF(Q{r}<=0,1,0)'),
 ('Days on book (defaulted deals)','business days, defaulted only','=IF(AE{r}=1,AD{r},"")'),
 ('Days to close','business days, non-defaulted Closed family','=IF(AND(AE{r}=0,AG{r}=1),AD{r},"")'),
 ('Size bucket','six advance buckets','=IF(N{r}<10000,"Under $10K",IF(N{r}<20000,"$10-20K",IF(N{r}<30000,"$20-30K",IF(N{r}<50000,"$30-50K",IF(N{r}<100000,"$50-100K","$100K+")))))'),
 ('Factor bucket','six factor buckets','=IF(M{r}<=1.35,"1.35 and under",IF(M{r}<=1.42,"1.36-1.42",IF(M{r}<=1.45,"1.43-1.45",IF(M{r}<1.49,"1.46-1.48",IF(M{r}<=1.49,"1.49","Above 1.49")))))'),
 ('Term bucket','cut on business days','=IF(AB{r}="","",IF(AB{r}<60,"Under 60",IF(AB{r}<90,"60-89",IF(AB{r}<120,"90-119",IF(AB{r}<150,"120-149","150+")))))'),
 ('Adv/rev bucket','with a "No revenue data" bucket','=IF(AL{r}="","No revenue data",IF(AL{r}<0.15,"Under 0.15x",IF(AL{r}<0.3,"0.15-0.30x",IF(AL{r}<0.5,"0.30-0.50x",IF(AL{r}<1,"0.50-1.00x","1.00x+")))))'),
 ('Industry','SIC2 looked up on Ref','=IF(I{r}="","No SIC",IFERROR(VLOOKUP(I{r},Ref!$A$2:$B$65,2,FALSE),"Unclassified"))'),
 ('Quarter','funding quarter','=IF(E{r}="","",TEXT(E{r},"YYYY")&"-Q"&ROUNDUP(MONTH(E{r})/3,0))'),
 ('Month','funding month','=IF(E{r}="","",TEXT(E{r},"YYYY-MM"))'),
 ('Year','funding year','=IF(E{r}="","",TEXT(E{r},"YYYY"))')]
planned=[('Open','from status mapping col 5 (new column in the mapping)'),('Age at as-of (bus. days)','NETWORKDAYS(funding date, Controls as-of date, holidays)'),('Seasoned','1 if age >= seasoning multiple x expected term (business days)'),('Eligible','seasoned x in-scope: the only rows that enter a rate'),('Renewal','1 if Renewal'),('m: position','1 if position = proposed position on Deal'),('m: new/renewal','1 if same as Deal'),('m: industry','1 if same as Deal'),('m: ISO','1 if same as Deal'),('m: state','1 if same as Deal'),('T1 NR+Ind+Pos','m:NR x m:Ind x m:Pos'),('T2 NR+ISO+Pos','m:NR x m:ISO x m:Pos'),('T3 NR+Ind','m:NR x m:Ind'),('T4 NR+ISO','m:NR x m:ISO'),('T5 NR+Pos','m:NR x m:Pos'),('T6 NR','m:NR'),('T7 Book','always 1'),('Best tier','lowest tier number the deal matches'),('In primary cohort','1 if the deal matches the primary tier chosen on Historical Score'),('Seq','running count of in-primary-cohort rows; feeds the Comparable Deals list')]
allh=raw_cols+[l[0] for l in live]+[p_[0] for p_ in planned]
for i,t in enumerate(allh):
    c=hs.cell(1,1+i,t); c.font=SUB; c.fill=MID if i<len(raw_cols) else (PatternFill('solid',fgColor='065F46') if i<len(raw_cols)+len(live) else DARK); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=BOX
hs.row_dimensions[1].height=42
hs.cell(2,1,'RAW from CRM export (values) →').font=NOTE
for i,(t,d,f) in enumerate(live): note(hs,f'{L(len(raw_cols)+1+i)}2','LIVE (house book): '+d)
for i,(t,d) in enumerate(planned): logic(hs,f'{L(len(raw_cols)+len(live)+1+i)}2',d)
hs.row_dimensions[2].height=80
keep=[0,1,2,3,4,5,6,7,9,11,12,13,14,15,16,17,18,19,20,21,23,24,25,26]
assert L(len(raw_cols)+1)=='Y' and L(len(raw_cols)+len(live))=='AX'
for i,r in enumerate(sample):
    rr=3+i
    for j,k in enumerate(keep):
        v=r[k]; v=str(v) if (k in (7,9) and v is not None) else v
        c=hs.cell(rr,1+j,v)
        if k in (4,5,26): c.number_format='mm/dd/yyyy'
        if k in (15,16,17,18,19): c.number_format='"$"#,##0'
        if k in (20,25): c.number_format='0.0%'
    for j,(t,d,f) in enumerate(live):
        c=hs.cell(rr,len(raw_cols)+1+j,f.format(r=rr)); c.font=N
        if t in ('Net cash','Principal lost'): c.number_format='"$"#,##0'
        if t in ('Collected %','Advance / revenue'): c.number_format='0.00'
note(hs,'A12','8 synthetic sample deals from Copy_of_House_Book_Analytics.xlsx. Green headers = helper columns carried over from House Book Analytics with their formulas live (re-pointed to this workbook\'s Controls and Ref). Dark headers = new columns for the matching engine, outlined only. The build replaces the samples with the full CRM export (1,067 rows) and fills every helper column down.')
hs.freeze_panes='E3'; widths(hs,{'A':12,'B':22,'C':22,'D':18,'K':14})
for i in range(len(live)+len(planned)): hs.column_dimensions[L(len(raw_cols)+1+i)].width=15
# ---------- Historical Score ----------
sc=wb.create_sheet('Historical Score'); widths(sc,{'A':7,'B':36,'C':44})
sc['A1']='HISTORICAL SCORE  -  comparable cohorts, credibility and the risk modifier  (outline)'; sc['A1'].font=Font(name='Calibri',size=14,bold=True)
note(sc,'A2','Every rate uses SEASONED deals only. Unseasoned deals are counted, never rated. State never enters the hierarchy.')
hdr(sc,'A4','DEAL BEING MATCHED','A4:C4')
for row,(t,d) in enumerate([('New deal or renewal','Deal profile input'),('Proposed position','Deal C42, or the override'),('Industry','Deal profile input'),('ISO','Deal profile input'),('State','Deal profile input (informational)'),('Seasoned-book principal lost rate','tier 7 row below - the baseline every cohort is compared to')],start=5):
    label(sc,f'B{row}',t); logic(sc,f'C{row}',d)
heads=['Tier','Cohort','Keys','Deals (all)','Seasoned','% seasoned','Funded $ (seasoned)','Open n','Open %','Default n','Default %','Principal lost rate','Collected on defaults','Return on funded','Avg adv/rev','Avg term (bd)','Avg factor','Avg deal size','Avg days on book','Renewal share','Seasoned enough?','Credibility Z','Blended lost rate','Relative index','Meets minimum?']
sub(sc,13,heads); sc.row_dimensions[13].height=40
how=['','','','count of in-window deals with the tier flag','count with flag x eligible','seasoned / all','sum advance (seasoned)','count open (seasoned)','open / seasoned','count default (seasoned)','default / seasoned','sum principal lost / funded (blank under the min-show count)','collected on defaulted / advance on defaulted','net cash / funded','average of adv/rev where revenue exists','average contracted term','average factor','funded / seasoned','average days on book','renewals / seasoned','Yes if seasoned >= min deals AND funded >= min $ AND open % <= max','n / (n + K)','Z x cohort lost + (1 - Z) x book lost','blended / book lost','1 if seasoned >= min deals and funded >= min $']
tiers=[('1','T1: new/renewal + industry + position','closest match'),('2','T2: new/renewal + ISO + position','ISO relationship at the same stack depth'),('3','T3: new/renewal + industry','industry survived controls'),('4','T4: new/renewal + ISO','ISO survived controls'),('5','T5: new/renewal + position','position only within the renewal split'),('6','T6: new/renewal only','the book split'),('7','T7: seasoned book','baseline')]
for i,(t,nme,keys) in enumerate(tiers):
    row=14+i; label(sc,f'A{row}',t,B); label(sc,f'B{row}',nme); label(sc,f'C{row}',keys)
    for j in range(3,25): c=sc.cell(row,j+1,'…'); c.font=LOGIC; c.fill=PLAN; c.alignment=CEN; c.border=BOX
row=21
for j,h in enumerate(how):
    if h: logic(sc,f'{L(j+1)}{row}',h)
sc.row_dimensions[row].height=90; label(sc,f'B{row}','How each column is computed',B)
sub(sc,23,['','SINGLE-DIMENSION EVIDENCE  (informational + largest-impact test)']+['']*23)
for i,(nme,keys) in enumerate([('Position only','m:position'),('New/renewal only','m:NR'),('Industry only','m:industry'),('ISO only','m:ISO'),('State only  (INFORMATIONAL - never in hierarchy)','m:state')]):
    row=24+i; label(sc,f'A{row}','S'+str(i+1),B); label(sc,f'B{row}',nme); label(sc,f'C{row}',keys)
    for j in range(3,25): c=sc.cell(row,j+1,'…'); c.font=LOGIC; c.fill=PLAN; c.alignment=CEN; c.border=BOX
hdr(sc,'A30','RESULT','A30:C30')
res=[('Primary tier','first tier (1..7) whose "Meets minimum?" = 1; tier 7 always qualifies'),('Primary cohort stats','the primary row: seasoned n / all, % seasoned, lost rate, open share, return, collected on defaults, Z, blended, index'),
     ('Modifier level (1 strong positive .. 6 severe)','count of modifier-table index bounds below the relative index, + 1'),('Confidence','Low if primary = tier 7 or Z < Medium line; High if Z >= High line and open share <= max; else Medium'),
     ('Modifier level after confidence cap','if Low: clamp within Neutral +/- max steps (Controls)'),('Largest-impact variable','the single-dimension row (position / NR / industry / ISO) with the largest |blended - book|'),
     ('ISO / industry / state / position sentences','"ISO X: n seasoned deals, lost a% vs b% book, Z c -> stronger / neutral / weaker than book" (state sentence always ends "informational only")'),
     ('Conflicting signals?','YES if ISO and industry sit on opposite sides of neutral'),('Warnings','baseline in use / cohort not seasoned-enough / low confidence capped / open share high')]
for i,(t,d) in enumerate(res):
    row=31+i; label(sc,f'B{row}',t,B); logic(sc,f'C{row}',d); sc.row_dimensions[row].height=32
for col in range(4,26): sc.column_dimensions[L(col)].width=11
sc.sheet_view.showGridLines=False; sc.freeze_panes='D14'

# ---------- Comparable Deals ----------
cd=wb.create_sheet('Comparable Deals'); widths(cd,{'A':5,'B':12,'C':26,'D':18,'E':11,'F':8,'G':24,'H':22,'I':6,'J':11,'K':7,'L':11,'M':11,'N':7,'O':11,'P':11,'Q':9,'R':9,'S':9,'T':9,'U':8,'V':7,'W':9,'X':44})
cd['A1']='COMPARABLE DEALS  -  every house-book deal in the primary cohort, by Contract ID  (outline)'; cd['A1'].font=Font(name='Calibri',size=14,bold=True)
logic(cd,'A2','Header line: "Primary cohort: <name> | <seasoned> / <all>". Rows: for k = 1..400, INDEX/MATCH the k-th row whose Seq = k on Historical Data. Blank when the cohort runs out.'); cd.merge_cells('A2:X2'); cd.row_dimensions[2].height=32
sub(cd,4,['#','Contract ID','DBA','Status','Funded','Position','Industry','ISO','State','New/renewal','Factor','Advance','Collected','% paid','Net cash','Principal lost','Term (bd)','Age (bd)','Days on book','Seasoned','Default','Open','Similarity tier','Why included'])
for k in range(1,9):
    row=4+k; cd.cell(row,1,k).font=NOTE
    for j in range(2,24): c=cd.cell(row,j,'…'); c.font=LOGIC; c.fill=PLAN; c.alignment=CEN; c.border=BOX
    logic(cd,f'X{row}','"<tier keys>  [seasoned | unseasoned, defaulted, open]"')
note(cd,'A14','Similarity tier 1 = closest. Unseasoned rows are shown but excluded from every rate. The underwriter can open any Contract ID in the CRM from here.')
cd.freeze_panes='C5'

# ---------- Offer Engine ----------
oe=wb.create_sheet('Offer Engine'); widths(oe,{'A':3,'B':44,'C':24,'D':60})
oe['B1']='OFFER ENGINE  -  base cash-flow offer -> historical adjustment -> adjusted offer  (outline)'; oe['B1'].font=Font(name='Calibri',size=14,bold=True)
sub(oe,3,['','Item','Base (cash flow only)','Adjusted (with history)'],start=1)
rows=[('Cash-flow max daily payment','Deal C59','same - this is the ceiling'),('Term (daily pmts)','Deal C61 (band term)','MAX(20, MIN(base, band-table row moved by the term step))'),('Factor','Deal D61 (band factor)','next approved tier by the factor step, never above 1.49'),('Fund $','Deal C62 (cash-flow max)','MIN(base x multiplier, max daily x adjusted term / adjusted factor), floor $500'),('Daily payment','fund x factor / term','same formula on adjusted numbers - always <= max daily'),('Modifier level / multiplier / steps / frequency','from Historical Score + Controls table','same'),('Frequency (final)','Deal E6','Weekly if the modifier requires it, else as proposed'),('Term & payment in final units','weeks = daily pmts / 5','same'),('Payback / net / months','fund x factor; fund x (1 - fee); term / 21.655','same')]
for i,(t,b,a) in enumerate(rows):
    row=4+i; label(oe,f'B{row}',t,B); logic(oe,f'C{row}',b); logic(oe,f'D{row}',a); oe.row_dimensions[row].height=30
hdr(oe,'B14','METRICS RE-CHECKED AT THE ADJUSTED OFFER  (must stay off RED / KNOCKOUT)','B14:D14')
for i,t in enumerate(['Total LVG','New holdback SP%','Daily debt / balance','Fund / revenue']):
    row=15+i; label(oe,f'B{row}',t); logic(oe,f'C{row}','recomputed with the adjusted payment / fund'); logic(oe,f'D{row}','GREEN / YELLOW / RED / KNOCKOUT against Controls; positions, trend, NSF and credit carried over unchanged')
hdr(oe,'B20','SAFETY PROOF (shown on every deal)','B20:D20')
for i,t in enumerate(['Adjusted fund <= cash-flow max fund','Adjusted daily payment <= cash-flow max daily','No offer-side metric RED / KNOCKOUT at the adjusted offer','Merchant knockouts untouched (history cannot remove them)','Red flags at the adjusted offer (offer-side + merchant-side) -> feeds the final decision']):
    row=21+i; label(oe,f'B{row}',t); logic(oe,f'C{row}','OK / FAIL')
oe.sheet_view.showGridLines=False

# ---------- Decision Summary ----------
ds=wb.create_sheet('Decision Summary'); widths(ds,{'A':3,'B':46,'C':22,'D':90})
ds['B1']='DECISION SUMMARY  (outline of the 20-item user-facing output)'; ds['B1'].font=Font(name='Calibri',size=14,bold=True)
sub(ds,3,['','Item','Value','How it will be produced'],start=1)
items=[('DECISION','APPROVE / CONDITIONAL / MANUAL REVIEW / DECLINE','DECLINE: merchant knockout, no clean room, adjusted fund < minimum, or red flags at the adjusted offer >= limit. MANUAL REVIEW: any active trigger. CONDITIONAL: any yellow/red at the adjusted offer, modifier Mild negative or worse, or a risk-based stip. Else APPROVE.'),
 ('Recommended funding amount','$','Offer Engine adjusted fund; detail shows cash-flow max, % used, requested amount'),('Recommended factor rate','1.40/1.45/1.49','adjusted factor; note if raised by history'),('Recommended term','weeks or daily pmts','adjusted term in final units; months; note if shortened'),('Payment','$ per week / day','with daily-equivalent, monthly, payback, net'),('Daily or weekly','Daily / Weekly','final frequency and why'),
 ('Holdback %','%','at the adjusted offer, with status and green line'),('Total leverage','%','at the adjusted offer'),('Funding-to-revenue','x','at the adjusted offer'),('Daily debt-to-balance','%','at the adjusted offer'),('Number of positions (incl. new)','n','Deal C42 and status; override noted'),
 ('Cash-flow risk score','x of N','Deal score, tier name, scaled score'),('Historical risk score / modifier','level name','relative index, multiplier, steps, confidence cap'),('Overall confidence','High / Medium / Low','Z, open share, primary tier'),('Number of comparable deals','seasoned / all','primary cohort'),('Percentage of comparable deals seasoned','%',''),('Comparable-cohort principal lost rate','%','vs seasoned book, blended, collected on defaults'),('Comparable-cohort return','%','seasoned deals only'),
 ('Which variable had the largest historical impact','name + numbers','largest |blended - book| among position / NR / industry / ISO'),('Which metric limited the offer','name','Deal D59 plus whether the term was capped by fund/revenue'),('ISO / industry / state / position evidence','sentences','from Historical Score; conflict flag'),('Required stipulations','text','standard + conditional rules'),('Manual-review triggers fired','text','blank = none'),('Main risks / positive factors','text','reds and greens named; renewal called out'),
 ("Underwriter's conclusion",'paragraph','"Cash flow supports up to $X (limit, term @ factor). The recommendation is $Y (Z% of the maximum) at f over t, $p/wk, because the primary comparable cohort (name) holds n seasoned deals with a principal lost rate of a% against b% for the seasoned book; at credibility Z that blends to c% (index i), which is <level>. Confidence is <conf>. <ISO sentence>. <Industry sentence>. <State sentence - informational only>. <triggers / stips>."')]
for i,(t,v,h) in enumerate(items):
    row=4+i; label(ds,f'B{row}',t,B); c=ds[f'C{row}']; c.value=v; c.font=NOTE; c.alignment=CEN; c.border=BOX; logic(ds,f'D{row}',h); ds.row_dimensions[row].height=30 if len(h)<110 else 60
ds.row_dimensions[4+len(items)-1].height=100
ds.sheet_view.showGridLines=False

# ---------- Checks ----------
ck=wb.create_sheet('Checks'); widths(ck,{'A':3,'B':66,'C':14,'D':60})
ck['B1']='CHECKS / AUDIT  (planned list)'; ck['B1'].font=Font(name='Calibri',size=14,bold=True)
sub(ck,3,['','Check','Result','Rule'],start=1)
checks=[('MISSING INPUTS',None),('Bank data, avg balance, new/renewal, ISO, industry, state, credit score, requested offer, factor is an approved tier','WARN when blank or not found in the house book'),
 ('FORMULA ERRORS',None),('Error count on Deal / Historical Data / Historical Score / Offer Engine / Decision Summary','FAIL if any'),
 ('SAFETY',None),('Adjusted fund <= cash-flow max; adjusted daily <= max daily; no offer-side red at adjusted offer; merchant knockouts untouched','FAIL if any proof fails'),
 ('COHORT QUALITY',None),('Primary cohort met the minimums; seasoned-enough; Z at least Medium; no ISO-vs-industry conflict; tighter tiers skipped; state informational','WARN'),
 ('DATA QUALITY (house book)',None),('Deals loaded; unmapped statuses; no revenue; no ISO / state / SIC; never-paid; Open but fully paid; Open past 1.5x term; weekly term > 39 weeks; position-3 concentration; Lowered Payments treatment; Written-Off collected most of payback','WARN / FAIL with counts'),
 ('RECONCILIATION to the original House Book Analytics workbook',None),('Total funded $34,481,626; collected $27,438,510; principal lost $5,270,163; lost rate 15.3%; payback = advance x factor; seasoned-book size; tier-7 funded = in-window seasoned funded; comparable list count = cohort matches','OK / FAIL')]
row=4
for t,rule in checks:
    if rule is None: hdr(ck,f'B{row}',t,f'B{row}:D{row}')
    else: label(ck,f'B{row}',t); c=ck[f'B{row}']; c.alignment=WRAP; logic(ck,f'C{row}','OK / WARN / FAIL'); logic(ck,f'D{row}',rule); ck.row_dimensions[row].height=45
    row+=1
ck.sheet_view.showGridLines=False

# ---------- README ----------
rd=wb.create_sheet('README',0); widths(rd,{'A':3,'B':130})
txt=["ASPIRE COMBINED UNDERWRITING MODEL  -  OUTLINE (brainstorm phase, formulas not yet built)",
"","WHAT THIS FILE IS","The Deal sheet sections A-E are the live, battle-tested cash-flow tool exactly as built (nothing above row 64 was touched). Every other tab is a blueprint: the layout is real, the values in Controls and Ref are real, and every calculated cell holds an italic purple description of the logic that will go there instead of a formula.",
"","THE THREE LAYERS","Layer 1  Deal A-E: seven metrics + (planned) credit score, green / yellow / red / knockout, score, cash-flow max daily payment (weakest of holdback, total leverage, daily-debt-to-balance), band term, band factor, cash-flow max fund C62. Nothing in Layers 2-3 can exceed it.",
"Layer 2  Historical Data + Historical Score: house-book deals (8 samples here; 1,067 in the build), seasoned = age >= 1.25 x contracted term. Seven matching tiers, new/renewal in every one, industry and ISO above position. Primary cohort = first tier with >= 10 seasoned deals and >= $150k. Credibility Z = n / (n + 30). Blended lost rate -> relative index -> modifier level; confidence caps it.",
"Layer 3  Offer Engine + Decision Summary: fund = MIN(cash-flow max x multiplier, max daily x adjusted term / adjusted factor). Term shortens by bands, factor rises by tiers, weekly can be required. All four offer-side metrics re-checked. Decision = DECLINE / MANUAL REVIEW / CONDITIONAL / APPROVE. Written conclusion built from the actual numbers.",
"","WHAT HISTORY CAN AND CANNOT DO","Can: reduce fund, shorten term, raise factor within 1.40 / 1.45 / 1.49, require weekly, add stips, send to manual review, or allow the full cash-flow max when evidence is strong and credible.","Cannot: remove a knockout, turn a green metric red, exceed the cash-flow max, or make a precise adjustment on a tiny or unseasoned cohort.",
"","WHAT THE HOUSE BOOK SUPPORTS (seasoned deals, after controls)","Robust: renewal vs new, construction trades, Westwood ISO. Composition, not the variable: United Secured Capital 8, position 3. Informational only: state, factor, term, deal size, weekly vs daily, advance/revenue. Never used: NSF count (post-funding).",
"","COLOUR KEY","Blue on yellow = input cell.   Italic purple on lavender = planned logic, to be replaced by a formula.   Dark header = section.   Green-font links to Controls come in the build.",
"","BUILD ORDER WHEN READY","1 Controls values -> 2 Historical Data helper columns -> 3 Historical Score tier stats -> 4 primary tier / modifier / confidence -> 5 Offer Engine -> 6 Decision Summary -> 7 Section F on Deal -> 8 Checks."]
for i,t in enumerate(txt):
    c=rd.cell(2+i,2,t); c.font=B if (t.isupper() or i==0) else N; c.alignment=Alignment(wrap_text=True,vertical='top')
rd.sheet_view.showGridLines=False
wb._sheets=[wb[n] for n in ['README','Deal','Decision Summary','Offer Engine','Historical Score','Comparable Deals','Historical Data','Controls','Checks','Ref']]
wb.active=1
wb.save(OUT); print('saved',OUT)
