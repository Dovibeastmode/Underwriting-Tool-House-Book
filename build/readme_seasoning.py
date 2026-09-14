import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT="README_Seasoning_and_Cohorts.xlsx"
INK="FF1F2937"; BAND="FF334155"; SOFT="FFF1F5F9"; WARN="FFFEF3C7"; NOTE="FFEFF6FF"
RULE=Side(style="thin", color="FFCBD5E1")

wb=openpyxl.Workbook(); ws=wb.active; ws.title="README"
ws.sheet_view.showGridLines=False
for col,w in [("A",2.5),("B",4),("C",112),("D",26),("E",46),("F",2),
              ("H",22),("I",16),("J",40),("P",20),("Q",12)]:
    ws.column_dimensions[col].width=w
ws.column_dimensions["P"].hidden=True
ws.column_dimensions["Q"].hidden=True

R=[1]
def nl(n=1): R[0]+=n

def title(txt, sub):
    r=R[0]; ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5)
    c=ws.cell(r,2,txt); c.font=Font(bold=True,size=20,color="FFFFFFFF"); c.alignment=Alignment(vertical="center")
    for cc in range(2,6): ws.cell(r,cc).fill=PatternFill("solid",fgColor=INK)
    ws.row_dimensions[r].height=34; nl()
    r=R[0]; ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5)
    ws.cell(r,2,sub).font=Font(size=10,italic=True,color="FF64748B"); nl(2)

def section(txt):
    nl(); r=R[0]; ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5)
    c=ws.cell(r,2,txt); c.font=Font(bold=True,size=12,color="FFFFFFFF")
    c.alignment=Alignment(vertical="center",indent=1)
    for cc in range(2,6): ws.cell(r,cc).fill=PatternFill("solid",fgColor=BAND)
    ws.row_dimensions[r].height=24; nl()

def head(a,b,c):
    r=R[0]
    for col,t in ((3,a),(4,b),(5,c)):
        x=ws.cell(r,col,t); x.font=Font(bold=True,size=9,color="FF475569"); x.border=Border(bottom=RULE)
    nl()

def line(text, sheet=None, addr=None, note=None, bullet=None, fill=None, bold=False):
    r=R[0]
    if bullet:
        b=ws.cell(r,2,bullet); b.font=Font(bold=True,size=10,color="FF64748B")
        b.alignment=Alignment(horizontal="right",vertical="top")
    c=ws.cell(r,3,text); c.font=Font(size=10,bold=bold,color=INK)
    c.alignment=Alignment(wrap_text=True,vertical="top")
    if sheet:
        ws.cell(r,16,sheet); ws.cell(r,17,addr or "A1")
        d=ws.cell(r,4)
        d.value=('=IF(IFERROR(VLOOKUP($P{r},$H$3:$I$9,2,FALSE),"")="",$P{r}&"!"&$Q{r},'
                 'HYPERLINK("#gid="&VLOOKUP($P{r},$H$3:$I$9,2,FALSE)&"&range="&$Q{r},$P{r}&"!"&$Q{r}))').format(r=r)
        d.font=Font(size=10,color="FF1D4ED8",underline="single",bold=True)
        d.alignment=Alignment(vertical="top")
    if note:
        e=ws.cell(r,5,note); e.font=Font(size=9,italic=True,color="FF64748B")
        e.alignment=Alignment(wrap_text=True,vertical="top")
    if fill:
        for cc in range(2,6): ws.cell(r,cc).fill=PatternFill("solid",fgColor=fill)
    nl()

# ---------------- jump table ----------------
ws["H1"]="JUMP-TO SETUP"; ws["H1"].font=Font(bold=True,size=11,color="FFFFFFFF")
for cc in range(8,11): ws.cell(1,cc).fill=PatternFill("solid",fgColor=BAND)
ws["H2"]="Sheet"; ws["I2"]="gid"; ws["J2"]="how to get it"
for cc in (8,9,10): ws.cell(2,cc).font=Font(bold=True,size=9,color="FF475569"); ws.cell(2,cc).border=Border(bottom=RULE)
for i,(s,hint) in enumerate([
    ("Deal","Click the tab. The URL ends #gid=1234567890 - paste that number here, once."),
    ("Historical Data",""),("Historical Score",""),("Controls",""),("Ref",""),("README","")]):
    ws.cell(3+i,8,s).font=Font(size=10)
    ws.cell(3+i,9).fill=PatternFill("solid",fgColor="FFFFF7CC")
    ws.cell(3+i,9).border=Border(bottom=RULE,top=RULE,left=RULE,right=RULE)
    if hint: ws.cell(3+i,10,hint).font=Font(size=9,italic=True,color="FF64748B")
ws["J9"]="Until a gid is filled in, the link shows as plain text - nothing breaks."
ws["J9"].font=Font(size=9,italic=True,color="FF64748B")

# ---------------- content ----------------
title("Seasoning & the Primary Combination",
      "How the house book decides which past deals count, and which of them resemble the deal in front of you.")

section("1.  SEASONING  -  WHY")
line("We only want to look at seasoned deals when historical data is adjusting the recommended offer. "
     "A deal funded last month has not had time to fail yet. If it sits in the same loss rate as a deal "
     "that ran its full term, it drags the loss rate down and makes the book look safer than it is.")

section("2.  SEASONING  -  HOW IT WORKS")
line("Work out the actual term in business days. MCA Track gives duration in months for a daily deal "
     "and in weeks for a weekly one, in the same column - so the unit has to be inferred.",
     "Historical Data","T","raw duration from the export", bullet="1")
line("The formula tells them apart by size: 12 or more means weeks, under 12 means months. Weeks get "
     "multiplied by 5. Months get multiplied by business days per month.",
     "Historical Data","AA","term, converted to business days", bullet="2")
line("Business days per month is editable - we use 21.655.", "Controls","B13", "260 business days / 12", bullet=None)
line("Work out how much time has actually passed. NETWORKDAYS counts business days from the funded date "
     "to the as-of date, skipping weekends and bank holidays.",
     "Historical Data","AB","age in business days", bullet="3")
line("The as-of date is 'today' as far as the book is concerned - set it to the date the data was pulled.",
     "Controls","B20", None)
line("Compare the two. If the age is smaller than the term, not enough time has passed and the deal is "
     "not seasoned. If the age is larger, it is.", bullet="4")
line("Add a cushion. The bar is not the term itself but the term multiplied by 1.25, so a seasoned deal "
     "has genuinely had time to go wrong rather than only just reaching the finish line.",
     "Controls","B22","seasoning multiple", bullet="5")
line("Override: any deal that has already resolved counts as seasoned regardless of age. I noticed closed "
     "deals that were fully paid off were being left out of the evaluation, which is counterintuitive - "
     "a finished deal is evidence whether it finished early or not.", fill=WARN, bullet="6")
line("Seasoned  =  resolved   OR   age >= 1.25 x term", bold=True, fill=SOFT)

section("3.  FINDING COMPARABLE DEALS")
line("Four columns check, row by row, whether that past deal matches what was entered on the Deal sheet. "
     "Each outputs 1 for a match and 0 for no match.")
head("Column","Jumps to","Checks")
line("AQ", "Historical Data","AQ","same position as the deal being underwritten")
line("AR", "Historical Data","AR","new deal vs renewal")
line("AS", "Historical Data","AS","same industry")
line("AT", "Historical Data","AT","same ISO")
line("Worked example - deal profile: Westwood / Eating & Drinking Places / Renewal / Position 1.", fill=SOFT)
line("We cannot just take every 1 in those columns. That would give us 'comparable' deals that match on "
     "Westwood but are a 7th position in construction. What matters is which deals match on the best "
     "COMBINATION of these, not on any one of them.", fill=WARN)

section("4.  SINGLE-DIMENSION EVIDENCE")
line("Before combining anything, each dimension is measured on its own. Column C counts the 1s for that "
     "one metric, multiplied by the eligibility flag so only seasoned, in-scope deals are counted.",
     "Historical Score","C23","count of matching deals")
line("Column D does the same but sums the funded dollars instead of counting deals. Principal lost rate "
     "and return on funded follow the same pattern.", "Historical Score","D23", None)
line("AP is the eligibility flag - seasoned AND inside the date window. Multiplying by it is what keeps "
     "unseasoned deals out of every number downstream.", "Historical Data","AP", None)

section("5.  THE PRIMARY COMBINATION")
line("There are 7 tiers, each combining a different set of metrics. They are ordered from least common to "
     "most common: Tier 1 is the hardest combination to match, Tier 7 the easiest.",
     "Historical Score","B14","the tier table")
line("The FIRST tier that clears the minimum becomes the primary combination, and the recommendation uses "
     "that tier's numbers. The search starts at Tier 1, so if Tier 1 qualifies we have found the most "
     "specific set of comparable deals available.", bold=True, fill=SOFT)
line("Tier 1 in our example is New/Renewal + ISO + Industry. Only 3 seasoned, in-scope deals match all "
     "three - an unlikely combination, which is exactly why it is Tier 1. Tier 6 is New/Renewal + Position "
     "and matches 14. The tiers get broader as you go down.", fill=SOFT)

section("6.  HOW A TIER IS COUNTED")
line("Take Tier 1, row 14. For each row of the book the formula asks:", "Historical Score","C14", None)
line("Is AP a 1?  Seasoned and in scope. If not, move to the next row.", bullet="1")
line("Is AR a 1?  New/renewal matches. Multiply the two.", bullet="2")
line("Does this tier actually require new/renewal?  That is what the flag table answers.",
     "Historical Score","K13","req NR / req ISO / req Ind / req Pos", bullet="3")
line("K14 is set to 1 by hand because Tier 1 requires new/renewal. Multiply the running result by K14. "
     "If the answer is still 1, the deal is counted.", "Historical Score","K14", None, bullet="4")
line("A flag of 0 makes that dimension drop out of the test entirely - so the same formula covers every "
     "tier, and a new tier is a row of 1s and 0s rather than a new formula.", fill=NOTE)
line("Columns F and G run the identical test, then divide - principal lost rate and return on funded "
     "instead of a count.", "Historical Score","F14", None)

section("7.  MEETING THE MINIMUM")
line("A combination with only a handful of deals behind it should not become the primary combination, "
     "however well it matches. Both thresholds are editable.")
head("Threshold","Jumps to","Set to")
line("Minimum seasoned deals", "Controls","B23","30")
line("Minimum funded dollars", "Controls","B24","$150,000")
line("The first tier whose 'Meets min?' reads Yes becomes the primary combination.",
     "Historical Score","E14", None)

section("8.  SUPPORTIVE, NEUTRAL OR NEGATIVE")
line("The whole outcome rests on one number - the index. It measures how far this combination's principal "
     "lost rate sits from the seasoned book as a whole.")
line("index   =   combination's principal lost rate   /   whole seasoned book's principal lost rate",
     bold=True, fill=SOFT)
line("In our example Tier 7 (New/renewal only) is the primary combination. Its principal lost rate is 8.5% "
     "across 127 deals - all seasoned renewals. The seasoned book runs at 16.6%. So the index is "
     "8.5 / 16.6 = 0.51.", "Historical Score","C33", None, fill=SOFT)
line("An index of 1.00 means this combination performs exactly like the book. Lower is better - it is "
     "losing less than the book. Higher is worse.")
head("Outcome","Jumps to","Index")
line("SUPPORTIVE", "Controls","B27","at or below 0.85")
line("NEUTRAL",    None,None,"between the two")
line("NEGATIVE",   "Controls","B28","above 1.15")
line("Both cut-offs are editable, so how much evidence it takes to move an offer is a setting, not "
     "something buried in a formula.")

wb.save(OUT); print("saved", OUT, "| rows:", R[0])
