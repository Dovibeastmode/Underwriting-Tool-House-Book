import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.utils import get_column_letter

SRC="/root/.claude/uploads/15c93af4-845a-5d84-bd2c-2f824a3823d2/2319ea44-Underwriting_Tool_Combined_with_House_book_analytics_2.xlsx"
OUT="Aspire_v10_README_and_AU_fix.xlsx"
LAST=1068

INK   = "FF1F2937"   # near-black slate
BAND  = "FF334155"   # section bar
SOFT  = "FFF1F5F9"   # light row fill
WARN  = "FFFEF3C7"   # amber note
BLUE  = "FFCCE5FF"   # "changed cell" highlight
RULE  = Side(style="thin", color="FFCBD5E1")

wb=openpyxl.load_workbook(SRC)

# ---------------- 1. fix the broken cohort column ----------------
hd=wb["Historical Data"]
DIMS=[("K","AR"),("L","AT"),("M","AS"),("N","AQ")]
for r in range(2,LAST+1):
    prod="*".join(f"({m}{r}*'Historical Score'!${f}$29+1-'Historical Score'!${f}$29)" for f,m in DIMS)
    c=hd[f"AU{r}"]; c.value=f"=IF(AD{r}=1,{prod},0)"; c.fill=PatternFill("solid",fgColor=BLUE)
hd["AU1"]="In primary cohort"; hd["AU1"].fill=PatternFill("solid",fgColor=BLUE)

# ---------------- 2. rebuild the README ----------------
del wb["README"]
ws=wb.create_sheet("README", 0)
ws.sheet_view.showGridLines=False
ws.column_dimensions["A"].width=3
ws.column_dimensions["B"].width=4
ws.column_dimensions["C"].width=104
ws.column_dimensions["D"].width=22
ws.column_dimensions["E"].width=52
for col in "FGH": ws.column_dimensions[col].width=2

R=[1]
def row(): return R[0]
def nl(n=1): R[0]+=n

def title(text, sub=""):
    r=row()
    ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5)
    c=ws.cell(r,2,text); c.font=Font(bold=True,size=20,color="FFFFFFFF"); c.alignment=Alignment(vertical="center")
    for cc in range(2,6): ws.cell(r,cc).fill=PatternFill("solid",fgColor=INK)
    ws.row_dimensions[r].height=34; nl()
    if sub:
        r=row(); ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5)
        c=ws.cell(r,2,sub); c.font=Font(size=10,italic=True,color="FF64748B"); nl()
    nl()

def section(text):
    nl()
    r=row()
    ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5)
    c=ws.cell(r,2,text); c.font=Font(bold=True,size=12,color="FFFFFFFF")
    c.alignment=Alignment(vertical="center",indent=1)
    for cc in range(2,6): ws.cell(r,cc).fill=PatternFill("solid",fgColor=BAND)
    ws.row_dimensions[r].height=24; nl()

def link(cellref, target):
    """cellref like 'Deal!C72' -> a clickable cell"""
    sheet,addr = target.split("!")
    loc = f"'{sheet}'!{addr}" if " " in sheet else f"{sheet}!{addr}"
    return cellref, loc

def line(text, cellref=None, note=None, bullet=None, wrap=True, fill=None, bold=False):
    r=row()
    if bullet:
        b=ws.cell(r,2,bullet); b.font=Font(bold=True,size=10,color="FF64748B"); b.alignment=Alignment(horizontal="right")
    c=ws.cell(r,3,text)
    c.font=Font(size=10,bold=bold,color=INK)
    c.alignment=Alignment(wrap_text=wrap,vertical="top")
    if cellref:
        label,loc = cellref
        d=ws.cell(r,4,label)
        d.hyperlink=Hyperlink(ref=f"D{r}", location=loc)
        d.font=Font(size=10,color="FF1D4ED8",underline="single",bold=True)
        d.alignment=Alignment(vertical="top")
    if note:
        e=ws.cell(r,5,note); e.font=Font(size=9,italic=True,color="FF64748B")
        e.alignment=Alignment(wrap_text=True,vertical="top")
    if fill:
        for cc in range(2,6): ws.cell(r,cc).fill=PatternFill("solid",fgColor=fill)
    if wrap and len(text)>100: ws.row_dimensions[r].height=None
    nl()

def head(a,b,c):
    r=row()
    for col,txt in ((3,a),(4,b),(5,c)):
        cc=ws.cell(r,col,txt); cc.font=Font(bold=True,size=9,color="FF475569")
        cc.border=Border(bottom=RULE)
    nl()

# ============================ CONTENT ============================
title("Aspire Underwriting Tool",
      "Cash-flow analysis + house-book evidence -> a recommended offer.   Blue cell references are clickable.")

section("1.  WHAT THIS TOOL DOES")
line("Combines a cash-flow analysis of the merchant's bank data with the performance of "
     "1,067 past deals in our own book, and returns a recommended offer.")
line("Every assumption and threshold is editable in one place - nothing is hard-coded in a formula.",
     link("Controls","Controls!A1"))

section("2.  THE THREE LAYERS")
head("Layer","Where","What it decides")
line("Layer 1 - Cash flow sets the ceiling. Bank data and your green/red thresholds "
     "produce the largest daily payment this merchant can carry.", link("Deal!B36","Deal!B36"),
     "Sections B, C, E")
line("Layer 2 - The house book finds the closest matching cohort of past deals and compares "
     "its loss rate to the book average.", link("Historical Score","Historical Score!A1"),
     "Cohort match + index")
line("Layer 3 - The index moves the sizing line, which re-sizes the offer up or down.",
     link("Deal!B75","Deal!B75"), "Section F")
line("Layer 1 always sets the ceiling. Layer 2 never invents evidence - if no cohort qualifies, "
     "the offer is pure cash flow.", fill=SOFT, bold=True)

section("3.  WORKFLOW")
line("Enter the deal profile: ISO, industry, new/renewal, credit score.", link("Deal!K12","Deal!K12"),
     "Drives the cohort match", bullet="1")
line("Enter bank data: revenue, average daily balance, negative days, NSFs - newest month first.",
     link("Deal!B11","Deal!B11"), "Rows 12-23", bullet="2")
line("Enter existing positions: payment, frequency.", link("Deal!B27","Deal!B27"), "Rows 28-33", bullet="3")
line("Set your thresholds if you haven't already. GREEN is the most you'd consider comfortable "
     "for that metric; RED is where it becomes a problem; KNOCKOUT is an automatic decline.",
     link("Deal!D37","Deal!D37"), "D = green, E = red, G = knockout", bullet="4")
line("Set how many reds force a decline, and the smallest fund you'd write.",
     link("Deal!C61","Deal!C61"), "C61 reds, E61 minimum $", bullet="5")
line("Read the answer in Section F.", link("Deal!B75","Deal!B75"), "Final decision + offer", bullet="6")

section("4.  SECTION F - READING THE ANSWER")
head("Row","Cell","What it means")
line("Cash flow only - what the bank data alone supports, at your green lines.",
     link("Deal!C77","Deal!C77"), "Reference point. Not the recommendation.")
line("RECOMMENDED FUND - the same calculation at the sizing line, after the house book has moved it.",
     link("Deal!C78","Deal!C78"), "This is the number to quote.")
line("HOUSE-BOOK OUTCOME - Supportive, Neutral or Negative.",
     link("Deal!C84","Deal!C84"), "Drives the sizing flex")
line("If the two funds are equal, the house book was Neutral or had no evidence.", fill=SOFT)

section("5.  HOW THE CASH-FLOW OFFER IS CALCULATED")
line("Three separate ceilings on the daily payment. Each one asks the same question a different way: "
     "how much can this merchant hand over per day before this metric leaves my comfort zone?")
line("The live example on the Deal sheet: revenue $100,000/mo, average balance $10,000, "
     "existing debt $5,000/week ($1,000/day).", fill=SOFT)
head("Ceiling","Cell","Calculation")
line("Total LVG - all debt against revenue, so existing positions are subtracted.",
     link("Deal!F68","Deal!F68"), "$100,000 x 30% - $21,655, / 21.655 = $385.36/day")
line("Holdback SP% - the new payment against revenue, ignoring existing debt.",
     link("Deal!F67","Deal!F67"), "$100,000 x 18% / 21.655 = $831.22/day")
line("Daily debt / balance - the total daily burden against the cash actually sitting there.",
     link("Deal!F69","Deal!F69"), "$10,000 x 18% - $1,000 = $800/day")
line("MAX DAILY takes the weakest of the three - all three have to hold at once. "
     "Here that is Total LVG at $385.36.", link("Deal!F70","Deal!F70"),
     "MIN of the three, floored at 0")
line("Why 21.655: business days per month (260 business days / 12). Weekly figures use 4.331 weeks/month.",
     link("Controls!A11","Controls!A11"), "Section 2 - conversion constants")

section("6.  FACTOR AND TERM")
line("A point score across the metric rows: GREEN = 1, YELLOW = 0.5, RED = 0.",
     link("Deal!C65","Deal!C65"))
line("The score picks a row from the pricing ladder. More greens = longer term and lower factor.",
     link("Controls!A3","Controls!A3"), "Section 1 - pricing bands")
line("Term is then capped three ways: the band's term cap, the 9-month maximum, "
     "and the fund/revenue green line.", link("Deal!C72","Deal!C72"), "MIN of the three, floor 20 payments")
line("FUND = daily payment x term / factor, rounded down to the nearest $500.",
     link("Deal!C73","Deal!C73"), "$385.36 x 190 / 1.40 = $52,000")

section("7.  HOW THE HOUSE BOOK ADJUSTS IT")
line("The tool finds the most specific cohort of past deals matching this merchant that still "
     "has enough deals behind it to be believable.", link("Historical Score!B13","Historical Score!B13"),
     "Tier table, most specific first")
line("A tier is defined by ticking which dimensions must match. 1 = must match, 0 = ignore.",
     link("Historical Score!K13","Historical Score!K13"), "The flag table")
line("Index = cohort loss rate / whole-book loss rate. 1.00 means this cohort performs exactly "
     "like the book. Below 1 is better than the book.", link("Historical Score!C33","Historical Score!C33"))
line("Supportive moves the sizing line toward the red line (bigger offer). "
     "Negative moves it below the green line (smaller offer).",
     link("Controls!A26","Controls!A26"), "Section 4 - thresholds and flex")
line("HAS EVIDENCE decides whether the house book gets a vote at all. If the match fell through "
     "to the whole book, the answer is NO and the offer is pure cash flow.",
     link("Historical Score!C37","Historical Score!C37"), "Read by 7 cells", fill=WARN)

section("8.  WHERE TO CHANGE THINGS")
head("I want to change...","Go to","Note")
line("Factor and term for each score band", link("Controls!A3","Controls!A3"), "Section 1")
line("Business days / weeks per month", link("Controls!A11","Controls!A11"), "Section 2")
line("Which deals count as evidence - date window, seasoning, cohort minimums",
     link("Controls!A19","Controls!A19"), "Section 3")
line("Supportive / Negative cut-offs and how far the sizing line moves",
     link("Controls!A26","Controls!A26"), "Section 4")
line("Which CRM statuses count as a default", link("Controls!A32","Controls!A32"), "Section 5")
line("Green / red / knockout thresholds per metric", link("Deal!D37","Deal!D37"), "On the Deal sheet")

section("9.  ADDING A COHORT TIER")
line("Insert a row inside the tier table - never at the bottom edge, or the ranges won't expand.",
     link("Historical Score!B14","Historical Score!B14"), bullet="1")
line("Type the cohort name in column B.", bullet="2")
line("Tick the flags in K:N. Leaving them blank means 'ignore every dimension', "
     "which returns the whole book.", link("Historical Score!K14","Historical Score!K14"), bullet="3")
line("Drag C, D, E, F, G, H down from the row above. The formula is identical in every row.", bullet="4")
line("Put it ABOVE 'New/renewal only'. That row matches almost everything, so anything below it "
     "is never reached.", link("Historical Score!B20","Historical Score!B20"), fill=WARN, bullet="5")
line("Check: the whole-book row must stay last, and the fallback row number should match your tier count.",
     link("Historical Score!D30","Historical Score!D30"), bullet="6")

section("10.  BEFORE YOU TRUST A CHANGE")
line("Any edit to the cohort engine: note the whole-book seasoned count (716) and the "
     "'New/renewal only' count before you start. If either moves, you have a typo.", fill=SOFT)
line("The sizing flex has a safety property - a flag of 0 multiplies by exactly 1, so adding a "
     "dimension with all flags at 0 cannot change any existing number.")

# borders + tidy
for r in range(1,row()+1):
    for c in range(2,6):
        cell=ws.cell(r,c)
        if cell.alignment and cell.alignment.wrap_text is None:
            pass
ws.freeze_panes="A2"
ws.sheet_properties.tabColor="1F2937"
wb.save(OUT)
print("saved",OUT,"| README rows:",row())
