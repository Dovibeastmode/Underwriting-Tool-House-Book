import sys, os, shutil, openpyxl
sys.path.insert(0,"build")
import v8_lib
from v8_lib import *

W="/tmp/claude-0/-home-user-Underwriting-Tool-House-Book/15c93af4-845a-5d84-bd2c-2f824a3823d2/scratchpad/mm"
shutil.rmtree(W,ignore_errors=True); os.makedirs(W,exist_ok=True)

# Historical Data match cols are contiguous AQ..AT in this order:
#   AQ = position, AR = new/renewal, AS = industry, AT = ISO
# so the flag columns must be re-ordered to MATCH that order.
FLAGS_REORDERED={   # (req Pos, req NR, req Ind, req ISO)
 14:(0,1,1,1),  # NR+ISO+Ind
 15:(1,1,1,0),  # NR+Ind+Pos
 16:(1,1,0,1),  # NR+ISO+Pos
 17:(0,1,1,0),  # NR+Ind
 18:(0,1,0,1),  # NR+ISO
 19:(1,1,0,0),  # NR+Pos
 20:(0,1,0,0),  # NR only
 21:(0,0,0,0),  # whole book
}
COH="--(MMULT(1-'Historical Data'!$AQ$2:$AT$1068,TRANSPOSE($K{r}:$N{r}))=0)"

def apply_mmult(wb):
    restore_baseline(wb["Historical Score"])
    apply_v8(wb)                      # start from the verified v8
    hs=wb["Historical Score"]; hd=wb["Historical Data"]
    for col,h in zip("KLMN",["req Pos","req NR","req Ind","req ISO"]): hs[f"{col}13"]=h
    for r,fl in FLAGS_REORDERED.items():
        for col,v in zip("KLMN",fl): hs[f"{col}{r}"]=v
        C=COH.format(r=r)
        hs[f"C{r}"]=f"=SUMPRODUCT('Historical Data'!$AP$2:$AP$1068*{C})"
        hs[f"D{r}"]=f"=SUMPRODUCT('Historical Data'!$AP$2:$AP$1068*{C}*'Historical Data'!$N$2:$N$1068)"
        hs[f"F{r}"]=f'=IF(D{r}=0,"",SUMPRODUCT(\'Historical Data\'!$AP$2:$AP$1068*{C}*\'Historical Data\'!$AK$2:$AK$1068)/D{r})'
        hs[f"G{r}"]=f'=IF(D{r}=0,"",SUMPRODUCT(\'Historical Data\'!$AP$2:$AP$1068*{C}*\'Historical Data\'!$AJ$2:$AJ$1068)/D{r})'
    # BA: same idea, one row at a time - no MMULT needed
    for r in range(2,1069):
        hd[f"BA{r}"]=(f'=IF(OR(AD{r}=0,\'Historical Score\'!$D$29=""),0,'
                      f"--(SUMPRODUCT(1-AQ{r}:AT{r},'Historical Score'!$K$29:$N$29)=0))")
    hs["C36"]='=IF(MAX($I$23:$I$26)=0,"none rated",INDEX($B$23:$B$26,MATCH(MAX($I$23:$I$26),$I$23:$I$26,0)))'

def v8_ref(wb):
    restore_baseline(wb["Historical Score"]); apply_v8(wb)
    wb["Historical Score"]["C36"]='=IF(MAX($I$23:$I$26)=0,"none rated",INDEX($B$23:$B$26,MATCH(MAX($I$23:$I$26),$I$23:$I$26,0)))'

v8_lib.PROFILES=[
 ("their example","New Deal",1,"Agricultural Services","United Secured Capital 8"),
 ("common new",   "New Deal",3,"Eating & Drinking Places","GREYSTONE"),
 ("renewal",      "Renewal", 2,"Trucking & Warehousing","BFA 3"),
 ("no iso/ind",   "New Deal",4,"",""),
 ("whole book",   "ZZZ",    99,"ZZZ","ZZZ"),
 ("blank",        "",       "", "",""),
]
print("v8 reference..."); ref=probe(SRC,"ref",v8_ref,W)
print("MMULT version..."); new=probe(SRC,"mm",apply_mmult,W)

def f(x): return f"{x:.8f}" if isinstance(x,float) else str(x)
bad=0
for k in ref:
    o,n=ref[k],new[k]; ok=True
    for key in ["D29","cohort","n","plr","rof","idx","outcome","flex","BAsum"]:
        if f(o[key])!=f(n[key]): ok=False; bad+=1; print(f"  {k}/{key}: {f(o[key])} vs {f(n[key])}")
    tiers = all(f(a)==f(b) for ta,tb in zip(o["tiers"],n["tiers"]) for a,b in zip(ta,tb))
    if not tiers:
        ok=False; bad+=1
        for i,(ta,tb) in enumerate(zip(o["tiers"],n["tiers"])):
            if [f(x) for x in ta]!=[f(x) for x in tb]:
                print(f"  {k}/tier row {14+i}: {[f(x) for x in ta]} vs {[f(x) for x in tb]}")
    print(f"{k:<16} {'OK - identical (result block + 8x5 tier table + BA sum)' if ok else 'MISMATCH'}")
print("\n"+("ALL RECONCILED" if bad==0 else f"{bad} mismatches"))
