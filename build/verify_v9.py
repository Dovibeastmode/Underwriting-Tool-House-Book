import sys, os, shutil, subprocess, openpyxl
sys.path.insert(0,"build")
from v9_main import SRC9, build, LAST

W="/tmp/claude-0/-home-user-Underwriting-Tool-House-Book/15c93af4-845a-5d84-bd2c-2f824a3823d2/scratchpad/v9"
shutil.rmtree(W,ignore_errors=True); os.makedirs(W,exist_ok=True)

PROFILES=[
 ("their example","New Deal",1,"Agricultural Services","United Secured Capital 8"),
 ("common new",   "New Deal",3,"Eating & Drinking Places","GREYSTONE"),
 ("renewal",      "Renewal", 2,"Trucking & Warehousing","BFA 3"),
 ("thin iso+ind", "New Deal",1,"Nonmetallic Minerals Mining","Fidelity Funding"),
 ("no iso/ind",   "New Deal",4,"",""),
 ("whole book",   "ZZZ",    99,"ZZZ","ZZZ"),
]
def recalc(f,od):
    os.makedirs(od,exist_ok=True)
    subprocess.run(["libreoffice","--headless","--convert-to","xlsx","--outdir",od,f],
                   check=True,capture_output=True,timeout=900)
    return os.path.join(od,os.path.basename(f))

def probe(tag, maker):
    out={}
    for i,(nm,nr,pos,ind,iso) in enumerate(PROFILES):
        wb=maker()
        hs=wb["Historical Score"]
        # LibreOffice cannot evaluate XLOOKUP - swap for verification only
        hs["C36"]='=IF(MAX($I$23:$I$26)=0,"none rated",INDEX($B$23:$B$26,MATCH(MAX($I$23:$I$26),$I$23:$I$26,0)))'
        for c,v in [("C5",nr),("C6",pos),("C7",ind),("C8",iso)]:
            hs[c]= v if v!="" else None
        f=os.path.join(W,f"{tag}{i}.xlsx"); wb.save(f)
        v=openpyxl.load_workbook(recalc(f,os.path.join(W,f"{tag}{i}o")),data_only=True)
        h=v["Historical Score"]; d=v["Deal"]
        out[nm]=dict(D29=h["D29"].value,cohort=h["C29"].value,n=h["C30"].value,plr=h["C31"].value,
            rof=h["C32"].value,idx=h["C33"].value,outcome=h["C34"].value,flex=h["C35"].value,
            F51=d["F51"].value,C56=d["C56"].value,C78=d["C78"].value,
            tiers=[(h[f"C{r}"].value,h[f"D{r}"].value,h[f"E{r}"].value,h[f"F{r}"].value,h[f"H{r}"].value) for r in range(14,22)],
            BA=sum((v["Historical Data"][f"BA{r}"].value or 0) for r in range(2,LAST+1)))
    return out

print("original (pre-v8)..."); old=probe("old", lambda: openpyxl.load_workbook(SRC9))
print("v9 scalable...");       new=probe("new", lambda: build(highlight=False)[0])

def f(x): return f"{x:.8f}" if isinstance(x,float) else str(x)
bad=0
for k in old:
    o,n=old[k],new[k]; ok=True; msgs=[]
    for key in ["D29","cohort","n","plr","rof","idx","outcome","flex","F51","C56","C78","BA"]:
        if f(o[key])!=f(n[key]): ok=False; bad+=1; msgs.append(f"{key}: {f(o[key])} -> {f(n[key])}")
    tsame=all(f(a)==f(b) for ta,tb in zip(o["tiers"],n["tiers"]) for a,b in zip(ta,tb))
    if not tsame:
        ok=False; bad+=1
        for i,(ta,tb) in enumerate(zip(o["tiers"],n["tiers"])):
            if [f(x) for x in ta]!=[f(x) for x in tb]:
                msgs.append(f"tier {14+i}: {[f(x) for x in ta]} -> {[f(x) for x in tb]}")
    print(f"{k:<16} {'OK - identical' if ok else 'DIFF'}")
    for m in msgs: print("      ",m)
print("\n"+("ALL RECONCILED" if bad==0 else f"{bad} differences"))
