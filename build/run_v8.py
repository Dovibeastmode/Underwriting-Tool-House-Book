import sys, os, shutil, openpyxl, json
sys.path.insert(0,"build")
from v8_lib import *

W="/tmp/claude-0/-home-user-Underwriting-Tool-House-Book/15c93af4-845a-5d84-bd2c-2f824a3823d2/scratchpad/v8"
shutil.rmtree(W,ignore_errors=True); os.makedirs(W,exist_ok=True)

def base_only(wb): restore_baseline(wb["Historical Score"])
def v8(wb):
    restore_baseline(wb["Historical Score"])   # same starting point
    apply_v8(wb)
    wb["Historical Score"]["C36"]='=IF(MAX($I$23:$I$26)=0,"none rated",INDEX($B$23:$B$26,MATCH(MAX($I$23:$I$26),$I$23:$I$26,0)))'

print("recalculating baseline (v7)...")
old=probe(SRC,"old",base_only,W)
print("recalculating v8...")
new=probe(SRC,"new",v8,W)

def fnum(x):
    if x is None: return "None"
    if isinstance(x,float): return f"{x:.6f}"
    return str(x)

bad=0
for name in old:
    o,n=old[name],new[name]
    print(f"\n=== {name}")
    for k in ["D29","cohort","n","plr","rof","idx","outcome","flex"]:
        same = fnum(o[k])==fnum(n[k])
        if not same: bad+=1
        print(f"   {k:<9} old={fnum(o[k]):<34} new={fnum(n[k]):<34} {'OK' if same else '*** DIFF ***'}")
    ts = all(fnum(a)==fnum(b) for ta,tb in zip(o["tiers"],n["tiers"]) for a,b in zip(ta,tb))
    if not ts: bad+=1
    print(f"   tier table (8 rows x 5 cols)  {'OK - identical' if ts else '*** DIFF ***'}")
    if not ts:
        for i,(ta,tb) in enumerate(zip(o["tiers"],n["tiers"])):
            if [fnum(x) for x in ta]!=[fnum(x) for x in tb]:
                print("      row",14+i,"old",[fnum(x) for x in ta]); print("                new",[fnum(x) for x in tb])
    print(f"   BA column sum  old={o['BAsum']}  new={n['BAsum']}  {'OK' if o['BAsum']==n['BAsum'] else '*** DIFF ***'}")
    if o['BAsum']!=n['BAsum']: bad+=1

print("\n"+("ALL RECONCILED" if bad==0 else f"{bad} MISMATCHES"))
