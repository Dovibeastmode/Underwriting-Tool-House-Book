# House Book — Session Log

Aspire Funding Platform · analytics workbook v11 → v15, presentation v1 → v11
Window used throughout unless stated: deals funded 11/01/2024 – 06/30/2026 (the deck) or – 08/25/2026 (the workbook).

---

## 1. Deliverables produced

| File | What it is |
|---|---|
| `House_Book_Analytics_v12.xlsx` | Remediation of every audit finding against v11 |
| `House_Book_Sample_8_Deals_v12.xlsx` | Matching 8-row sample |
| `House_Book_Analytics_v13.xlsx` | Threshold split + Scenario tab (Scenario later removed) |
| `House_Book_Analytics_v14.xlsx` | Scenario deleted, dimension takeaways added |
| `House_Book_Analytics_v15.xlsx` | Principal lost rate replaces principal default exposure |
| `Aspire_House_Book_Dimensions_v11.pptx` | 14-slide deck: title, decided book, TL;DR, 11 dimensions |

---

## 2. Metric definitions as they now stand

**Principal lost rate** (replaced principal default exposure in v15)
`principal lost on defaulted deals ÷ total funded`, where principal lost per deal is `MAX(advance − collected, 0)`.
Full book 15.3%; deck window, all deals, 17.0%; deck window, decided deals only, 26.6%.
Reads as: of every dollar funded, this share has not come back. Collections to date are netted in.

**Return on funded capital**
`(collected − advance) ÷ advance`. On the deck it excludes deals with status Open.

**Principal default exposure** (removed in v15)
`funded on defaulted deals ÷ total funded`. Measured how often we picked a bad deal, weighted by size; ignored collections entirely, which is why it was replaced.

**Collected on defaults** (was "recovery rate")
`collected on defaulted deals ÷ funded on defaulted deals`. Renamed because most of that money is regular debits before default, not recovered by a collections effort — and because values above 100% read as absurd under the word "recovery."

**Closed family** — Closed, Closed/NSF, Closed/NSF Unpaid. Read from the flag in Data!AJ, never from a status string.

**Resolved** — Closed family + Written-Off + Bankruptcy + Settlement Agreement. Excludes Legal and In House Collection, which are still collecting.

**Decided / ex-open** — every status except Open. This is the deck's basis. Note it is *not* the same as finished: it counts Legal and In House Collection as complete.

---

## 3. Headline figures

### Full book (11/01/24 – 08/25/26)
- 1,067 deals · $34,481,626 funded · $27,438,510 collected · net **−$7,043,116**
- Principal lost rate 15.3% · collected on defaults 40.9% · $5,270,163 lost
- Return on funded capital −20.4% · on resolved capital +37.6%
- Commission $3,895,665 gross · 15 deals never paid ($276,500)

### Deck window (11/01/24 – 06/30/26)
- 994 deals · $30,540,626 funded · principal lost rate 17.0%

### The decided book (deck window, excluding 234 open deals)
- 760 deals · $19,467,626 funded · $18,651,252 collected (66.5% of contracted payback)
- Net cash **−$816,374** · return −4.2% · principal lost rate 26.6% ($5,184,923)

**Where it splits:**

| Segment | Deals | Funded | Share | Net cash | Return | Collected, % of payback |
|---|---|---|---|---|---|---|
| Paid off (Closed family) | 387 | $10,228,126 | 52.5% | +$4,384,377 | +42.9% | 100.0% |
| In collections (Legal + In House) | 317 | $7,849,500 | 40.3% | −$4,817,047 | −61.4% | 26.5% |

The two segments are 93% of decided dollars. The entire result is the gap between them.

### The collections question
Legal + In House have collected 26.5% of contracted payback. Deals that reached a terminal status ended at 36% (Bankruptcy), 50% (Settlement), 64% (Written-Off). Recovery on live collections rises with age: 16.4% of payback under 6 months, 29.4% at 6–12 months, 36.3% at 12+ months.

If collections reach 50% of payback, that is **$2,695,327 more collected** — net cash **+$1,878,953**, return **+9.7%**. So −4.2% is a floor, not a verdict, and the difference between 26.5% and 50% is the difference between the house book losing money and making it.

---

## 4. Findings by dimension (deck window)

| Dimension | Best | Worst | Action |
|---|---|---|---|
| Position | 7th (2.9% lost), 1st (+$199K) | 3rd — 466 deals, 48% of dollars, −$1.13M | Stop 3rd position to new merchants; it exceeds the whole book's loss |
| New vs renewal | Renewal: 12.1% lost, +$346K | New: 18.4%, −$1.16M | Fund renewals first |
| ISO | Westwood 1.8% lost, +$366K | PMF −$474K, USC8 −$321K | Grow Westwood/Premium; review PMF and USC8 |
| Industry | Personal Services +$139K, Legal Services +$168K | General Building −$407K, Special Trade −$354K | Tighten construction trades |
| Funding year | 2025 +$660K | 2026 −$1.45M | Same lost rate both years (16.9 / 17.0) — the gap is unfinished collections |
| Factor | 1.46–1.48 (thin, 42 deals) | Above 1.49 at 33.4% lost | Higher pricing is a warning sign, not a fix |
| Deal size | — | — | All buckets −8% to 0%; manage as concentration, not credit |
| Advance to revenue | — | Under 0.15x at 26.0% lost (backwards) | Verify the revenue data before using as a screen |
| State | — | — | No state carries the volume to act on |
| Term | — | — | The gradient is age and size, not term |
| Payment frequency | — | — | Weekly deals are just larger and longer |

---

## 5. The confounds — read this before quoting any single-variable finding

**Age is the dominant confound, not size.** 46% of 150+ day deals and 55% of $100K+ deals were still open on the full book. An open deal has not finished paying, so its return is negative by construction.

Controlled against a settled window (deals funded to 12/31/25):

| Finding | Full book | Settled window | Verdict |
|---|---|---|---|
| Shorter term is better | −6.5% → −26.4% | reverses: −1.0% → +15.0% | Did not survive |
| Smaller deals are better | −4.9% → −28.4% | flat: 0.0% → +6.4% | Did not survive |
| Position 3 is worst | −28.1% | worst in every size and term bucket | **Survived** |
| Renewals are worse | −25.5% | reverses: +22.0% vs +3.9% | **Survived, sign flipped** |

**Neither metric is age-proof.** Return excluding open deals is still contaminated through the collections book — Legal deals under 6 months show 76% lost against advance, at 12+ months 49.2%, same decisions. Principal lost rate falls the same way as recoveries arrive. The only real control is comparing vintages or restricting the window.

**Seasoning control used in v13** (since removed with the Scenario tab): a deal is seasoned once its age in business days ≥ 1.25 × contracted term.

---

## 6. Controls reference

| Cell | Control | Scope |
|---|---|---|
| B5 / B6 | Date window | Everything. Also the practical seasoning control: set B6 to 12/31/25 to read the settled book |
| B17 | Selection mode | Industry, State, ISO ranking only |
| B18 | Top N | Industry, State, ISO only |
| B19 | Minimum deals before a **rate** is shown | Everywhere. Blanks rate cells; counts and dollars still show |
| B20 | Minimum deals for a category to **appear at all** | Industry, State, ISO only (added v13) |
| B21 | Exclude unknown/blank categories | No SIC, No state, No ISO, No revenue data |
| B22 | Minimum deals before a category is named in a Verdict sentence | Verdict narrative only |

**B19 vs B20:** B20 decides whether a category exists on the chart; B19 decides whether it gets to show a percentage. B19 exists because exhaustive dimensions (position, size, status) cannot drop a bucket without the block failing to add up to the book.

Unranked blocks — position, deal size, factor, term, month/quarter/year, advance-to-revenue, new vs renewal, status, payment frequency — respond only to the date window and B19.

---

## 7. Calc master block columns

Every row is one category; every column asks the Data tab one question, scoped to the date window.

A label · B funded $ · C collected $ · **D net (C−B)** · E payback · F commission · G principal lost rate · H collected on defaults · **I return (D/B)** · J deals · **K share of book (B/BT6)** · **L avg deal size (B/J)** · M avg term business days · N avg term payments · O avg days on book, defaulted · P avg days to close · Q renewal share by count · R renewal share by dollars · S/T/U book reference lines (copies of BT2/BT3/BT4) · V default rate by count · W chart label (`name (deals)`).

Rates (G, H, I, L, M, N, O, P) blank below B19; counts and dollars always show. Bold columns are derived from neighbours rather than re-running SUMIFS.

Book-level figures live once in **BS1:BT6**; every reference line and Verdict tile reads them.

---

## 8. Audit findings fixed (v11 → v12)

**Broke the numbers**
1. Verdict B22/B23 used exact-match `"Closed"`, dropping $104,834 of NSF deals → repointed to the Closed-family flag.
2. Verdict B43/C43 and the industry self-check read Data!AX (advance/revenue bucket) instead of AY (industry) → "No SIC" reported 0 deals forever and the self-check reported a permanent false mismatch masked by a `MAX(0,…)` wrapper.
3. Data!AI (Resolved) was computed and read by nothing → added Return on resolved capital and Return on funded capital tiles.

**Misleading**
4. Commission "toggle" language in five places, describing a toggle that does not and must not exist.
5. Shipped window ended 02/28/26, showing +$181,686 on a book that lost $7.0M.
6. Narrative named categories below the naming threshold; top-10 hard-coded.
7. Degenerate sentences naming one category as both best and worst.
8. Float-fragile self-checks reading MATCH in one engine and MISMATCH in another.
9. Map Data had no minimum-deal suppression (VT: 1 deal, 100%), no recovery, no "No state" row, and NY was missing entirely.
10–16. Deal counts absent from rate charts; unlabelled reference lines; tautological status columns; blanks toggle missing "No revenue data"; three non-Reserve-Bank holidays in the ACH table; stale notes; undocumented sample divergence.

**Never audited (Part 4)** — accessibility failed on chart colour only (fixed to greyscale-safe fills); Calc helper columns unheaded (fixed); four columns re-ran SUMIFS for what neighbours already gave (fixed); one duplicate figure on Verdict (removed).

---

## 9. Corrections made during the session

Recorded because each one was quoted before it was checked:

- **2024 lost rate**: said 25.1% / $15,290 → actually **46.4% / $28,303**. Also said all 6 deals defaulted → **5 of 6**; the sixth closed paid-in-full, and two of the five collected more than their advance.
- **No-revenue deals**: said 149 deals / $3,928,626 → actually **152 deals / $2,212,626**. Only 4 are open (77 Closed, 58 Legal, 8 Written-Off, 2 Closed/NSF Unpaid, 2 In House Collection, 1 Bankruptcy). *The note typed onto the advance-to-revenue slide still carries the wrong figure.*
- **"Renewals are the best paper"** cited on the renewal-share chart: that chart measures share, not performance, and the claim came from the deleted Scenario tab's seasoned view. On the full book renewals were worse.
- **"Lost rate is immune to age"**: wrong. It falls as defaulted deals keep collecting — Legal deals show 76% lost under 6 months, 49% at 12+.
- **"Return is contaminated by open deals"**: wrong as stated, since the deck's return column already excludes them. The contamination runs through the collections book instead.
- **Resolved as the honest alternative basis**: wrong. Resolved holds only 34 defaults (6.8% of its dollars) against 352 on the ex-open basis, so +37.7% is a survivorship figure.

---

## 10. Open items

1. **Hand-edited v12 was never uploaded** — Controls values and typed text could not be diffed. Going forward: user owns Controls values and typed text, Claude owns structure and formulas.
2. **Write-off has no definition.** All 12 Written-Off deals collected more than their advance (114% of principal), so the status marks "stopped pursuing," not "lost." Collections workflow needs this defined before the bucket can be read.
3. **154 deals have no revenue figure** — advance-to-revenue is the strongest screen and 14% of the book cannot be scored on it.
4. **Advance-to-revenue gradient inverted** in the 06/30 window (under-0.15x worst at 26.0%). Either the revenue data is wrong on those 125 deals or something real is happening with small-advance merchants.
5. **PMF ramp** — 76 deals total, but 40 of them written May–July 2026. Half the relationship is three months old.
6. **The generator was never updated.** v12–v15 are patched files. A rebuild from `gen.py` reverts every change in this session.
7. **v15 metric change is workbook-only** — if pasting Calc into an existing file, the Controls B17 dropdown list must be edited to the new mode names or those two modes silently do nothing.
8. **Three implausible durations** — 5674174148 (56 weeks, $160,000), 5674178327 (100 weeks, $25,000), 5674178331 (90 weeks, $20,000).

---

## 11. What to do next

**Underwriting**: no 3rd position to new merchants; renewals first; construction trades tightened; PMF and United Secured Capital 8 under review before further funding. Do not set term, size, or factor caps — the data does not support them.

**Collections**: this is where the book is decided. $7.85M sits in Legal and In House Collection at 26.5% of payback collected, of which roughly $4.9M is principal. Comparable terminal outcomes land at 36–64%. That gap is the argument for the collections workflow.

**Reporting**: use the date window as the seasoning control — B6 at 12/31/25 answers "what should we write," B6 at today answers "what are we exposed to." Those are two different meetings.
