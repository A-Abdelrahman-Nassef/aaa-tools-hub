# NRM 1 Summary (BUA) — sheet anatomy

The structure of the elemental summary in an AAA Pre-Tender Estimate / Tender
Evaluation workbook. The scripts detect all of this rather than assuming it, but
you need the vocabulary to read what they print and to spot a file that differs.

## Columns

```
B  NRM code (#)          e.g. 0.1.1, 8.8.3 — the lookup key, stored as text
C  Description           e.g. "Facilitating Works", "Substructure"
then one 3-column block per asset:
   Amount     XLOOKUP into that asset's tab
   Cost/m²    Amount / that asset's BUA
   %          Amount / that group's grand total
then a narrow spacer column, then defaults
```

First asset is D/E/F, second G/H/I, third J/K/L, and so on in steps of three.
A new group's columns are inserted before the spacer, which shifts right.

## Header rows

| Row | Holds |
|---|---|
| 1 | Column-width markers (plain numbers, not data) |
| 2 | Asset name merged across Amount:Cost/m², and **that asset's BUA in the % cell** |
| 3 | `Amount` / `Cost/m²` / `%` |
| 4 | `SAR` / `SAR/m²` / `%` — also the autofilter row |

Row 2's % cell is the divisor for the whole Cost/m² column. It carries a
percentage number format inherited from the template, so a BUA of 18,448.22
displays as `1,844,822%`. That is cosmetic and pre-existing — do not "fix" it
without asking, since changing it alters a cell the user reads.

## Body

- ~163 NRM code rows carrying the lookups, at outline level 1 or 2, mostly hidden
- ~48 section rows with `SUBTOTAL(9,…)` over their own block
- ~10 single-cell `SUBTOTAL(9,X)` rows
- `Total (Excl. Preliminaries)` near row 244: `=SUBTOTAL(9,<col>5:<col>241)`

Nested `SUBTOTAL(9,…)` ignores inner SUBTOTALs, so the total counts each amount
once. Keep SUBTOTAL, never SUM.

## Formulas, for a group whose Amount column is `A`, Cost/m² `C`, % `P`

```
A{r}  =_xlfn.XLOOKUP($B{r},'<tab>'!$A:$A,'<tab>'!$K:$K,0,0,1)
C{r}  =A{r}/P${hdr}
P{r}  =IFERROR(A{r}/A${total},0)
```

`_xlfn.` is how Excel stores functions newer than the file format; it is correct
on disk and must be written that way.

The lookup reads column A of the asset tab (its NRM Code column) and returns
either column J (Estimated Amount, priced on Designer quantities) or column K
(Pre-Tender Estimate, priced on AAA quantities). Which one is a decision for the
user, not a default.

## Known quirks in the wild

- **Two groups reading one tab.** A duplicated template group keeps the first
  asset's tab reference while its header says otherwise. Both columns then show
  identical figures.
- **Latent `#REF!`** around row 241: `=IF(ISBLANK(D241),"",D241*#REF!)`. Blank
  while its input is empty, broken the moment it is filled. Never copy it into a
  new group; leave those cells empty and report it.
- **Renaming an asset tab breaks 163 lookups per group.** Excel repoints them
  when you rename in the UI; editing XML does not. Step 01 sweeps for this.

## Package parts this sheet touches

| Part | Why |
|---|---|
| `xl/worksheets/sheetN.xml` | `cols`, `mergeCells`, `dimension`, per-row `spans`, `autoFilter`, outline `hidden` rows |
| `xl/sharedStrings.xml` | Asset labels. `count` = references, `uniqueCount` = `<si>` elements; both must match reality or Excel rejects the file |
| `xl/workbook.xml` | `Print_Area` for this sheet, keyed by positional `localSheetId`; set `fullCalcOnLoad="1"` after editing |
| `xl/calcChain.xml` | Stale after surgery — delete it if present and let Excel rebuild |
