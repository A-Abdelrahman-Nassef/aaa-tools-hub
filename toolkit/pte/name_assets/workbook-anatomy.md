# Workbook anatomy

The structure an AAA Pre-Tender Estimate / Tender Evaluation workbook shares.
Read this before the first build; the scripts detect all of it rather than
assuming it, but you need the vocabulary to read what they print and to spot
when a file departs from the pattern.

## Contents
1. The sheets
2. Asset tab column grammar
3. Asset tab row structure
4. Asset Summary grammar
5. The link mapping between them
6. Weight columns
7. Package-level parts that matter

## 1. The sheets

| Sheet | Role |
|---|---|
| Asset Summary | One row per asset, each linking to that asset tab's Total Amounts row. First tab. |
| NRM 1 Summary (BUA) | An NRM1 elemental roll-up. Reads the asset tabs by NRM code. Not built by this skill, but it references them, so renames must follow. |
| Asset #01 (or similar) | The template tender evaluation form. This is what gets renamed and cloned. |
| Hidden legacy tab | Often a stale copy of the form, e.g. `TenderEvaluation_Form_Asse (11)`. Leave it. |

Identify the template by what the summary links to, not by its name. Names vary
between projects.

## 2. Asset tab column grammar

Leading analysis columns, then description columns, then repeating 4-column
blocks — one for AAA, then one per tenderer:

```
A NRM Code | B Weight/Asset | C Weight/Division | D Sr. | E BOQ Item Description |
F Unit | G Designer QTY | H-K AAA | L-O Tenderer #01 | P-S #02 | ... AV-AY #10
```

Each block is `QTY, Rate, Designer-QTY Amount, AAA-QTY Amount`:

| Block | QTY | Rate | 3rd column | 4th column |
|---|---|---|---|---|
| AAA | H | I | J Estimated Amount [Designer QTY] | K Pre-Tender Estimate Amount |
| Tenderer #n | q | q+1 | q+2 As-Submitted Amount | q+3 As-Evaluated Amount |

Amount formulas, for a block whose QTY column is `q`:

- `q+2` = `=$G{row}*{q+1}{row}` — priced against the **Designer** quantity
- `q+3` = `=$H{row}*{q+1}{row}` — priced against the **AAA** quantity

The two QTY anchors are absolute-column so the formulas copy across cleanly.
QTY and Rate are input cells, left blank for the team.

Block count is detected from the 4-column merges on the group-header row (row
4). Do not hard-code 11.

## 3. Asset tab row structure

- Rows 1-5: a width-marker row, spacers, and the two header rows (4 = group
  labels, 5 = sub-labels).
- Body: division headers (Sr. = a letter, Description = `Division NN : Name`)
  each holding `=SUBTOTAL(9,<col><first>:<col><last>)` over their own block,
  with items, section headings and blanks beneath at outline level 1.
- **Total Amounts** row at the end: `=SUBTOTAL(9,<col><body first>:<col><body last>)`
  over the whole body.

Nested `SUBTOTAL(9,…)` ignores inner SUBTOTALs, so the grand total counts item
values once and never double-counts the division subtotals. **Keep SUBTOTAL,
never SUM.**

Each tab has its **own** Total Amounts row — tabs differ in length once real
BOQs are loaded. Detect it; never assume a fixed row number.

## 4. Asset Summary grammar

The condensed form — QTY and Rate dropped, **2 amount columns per block**:

```
A Weight/PTE | B Sr. | C BOQ Item Description (the asset name) |
D Client BUA | E Client Cost/m2 | F IPMS 1 | G Total BUA AAA |
H Building BUA AAA | I Basement BUA AAA | J AAA Cost/m2 |
K-L AAA | M-N Tenderer #01 | ... AE-AF #10
```

The body alternates **spacer row, asset row, spacer row, asset row** and ends
with a spacer then `Total Amounts`. Keep that alternation when adding rows or
the banding breaks.

Per-asset-row formulas:

- `E` = `=IFERROR(L{r}/$D{r}, 0)` — cost per m² on the client's BUA
- `G` = `=SUM(H{r}:I{r})` — building + basement BUA
- `J` = `=IFERROR(L{r}/$G{r}, 0)` — cost per m² on AAA's BUA
- `A` = `=IFERROR(IF($L{r}=0,"",$L{r}/$L${total}),"")` — share of total PTE

Total row: `=SUBTOTAL(9,<col>{first body row}:<col>{last body row})`.

## 5. The link mapping

Summary block *k* pulls form block *k*'s two amount columns. With the standard
grammar that is:

| Summary | Form | Meaning |
|---|---|---|
| K | J | AAA Estimated Amount |
| L | K | AAA Pre-Tender Estimate |
| M, N | N, O | Tenderer #01 As-Submitted / As-Evaluated |
| O, P | R, S | Tenderer #02 |
| … | … | each further block is +2 on the summary, +4 on the form |
| AE, AF | AX, AY | Tenderer #10 |

Every link is `='<Asset tab>'!$<col>${that tab's Total Amounts row}`.

Rather than trusting this table, learn the mapping from whichever summary row is
already linked — it is the authoritative record of how this particular file is
wired, and it survives layout variations.

## 6. Weight columns

- **Weight/Asset** (B) = item's AAA Pre-Tender amount ÷ the asset's grand-total PTE
- **Weight/Division** (C) = item's AAA Pre-Tender amount ÷ its division's PTE subtotal
- **Weight/PTE** on the summary (A) = asset's PTE ÷ the summary grand total

The guard matters. A template may write `=IF(ISBLANK(K{r}),"",…)`, but `K{r}`
holds a formula, so `ISBLANK` is false even before any rate is entered — the
cell reads 0 and the division gives `#DIV/0!`, which cascades into the SUBTOTALs.
Use a **value** guard instead: `=IFERROR(IF($K{r}=0,"",$K{r}/$K${total}),"")`.
Identical once priced, blank rather than an error while unpriced.

When the summary's total row moves, every weight denominator on it must move
too — that is an absolute reference (`$L$28` → `$L$38`), invisible to any
range-based fix-up.

## 7. Package-level parts that matter

| Part | Why you care |
|---|---|
| `xl/workbook.xml` | Tab order, `sheetId`, `state="hidden"`, and thousands of `definedName` entries. `localSheetId` is a **positional index** into the tab order — inserting tabs shifts it. |
| `xl/_rels/workbook.xml.rels` | Maps each sheet's `r:id` to its `worksheets/sheetN.xml` part. The sheet part number does not match tab position. |
| `[Content_Types].xml` | Every new sheet part needs an `Override` entry or the file will not open. |
| `xl/worksheets/_rels/sheetN.xml.rels` | Printer settings, VML drawings. Clones need their own. |
| `xl/sharedStrings.xml` | Labels live here. `count` is total references, `uniqueCount` is the number of `si` elements — both must match reality. |
| `xl/calcChain.xml` | A cached dependency graph. Delete it after surgery and set `fullCalcOnLoad="1"` on `calcPr`. |
| `docProps/app.xml` | `TitlesOfParts` lists sheet names. Excel rewrites it on save; only update the renamed entry, and keep the vector `size` consistent with the elements actually present. |
| `xl/webextensions/`, `xl/drawings/` | Present in real files, dropped by openpyxl. Another reason to edit the XML. |
