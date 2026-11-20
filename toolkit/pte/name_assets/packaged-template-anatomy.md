# Packaged template anatomy (CLIENT / PROJECT R00 family)

The newer workbook. Recognise it by the two template tabs it ships:
`Asset #01 {no package}` and `Asset #02 {with package}`.

## Sheets and the chain between them

```
Scheduale of Assets        asset name (col D) + area data (E..R)
        |  name lookup
Elemental Summary (NRM01)  a 3-column group per asset; BUA pulled from the
        |                  Scheduale by INDEX/MATCH on the name
Asset Summary              a row per asset, linked to that asset tab's own
        |                  Total Amounts row
Project Summary            the same rows x a Repetition factor
```

Also present and not touched by this skill: `Project Information`,
`Pricing References`.

Because the elemental sheet finds each asset's BUA by **matching the name**
against `Scheduale of Assets` column D, the asset name must be written
identically in the tab name, the Scheduale, the Asset Summary, the Project
Summary and the elemental header. One build writes all five, which is the point.

## The packaged asset tab

A package block is exactly **99 rows** and repeats:

```
P        package header    D = letter (A, B, C...), E = package name
                           SUBTOTAL over P+1 : P+97
P+2      division 01       D = "<letter>.01", SUBTOTAL over P+3 : P+5
P+6      division 02       ... 24 divisions, every 4 rows, last at P+94
P+98     blank separator
P+99     next package header
8+99N    Total Amounts     SUBTOTAL over row 6 : the row above
```

First package header is row 8. An asset with N packages ends at row `8 + 99N`,
so every asset tab has a **different** Total Amounts row and the Asset Summary
link must be built per asset, never assumed.

An asset with no packages uses `Asset #01 {no package}` instead: divisions
directly under the header, total at row 105.

## Column grammar (both asset templates)

```
A NRM Code | B Weight/Asset | C Weight/Division | D Sr. | E Description |
F Unit | G Designer QTY | H-K AAA | L-O Tenderer #01 | ... AV-AY #10 |
BB-BO  pricing analysis (Pricing Basis, Base Price, adjustments, Dry Cost...)
```

AAA block: `H` QTY, `I` Rate, `J = $G*I` Estimated (Designer quantities),
`K = $H*I` Pre-Tender Estimate (AAA quantities). Each tenderer block repeats
that shape every 4 columns.

`Asset Summary` carries two amount columns per block (I..AD), mapping to the 3rd
and 4th column of each asset-tab block. Learn that mapping from whichever
summary row is already linked rather than hard-coding it.

## Defects this family shipped with

All corrected on build; all were in the blank template as delivered.

| Where | What | Effect |
|---|---|---|
| Packaged tab, total row | `SUBTOTAL(9,J6:J106)` | covered package 01 only, so every later package was missing from the asset total |
| Elemental amount cells | `XLOOKUP(...,0,0,1)` | returns the first match, so an NRM code appearing once per package counted once |
| Packaged tab, package header | weight columns `SUBTOTAL(9,B9:B11)` | three rows, inherited from the flat template; should span the package |
| Asset & Project Summary | `SUBTOTAL(9,A6:A47)` | weight total stopped 10 rows short, dropping assets 21-25 |
| Asset & Project Summary | `B28 = 1` | Sr. number should be 11 |

## Sample data the template ships

| Tab | Cells | Value if cloned |
|---|---|---|
| `Asset #01 {no package}` | A21 = `1.1.1`, G21 = 800, H21 = 1000, I21 = 2000 | 1,600,000 estimated / 2,000,000 PTE |
| `Asset #02 {with package}` | A23/A122 = `1.1.1`, H = 2500, I = 2000 | 5,000,000 **per package** |

Four packages therefore start life holding 20,000,000 of cost that nobody
entered. The build clears these input cells. The template also ships cached
formula results computed from them, which stay stale after the clear, so every
cached result is dropped too and the file recalculates on open.
