# Anatomy of a division-split Designer BOQ

What these workbooks look like, and the quirks worth expecting. Read this when a
file doesn't behave like the scripts assume.

## Tab layout

```
COVER SHEET
TOTAL SUMMARY              <- sums the division summaries
DIVISION 3                 <- the bill itself: items, quantities, rates
DIVISION 3 SUMMARY         <- page totals collected for that division
DIVISION 4
DIVISION 4 SUMMARY
...
DIVISION 31
DIVISION 31 SUMMARY
```

Divisions follow CSI numbering and are not contiguous: 3, 4, 5, 6, 7, 8, 9, then
22, 23, 26, 27, 31. Not every file carries every division — a unit type with no
WPC cladding simply has no Division 6 tab at all.

Tab names sometimes carry a trailing space (`DIVISION 23 SUMMARY `). Match on
the stripped, upper-cased name; never on an exact string.

## Inside a division tab

Six columns:

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| ITEM | DESCRIPTION | UNIT | QTY | RATE | AMOUNT EGP |

- Row 1 is the header, and each tab sets `print_title_rows = $1:$1` so it repeats
  on every printed page.
- Item rows carry a reference (`A.1`, `B.2`), a unit, a quantity, an empty RATE
  (the file is Issued For Tender, so rates are the tenderer's job) and an AMOUNT
  formula — either `=E7*$D7` or `=D7*E7`, both appear.
- Section rows carry only a description in column B: the CSI code and title
  (`03 30 00 Cast in Place Concrete`), preambles, or a `(Cont.)` repeat of the
  section title after a page total.
- `TOTAL OF PAGE` rows sum their own page range only, e.g. `=SUM(F59:F92)`.
  There is no division grand total on the detail tab — that lives on the
  division summary.
- Blank rows are deliberate padding that pushes a page total to the bottom of
  its printed page. Keep them; they carry row heights and borders.
- Descriptions are long wrapped text, so row heights vary from 18 to 180 points.
  Copying heights is not cosmetic — without them the sheet is unreadable.

No merged cells in the files seen so far, but the build handles them anyway.

## Quirks seen in the wild

- **Wrong currency in a header.** `AMOUNT SAR` on a Division 6 tab in an
  Egyptian BOQ, a leftover from a Saudi template. The other tabs said EGP.
- **A section title in a different format.** `DIVISION 06 Wood, Plastics and
  Composites` where every other tab uses `06 00 00 <title>`.
- **Stray scratch numbers** in columns G, H or I beside an item — a designer's
  working calculation left in the file. They sit outside the print area. Carry
  them over rather than dropping them, and tell the user where they are.
- **Rate-only items**, where AMOUNT holds the text `(rate only)` or `Rate Only`
  instead of a formula. They must not be swept into any total.
- **Junk defined names** inherited from decades of copied templates
  (`__123Graph_ACURRENT`, `aaa`). The build drops them, since they'd otherwise
  point at deleted tabs.
- **Huge phantom column counts** — a tab reporting 16,384 columns because of
  stray formatting. Always compute the extent from values, never from
  `ws.max_column`.
- **Identical files under different names.** Mirrored unit types (`VILLA_V` and
  `VILLA_VM`) can be byte-identical across every division. Legitimate for a
  mirror, wrong if the two units actually differ.

## Why the combined sheet is built the way it is

- **One header at row 1**, frozen and set as the repeating print title, rather
  than twelve copies. Each source tab already did exactly this with its own row
  1; keeping twelve would print a redundant header mid-page.
- **A banner row per division** because a 900-row sheet with no tab boundaries is
  hard to navigate. It's the only added content and is safe to delete.
- **A page break per division** so the printed output still matches the source:
  same page count as the original tabs printed one after another.
- **Formulas translated, not flattened.** Copying amounts as values would look
  identical on screen and silently break the moment a rate is entered.
- **The source workbook is left alone.** Its summaries and TOTAL SUMMARY still
  reference the division tabs; the combined sheet ships as a separate file.
