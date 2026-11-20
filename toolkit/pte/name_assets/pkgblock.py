#!/usr/bin/env python3
"""
Build asset tabs on the CLIENT/PROJECT R00 template, with per-asset packages.

The packaged template holds two 99-row package blocks:

    P        package header   D = letter, E = name, SUBTOTAL over P+1 : P+97
    P+2      division 01      D = "<letter>.01", SUBTOTAL over P+3 : P+5
    P+6      division 02      ... 24 divisions, every 4 rows, last at P+94
    P+98     blank separator
    ...      next package at P+99
    8+99N    Total Amounts

An asset with N packages gets N blocks. Every cell is written with an EXPLICIT
formula taken from the template's resolved text with the row numbers remapped -
cloning cells verbatim would copy their shared-formula declarations and make the
new block evaluate the first one.

Two defects in the template are corrected while building:
  * the grand total read SUBTOTAL(9,J6:J106), covering package 01 only; it now
    spans the whole body
  * the package header's weight columns read SUBTOTAL(9,B9:B11) - three rows,
    a leftover from the unpackaged template - they now span the package
"""
import re
from xml.sax.saxutils import escape as xesc

BLOCK = 99           # rows per package block
FIRST = 8            # first package header row
DIVS = 24            # divisions per package
TPL_TOTAL = 206      # total row in the 2-package template
TPL_END = FIRST + BLOCK - 1          # 106, last row of template block 1
AMOUNT_COLS = None   # discovered from the total row


def colnum(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def rowmap(r, i, total):
    """Where template row r lands in package block i."""
    if FIRST <= r <= TPL_END:
        return r + BLOCK * i
    if r == TPL_TOTAL:
        return total
    return r


def remap(formula, i, total):
    """Shift every row reference in a formula into package block i."""
    return re.sub(r'(\$?)([A-Z]{1,3})(\$?)(\d+)',
                  lambda m: '%s%s%s%d' % (m.group(1), m.group(2), m.group(3),
                                          rowmap(int(m.group(4)), i, total)),
                  formula)


def cells_of(rowxml):
    """[(ref, col, attrs, body or None)] for each cell in a row element."""
    out = []
    for m in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', rowxml, re.S):
        out.append((m.group(1), int(m.group(2)), m.group(3), m.group(4)))
    return out


def build_asset_rows(tpl_rows, tpl_formulas, tpl_values, packages, clear_seed=True):
    """
    tpl_rows      {row number: row XML} from the packaged template
    tpl_formulas  {(col, row): resolved formula text} via openpyxl
    tpl_values    {(col, row): cell value} via openpyxl - literals are shared
                  strings in the XML, so their text is only visible from here
    packages      list of package names
    clear_seed    blank the sample NRM code / quantities / rates the template
                  ships on its demo item rows; cloning them would multiply that
                  phantom money by the number of packages
    returns       (xml for all body rows, total row number)
    """
    n = len(packages)
    total = FIRST + BLOCK * n
    out = []

    # rows 1..7 stay exactly as they are
    for r in sorted(k for k in tpl_rows if k < FIRST):
        out.append(tpl_rows[r])

    for i, pkgname in enumerate(packages):
        letter = chr(65 + i)
        for r in range(FIRST, TPL_END + 1):
            if r not in tpl_rows:
                continue
            newr = r + BLOCK * i
            rx = tpl_rows[r]
            rx = re.sub(r'(<row r=")\d+(")', r'\g<1>%d\g<2>' % newr, rx)
            rebuilt = ''
            for col, _r, attrs, body in cells_of(tpl_rows[r]):
                ref = '%s%d' % (col, newr)
                f = tpl_formulas.get((col, r))
                if isinstance(f, str) and f.startswith('='):
                    nf = remap(f[1:], i, total)
                    # package header: the weight columns must span the package,
                    # not the three rows they inherited from the flat template
                    if r == FIRST and col in ('B', 'C'):
                        nf = 'SUBTOTAL(9,%s%d:%s%d)' % (col, newr + 1, col, newr + BLOCK - 2)
                    rebuilt += '<c r="%s"%s><f>%s</f><v>0</v></c>' % (ref, attrs, xesc(nf))
                    continue
                # sample data on a demo item row (J is a QTY x rate formula there)
                if (clear_seed and col in ('A', 'G', 'H', 'I')
                        and str(tpl_formulas.get(('J', r), '')).startswith('=$G')):
                    rebuilt += '<c r="%s"%s/>' % (ref, re.sub(r'\st="[a-z]+"', '', attrs))
                    continue
                # literal cells: package name, division code, or plain text
                if r == FIRST and col == 'D':
                    rebuilt += '<c r="%s"%s t="inlineStr"><is><t>%s</t></is></c>' % (
                        ref, attrs.replace(' t="s"', '').replace(' t="inlineStr"', ''),
                        xesc(letter))
                elif r == FIRST and col == 'E':
                    rebuilt += '<c r="%s"%s t="inlineStr"><is><t>%s</t></is></c>' % (
                        ref, attrs.replace(' t="s"', '').replace(' t="inlineStr"', ''),
                        xesc(pkgname))
                elif col == 'D' and re.fullmatch(r'A\.\d\d',
                                                 str(tpl_values.get((col, r), '')).strip()):
                    code = str(tpl_values[(col, r)]).strip().split('.')[1]
                    rebuilt += '<c r="%s"%s t="inlineStr"><is><t>%s.%s</t></is></c>' % (
                        ref, attrs.replace(' t="s"', '').replace(' t="inlineStr"', ''),
                        letter, code)
                elif body is None:
                    rebuilt += '<c r="%s"%s/>' % (ref, attrs)
                else:
                    rebuilt += '<c r="%s"%s>%s</c>' % (ref, attrs, body)
            rx = re.sub(r'(<row [^>]*>).*(</row>)', lambda m: m.group(1) + rebuilt + m.group(2),
                        rx, flags=re.S)
            out.append(rx)

    # total row, spanning the whole body this time
    trx = tpl_rows[TPL_TOTAL]
    trx = re.sub(r'(<row r=")\d+(")', r'\g<1>%d\g<2>' % total, trx)
    rebuilt = ''
    for col, _r, attrs, body in cells_of(tpl_rows[TPL_TOTAL]):
        ref = '%s%d' % (col, total)
        f = tpl_formulas.get((col, TPL_TOTAL))
        if isinstance(f, str) and f.startswith('=SUBTOTAL'):
            rebuilt += '<c r="%s"%s><f>SUBTOTAL(9,%s6:%s%d)</f><v>0</v></c>' % (
                ref, attrs, col, col, total - 1)
        elif body is None:
            rebuilt += '<c r="%s"%s/>' % (ref, attrs)
        else:
            rebuilt += '<c r="%s"%s>%s</c>' % (ref, attrs, body)
    trx = re.sub(r'(<row [^>]*>).*(</row>)', lambda m: m.group(1) + rebuilt + m.group(2),
                 trx, flags=re.S)
    out.append(trx)
    return ''.join(out), total
