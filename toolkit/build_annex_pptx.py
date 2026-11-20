"""build_annex_pptx.py - "Annex D – Areas Calculation" as a PowerPoint deck.

usage:
  python build_annex_pptx.py --out "<Project> - Annex D - Areas Calculation - R01.pptx" --project "<Project>"
         [--client "<Client>"] [--rev R01] [--letter D] [--date "17 November 2026"]
         [--template templates/<AAA template>.pptx] [--workbook "<workbook file name>"] [--stills <dir>]
         asset1.json [asset2.json ...]

Inputs are the same asset JSONs (out/data/<slug>.json) the Areas workbook is built from, so every number in the deck
is the workbook's number. Stills: <stills>/<slug>/compare.png (render_stills.py on the asset's model).

Slides: cover; AAA BUA vs IPMS 1; one slide per asset (3D still + area table by level); summary of all assets;
the template's back cover when it has one.

Template: an AAA PowerPoint deck or template (.pptx or .potx).
- If it has slides, they are the brand. Slide 1 is the cover: its texts are replaced ("ANNEX X", the title, the
  CLIENT:/PROJECT:/TITLE:/DATE: lines). Slide 2 is the page frame, copied for every content slide (header bar, logo,
  section label, title, "PAGE |" footer); whatever sits between the title and the footer is dropped. The last slide is
  the back cover; its long text becomes the closing note. Other slides are dropped. A finished annex, such as the
  Sample Villa Annex D, works as the template.
- Without slides its layouts are used: cover = "Title Slide"/"Cover" (else the first); content = "Title Only" (else
  "Title and Content", else the layout with a title and the fewest placeholders).
Without a template a plain 16:9 deck is drawn in AAA navy.

After saving, the deck is re-opened and every total on the asset and summary slides is checked against the JSONs;
any difference fails the run (exit 1).
"""
import argparse, copy, io, json, re, sys, zipfile
from datetime import date
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

NAVY = RGBColor(0x1B, 0x23, 0x5E)
TABLE_HEAD = RGBColor(0x1F, 0x38, 0x64)
INK = RGBColor(0x1E, 0x24, 0x2B)
MUTED = RGBColor(0x56, 0x61, 0x6C)
TOTAL_FILL = RGBColor(0xE4, 0xE9, 0xF0)
BAND_FILL = RGBColor(0xF3, 0xF5, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SKIN, BUA, IPMS = RGBColor(0x3F, 0xA2, 0x7B), RGBColor(0x4F, 0x7C, 0xC0), RGBColor(0xE3, 0x9A, 0x2F)
R_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

IN_BUA = {'cov': True, 'bal': True, 'unc': False, 'perg': False, 'shel': False, 'nb': False, 'plant': False}
IN_IPMS = {'cov': True, 'bal': True, 'unc': True, 'perg': True, 'shel': True, 'nb': True, 'plant': False}
OPEN_CODE = {'bal': 'H1.1', 'cov': 'H1.2', 'shel': 'H1.3', 'unc': 'H1.4', 'perg': 'H1.4', 'nb': 'A1', 'plant': '–'}

RULES = [
    ('External skin area (outer face of external walls), every enclosed level', 'Included', 'Included'),
    ('H1.1 Balconies, H1.2 covered terraces', 'Included', 'Included'),
    ('H1.4 Uncovered terraces', 'Excluded', 'Included'),
    ('H1.4 Terraces under an open-slat pergola (treated as uncovered)', 'Excluded', 'Included'),
    ('H1.3 Sheltered areas', 'Excluded', 'Included'),
    ('A1 Notional boundary areas', 'Excluded', 'Included'),
    ('Plant / equipment roofs, maintenance access only', 'Not counted', 'Not counted'),
    ('B1.3 Atrium void', 'Counted once, on one floor', 'Counted once, on one floor'),
]


def fmt(v):
    return '–' if v is None or abs(v) < 0.005 else f'{v:,.2f}'


# ---------------- template ----------------
def open_template(path):
    """Returns (presentation, template used, brand slides {'cover'|'page'|'closing': (slide, sldId)})."""
    if not path:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        return prs, False, {}
    p = Path(path)
    if p.suffix.lower() == '.potx':          # python-pptx opens .pptx only: patch the main content type in memory
        src, buf = zipfile.ZipFile(p), io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as dst:
            for it in src.infolist():
                data = src.read(it.filename)
                if it.filename == '[Content_Types].xml':
                    data = data.replace(b'presentationml.template.main+xml', b'presentationml.presentation.main+xml')
                dst.writestr(it, data)
        buf.seek(0)
        prs = Presentation(buf)
    else:
        prs = Presentation(str(p))
    slides, ids = list(prs.slides), list(prs.slides._sldIdLst)
    brand = {}                                # cover = first slide, back cover = last, page frame = second
    if slides:
        brand['cover'] = (slides[0], ids[0])
    if len(slides) >= 2:
        brand['closing'] = (slides[-1], ids[-1])
    if len(slides) >= 3:
        brand['page'] = (slides[1], ids[1])
    keep = {id(v[1]) for v in brand.values()}
    for sid in ids:                           # drop every other slide; masters and layouts stay
        if id(sid) not in keep:
            drop_slide(prs, sid)
    return prs, True, brand


def drop_slide(prs, sid):
    prs.part.drop_rel(sid.rId)
    prs.slides._sldIdLst.remove(sid)


def pick_layout(prs, kind):
    layouts = list(prs.slide_layouts)
    want = ('title slide', 'cover') if kind == 'cover' else ('title only', 'title and content', 'title, content')
    for w in want:
        for l in layouts:
            if w in (l.name or '').lower():
                return l
    if kind == 'cover':
        return layouts[0]
    with_title = [l for l in layouts if any(ph.placeholder_format.type in (1, 3) for ph in l.placeholders)]
    return min(with_title, key=lambda l: len(l.placeholders)) if with_title else layouts[-1]


def is_title(sh):
    return sh.is_placeholder and sh.placeholder_format.type in (1, 3)       # TITLE, CENTER_TITLE


def text_of(sh):
    return sh.text_frame.text.strip() if sh.has_text_frame else ''


def set_text(sh, text):
    """Replace a shape's text, keeping the formatting of its first run."""
    tf = sh.text_frame
    p0 = tf.paragraphs[0]
    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    runs = p0.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p0.add_run().text = text


def font_of(sh):
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if r.font.name:
                return r.font.name
    return None


def remove_shape(sh):
    sh._element.getparent().remove(sh._element)


class Deck:
    def __init__(self, a):
        self.a = a
        self.prs, self.templ, self.brand = open_template(a.template)
        self.W, self.H = self.prs.slide_width, self.prs.slide_height
        self.mx = int(self.W * 0.045)
        if 'page' in self.brand:             # line the content up with the template's title
            t = next((sh for sh in self.brand['page'][0].shapes if is_title(sh)), None)
            if t is not None:
                self.mx = t.left
        self.page = 0
        self.header = ' – '.join(x for x in (a.client, a.project, f'Annex {a.letter}', 'Areas Calculation') if x).upper()
        self.font = None                     # the template's font, also used for the text and tables added here
        for key in ('page', 'cover'):
            if key in self.brand:
                shapes = sorted((sh for sh in self.brand[key][0].shapes if text_of(sh)), key=lambda sh: not is_title(sh))
                self.font = next((f for f in map(font_of, shapes) if f), None)
                if self.font:
                    break

    # -- helpers
    def text(self, sl, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP):
        tb = sl.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = 0
        for i, p in enumerate(paras):
            if isinstance(p, str):
                p = dict(t=p)
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = p.get('align', PP_ALIGN.LEFT)
            para.space_after = Pt(p.get('after', 4))
            r = para.add_run()
            r.text = p['t']
            f = r.font
            f.size = Pt(p.get('size', 14))
            f.bold = p.get('bold', False)
            f.color.rgb = p.get('color', INK)
            if self.font:
                f.name = self.font
        return tb

    def rect(self, sl, x, y, w, h, fill):
        s = sl.shapes.add_shape(1, x, y, w, h)
        s.fill.solid()
        s.fill.fore_color.rgb = fill
        s.line.fill.background()
        s.shadow.inherit = False
        return s

    def clear_placeholders(self, sl, keep=()):
        for ph in list(sl.placeholders):
            if ph.placeholder_format.idx not in keep:
                remove_shape(ph)

    def picture(self, sl, img, x, y, w, h):
        with Image.open(img) as im:
            iw, ih = im.size
        s = min(w / iw, h / ih)
        pw, ph = int(iw * s), int(ih * s)
        pic = sl.shapes.add_picture(str(img), x + (w - pw) // 2, y + (h - ph) // 2, pw, ph)
        pic.line.color.rgb = RGBColor(0xCB, 0xD1, 0xD7)
        pic.line.width = Pt(0.75)
        return pic

    def table(self, sl, rows, x, y, w, widths, total_rows=(), size=12, row_h=Pt(24), numeric_from=1, name=None):
        nr, nc = len(rows), len(rows[0])
        gs = sl.shapes.add_table(nr, nc, x, y, w, row_h * nr)
        if name:
            gs.name = name
        t = gs.table
        tw = sum(widths)
        for j, cw in enumerate(widths):
            t.columns[j].width = int(w * cw / tw)
        for i, row in enumerate(rows):
            t.rows[i].height = row_h
            for j, v in enumerate(row):
                c = t.cell(i, j)
                c.margin_left = c.margin_right = Pt(6)
                c.margin_top = c.margin_bottom = Pt(2)
                c.vertical_anchor = MSO_ANCHOR.MIDDLE
                c.fill.solid()
                head, tot = i == 0, i in total_rows
                c.fill.fore_color.rgb = TABLE_HEAD if head else TOTAL_FILL if tot else (WHITE if i % 2 else BAND_FILL)
                tf = c.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.RIGHT if j >= numeric_from and not (head and j < numeric_from) else PP_ALIGN.LEFT
                r = p.add_run()
                r.text = str(v)
                r.font.size = Pt(size - 1 if head else size)
                r.font.bold = head or tot
                r.font.color.rgb = WHITE if head else INK
                if self.font:
                    r.font.name = self.font
        return t

    def cover_lines(self):
        a = self.a
        lines = [f'Client: {a.client}' if a.client else None, f'Project: {a.project}', f'Revision: {a.rev}', f'Date: {a.date}']
        return [l for l in lines if l]

    # -- brand slides from the template
    def brand_cover(self, sl):
        a = self.a
        fields = {'CLIENT': a.client, 'PROJECT': a.project, 'TITLE': f'Annex {a.letter} – Areas Calculation – {a.rev}',
                  'DATE': a.date, 'REVISION': a.rev, 'REV': a.rev}
        titled = filled = False
        texts = [sh for sh in sl.shapes if text_of(sh)]
        for sh in texts:
            t = text_of(sh)
            m = re.match(r'(CLIENT|PROJECT|TITLE|DATE|REVISION|REV)\s*:', t, re.I)
            if is_title(sh):
                set_text(sh, 'AREAS CALCULATION')
                titled = True
            elif re.fullmatch(r'ANNEX\s+[A-Z0-9]{1,3}', t, re.I):
                set_text(sh, f'ANNEX {a.letter}')
            elif m:
                v = fields[m.group(1).upper()]
                if v:
                    set_text(sh, f'{m.group(1).upper()}: {v.upper()}')
                else:
                    remove_shape(sh)          # no client given: leave the line out
                filled = True
        if not titled and texts:              # no title placeholder: the largest text is the title
            big = max(texts, key=lambda sh: max((r.font.size or 0 for p in sh.text_frame.paragraphs for r in p.runs), default=0))
            if not re.match(r'(ANNEX|CLIENT|PROJECT|TITLE|DATE|REV)', text_of(big), re.I):
                set_text(big, 'AREAS CALCULATION')
        if not filled:
            self.text(sl, self.mx, int(self.H * 0.62), int(self.W * 0.6), int(self.H * 0.25),
                      [dict(t=l.upper(), size=14, color=MUTED) for l in self.cover_lines()])

    def brand_closing(self, sl):
        wb = f' ({self.a.workbook})' if self.a.workbook else ''
        note = (f'This annex accompanies the report it is issued with and should be read together with it. Every area in it '
                f'is taken from the AAA Areas Calculations workbook{wb}, where the room-by-room schedule is given.')
        longest = max((sh for sh in sl.shapes if len(text_of(sh)) > 40), key=lambda sh: len(text_of(sh)), default=None)
        if longest is not None:
            set_text(longest, note)

    def brand_page(self, num, title):
        """Copy the template's page frame (slide 2) and fill it; returns (slide, top, bottom) of the free area."""
        src = self.brand['page'][0]
        shapes = list(src.shapes)
        t_src = next((sh for sh in shapes if is_title(sh)), None)
        foot = next((sh for sh in shapes if text_of(sh).upper().startswith('PAGE')), None)
        body_top = t_src.top + t_src.height if t_src is not None else int(self.H * 0.2)
        body_bot = foot.top if foot is not None else int(self.H * 0.92)
        slack = Inches(0.05)
        sl = self.prs.slides.add_slide(src.slide_layout)
        self.clear_placeholders(sl)
        bg = src._element.cSld.find(qn('p:bg'))
        if bg is not None:
            sl._element.cSld.insert(0, copy.deepcopy(bg))
        tree = sl.shapes._spTree
        for sh in shapes:
            if sh.top >= body_top - slack and sh.top + sh.height <= body_bot + slack:
                continue                      # the template's own content: dropped
            el = copy.deepcopy(sh._element)
            for node in el.iter():            # carry over pictures and links
                for k, v in list(node.attrib.items()):
                    if k.startswith(R_NS) and v in src.part.rels:
                        rel = src.part.rels[v]
                        node.set(k, sl.part.relate_to(rel.target_ref, rel.reltype, is_external=True) if rel.is_external
                                 else sl.part.relate_to(rel.target_part, rel.reltype))
            tree.insert_element_before(el, 'p:extLst')
        for sh in sl.shapes:
            t = text_of(sh)
            if is_title(sh):
                set_text(sh, f'{num} {title}')
            elif t.upper().startswith('PAGE'):
                set_text(sh, f'PAGE | {self.a.letter}-{self.page}')
            elif t.startswith('|'):
                set_text(sh, f'| {self.header}')
            elif (t_src is not None and re.match(r'[A-Z]{1,2}\s+[A-Z]', t) and sh.top + sh.height <= t_src.top + slack
                  and sh.top >= t_src.top - Inches(0.6)):
                set_text(sh, f'{self.a.letter} AREAS CALCULATION')      # section label above the title
        return sl, body_top + Inches(0.25), body_bot - Inches(0.12)

    # -- slides
    def cover(self):
        a = self.a
        if 'cover' in self.brand:
            self.brand_cover(self.brand['cover'][0])
            return
        sl = self.prs.slides.add_slide(pick_layout(self.prs, 'cover') if self.templ else self.prs.slide_layouts[6])
        title = f'Annex {a.letter} – Areas Calculation'
        lines = self.cover_lines()
        if self.templ and sl.shapes.title is not None:
            sl.shapes.title.text = title
            sub = [ph for ph in sl.placeholders if ph.placeholder_format.idx != 0 and ph.placeholder_format.type in (4, 2, 7)]
            keep = [0]
            if sub:
                sub[0].text_frame.text = '\n'.join(lines)
                keep.append(sub[0].placeholder_format.idx)
            else:
                self.text(sl, self.mx, int(self.H * 0.62), int(self.W * 0.6), int(self.H * 0.25),
                          [dict(t=l, size=14, color=MUTED) for l in lines])
            self.clear_placeholders(sl, keep)
            return
        self.rect(sl, 0, 0, self.W, self.H, NAVY)
        self.text(sl, self.mx, int(self.H * 0.08), int(self.W * 0.5), Inches(0.6),
                  [dict(t='AAA  |  CONSULTANCY', size=20, bold=True, color=WHITE)])
        self.text(sl, self.mx, int(self.H * 0.30), int(self.W * 0.8), Inches(0.4),
                  [dict(t=f'ANNEX {a.letter}', size=14, bold=True, color=RGBColor(0xAF, 0xBC, 0xD6))])
        self.text(sl, self.mx, int(self.H * 0.36), int(self.W * 0.8), Inches(1.4),
                  [dict(t='AREAS CALCULATION', size=44, bold=True, color=WHITE)])
        self.text(sl, self.mx, int(self.H * 0.62), int(self.W * 0.6), int(self.H * 0.25),
                  [dict(t=l.upper(), size=13, color=WHITE) for l in lines])

    def content(self, title):
        """New content slide; returns (slide, top of the free area, bottom of the free area)."""
        self.page += 1
        num = f'{self.a.letter}.{self.page}'
        if 'page' in self.brand:
            return self.brand_page(num, title)
        if self.templ:
            sl = self.prs.slides.add_slide(pick_layout(self.prs, 'content'))
            t = sl.shapes.title
            if t is not None:
                t.text = f'{num}  {title}'
                top = t.top + t.height + Inches(0.15)
            else:
                self.text(sl, self.mx, Inches(0.5), self.W - 2 * self.mx, Inches(0.7), [dict(t=f'{num}  {title}', size=26, bold=True)])
                top = Inches(1.35)
            self.clear_placeholders(sl, keep=(0,))
            return sl, top, self.H - Inches(0.75)
        sl = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.rect(sl, self.mx, Inches(0.28), self.W - 2 * self.mx - Inches(1.3), Inches(0.26), NAVY)
        self.text(sl, self.mx + Inches(0.1), Inches(0.28), self.W - 2 * self.mx - Inches(1.5), Inches(0.26),
                  [dict(t='| ' + self.header, size=9, bold=True, color=WHITE)], anchor=MSO_ANCHOR.MIDDLE)
        self.text(sl, self.W - self.mx - Inches(1.15), Inches(0.22), Inches(1.15), Inches(0.4),
                  [dict(t='AAA\nCONSULTANCY', size=8, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)])
        self.text(sl, self.mx, Inches(0.72), self.W - 2 * self.mx, Inches(0.3),
                  [dict(t=f'{self.a.letter}  AREAS CALCULATION', size=10, bold=True, color=TABLE_HEAD)])
        self.text(sl, self.mx, Inches(0.98), self.W - 2 * self.mx, Inches(0.6), [dict(t=f'{num} {title}', size=26, bold=True)])
        self.rect(sl, self.mx, self.H - Inches(0.42), self.W - 2 * self.mx, Inches(0.22), NAVY)
        self.text(sl, self.W - self.mx - Inches(2.0), self.H - Inches(0.42), Inches(1.9), Inches(0.22),
                  [dict(t=f'PAGE | {self.a.letter}-{self.page}', size=9, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)],
                  anchor=MSO_ANCHOR.MIDDLE)
        return sl, Inches(1.75), self.H - Inches(0.6)

    def basis(self, datas):
        sl, top, bot = self.content('AAA BUA and IPMS 1')
        lw = int((self.W - 2 * self.mx) * 0.36)
        wb = f' ({self.a.workbook})' if self.a.workbook else ''
        void = any(d['totals'].get('void_removed') for d in datas)
        paras = [
            dict(t='Two measures are reported for every asset.', size=15, bold=True, after=10),
            dict(t='External skin area: the enclosed floor area of each level, measured to the outer face of the external walls.', size=13, after=8),
            dict(t='AAA BUA = external skin area + balconies and covered terraces.', size=13, bold=True, color=BUA, after=8),
            dict(t='IPMS 1 = AAA BUA + uncovered terraces, terraces under open pergolas, sheltered and notional boundary areas.',
                 size=13, bold=True, color=IPMS, after=8),
            dict(t='Plant roofs with maintenance access only are not counted by either measure.', size=13, after=8),
            dict(t=('An atrium void is counted once, on one floor, and removed from the floors above.' if void else
                    'Atrium voids, where present, are counted once, on one floor.'), size=13, after=8),
            dict(t=f'Every figure in this annex is taken from the AAA Areas Calculations workbook{wb}, '
                   'where the room-by-room schedule is given.', size=12, color=MUTED),
        ]
        self.text(sl, self.mx, top, lw, bot - top, paras)
        x = self.mx + lw + Inches(0.4)
        self.table(sl, [('Component', 'AAA BUA', 'IPMS 1')] + RULES, x, top, self.W - self.mx - x, (5.2, 1.9, 1.9),
                   size=12, row_h=Pt(30), numeric_from=9)

    def asset(self, d, still):
        sl, top, bot = self.content(d['asset'])
        sch, lv = d['sched'], d['levels']
        void = any(sch[L['id']]['void_removed'] for L in lv)
        head = ['Level'] + (['Void removed'] if void else []) + ['Ext. skin', '+ Balconies, covered', 'AAA BUA', '+ IPMS 1 only', 'IPMS 1']
        rows, tot = [head], dict(vr=0, skin=0, b=0, bua=0, i=0, ipms=0)
        for L in lv:
            s = sch[L['id']]
            badd, iadd = round(s['bua'] - s['skin'], 2), round(s['ipms'] - s['bua'], 2)
            rows.append([L['name']] + ([fmt(s['void_removed'])] if void else []) +
                        [fmt(s['skin']), fmt(badd), fmt(s['bua']), fmt(iadd), fmt(s['ipms'])])
            for k, v in (('vr', s['void_removed']), ('skin', s['skin']), ('b', badd), ('bua', s['bua']), ('i', iadd), ('ipms', s['ipms'])):
                tot[k] += v
        T = d['totals']
        for k in ('skin', 'bua', 'ipms'):
            if abs(tot[k] - T[k]) > 0.011:
                raise SystemExit(f"{d['asset']}: level {k} adds to {tot[k]:.2f}, totals say {T[k]:.2f}")
        rows.append(['Total'] + ([fmt(T['void_removed'])] if void else []) +
                    [fmt(T['skin']), fmt(T['bua'] - T['skin']), fmt(T['bua']), fmt(T['ipms'] - T['bua']), fmt(T['ipms'])])

        iw = int((self.W - 2 * self.mx) * 0.42)
        legend_h = Inches(0.35)
        if still and Path(still).exists():
            self.picture(sl, still, self.mx, top, iw, bot - top - legend_h)
        else:
            self.text(sl, self.mx, top, iw, Inches(0.5), [dict(t='3D still not rendered (offline model missing).', size=12, color=MUTED)])
        lx = self.mx
        for c, t in ((SKIN, 'Inside the external skin'), (BUA, 'Added by AAA BUA'), (IPMS, 'IPMS 1 only')):
            self.rect(sl, lx, bot - Inches(0.22), Inches(0.14), Inches(0.14), c)
            self.text(sl, lx + Inches(0.2), bot - Inches(0.27), Inches(1.9), Inches(0.25), [dict(t=t, size=10, color=MUTED)])
            lx += Inches(1.85)

        x = self.mx + iw + Inches(0.35)
        w = self.W - self.mx - x
        widths = (2.2,) + ((1.4,) if void else ()) + (1.4, 1.6, 1.4, 1.5, 1.4)
        rh = Pt(26)
        self.table(sl, rows, x, top, w, widths, total_rows=(len(rows) - 1,), size=11, row_h=rh, name=f"totals:{d['asset']}")
        y = top + rh * len(rows) + Inches(0.25)
        notes = [dict(t='Outside the external skin (m²)', size=11, bold=True, after=3)]
        names = {L['id']: L['name'] for L in lv}
        ops = [o for o in d['opens'] if o['key'] != 'void']     # removed void floors: covered by the void note below
        for o in ops[:9]:
            how = 'AAA BUA and IPMS 1' if IN_BUA[o['key']] else 'IPMS 1 only' if IN_IPMS[o['key']] else 'not counted'
            notes.append(dict(t=f"{names.get(o['level'], o['level'])} · {OPEN_CODE[o['key']]} {o['name']}: {fmt(o['area'])}, {how}",
                              size=10, color=MUTED, after=1))
        if len(ops) > 9:
            notes.append(dict(t=f'… and {len(ops) - 9} more, listed in the workbook.', size=10, color=MUTED))
        if not ops:
            notes.append(dict(t='None.', size=10, color=MUTED))
        if void:
            notes.append(dict(t=f"Atrium void: counted once; {fmt(T['void_removed'])} m² removed from the floors above.", size=10,
                              color=MUTED, after=1))
        src = d.get('meta', {}).get('sources')
        if src:
            notes.append(dict(t=f'Source: {src}', size=9, color=MUTED))
        self.text(sl, x, y, w, bot - y, notes)

    def summary(self, datas):
        sl, top, bot = self.content('Summary of all assets')
        rep = all(d.get('report_bua') for d in datas)
        head = ['Asset', 'Ext. skin (m²)', 'AAA BUA (m²)', 'IPMS 1 (m²)', 'Report BUA (m²)', 'Skin vs report']
        rows = [head]
        S = dict(skin=0, bua=0, ipms=0, rep=0)
        for d in datas:
            T, rb = d['totals'], d.get('report_bua')
            rows.append([d['asset'], fmt(T['skin']), fmt(T['bua']), fmt(T['ipms']), fmt(rb) if rb else '–',
                         f"{(T['skin'] / rb - 1) * 100:+.1f}%" if rb else '–'])
            for k in ('skin', 'bua', 'ipms'):
                S[k] += T[k]
            S['rep'] += rb or 0
        rows.append(['Total, all assets', fmt(S['skin']), fmt(S['bua']), fmt(S['ipms']), fmt(S['rep']) if rep else '–',
                     f"{(S['skin'] / S['rep'] - 1) * 100:+.1f}%" if rep and S['rep'] else '–'])
        w = self.W - 2 * self.mx
        rh = Pt(30)
        self.table(sl, rows, self.mx, top, w, (3.6, 1.5, 1.5, 1.5, 1.6, 1.4), total_rows=(len(rows) - 1,), size=13, row_h=rh,
                   name='totals:summary')
        y = top + rh * len(rows) + Inches(0.35)
        wb = f'"{self.a.workbook}"' if self.a.workbook else 'the AAA Areas Calculations workbook'
        self.text(sl, self.mx, y, w, bot - y, [
            dict(t=f'These totals are the same as the Summary sheet of {wb}, built from the same measurement.', size=13, after=6),
            dict(t='Report BUA is the figure used in the cost estimate; the check is the measured external skin against it (target ±2%).',
                 size=12, color=MUTED),
        ])
        return S

    def save(self, out):
        if 'page' in self.brand:              # the page frame was only a pattern
            drop_slide(self.prs, self.brand['page'][1])
        if 'closing' in self.brand:           # back cover last
            sl, sid = self.brand['closing']
            self.brand_closing(sl)
            lst = self.prs.slides._sldIdLst
            lst.remove(sid)
            lst.append(sid)
        self.prs.save(out)


def verify(out, datas):
    """Re-open the deck and check every total row against the JSONs."""
    prs = Presentation(out)
    tables = {}
    for sl in prs.slides:
        for sh in sl.shapes:
            if sh.has_table and sh.name.startswith('totals:'):
                tables[sh.name[7:]] = [[c.text for c in r.cells] for r in sh.table.rows]
    num = lambda s: 0.0 if s in ('–', '') else float(s.replace(',', ''))
    bad = []
    for i, d in enumerate(datas):
        T = d['totals']
        t = tables.get(d['asset'])
        if not t:
            bad.append(f"no table for {d['asset']}")
            continue
        last = t[-1]
        for k, v in zip(('ipms', 'bua', 'skin'), (last[-1], last[-3], last[-5])):
            if abs(num(v) - T[k]) > 0.005:
                bad.append(f"{d['asset']} slide {k}: {v} vs {T[k]}")
        row = tables['summary'][i + 1]
        for k, v in zip(('skin', 'bua', 'ipms'), row[1:4]):
            if abs(num(v) - T[k]) > 0.005:
                bad.append(f"summary {d['asset']} {k}: {v} vs {T[k]}")
    if bad:
        raise SystemExit('annex totals differ from the workbook data: ' + '; '.join(bad))
    return len(prs.slides)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('data', nargs='+')
    ap.add_argument('--out', required=True)
    ap.add_argument('--project', required=True)
    ap.add_argument('--client', default='')
    ap.add_argument('--rev', default='R01')
    ap.add_argument('--letter', default='D')
    ap.add_argument('--date', default=f'{date.today().day} {date.today():%B %Y}')
    ap.add_argument('--template', default='')
    ap.add_argument('--workbook', default='')
    ap.add_argument('--stills', default='')
    a = ap.parse_args()
    datas = [json.loads(Path(p).read_text(encoding='utf-8')) for p in a.data]
    deck = Deck(a)
    deck.cover()
    deck.basis(datas)
    for p, d in zip(a.data, datas):
        deck.asset(d, Path(a.stills) / Path(p).stem / 'compare.png' if a.stills else None)
    deck.summary(datas)
    deck.save(a.out)
    n = verify(a.out, datas)
    used = ('brand slides from ' if deck.brand else 'layouts from ') + Path(a.template).name if a.template else 'none (plain AAA navy)'
    print(f"annex: {n} slides, template {used}, totals match the workbook data -> {a.out}")


if __name__ == '__main__':
    main()
