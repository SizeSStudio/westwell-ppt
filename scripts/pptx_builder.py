#!/usr/bin/env python3
"""
Westwell PPT Builder  v2.0
──────────────────────────
WestwellPPT class with slide-type methods that match the actual visual
language of Westwell solution decks (airport, factory, seaport references).

Design philosophy (v2):
  • Titles: top-left, left-aligned, navy, ~22pt bold — never centered
  • Thin teal horizontal rule separates title from content
  • Large text / large image / large whitespace — "presentation style"
  • Dark slides for: cover, chapter, statement moments
  • Layout placeholders are always suppressed; content added as free textboxes

Usage:
    import sys; sys.path.insert(0, '~/.claude/skills/westwell-ppt')
    from scripts.pptx_builder import WestwellPPT, preview_pptx

    ppt = WestwellPPT(template='…/PPTTemplate.potx', output='out.pptx')
    ppt.cover("Title", "Subtitle", "2026年4月")
    ppt.agenda(["Chapter 1", "Chapter 2", "Chapter 3"])
    ppt.chapter("01", "Chapter Title", "Supporting line")
    ppt.statement("核心论点：一句话定义这一章的价值")
    ppt.bullets("结论性标题", ["要点A", "要点B", "要点C"])
    ppt.text_image("结论性标题", "Body text…", "/path/image.png")
    ppt.stats("结论性标题", [("95%","指标","说明"), ("<1s","指标","说明")])
    ppt.two_col("标题", "左栏头", "左栏内容", "右栏头", "右栏内容")
    ppt.table("标题", ["Col1","Col2"], [["r1c1","r1c2"]])
    ppt.image("标题", "/path/image.png", "caption")
    ppt.end("期待与您深入探讨")
    ppt.save()
"""

import os, subprocess, tempfile, math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

# ── Colours ───────────────────────────────────────────────────────────────────
C_DARK   = RGBColor(0x27, 0x43, 0xC6)   # #2743C6  dark blue background
C_NAVY   = RGBColor(0x00, 0x1B, 0xA6)   # #001BA6  title / heading / primary text on light
C_TEAL   = RGBColor(0x00, 0xCA, 0xD4)   # #00CAD4  accent / rule / numbers (non-text)
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
C_LGRAY  = RGBColor(0xF0, 0xF4, 0xFF)   # very light blue-grey, card bg
C_MGRAY  = RGBColor(0xD0, 0xD8, 0xF0)   # medium gray, card borders
C_GRAY   = RGBColor(0x55, 0x55, 0x66)   # secondary body text
C_BLACK  = RGBColor(0x1A, 0x1A, 0x2E)   # near-black body text
C_ALTROW = RGBColor(0xE4, 0xEC, 0xFB)   # table alternate row

# ── Geometry (inches) ─────────────────────────────────────────────────────────
SW, SH = 13.333, 7.500      # slide canvas

# Standard content slide title
TL, TT   = 0.906, 0.48      # title left, top
TW, TH   = 11.521, 1.05     # title width (full-bleed reference), height
# Safe title textbox width: caps title so it does not collide with the
# "Make a WELL Change" brand label in the top-right of every content slide.
# Leave ≈2.5" of right margin for the brand zone.
TITLE_SAFE_W     = 9.00
TITLE_MAX_CHARS  = 25       # soft limit per line — warning, not an error

# Teal rule below title (sits just below the visible glyph line at 26pt).
# Previous value 1.58 left a visible gap between title text and rule;
# 1.15 tucks the rule right under the text baseline for a tighter title block.
RULE_Y   = 1.15              # y-position of horizontal rule
RULE_H   = 0.05              # rule height (thin line)

# Content area (below rule, above WMF decoration)
CL       = 0.906
CT       = 1.80              # content top (below rule, with breathing room)
CR       = 12.427
CB       = 5.70              # conservative bottom (WMF circles at 5.933")
CW       = CR - CL           # 11.521"
CH       = CB - CT           # 3.90"

# Cover geometry
COV_TL, COV_TT = 0.689, 1.70   # cover title (left half, moved up)
COV_TW, COV_TH = 8.754, 2.705
COV_SUB_Y      = 4.60           # cover subtitle / meta area start

# ── Fonts ─────────────────────────────────────────────────────────────────────
F_TITLE  = 'Encode Sans'
F_BODY   = '思源黑体'
F_ACCENT = 'Encode Sans'


def I(x):
    return Inches(x)


# ── Preview helper ────────────────────────────────────────────────────────────

def preview_pptx(pptx_path: str, out_dir: str, dpi_scale: float = 1.5):
    """
    Convert pptx → PDF (via soffice) → PNG per slide (via pymupdf/fitz).
    Saves files to out_dir as <stem>_slide_NN.png.
    """
    import fitz
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(pptx_path))[0]

    # Convert to PDF
    pdf_path = os.path.join(out_dir, stem + '_preview.pdf')
    subprocess.run(
        ['soffice', '--headless', '--convert-to', 'pdf',
         '--outdir', out_dir, pptx_path],
        check=True, capture_output=True
    )
    # soffice writes to same dir with same stem
    converted = os.path.join(out_dir, stem + '.pdf')
    if os.path.exists(converted) and converted != pdf_path:
        os.rename(converted, pdf_path)

    # Render pages
    doc = fitz.open(pdf_path)
    mat = fitz.Matrix(dpi_scale * 96 / 72, dpi_scale * 96 / 72)
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=mat)
        pix.save(os.path.join(out_dir, f'{stem}_slide_{i+1:02d}.png'))


# ── Main class ────────────────────────────────────────────────────────────────

class WestwellPPT:
    """Westwell-branded PPTX builder. v2.0 — Westwell visual language."""

    def __init__(self, template: str, output: str):
        self.output = output
        self._pptx_path = self._ensure_pptx(os.path.expanduser(template))
        self.prs = Presentation(self._pptx_path)
        self._layout_map = {l.name: l for l in self.prs.slide_master.slide_layouts}
        self._clear_slides()

    # ── Init helpers ──────────────────────────────────────────────────────────

    def _ensure_pptx(self, path: str) -> str:
        if path.lower().endswith('.pptx'):
            return path
        tmp = tempfile.mkdtemp()
        subprocess.run(
            ['soffice', '--headless', '--convert-to', 'pptx',
             '--outdir', tmp, path],
            check=True, capture_output=True
        )
        base = os.path.splitext(os.path.basename(path))[0]
        out = os.path.join(tmp, base + '.pptx')
        if not os.path.exists(out):
            raise FileNotFoundError(f'Conversion failed: {out}')
        return out

    def _clear_slides(self):
        lst = self.prs.slides._sldIdLst
        for el in list(lst):
            rid = el.get(qn('r:id'))
            lst.remove(el)
            try:
                self.prs.part.drop_rel(rid)
            except Exception:
                pass

    def _layout(self, name: str):
        if name in self._layout_map:
            return self._layout_map[name]
        return self._layout_map.get(
            'custom slide1-light',
            list(self._layout_map.values())[0]
        )

    # ── Low-level drawing primitives ──────────────────────────────────────────

    def _suppress_placeholders(self, slide):
        """Remove all layout placeholders so none show 'Click to add…'."""
        for ph in list(slide.placeholders):
            sp = ph._element
            parent = sp.getparent()
            if parent is not None:
                parent.remove(sp)

    def _textbox(self, slide, l, t, w, h, text,
                 size=18, bold=False, color=None,
                 align=PP_ALIGN.LEFT, font=None, wrap=True,
                 line_spacing=None):
        color = color or C_BLACK
        font  = font  or F_BODY
        tb = slide.shapes.add_textbox(I(l), I(t), I(w), I(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        p = tf.paragraphs[0]
        p.alignment = align
        if line_spacing:
            from pptx.util import Pt as _Pt
            from pptx.oxml.ns import qn as _qn
            from lxml import etree as _et
            pPr = p._pPr
            if pPr is None:
                pPr = _et.SubElement(p._p, _qn('a:pPr'))
            lnSpc = _et.SubElement(pPr, _qn('a:lnSpc'))
            spcPts = _et.SubElement(lnSpc, _qn('a:spcPts'))
            spcPts.set('val', str(int(line_spacing * 100)))
        run = p.add_run()
        run.text = text
        run.font.name  = font
        run.font.size  = Pt(size)
        run.font.bold  = bold
        run.font.color.rgb = color
        return tb

    def _rich_textbox(self, slide, l, t, w, h, text,
                      size=18, color=None, font=None,
                      align=PP_ALIGN.LEFT, wrap=True):
        """
        Textbox with inline **bold** markdown support.
        Any text wrapped in **...** is rendered bold; the rest is regular weight.
        Falls back to _textbox if no markers are found.
        """
        import re
        color = color or C_BLACK
        font  = font  or F_BODY
        parts = re.split(r'\*\*(.*?)\*\*', text)
        if len(parts) == 1:
            # No bold markers — use simple textbox
            return self._textbox(slide, l, t, w, h, text,
                                 size=size, color=color, font=font,
                                 align=align, wrap=wrap)
        tb = slide.shapes.add_textbox(I(l), I(t), I(w), I(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        p = tf.paragraphs[0]
        p.alignment = align
        for i, part in enumerate(parts):
            if not part:
                continue
            run = p.add_run()
            run.text = part
            run.font.name  = font
            run.font.size  = Pt(size)
            run.font.bold  = (i % 2 == 1)   # odd index = between ** **
            run.font.color.rgb = color
        return tb

    def _para_textbox(self, slide, l, t, w, h, paras, wrap=True):
        """Multi-paragraph textbox. paras = list of dicts."""
        tb = slide.shapes.add_textbox(I(l), I(t), I(w), I(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        for i, para in enumerate(paras):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = para.get('align', PP_ALIGN.LEFT)
            if 'space_before' in para:
                p.space_before = Pt(para['space_before'])
            if 'space_after' in para:
                p.space_after = Pt(para['space_after'])
            run = p.add_run()
            run.text = para.get('text', '')
            run.font.name  = para.get('font', F_BODY)
            if 'size'  in para: run.font.size  = Pt(para['size'])
            if 'bold'  in para: run.font.bold  = para['bold']
            if 'color' in para: run.font.color.rgb = para['color']
        return tb

    def _rect(self, slide, l, t, w, h, fill=None, line=None, line_w=0,
              radius=0):
        """Filled rectangle. radius > 0 gives rounded corners (approx via shape)."""
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        if radius > 0:
            # Use RoundedRectangle (shape type 5)
            shape = slide.shapes.add_shape(5, I(l), I(t), I(w), I(h))
            # Set corner radius via XML adj
            sp_el = shape._element
            spPr = sp_el.find(qn('p:spPr'))
            if spPr is not None:
                prstGeom = spPr.find(qn('a:prstGeom'))
                if prstGeom is not None:
                    avLst = prstGeom.find(qn('a:avLst'))
                    if avLst is None:
                        avLst = etree.SubElement(prstGeom, qn('a:avLst'))
                    gd = etree.SubElement(avLst, qn('a:gd'))
                    # radius as fraction of shorter dimension, in 1/100000 units
                    r_val = int(min(radius / min(w, h), 0.5) * 100000)
                    gd.set('name', 'adj')
                    gd.set('fmla', f'val {r_val}')
        else:
            shape = slide.shapes.add_shape(1, I(l), I(t), I(w), I(h))

        if fill is not None:
            shape.fill.solid()
            shape.fill.fore_color.rgb = fill
        else:
            shape.fill.background()
        if line is not None:
            shape.line.color.rgb = line
            if line_w:
                shape.line.width = Pt(line_w)
        else:
            shape.line.fill.background()
        return shape

    def _img_size(self, img_path: str, max_w: float, max_h: float):
        """
        Return (w, h) in inches that fit within max_w × max_h, preserving aspect ratio.
        Reads PNG dimensions from the file header (no PIL dependency).
        Falls back to (max_w, max_h) if the file can't be parsed.
        """
        import struct
        try:
            with open(img_path, 'rb') as f:
                header = f.read(24)
            ext = os.path.splitext(img_path)[1].lower()
            if ext == '.png' and header[:8] == b'\x89PNG\r\n\x1a\n':
                iw = struct.unpack('>I', header[16:20])[0]
                ih = struct.unpack('>I', header[20:24])[0]
            elif ext in ('.jpg', '.jpeg'):
                # JPEG: scan for SOF marker
                with open(img_path, 'rb') as f:
                    data = f.read(65536)
                pos = 2
                iw = ih = 0
                while pos < len(data) - 8:
                    marker = data[pos:pos+2]
                    if marker[0] != 0xFF:
                        break
                    length = struct.unpack('>H', data[pos+2:pos+4])[0]
                    if marker[1] in (0xC0, 0xC1, 0xC2):
                        ih = struct.unpack('>H', data[pos+5:pos+7])[0]
                        iw = struct.unpack('>H', data[pos+7:pos+9])[0]
                        break
                    pos += 2 + length
                if not iw or not ih:
                    return max_w, max_h
            else:
                return max_w, max_h
            scale = min(max_w / iw, max_h / ih)
            return iw * scale, ih * scale
        except Exception:
            return max_w, max_h

    def _hline(self, slide, l, t, w, color=C_TEAL, weight=0.04):
        """Thin horizontal rule (teal by default)."""
        self._rect(slide, l, t, w, weight, fill=color)

    def _title(self, slide, text, dark=False, size=26):
        """Standard content-slide title: left-aligned, top-left, +rule below.

        Title width is capped at TITLE_SAFE_W so long titles do not collide
        with the top-right 'Make a WELL Change' brand label. Titles that
        exceed TITLE_MAX_CHARS per line emit a warning — they'll still
        render, but the author should shorten or split them.
        """
        # Soft length check (first line only — users can split via \n).
        import sys as _sys
        for line in text.split('\n'):
            if len(line) > TITLE_MAX_CHARS:
                print(
                    f'[westwell-ppt] WARN: title line has {len(line)} chars '
                    f'(soft max {TITLE_MAX_CHARS}):\n'
                    f'                     {line!r}\n'
                    f'                     Consider shortening or using \\n '
                    f'to wrap earlier.',
                    file=_sys.stderr,
                )
                break
        color = C_WHITE if dark else C_NAVY
        self._textbox(slide, TL, TT, TITLE_SAFE_W, TH,
                      text, size=size, bold=True, color=color,
                      font=F_TITLE, align=PP_ALIGN.LEFT)
        rule_color = C_TEAL if not dark else C_TEAL
        # Short accent rule (left-anchored, ~1.2" wide) — Westwell signature
        self._rect(slide, TL, RULE_Y, 1.20, RULE_H, fill=rule_color)
        # Soft long rule under title, light gray
        self._rect(slide, TL + 1.20, RULE_Y + 0.012,
                   TW - 1.20, 0.015,
                   fill=C_MGRAY if not dark else RGBColor(0x3A, 0x56, 0xD0))

    def _new_slide(self, layout_name: str, dark=False,
                   density: str = 'compact') -> object:
        """Add slide, suppress all placeholders.

        density controls which content master is used:
          • 'compact'  → `custom slide2-{mode}` (slim bottom decoration,
            the Westwell-preferred default — lets content breathe without
            the heavy decorative footer competing visually)
          • 'standard' → `custom slide1-{mode}` (full decorative bottom;
            use sparingly when you want extra visual rhythm on an
            otherwise sparse page)

        Default is 'compact' because the full decorative bottom tends to
        out-weigh the content on most pages. Only pass density='standard'
        when you explicitly want the heavier footer for visual variety.
        """
        if density == 'standard':
            name = 'custom slide1-dark' if dark else layout_name
        else:
            name = 'custom slide2-dark' if dark else 'custom slide2-light'
        slide = self.prs.slides.add_slide(self._layout(name))
        self._suppress_placeholders(slide)
        return slide

    # ── Public slide methods ───────────────────────────────────────────────────

    def cover(self, title: str, subtitle: str = '',
              context: str = '', date: str = ''):
        """
        Cover slide. Dark blue background, Westwell geometric art right side.
        Title occupies left ~65% of the slide.
        All layout placeholders are suppressed — content added as free textboxes.
        """
        slide = self.prs.slides.add_slide(self._layout('标题幻灯片'))
        self._suppress_placeholders(slide)

        # Main title (left half, upper-center)
        self._textbox(slide, COV_TL, COV_TT, COV_TW, COV_TH,
                      title, size=40, bold=True, color=C_WHITE,
                      font=F_TITLE, align=PP_ALIGN.LEFT)

        # Thin teal rule below title
        lines = title.count('\n') + 1
        rule_y = COV_TT + lines * 0.68 + 0.15
        self._hline(slide, COV_TL, rule_y, 5.5, C_TEAL, 0.05)

        # Subtitle / context / date stacked below rule
        y = COV_SUB_Y
        if subtitle:
            self._textbox(slide, COV_TL, y, COV_TW, 0.55,
                          subtitle, size=18, color=C_LGRAY,
                          font=F_BODY, align=PP_ALIGN.LEFT)
            y += 0.58
        if context:
            self._textbox(slide, COV_TL, y, COV_TW, 0.40,
                          context, size=14, color=C_TEAL,
                          font=F_TITLE, align=PP_ALIGN.LEFT)
            y += 0.43
        if date:
            self._textbox(slide, COV_TL, y, COV_TW, 0.35,
                          date, size=13, color=C_LGRAY,
                          font=F_TITLE, align=PP_ALIGN.LEFT)
        return slide

    def agenda(self, items: list):
        """
        Table-of-contents slide.
        items: list of strings OR (num, title) OR (num, title, subtitle) tuples.
        Layout: 2-row card grid on white background.
        """
        slide = self._new_slide('custom slide1-light')

        # "目录" header
        self._textbox(slide, TL, TT, 3.0, TH,
                      '目 录', size=26, bold=True, color=C_NAVY,
                      font=F_TITLE, align=PP_ALIGN.LEFT)
        self._hline(slide, TL, RULE_Y, TW, C_TEAL, 0.04)

        n = len(items)
        cols = min(3, n)
        rows = math.ceil(n / cols)

        GAP  = 0.20
        card_w = (CW - GAP * (cols - 1)) / cols
        card_h = (CH - GAP * (rows - 1)) / rows

        for i, item in enumerate(items):
            if isinstance(item, str):
                num = f'{i+1:02d}'
                ttl = item
                sub = ''
            elif len(item) >= 3:
                num, ttl, sub = str(item[0]), str(item[1]), str(item[2])
            else:
                num = str(item[0]) if len(item) > 0 else f'{i+1:02d}'
                ttl = str(item[1]) if len(item) > 1 else ''
                sub = ''

            col_i = i % cols
            row_i = i // cols
            x = CL + col_i * (card_w + GAP)
            y = CT + row_i * (card_h + GAP)

            # Card background (rounded rect)
            self._rect(slide, x, y, card_w, card_h,
                       fill=C_LGRAY, radius=0.12)
            # Left teal accent strip
            self._rect(slide, x, y, 0.06, card_h, fill=C_TEAL)
            # Chapter number
            self._textbox(slide, x + 0.18, y + 0.18, card_w - 0.28, 0.70,
                          num, size=44, bold=True, color=C_TEAL,
                          font=F_ACCENT, align=PP_ALIGN.LEFT)
            # Chapter title
            self._textbox(slide, x + 0.18, y + 0.80, card_w - 0.28,
                          card_h - 0.90,
                          ttl, size=16, bold=True, color=C_NAVY,
                          font=F_BODY, align=PP_ALIGN.LEFT)
            if sub:
                self._textbox(slide, x + 0.18, y + card_h - 0.45,
                              card_w - 0.28, 0.40,
                              sub, size=11, color=C_GRAY,
                              font=F_BODY, align=PP_ALIGN.LEFT)
        return slide

    def chapter(self, num: str, title: str,
                subtitle: str = '', highlight: str = ''):
        """
        Chapter separator. Uses 'agenda slide 2' layout — dark blue + colorful
        right-side geometric art inherited from template.
        Content occupies the left ~55% of the slide so the art stays visible.
        """
        slide = self.prs.slides.add_slide(self._layout('agenda slide 2'))
        self._suppress_placeholders(slide)

        # Constrain content to left side so template's right-side art stays visible.
        # Art extends into bottom-left too, so keep text above y~4.6.
        cx_l = 0.906
        cx_w = 5.30

        # Large chapter number — white on dark bg (text-color rule:
        # navy / black / white only).
        self._textbox(slide, cx_l, 0.90, 3.0, 1.50,
                      num, size=100, bold=True, color=C_WHITE,
                      font=F_ACCENT, align=PP_ALIGN.LEFT)

        # Thin teal rule
        self._hline(slide, cx_l, 2.55, 4.20, C_TEAL, 0.06)

        # Chapter title — white, below number
        lines = title.split('\n')
        self._para_textbox(
            slide, cx_l, 2.75, cx_w, 1.80,
            [{'text': ln, 'size': 30, 'bold': True,
              'color': C_WHITE, 'font': F_TITLE,
              'space_after': 4}
             for ln in lines]
        )
        # subtitle/highlight intentionally ignored: template art occupies
        # the lower-left region and any text there becomes unreadable.
        return slide

    def statement(self, text: str, label: str = '', dark: bool = True,
                  density: str = 'compact'):
        """
        Full-slide statement. A single bold claim, centered.
        Use for pivotal moments: the core thesis, a key data insight,
        or the opening of a chapter.
        label: small caption in top-left (e.g. section context)
        """
        slide = self._new_slide(
            'custom slide1-light', dark=dark, density=density)
        body_color = C_WHITE if dark else C_NAVY

        if label:
            # Centered caption above the main statement, with teal brackets
            self._textbox(slide, TL, 1.80, TW, 0.45,
                          f'— {label} —', size=15, color=C_TEAL,
                          font=F_TITLE, align=PP_ALIGN.CENTER)

        # Each line as its own centered paragraph — avoids alignment drift
        lines = [l for l in text.split('\n') if l.strip()]
        stmt_t = 2.70 if label else 2.40
        self._para_textbox(
            slide, TL, stmt_t, TW, 3.20,
            [{'text': ln, 'size': 36, 'bold': True,
              'color': body_color, 'font': F_BODY,
              'align': PP_ALIGN.CENTER, 'space_after': 18}
             for ln in lines]
        )
        # Short teal accent rule below statement
        self._rect(slide, SW / 2 - 0.6, 5.20, 1.2, 0.05, fill=C_TEAL)
        return slide

    def bullets(self, title: str, points: list,
                dark: bool = False, body_size: int = 19,
                density: str = 'compact'):
        """
        Bulleted list slide.
        points: list of strings. Each is one bullet.
        Design: left-aligned title + teal rule + clean bullet items.
        density: 'standard' (default) or 'compact' — use compact for 5+
        bullets or when the full bottom decoration would crowd content.
        """
        slide = self._new_slide(
            'custom slide1-light', dark=dark, density=density)
        self._title(slide, title, dark=dark)

        body_color = C_WHITE if dark else C_BLACK
        n = len(points)
        gap = 0.10
        item_h = min(0.72, (CH - gap * (n - 1)) / n)

        for i, pt in enumerate(points):
            y = CT + i * (item_h + gap)
            # Teal square marker — vertically centred in item height
            marker_size = 0.10
            marker_y = y + (item_h - marker_size) / 2
            self._rect(slide, CL, marker_y, marker_size, marker_size,
                       fill=C_TEAL)
            # Bullet text — supports **bold** inline markers
            self._rich_textbox(slide, CL + 0.22, y, CW - 0.24, item_h,
                               pt, size=body_size, color=body_color,
                               font=F_BODY)
        return slide

    def text_image(self, title: str, body: str, img_path: str = None,
                   dark: bool = False, body_size: int = 16,
                   density: str = 'compact'):
        """
        Left text (32%) + right image (64%), side by side.

        Extended content area (CB_IMG=6.60) gives both columns 4.80" of
        vertical room — ~23% taller than the standard CB. The image
        column is deliberately wide so images (architecture diagrams,
        flowcharts) are rendered as large as their aspect ratio allows.

        img_path=None → styled placeholder box.
        Supports **bold** inline markers in body text.
        Image is aspect-ratio-aware (never distorted).
        """
        slide = self._new_slide(
            'custom slide1-light', dark=dark, density=density)
        self._title(slide, title, dark=dark)

        body_color = C_WHITE if dark else C_BLACK

        # Extended content area — match image() so side-by-side layouts
        # don't waste the 0.9" below the standard CB.
        CB_IMG = 6.60
        content_h = CB_IMG - CT            # 4.80

        text_w = CW * 0.32                 # 3.69" (narrower than before)
        gap    = CW * 0.03                 # 0.35"
        img_l  = CL + text_w + gap         # 5.00
        img_w  = CW - text_w - gap         # 7.48" (wider than before)

        # Body text — supports **bold** markers
        self._rich_textbox(slide, CL, CT, text_w, content_h,
                           body, size=body_size, color=body_color,
                           font=F_BODY)

        # Light card backdrop behind image (only on light slides, image present)
        if img_path and os.path.exists(img_path) and not dark:
            self._rect(slide, img_l, CT, img_w, content_h,
                       fill=C_LGRAY, radius=0.10)

        # Image (right) — aspect-ratio-aware with inner padding
        if img_path and os.path.exists(img_path):
            pad = 0.22
            w, h = self._img_size(img_path, img_w - pad * 2, content_h - pad * 2)
            x_off = (img_w - w) / 2
            y_off = (content_h - h) / 2
            slide.shapes.add_picture(
                img_path, I(img_l + x_off), I(CT + y_off), I(w), I(h))
        else:
            # Styled placeholder
            ph_fill = RGBColor(0x1A, 0x2E, 0x7A) if dark else C_LGRAY
            self._rect(slide, img_l, CT, img_w, content_h,
                       fill=ph_fill, radius=0.12)
            self._rect(slide, img_l, CT, img_w, 0.06, fill=C_TEAL)
            lbl_color = C_TEAL
            self._textbox(slide, img_l, CT + content_h / 2 - 0.35, img_w, 0.50,
                          '[ 示意图 ]', size=15, bold=False,
                          color=lbl_color, font=F_TITLE,
                          align=PP_ALIGN.CENTER)
        return slide

    def stats(self, title: str, stats: list, dark: bool = False,
              density: str = 'compact'):
        """
        KPI / metric slide.
        stats: list of (value, label, description) tuples. 2–4 items.
        Large teal value + navy label + gray description.

        Value font size auto-shrinks so the longest value fits cleanly:
        mixed values like ("55万", "15–20 分钟", "数十架") all render at
        the same reduced size instead of the long one shrinking alone.

        density: 'standard' or 'compact' — use compact for info-heavy KPI
        pages where the bottom Westwell decoration would compete.
        """
        slide = self._new_slide(
            'custom slide1-light', dark=dark, density=density)
        self._title(slide, title, dark=dark)

        # Auto-shrink value font size so all values render at the same
        # scale (longest one dictates). Chinese and ASCII mixed values are
        # treated by character count — conservative but visually reliable.
        max_len = max(len(str(v)) for v, _, _ in stats)
        if   max_len <= 4:  val_size = 54
        elif max_len <= 6:  val_size = 42
        elif max_len <= 8:  val_size = 34
        elif max_len <= 10: val_size = 28
        else:               val_size = 24

        n = len(stats)
        gap = 0.20
        card_w = (CW - gap * (n - 1)) / n
        card_h = CH

        for i, (val, lbl, desc) in enumerate(stats):
            x = CL + i * (card_w + gap)
            # Light card background on light slides
            if not dark:
                self._rect(slide, x, CT, card_w, card_h,
                           fill=C_LGRAY, radius=0.10)

            # Thin top accent
            self._rect(slide, x, CT, card_w, 0.06, fill=C_TEAL)

            # Value — large KPI number. Text-color rule: navy on light,
            # white on dark (no teal text). wrap=False so values like
            # "15–20 min" stay on a single line.
            val_color = C_WHITE if dark else C_NAVY
            self._textbox(slide, x, CT + 0.25, card_w, 1.30,
                          val, size=val_size, bold=True, color=val_color,
                          font=F_ACCENT, align=PP_ALIGN.CENTER, wrap=False)
            # Label — smaller supporting text, same color rule as value
            lbl_color = C_WHITE if dark else C_NAVY
            self._textbox(slide, x, CT + 1.62, card_w, 0.45,
                          lbl, size=17, bold=True, color=lbl_color,
                          font=F_BODY, align=PP_ALIGN.CENTER)
            # Thin separator
            self._hline(slide, x + 0.55, CT + 2.18, card_w - 1.1,
                        C_TEAL if not dark else C_WHITE, 0.025)
            # Description
            self._textbox(slide, x + 0.20, CT + 2.32, card_w - 0.40, 1.50,
                          desc, size=13, color=C_GRAY if not dark else C_LGRAY,
                          font=F_BODY, align=PP_ALIGN.CENTER, wrap=True)
        return slide

    def two_col(self, title: str,
                left_head: str, left_body: str,
                right_head: str, right_body: str,
                dark: bool = False, density: str = 'compact'):
        """
        Two-column layout. Heading + body in each column.
        A thin vertical teal rule separates the columns.
        density: 'standard' or 'compact' — use compact when each column
        has 6+ lines of body text.
        """
        slide = self._new_slide(
            'custom slide1-light', dark=dark, density=density)
        self._title(slide, title, dark=dark)

        col_w = (CW - 0.30) / 2   # gap between cols = 0.30"
        r_x   = CL + col_w + 0.30

        for x, head, body in [(CL, left_head, left_body),
                               (r_x, right_head, right_body)]:
            # Card background
            card_color = RGBColor(0x1A, 0x2E, 0x7A) if dark else C_LGRAY
            self._rect(slide, x, CT, col_w, CH, fill=card_color, radius=0.10)
            # Top accent strip (teal — non-text accent is allowed)
            self._rect(slide, x, CT, col_w, 0.06, fill=C_TEAL)
            # Column heading — text-color rule: navy on light, white on dark
            head_color = C_WHITE if dark else C_NAVY
            self._textbox(slide, x + 0.28, CT + 0.22, col_w - 0.40, 0.55,
                          head, size=19, bold=True, color=head_color,
                          font=F_BODY, align=PP_ALIGN.LEFT)
            # Thin rule under heading
            self._hline(slide, x + 0.28, CT + 0.82, col_w - 0.56, C_TEAL, 0.022)
            # Column body — supports **bold** markers
            body_color = C_LGRAY if dark else C_BLACK
            self._rich_textbox(slide, x + 0.28, CT + 1.00, col_w - 0.40,
                               CH - 1.15, body, size=17, color=body_color,
                               font=F_BODY)
        return slide

    def three_col(self, title: str, columns: list, dark: bool = False,
                  density: str = 'compact'):
        """
        Three-column insight card layout.
        columns: list of 3 dicts {"head": str, "body": str}
        Best for: executive summary pillars, phase overview, 3-way comparison.
        Supports **bold** inline markers in body text.
        """
        slide = self._new_slide(
            'custom slide1-light', dark=dark, density=density)
        self._title(slide, title, dark=dark)

        GAP    = 0.18
        col_w  = (CW - GAP * 2) / 3

        for i, col in enumerate(columns[:3]):
            x = CL + i * (col_w + GAP)
            card_color = RGBColor(0x1A, 0x2E, 0x7A) if dark else C_LGRAY
            # Card background
            self._rect(slide, x, CT, col_w, CH, fill=card_color, radius=0.10)
            # Teal top accent strip
            self._rect(slide, x, CT, col_w, 0.06, fill=C_TEAL)
            # Column heading — text-color rule: navy on light, white on dark
            head_color = C_WHITE if dark else C_NAVY
            self._textbox(slide, x + 0.24, CT + 0.22, col_w - 0.36, 0.55,
                          col.get('head', ''), size=18, bold=True,
                          color=head_color, font=F_BODY, align=PP_ALIGN.LEFT)
            # Rule under heading
            self._hline(slide, x + 0.24, CT + 0.82, col_w - 0.48, C_TEAL, 0.022)
            # Body text — supports **bold** markers
            body_color = C_LGRAY if dark else C_BLACK
            self._rich_textbox(slide, x + 0.24, CT + 1.00, col_w - 0.36,
                               CH - 1.15, col.get('body', ''),
                               size=16, color=body_color, font=F_BODY)
        return slide

    def table(self, title: str, headers: list, rows: list,
              dark: bool = False, bold_col: int = None,
              density: str = 'compact', body_size: int = 11):
        """
        Data table. Navy header row, alternating body rows.

        body_size: body cell font size (pt). Default 11 is tuned for
        dense tables (6+ cols × 8+ rows). For sparse tables (3–5 cols ×
        4–6 rows) bump to 14–16 so text has visual weight matching its
        generous cell space. Header row scales to body_size + 2.

        density: 'standard' or 'compact' — compact uses a slimmer bottom
        decoration so 6+ row tables do not crowd the footer.
        """
        slide = self._new_slide(
            'custom slide1-light', dark=dark, density=density)
        self._title(slide, title, dark=dark)

        n_rows = 1 + len(rows)
        n_cols = len(headers)
        tbl = slide.shapes.add_table(
            n_rows, n_cols, I(CL), I(CT), I(CW), I(CH - 0.05)
        ).table

        col_w = I(CW) / n_cols
        for col in tbl.columns:
            col.width = int(col_w)

        header_size = body_size + 2

        # Header row
        for ci, hdr in enumerate(headers):
            cell = tbl.cell(0, ci)
            cell.text = hdr
            p = cell.text_frame.paragraphs[0]
            p.font.name = F_TITLE
            p.font.size = Pt(header_size)
            p.font.bold = True
            p.font.color.rgb = C_WHITE
            p.alignment = PP_ALIGN.CENTER
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_NAVY

        # Data rows — text-color rule: navy (bold_col) / black (normal),
        # no teal text.
        for ri, row_data in enumerate(rows):
            bg = C_WHITE if ri % 2 == 0 else C_ALTROW
            for ci, val in enumerate(row_data):
                cell = tbl.cell(ri + 1, ci)
                cell.text = str(val)
                p = cell.text_frame.paragraphs[0]
                p.font.name = F_BODY
                p.font.size = Pt(body_size)
                p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
                if bold_col is not None and ci == bold_col:
                    p.font.bold = True
                    p.font.color.rgb = C_NAVY
                else:
                    p.font.color.rgb = C_BLACK
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg
        return slide

    def image(self, title: str, img_path: str,
              caption: str = '', body: str = '',
              dark: bool = False, density: str = 'compact'):
        """
        Image slide with optional narrative body below.

          • caption — short single-line source note (12pt gray, centered).
          • body    — multi-line narrative explaining what the image shows
                      and which concrete scenario it addresses (15pt,
                      left-aligned, supports **bold** markers). Prefer
                      body over caption when the image needs context.

        Layout:
          • Image is TOP-aligned (not vertically centered) so large
            images get maximum height and the text below reads as a
            natural continuation, not a stranded afterthought.
          • Content area extends down to CB_IMG (~6.60"), deeper than
            the standard CB, because compact layout leaves that region
            clean and an image is the whole point of this slide.
        """
        slide = self._new_slide(
            'custom slide1-light', dark=dark, density=density)
        self._title(slide, title, dark=dark)

        CB_IMG = 6.60             # extended bottom for image slides
        content_h = CB_IMG - CT   # 4.80" vs standard 3.90"

        if body:
            bottom_h = 1.05       # 3 lines at 15pt + breathing room
        elif caption:
            bottom_h = 0.35
        else:
            bottom_h = 0.0
        avail_h = content_h - bottom_h

        if img_path and os.path.exists(img_path):
            w, h = self._img_size(img_path, CW, avail_h)
            # Horizontally center; TOP-align vertically
            x_off = (CW - w) / 2
            slide.shapes.add_picture(
                img_path, I(CL + x_off), I(CT), I(w), I(h))

        if body:
            body_color = C_WHITE if dark else C_BLACK
            self._rich_textbox(
                slide, CL, CB_IMG - bottom_h, CW, bottom_h,
                body, size=15, color=body_color, font=F_BODY,
            )
        elif caption:
            self._textbox(
                slide, CL, CB_IMG - 0.32, CW, 0.32,
                caption, size=12, color=C_GRAY,
                font=F_BODY, align=PP_ALIGN.CENTER,
            )
        return slide

    def text(self, title: str, body: str, dark: bool = False,
             body_size: int = 20, density: str = 'compact'):
        """
        Simple text slide. Title + rule + flowing body text.
        Use for a single statement backed by a prose explanation.
        """
        slide = self._new_slide(
            'custom slide1-light', dark=dark, density=density)
        self._title(slide, title, dark=dark)

        body_color = C_WHITE if dark else C_BLACK
        self._textbox(slide, CL, CT, CW, CH,
                      body, size=body_size, color=body_color,
                      font=F_BODY, wrap=True)
        return slide

    def end(self, text: str = 'THANK YOU', subtitle: str = ''):
        """
        End slide. Uses 'end slide' layout (dark blue + left colorful art).
        Text is centered on the right two-thirds of the slide.
        """
        slide = self.prs.slides.add_slide(self._layout('end slide'))
        self._suppress_placeholders(slide)

        lines = text.split('\n')
        cx = SW / 2
        self._para_textbox(
            slide, cx - 3.5, 2.50, 7.0, 2.0,
            [{'text': ln, 'size': 34, 'bold': True,
              'color': C_WHITE, 'font': F_TITLE,
              'align': PP_ALIGN.CENTER, 'space_after': 8}
             for ln in lines]
        )
        if subtitle:
            self._textbox(slide, cx - 3.5, 4.55, 7.0, 0.60,
                          subtitle, size=16, color=C_LGRAY,
                          font=F_BODY, align=PP_ALIGN.CENTER)
        return slide

    def save(self) -> str:
        os.makedirs(os.path.dirname(os.path.abspath(self.output)),
                    exist_ok=True)
        self.prs.save(self.output)
        return self.output
