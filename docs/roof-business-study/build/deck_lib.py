# -*- coding: utf-8 -*-
"""屋根ビジネス研究会 学習資料 共通描画ライブラリ"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import copy

NAVY   = RGBColor(0x00, 0x20, 0x60)
BLUE   = RGBColor(0x00, 0x70, 0xC0)
ORANGE = RGBColor(0xFF, 0xC0, 0x00)
RED    = RGBColor(0xFF, 0x00, 0x00)
DKRED  = RGBColor(0xC0, 0x00, 0x00)
BLACK  = RGBColor(0x00, 0x00, 0x00)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)

FONT = 'メイリオ'
SW, SH = 10.0, 7.5          # inches (4:3)
MARGIN = 0.3
CW = SW - MARGIN * 2        # content width 9.4

_registry = []              # (slide_idx, shape, box_w_pt, box_h_pt) for overflow check


def new_deck():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    return prs


def _set_font(run, size, color, bold=False, font=FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:latin', 'a:ea', 'a:cs'):
        for e in rPr.findall(qn(tag)):
            rPr.remove(e)
        el = rPr.makeelement(qn(tag), {'typeface': font})
        rPr.append(el)


def textbox(slide, l, t, w, h, lines, size=12, color=BLACK, bold=False,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.15, track=True):
    """lines: str or list of str or list of (text, size, color, bold) tuples"""
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = anchor
    if isinstance(lines, str):
        lines = [lines]
    first = True
    for ln in lines:
        if isinstance(ln, tuple):
            txt, sz, col, bd = (list(ln) + [size, color, bold])[:4]
        else:
            txt, sz, col, bd = ln, size, color, bold
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = spacing
        r = p.add_run()
        r.text = txt
        _set_font(r, sz, col, bd)
    if track:
        _registry.append((len(slide.shapes.parent.slides._sldIdLst) if False else None,
                          tb, w * 72, h * 72, lines, size))
    return tb


def rect(slide, l, t, w, h, fill=None, line=None, line_w=1.25, shape=MSO_SHAPE.RECTANGLE, radius=None):
    sp = slide.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    sp.text_frame.text = ''
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    return sp


def label(slide, l, t, w, h, text, size=12, color=BLACK, bold=False,
          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, fill=None, line=None, spacing=1.1):
    if fill is not None or line is not None:
        rect(slide, l, t, w, h, fill, line)
    return textbox(slide, l, t, w, h, text, size, color, bold, align, anchor, spacing)


def arrow(slide, l, t, w, h, color=BLUE, direction='right'):
    shp = MSO_SHAPE.RIGHT_ARROW if direction == 'right' else MSO_SHAPE.DOWN_ARROW
    sp = slide.shapes.add_shape(shp, Inches(l), Inches(t), Inches(w), Inches(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = color
    sp.line.fill.background(); sp.shadow.inherit = False
    return sp


def base_slide(prs, chapter, title, page, title_size=24):
    """共通ヘッダー・タイトル帯・フッターを持つスライドを返す"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    # ヘッダー
    textbox(s, MARGIN, 0.07, CW, 0.24,
            [('Roof Reform-Manual　' + chapter, 9, NAVY, False)],
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, track=False)
    line = rect(s, MARGIN, 0.30, CW, 0.012, fill=NAVY)
    # タイトル帯
    if title:
        rect(s, MARGIN, 0.38, CW, 0.62, fill=NAVY)
        textbox(s, MARGIN + 0.12, 0.38, CW - 0.24, 0.62, title, title_size, WHITE, True,
                PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE, track=False)
    _footer(s, page)
    return s


def _footer(s, page):
    textbox(s, MARGIN, 7.12, 6.0, 0.26,
            [('Copyright©2020 Funai Consulting Inc. All rights reserved.', 9, NAVY, False)],
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, track=False)
    textbox(s, SW - MARGIN - 1.0, 7.12, 1.0, 0.26, [(str(page), 9, NAVY, False)],
            align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, track=False)


def source(s, text):
    textbox(s, MARGIN, 6.80, CW, 0.26, [('出典：' + text, 9, BLUE, False)],
            align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, track=False)


def phase_tabs(s, tabs, active, top=1.08, left=MARGIN, width=CW, h=0.34, size=11):
    n = len(tabs)
    gap = 0.06
    w = (width - gap * (n - 1)) / n
    for i, t in enumerate(tabs):
        on = (i == active)
        label(s, left + i * (w + gap), top, w, h, t, size,
              WHITE if on else BLUE, True, fill=BLUE if on else WHITE,
              line=BLUE)


def table(s, l, t, w, rows, col_w=None, header_fill=NAVY, size=11, row_h=0.30,
          header_size=11, first_col_fill=None, align_first=PP_ALIGN.LEFT):
    """rows[0] をヘッダーとして描画。col_w は比率リスト。"""
    nr, nc = len(rows), len(rows[0])
    tbl_shape = s.shapes.add_table(nr, nc, Inches(l), Inches(t), Inches(w), Inches(row_h * nr))
    tbl = tbl_shape.table
    tbl.first_row = False
    tbl.horz_banding = False
    if col_w:
        tot = sum(col_w)
        for i, cwr in enumerate(col_w):
            tbl.columns[i].width = Emu(int(Inches(w) * cwr / tot))
    for ri, row in enumerate(rows):
        tbl.rows[ri].height = Inches(row_h)
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.margin_left = cell.margin_right = Inches(0.05)
            cell.margin_top = cell.margin_bottom = Inches(0.01)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if ri == 0:
                cell.fill.fore_color.rgb = header_fill
            elif ci == 0 and first_col_fill is not None:
                cell.fill.fore_color.rgb = first_col_fill
            else:
                cell.fill.fore_color.rgb = WHITE
            tf = cell.text_frame
            tf.word_wrap = True
            parts = val.split('\n') if isinstance(val, str) else [str(val)]
            for pi, ptxt in enumerate(parts):
                p = tf.paragraphs[0] if pi == 0 else tf.add_paragraph()
                p.line_spacing = 1.0
                if ri == 0:
                    p.alignment = PP_ALIGN.CENTER
                elif ci == 0:
                    p.alignment = align_first
                else:
                    p.alignment = PP_ALIGN.LEFT
                r = p.add_run(); r.text = ptxt
                if ri == 0:
                    _set_font(r, header_size, WHITE, True)
                elif ci == 0 and first_col_fill is not None:
                    _set_font(r, size, WHITE, True)
                else:
                    _set_font(r, size, BLACK, False)
    return tbl


def bullets(txts, mark='・'):
    return [mark + t for t in txts]
