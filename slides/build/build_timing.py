# -*- coding: utf-8 -*-
"""時期訴求パート（「今が一番お得」をつくる）研修資料ビルダー

必要性訴求パート・経済メリット訴求パートに続く3本目。
アプローチブック Paragraph 03（p.23〜p.26）＝ 営業ルール⑬⑭⑮ を扱う。
講義25分＋ロープレ30分＋まとめ5分。

勝ち筋は、お客様の頭を「導入するかどうか」から「いつ導入するのが賢いか」へ移すこと。
総額（販売金額）はこのパートでもまだ出さない。

    TRAINING_ASSETS=assets python3.13 slides/build/build_timing.py

画像
  <ASSETS>/approachbook/ab23.jpg 〜 ab26.jpg   アプローチブック誌面（キーメッセージ帯込み）

内容の出典：docs/知識_アプローチブック.md ／ docs/知識_営業ルール20.md ／
「第4回 時期・料金訴求」7〜14・23・26・27枚目 ／「創蓄セット販売 強化合宿」
"""
import copy, os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
MSO_SHAPE_TYPE_AUTO = MSO_SHAPE_TYPE.AUTO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from pptx.opc.packuri import PackURI

REPO   = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
ASSETS = os.environ.get('TRAINING_ASSETS', os.path.join(REPO, 'assets'))
AB     = os.path.join(ASSETS, 'ab')
TPL    = os.path.join(REPO, 'templates', '研修資料_フォーマット見本.pptx')
OUT    = os.path.join(REPO, 'slides', '2026営業研修_時期訴求パート_今が一番お得をつくる.pptx')

CLIENT  = '株式会社山岸'
DURATION = 60            # 研修の所要時間（分）

GOTHIC='HGPｺﾞｼｯｸE'; MINCHO='HGP明朝E'; UD='BIZ UDPゴシック'; YU='游ゴシック'
RED='FF0000'; BLACK='000000'; WHITE='FFFFFF'; GREEN='9BBB59'; YELLOW='FFFF00'
GRAY='595959'; BLUE='0070C0'; LTGREEN='EAF1DD'; LTGRAY='F2F2F2'; LTYEL='FFF2CC'
LTBLUE='DEEBF7'; ORANGE='F79646'

# ---------------------------------------------------------------- helpers
def _rpr_font(run, name):
    rPr = run._r.get_or_add_rPr()
    rPr.get_or_add_latin().set('typeface', name)
    latin = rPr.find(qn('a:latin'))
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = parse_xml('<a:ea xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>')
        latin.addnext(ea)
    ea.set('typeface', name)

def _highlight(run, color):
    rPr = run._r.get_or_add_rPr()
    hl = parse_xml(
        '<a:highlight xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:srgbClr val="%s"/></a:highlight>' % color)
    latin = rPr.find(qn('a:latin'))
    if latin is not None:
        latin.addprevious(hl)
    else:
        rPr.append(hl)

def run(p, text, size=16, bold=False, color=BLACK, font=GOTHIC, hl=None):
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold
    r.font.color.rgb = RGBColor.from_string(color)
    if hl: _highlight(r, hl)
    _rpr_font(r, font)
    return r

def tb(slide, x, y, w, h, wrap=True, anchor=MSO_ANCHOR.TOP, margin=0.0):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(margin)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    return box, tf

def para(tf, first=False, space_after=4, align=PP_ALIGN.LEFT, line=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.space_after = Pt(space_after); p.alignment = align
    if line: p.line_spacing = line
    return p

def lines(tf, items, size=16, color=BLACK, bold=False, space=5, font=GOTHIC, line=None, align=PP_ALIGN.LEFT):
    """items: str | (str,dict)"""
    for i, it in enumerate(items):
        opts = {}
        if isinstance(it, tuple): it, opts = it
        p = para(tf, first=(i == 0), space_after=opts.pop('space', space),
                 align=opts.pop('align', align), line=opts.pop('line', line))
        run(p, it, size=opts.pop('size', size), bold=opts.pop('bold', bold),
            color=opts.pop('color', color), font=opts.pop('font', font),
            hl=opts.pop('hl', None))

def rect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, lw=1.0,
         anchor=MSO_ANCHOR.TOP):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill: s.fill.solid(); s.fill.fore_color.rgb = RGBColor.from_string(fill)
    else: s.fill.background()
    if line:
        s.line.color.rgb = RGBColor.from_string(line); s.line.width = Pt(lw)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    tf = s.text_frame; tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.08)
    tf.margin_top = tf.margin_bottom = Inches(0.04)
    return s, tf

def pic(slide, path, x, y, w, border=True):
    p = slide.shapes.add_picture(path, Inches(x), Inches(y), width=Inches(w))
    if border:
        p.line.color.rgb = RGBColor.from_string('808080'); p.line.width = Pt(1)
    return p

def table(slide, x, y, w, col_w, rows, font_size=11, header=True,
          header_fill='000000', header_color=WHITE, row_h=0.30, head_h=0.30):
    n_r, n_c = len(rows), len(col_w)
    shp = slide.shapes.add_table(n_r, n_c, Inches(x), Inches(y), Inches(w),
                                 Inches(head_h + row_h * (n_r - 1)))
    t = shp.table
    t.first_row = header; t.horz_banding = False
    for i, cw in enumerate(col_w): t.columns[i].width = Inches(cw)
    t.rows[0].height = Inches(head_h)
    for i in range(1, n_r): t.rows[i].height = Inches(row_h)
    for ri, rowdata in enumerate(rows):
        for ci, cell in enumerate(rowdata):
            txt, opts = (cell, {}) if isinstance(cell, str) else cell
            c = t.cell(ri, ci)
            c.margin_left = c.margin_right = Inches(0.05)
            c.margin_top = c.margin_bottom = Inches(0.02)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            fill = opts.get('fill', header_fill if (header and ri == 0) else None)
            if fill: c.fill.solid(); c.fill.fore_color.rgb = RGBColor.from_string(fill)
            else: c.fill.background()
            tf = c.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = opts.get('align',
                PP_ALIGN.CENTER if (header and ri == 0) else PP_ALIGN.LEFT)
            run(p, txt, size=opts.get('size', font_size),
                bold=opts.get('bold', header and ri == 0),
                color=opts.get('color', header_color if (header and ri == 0) else BLACK),
                font=opts.get('font', YU))
    return t

# ---------------------------------------------------------------- slide frame
prs = Presentation(TPL)
BLANK = None
for l in prs.slide_masters[0].slide_layouts:
    if l.name == '白紙': BLANK = l
assert BLANK is not None

sldIdLst = prs.slides._sldIdLst
orig_ids = list(sldIdLst)
cover_id, colophon_id = orig_ids[0], orig_ids[6]
for sid in orig_ids[1:6]:                      # drop template samples 2-6
    rId = sid.get(qn('r:id')); sldIdLst.remove(sid)
    prs.part.drop_rel(rId)
# python-pptx names new slides by len(sldIdLst)+1 -> move colophon out of that range
prs.part.related_part(colophon_id.get(qn('r:id'))).partname = PackURI('/ppt/slides/slide900.xml')

page_no = [0]

def new_slide(heading=None):
    s = prs.slides.add_slide(BLANK)
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    if heading is not None:
        bar = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(0.157), Inches(0.757), Inches(0.140), Inches(0.512))
        bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor.from_string(BLACK)
        bar.line.color.rgb = RGBColor.from_string('D9D9D9'); bar.line.width = Pt(0.75)
        bar.shadow.inherit = False
        box, tf = tb(s, 0.348, 0.736, 9.9, 0.572)
        lines(tf, [heading], size=28, font=MINCHO)
    page_no[0] += 1
    box, tf = tb(s, 8.31, 7.18, 2.30, 0.33)
    lines(tf, [str(page_no[0])], size=12, color=WHITE, font=GOTHIC, align=PP_ALIGN.RIGHT)
    return s

def notes(s, text):
    s.notes_slide.notes_text_frame.text = text

# ---------------------------------------------------------------- 表紙
cover = prs.slides[0]
for shp in cover.shapes:
    if shp.shape_type == MSO_SHAPE_TYPE_AUTO and shp.name == '正方形/長方形 4':
        sp = shp._element.spPr
        fill = sp.find(qn('a:solidFill'))
        for ch in list(fill): fill.remove(ch)
        fill.append(parse_xml(
            '<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'val="E46C0A"><a:alpha val="80000"/></a:srgbClr>'))
    if shp.has_text_frame and shp.name == 'テキスト ボックス 5':
        tf = shp.text_frame; tf.clear()
        lines(tf, [
            ('%s　御中' % CLIENT, {'size': 26, 'bold': True, 'color': WHITE, 'font': UD, 'space': 10}),
            ('時期訴求パート', {'size': 48, 'bold': True, 'color': WHITE, 'font': UD, 'space': 8}),
            ('「今が一番お得」をつくる', {'size': 32, 'bold': True, 'color': WHITE, 'font': UD, 'space': 10}),
            ('〜太陽光＋蓄電池セット販売〜', {'size': 22, 'bold': True, 'color': WHITE, 'font': UD}),
        ])
    if shp.has_text_frame and shp.name == 'テキスト ボックス 6':
        tf = shp.text_frame; tf.clear()
        lines(tf, [('2026年度　所要%d分　株式会社船井総合研究所' % DURATION,
                    {'size': 14, 'color': WHITE, 'font': UD})],
              align=PP_ALIGN.RIGHT)

# ---------------------------------------------------------------- 上部ヘッダー（スライドマスター）
for _m in prs.slide_masters:
    for _shp in _m.shapes:
        if _shp.has_text_frame and '住宅用太陽光' in _shp.text_frame.text:
            _ps = _shp.text_frame.paragraphs[0]
            _rs = _ps.runs
            if _rs:
                _rs[0].text = '%s様　住宅用太陽光・蓄電池研修' % CLIENT
                for _r in _rs[1:]:
                    _r._r.getparent().remove(_r._r)

page_no[0] = 1



# ---------------------------------------------------------------- page helper
def page_slide(s, chapter, title, img, caption, goal, says, conc, caution=None,
               adv=None, panel=None, panel_table=None, img_w=4.95,
               conc_label='緑帯の結論（そのまま言い切る）　'):
    """アプローチブック1ページ＝1枚。

    says : 【基本】これだけ言えば商談として成立する行（初めての人はここだけ追う）
    adv  : 【応用】お客様に合わせて変える・切り返す行（経験者向け）
    panel / panel_table : 誌面画像が無いページで、誌面の中身を再現する左パネル
    """
    box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
    lines(tf, [(title, {'size': 23, 'bold': True})])
    if img:
        # 誌面は下端のキーメッセージ帯まで含めて丸ごと貼る
        ph = pic(s, os.path.join(ASSETS, img), 0.45, 1.90, img_w)
        ih = ph.height / 914400.0
        top = 1.90 + max(0.0, (3.45 - ih)) / 2.0      # 左カラムの中央に寄せる
        ph.top = Emu(int(top * 914400))
        box, tf = tb(s, 0.45, min(top + ih + 0.04, 5.44), img_w, 0.28)
        lines(tf, [(caption, {'size': 11, 'color': GRAY})])
    elif panel or panel_table:
        ph = 3.34
        _, tf = rect(s, 0.45, 1.92, 4.85, ph, fill=LTYEL, line='BF8F00')
        head, body = panel if panel else ('アプローチブックの誌面', [])
        lines(tf, [(head, {'size': 13, 'bold': True, 'color': '7F6000', 'space': 7})])
        for it in body:
            opts = {}
            if isinstance(it, tuple): it, opts = it
            q = para(tf, space_after=opts.pop('space', 6))
            run(q, it, size=opts.pop('size', 13), bold=opts.pop('bold', False),
                color=opts.pop('color', BLACK))
        if panel_table:
            cw, trows, ty = panel_table
            table(s, 0.62, ty, sum(cw), cw, trows, font_size=11.5,
                  row_h=0.29, head_h=0.27)
        box, tf = tb(s, 0.45, 5.32, 4.85, 0.28)
        lines(tf, [(caption, {'size': 11, 'color': GRAY})])
    _, tf = rect(s, 5.65, 1.92, 4.75, 0.95, fill=LTBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    lines(tf, [('引き出す反応（お客様の台詞）', {'size': 11, 'bold': True, 'color': BLUE, 'space': 4})])
    for q in goal:
        p = para(tf, space_after=2); run(p, '「%s」' % q, size=14, bold=True)
    # 基本：初めての人はここだけ追えばよい
    _, tf = rect(s, 5.65, 2.95, 4.75, 1.58, fill=LTGRAY)
    p = para(tf, first=True, space_after=5)
    run(p, '基本', size=11, bold=True, color=WHITE, hl='4F6228')
    run(p, '　これだけ言えば成立します', size=11, bold=True, color=GRAY)
    for it in says:
        opts = {}
        if isinstance(it, tuple): it, opts = it
        q = para(tf, space_after=opts.pop('space', 4))
        run(q, it, size=opts.pop('size', 12), bold=opts.pop('bold', False),
            color=opts.pop('color', BLACK))
    # 応用：経験者はここまで
    _, tf = rect(s, 5.65, 4.59, 4.75, 1.19, fill=LTYEL)
    p = para(tf, first=True, space_after=5)
    run(p, '応用', size=11, bold=True, color=WHITE, hl='C55A11')
    run(p, '　経験者はここまで', size=11, bold=True, color=GRAY)
    for it in (adv or []):
        opts = {}
        if isinstance(it, tuple): it, opts = it
        q = para(tf, space_after=opts.pop('space', 3))
        run(q, it, size=opts.pop('size', 12), bold=opts.pop('bold', False),
            color=opts.pop('color', BLACK))
    _, tf = rect(s, 0.45, 5.80, 9.95, 0.58, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
    p = para(tf, first=True, align=PP_ALIGN.CENTER)
    run(p, conc_label, size=12, bold=True, color=GRAY)
    run(p, conc, size=15, bold=True, hl=YELLOW)
    if caution:
        _, tf = rect(s, 0.45, 6.48, 9.95, 0.50, fill='FCE4E4', line=RED,
                     anchor=MSO_ANCHOR.MIDDLE)
        p = para(tf, first=True, align=PP_ALIGN.CENTER)
        run(p, '注意　', size=13, bold=True, color=RED)
        run(p, caution, size=14, bold=True)

def chapter_slide(s, num, title, pages, minutes, fill, goal, points):
    box, tf = tb(s, 0.45, 1.45, 9.95, 0.55)
    lines(tf, [('%s　%s' % (num, title), {'size': 30, 'bold': True})])
    _, tf = rect(s, 0.45, 2.20, 4.95, 1.30, fill=fill, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [
        ('使うページ', {'size': 13, 'bold': True, 'color': GRAY, 'space': 5}),
        (pages, {'size': 19, 'bold': True}),
    ])
    _, tf = rect(s, 5.60, 2.20, 4.80, 1.30, fill=fill, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [
        ('かける時間', {'size': 13, 'bold': True, 'color': GRAY, 'space': 5}),
        (minutes, {'size': 19, 'bold': True}),
    ])
    _, tf = rect(s, 0.45, 3.70, 9.95, 1.05, fill=None, line=RED, lw=1.5,
                 anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [
        ('この切り口のゴール', {'size': 13, 'bold': True, 'color': GRAY,
                     'align': PP_ALIGN.CENTER, 'space': 4}),
        (goal, {'size': 20, 'bold': True, 'color': RED, 'align': PP_ALIGN.CENTER}),
    ])
    _, tf = rect(s, 0.45, 5.00, 9.95, 1.85, fill=LTGRAY)
    lines(tf, [('めくるときのポイント', {'size': 15, 'bold': True, 'space': 8})])
    for it in points:
        p = para(tf, space_after=6); run(p, '・' + it, size=15)


def page_slide(s, chapter, title, img, caption, goal, says, conc, caution=None,
               adv=None, panel=None, panel_table=None, img_w=4.95,
               conc_label='緑帯の結論（そのまま言い切る）　'):
    """アプローチブック1ページ＝1枚。

    says : 【基本】これだけ言えば商談として成立する行（初めての人はここだけ追う）
    adv  : 【応用】お客様に合わせて変える・切り返す行（経験者向け）
    panel / panel_table : 誌面画像が無いページで、誌面の中身を再現する左パネル
    """
    box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
    lines(tf, [(title, {'size': 23, 'bold': True})])
    if img:
        # 誌面は下端のキーメッセージ帯まで含めて丸ごと貼る
        ph = pic(s, os.path.join(ASSETS, img), 0.45, 1.90, img_w)
        ih = ph.height / 914400.0
        top = 1.90 + max(0.0, (3.45 - ih)) / 2.0      # 左カラムの中央に寄せる
        ph.top = Emu(int(top * 914400))
        box, tf = tb(s, 0.45, min(top + ih + 0.04, 5.44), img_w, 0.28)
        lines(tf, [(caption, {'size': 11, 'color': GRAY})])
    elif panel or panel_table:
        ph = 3.34
        _, tf = rect(s, 0.45, 1.92, 4.85, ph, fill=LTYEL, line='BF8F00')
        head, body = panel if panel else ('アプローチブックの誌面', [])
        lines(tf, [(head, {'size': 13, 'bold': True, 'color': '7F6000', 'space': 7})])
        for it in body:
            opts = {}
            if isinstance(it, tuple): it, opts = it
            q = para(tf, space_after=opts.pop('space', 6))
            run(q, it, size=opts.pop('size', 13), bold=opts.pop('bold', False),
                color=opts.pop('color', BLACK))
        if panel_table:
            cw, trows, ty = panel_table
            table(s, 0.62, ty, sum(cw), cw, trows, font_size=11.5,
                  row_h=0.29, head_h=0.27)
        box, tf = tb(s, 0.45, 5.32, 4.85, 0.28)
        lines(tf, [(caption, {'size': 11, 'color': GRAY})])
    _, tf = rect(s, 5.65, 1.92, 4.75, 0.95, fill=LTBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    lines(tf, [('引き出す反応（お客様の台詞）', {'size': 11, 'bold': True, 'color': BLUE, 'space': 4})])
    for q in goal:
        p = para(tf, space_after=2); run(p, '「%s」' % q, size=14, bold=True)
    # 基本：初めての人はここだけ追えばよい
    _, tf = rect(s, 5.65, 2.95, 4.75, 1.58, fill=LTGRAY)
    p = para(tf, first=True, space_after=5)
    run(p, '基本', size=11, bold=True, color=WHITE, hl='4F6228')
    run(p, '　これだけ言えば成立します', size=11, bold=True, color=GRAY)
    for it in says:
        opts = {}
        if isinstance(it, tuple): it, opts = it
        q = para(tf, space_after=opts.pop('space', 4))
        run(q, it, size=opts.pop('size', 12), bold=opts.pop('bold', False),
            color=opts.pop('color', BLACK))
    # 応用：経験者はここまで
    _, tf = rect(s, 5.65, 4.59, 4.75, 1.19, fill=LTYEL)
    p = para(tf, first=True, space_after=5)
    run(p, '応用', size=11, bold=True, color=WHITE, hl='C55A11')
    run(p, '　経験者はここまで', size=11, bold=True, color=GRAY)
    for it in (adv or []):
        opts = {}
        if isinstance(it, tuple): it, opts = it
        q = para(tf, space_after=opts.pop('space', 3))
        run(q, it, size=opts.pop('size', 12), bold=opts.pop('bold', False),
            color=opts.pop('color', BLACK))
    _, tf = rect(s, 0.45, 5.80, 9.95, 0.58, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
    p = para(tf, first=True, align=PP_ALIGN.CENTER)
    run(p, conc_label, size=12, bold=True, color=GRAY)
    run(p, conc, size=15, bold=True, hl=YELLOW)
    if caution:
        _, tf = rect(s, 0.45, 6.48, 9.95, 0.50, fill='FCE4E4', line=RED,
                     anchor=MSO_ANCHOR.MIDDLE)
        p = para(tf, first=True, align=PP_ALIGN.CENTER)
        run(p, '注意　', size=13, bold=True, color=RED)
        run(p, caution, size=14, bold=True)

def chapter_slide(s, num, title, pages, minutes, fill, goal, points):
    box, tf = tb(s, 0.45, 1.45, 9.95, 0.55)
    lines(tf, [('%s　%s' % (num, title), {'size': 30, 'bold': True})])
    _, tf = rect(s, 0.45, 2.20, 4.95, 1.30, fill=fill, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [
        ('使うページ', {'size': 13, 'bold': True, 'color': GRAY, 'space': 5}),
        (pages, {'size': 19, 'bold': True}),
    ])
    _, tf = rect(s, 5.60, 2.20, 4.80, 1.30, fill=fill, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [
        ('かける時間', {'size': 13, 'bold': True, 'color': GRAY, 'space': 5}),
        (minutes, {'size': 19, 'bold': True}),
    ])
    _, tf = rect(s, 0.45, 3.70, 9.95, 1.05, fill=None, line=RED, lw=1.5,
                 anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [
        ('この切り口のゴール', {'size': 13, 'bold': True, 'color': GRAY,
                     'align': PP_ALIGN.CENTER, 'space': 4}),
        (goal, {'size': 20, 'bold': True, 'color': RED, 'align': PP_ALIGN.CENTER}),
    ])
    _, tf = rect(s, 0.45, 5.00, 9.95, 1.85, fill=LTGRAY)
    lines(tf, [('めくるときのポイント', {'size': 15, 'bold': True, 'space': 8})])
    for it in points:
        p = para(tf, space_after=6); run(p, '・' + it, size=15)


# ================================================================ 2. ゴールと進め方
s = new_slide('はじめに')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('このパートのゴールと、今日の進め方', {'size': 24, 'bold': True})])
_, tf = rect(s, 0.45, 1.92, 9.95, 0.62, fill=LTGRAY, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '講義 25分', size=16, bold=True, color=WHITE, hl='E46C0A')
run(p, '　→　', size=15, color=GRAY)
run(p, 'ロールプレイング 30分', size=16, bold=True, color=WHITE, hl='E46C0A')
run(p, '　→　', size=15, color=GRAY)
run(p, 'まとめ 5分', size=16, bold=True, color=WHITE, hl='E46C0A')
_, tf = rect(s, 0.45, 2.72, 4.80, 2.20, fill=LTGRAY, line='A6A6A6')
p = para(tf, first=True, space_after=8)
run(p, '基本', size=14, bold=True, color=WHITE, hl='4F6228')
run(p, '　初めての方・数回の方', size=14, bold=True)
for it in ['① p.23からp.26まで順に通せる',
           '② 「機会損失」という言葉を使える',
           '③ 10年待つといくら損かを言える']:
    q = para(tf, space_after=6); run(q, it, size=14)
_, tf = rect(s, 5.60, 2.72, 4.80, 2.20, fill=LTYEL, line='BF8F00')
p = para(tf, first=True, space_after=8)
run(p, '応用', size=14, bold=True, color=WHITE, hl='C55A11')
run(p, '　契約率を上げたい方', size=14, bold=True)
for it in ['① 「待ちたい」3つを聞き分けて潰せる',
           '② 「なぜ安くならないか」を根拠で答えられる',
           '③ 二者択一話法で締められる']:
    q = para(tf, space_after=6); run(q, it, size=14)
_, tf = rect(s, 0.45, 5.18, 9.95, 0.72, fill=None, line=RED, lw=1.5,
             anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [('扱うのは アプローチブック p.23 〜 p.26（Paragraph 03 を丸ごと）',
            {'size': 19, 'bold': True, 'align': PP_ALIGN.CENTER})])
_, tf = rect(s, 0.45, 6.14, 9.95, 0.76, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '「基本」だけで商談は成立します。', size=16, bold=True, hl=YELLOW)
run(p, '　「応用」は余力のある人だけ。', size=14)
notes(s, '・講義は25分。残りは全部ロープレに使うと最初に宣言する')

# ================================================================ 3. このパートの勝ち筋
s = new_slide('はじめに')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('このパートの勝ち筋', {'size': 24, 'bold': True})])
_, tf = rect(s, 0.45, 1.90, 4.80, 1.10, fill='FCE4E4', line=RED, anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [
    ('✕　導入するか / しないか', {'size': 17, 'bold': True, 'color': RED,
                        'align': PP_ALIGN.CENTER, 'space': 5}),
    ('この土俵にいる限り、断る理由は', {'size': 13, 'align': PP_ALIGN.CENTER, 'space': 2}),
    ('いくらでも出てきます', {'size': 13, 'align': PP_ALIGN.CENTER}),
])
_, tf = rect(s, 5.60, 1.90, 4.80, 1.10, fill=LTGREEN, line='70A040', anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [
    ('○　今やるか / 10年後にやるか', {'size': 17, 'bold': True,
                          'align': PP_ALIGN.CENTER, 'space': 5}),
    ('どちらを選んでも「やる」前提に', {'size': 13, 'align': PP_ALIGN.CENTER, 'space': 2}),
    ('なっています', {'size': 13, 'align': PP_ALIGN.CENTER}),
])
box, tf = tb(s, 0.45, 3.14, 9.95, 0.36)
lines(tf, [('前回の経済メリットで「1ヶ月でいくら減るか」は出し切っています。あとは時間軸に掛けるだけです',
            {'size': 14, 'color': GRAY})])
rows = [
    ['', '言い方', '数字'],
    [('基本', {'bold': True, 'fill': LTGRAY}),
     ('誌面どおり「10年間、電気代を払い続けるだけです」', {'fill': LTGRAY}),
     ('10年間で 240万円', {'bold': True, 'fill': LTGRAY, 'color': RED})],
    [('応用', {'bold': True, 'fill': LTYEL}),
     ('「前回お出しした削減額が、10年まるごと手に入りません」', {'fill': LTYEL}),
     ('12,952円 × 120ヶ月\n＝ 約155万円', {'bold': True, 'fill': LTYEL, 'color': RED})],
]
table(s, 0.45, 3.56, 9.95, [1.20, 5.90, 2.85], rows, font_size=13.5, row_h=0.84, head_h=0.32)
_, tf = rect(s, 0.45, 5.70, 9.95, 0.72, fill=LTBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
             anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '誌面は「月20,000円のご家庭」の例。', size=13, bold=True, color=BLUE)
run(p, '　お客様の検針票の実額が分かっていれば、その場で計算し直す', size=14, bold=True)
_, tf = rect(s, 0.45, 6.54, 9.95, 0.42, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, 'お客様の頭の中を「導入するか」から「いつ導入するのが賢いか」へ切り替える',
    size=16, bold=True, hl=YELLOW)
notes(s, '・営業マニュアルのルール⑬は「月12,000円で10年144万円」。2026年版の誌面は240万円。誌面に合わせる')

# ================================================================ 4. 使う資料の地図
s = new_slide('はじめに')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('p.23〜p.26 で、どのページで何を言うか', {'size': 24, 'bold': True})])
rows = [
    ['ルール', 'ページ', '見せるもの', '引き出す反応'],
    [('⑬', {'bold': True, 'fill': LTGREEN}), ('p.23', {'bold': True, 'fill': LTGREEN}),
     ('10年後に設置 vs すぐ設置（時間軸の比較）', {'bold': True, 'fill': LTGREEN}),
     ('今、設置した方が将来的にお得なのですね', {'fill': LTGREEN})],
    [('⑭', {'bold': True, 'fill': LTYEL}), ('p.24', {'bold': True, 'fill': LTYEL}),
     ('消費者物価指数（6年で12.7％上昇）', {'bold': True, 'fill': LTYEL}),
     ('待っても安くなるどころか上がるのですね', {'fill': LTYEL})],
    [('⑮', {'bold': True, 'fill': LTBLUE}), ('p.25', {'fill': LTBLUE}),
     ('実質負担額＝設置費用−メリット／賃貸 vs 持ち家', {'fill': LTBLUE}),
     ('長期的には、財産になるのですね', {'fill': LTBLUE})],
    [('', {'fill': LTBLUE}), ('p.26', {'bold': True, 'fill': LTBLUE}),
     ('生涯比較（15年で財産・以降は光熱費0円）', {'bold': True, 'fill': LTBLUE}),
     ('早く設置したら、その分お得ですね', {'fill': LTBLUE})],
]
table(s, 0.45, 1.94, 9.95, [0.90, 1.00, 4.15, 3.90], rows, font_size=12.5, row_h=0.62, head_h=0.32)
_, tf = rect(s, 0.45, 4.78, 4.80, 1.30, fill=LTGRAY)
lines(tf, [
    ('4ページで約8分', {'size': 15, 'bold': True, 'space': 6}),
    ('p.23 と p.26 が山場。p.24・p.25 は', {'size': 13.5, 'space': 3}),
    ('聞かれたときの裏づけとして通す。', {'size': 13.5, 'space': 5}),
    ('ページを行き来しない。', {'size': 13.5, 'bold': True}),
])
_, tf = rect(s, 5.60, 4.78, 4.80, 1.30, fill=None, line=RED, lw=1.5)
lines(tf, [
    ('このパートでも総額は出しません', {'size': 15, 'bold': True, 'color': RED, 'space': 6}),
    ('出すのは「待つといくら損か」まで。', {'size': 13.5, 'space': 3}),
    ('見積り・総額は次のパートです。', {'size': 13.5}),
])
_, tf = rect(s, 0.45, 6.30, 9.95, 0.62, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, 'アプローチブック Paragraph 03「お得に導入できるのは？」を丸ごと通します',
    size=16, bold=True, hl=YELLOW)

# ================================================================ 5. p.23
s = new_slide('① 機会損失')
page_slide(s, '⑬', 'p.23　10年後に設置する vs すぐ設置する',
    'approachbook/ab23.jpg', 'アプローチブック p.23　導入時期について（10年後 vs すぐ）',
    ['今、設置した方が将来的にお得なのですね'],
    [('上下2本の帯を、左から指でなぞる', {'bold': True}),
     '「上が10年後に設置した場合。最初の10年は電気代だけ」',
     ('「10年間で240万円。これは戻ってきません」', {'bold': True, 'color': RED}),
     '「下がすぐ設置。15年あたりで経済効果が出ます」'],
    '設置を遅くするにつれ、経済メリットは少なくなります！',
    '矢印の種類を全部説明しない。「払うだけの10年」と「経済効果が発生」の2つだけ',
    adv=['誌面は月20,000円の例。実額が分かれば計算し直す',
         '売電単価も下がる。「待つ＝デメリットしかない」と言い切る'])

# ================================================================ 6. 機会損失を自分事にする3手
s = new_slide('① 機会損失')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('「機会損失」を自分事にする3手', {'size': 24, 'bold': True})])
box, tf = tb(s, 0.45, 1.86, 9.95, 0.32)
lines(tf, [('240万円を「へえ」で終わらせないための順番です。③まで行くとお客様が自分で言い始めます',
            {'size': 14, 'color': GRAY})])
steps = [
    ('①', '言葉を置く', LTGREEN,
     '「導入を待つということは、この経済効果を先送りにする、つまり',
     '『機会損失』が生じているのと同じなんです」'),
    ('②', '数字を置き換える', LTYEL,
     '「〇〇様は月々約20,000円ですので、10年間ですと',
     '240万円の機会損失になりますね」'),
    ('③', '質問を投げる', LTBLUE,
     '「もしこの240万円を別のことに使えるとしたら、',
     '何に使いたいですか？」'),
]
y = 2.30
for n, t, fill, l1, l2 in steps:
    _, tf = rect(s, 0.45, y, 0.60, 1.14, fill=fill, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [(n, {'size': 24, 'bold': True, 'align': PP_ALIGN.CENTER})])
    _, tf = rect(s, 1.15, y, 2.30, 1.14, fill=fill, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [(t, {'size': 16, 'bold': True, 'align': PP_ALIGN.CENTER})])
    _, tf = rect(s, 3.55, y, 6.85, 1.14, fill=LTGRAY, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [
        (l1, {'size': 14, 'space': 3}),
        (l2, {'size': 14, 'bold': True}),
    ])
    y += 1.26
_, tf = rect(s, 0.45, 6.10, 4.80, 0.82, fill=LTGRAY, line='A6A6A6')
p = para(tf, first=True, space_after=5)
run(p, '基本', size=12, bold=True, color=WHITE, hl='4F6228')
q = para(tf, space_after=0)
run(q, '①と②まで。「機会損失」と言えれば合格', size=13.5, bold=True)
_, tf = rect(s, 5.60, 6.10, 4.80, 0.82, fill=LTYEL, line='BF8F00')
p = para(tf, first=True, space_after=5)
run(p, '応用', size=12, bold=True, color=WHITE, hl='C55A11')
q = para(tf, space_after=0)
run(q, '③まで投げて、お客様に言わせる', size=13.5, bold=True)
notes(s, '・③の答え（旅行・教育費など）が出たら、それを後のクロージングで使う')

# ================================================================ 7. p.24
s = new_slide('② 待っても安くならない')
page_slide(s, '⑭', 'p.24　待っても安くならないの？',
    'approachbook/ab24.jpg', 'アプローチブック p.24　消費者物価指数の推移',
    ['待っても物価は安くなるどころか、上がっていく一方なのですね'],
    [('折れ線は追わない。赤い矢印と右上の数字だけを指す', {'bold': True}),
     '「2020年1月を100とすると、2026年3月は112.7」',
     ('「6年で物価は12.7％上がっています」', {'bold': True, 'color': RED})],
    '待っていても必ずしも安くなるとは限らないのです',
    '「絶対に安くならない」とは言わない。「限らない」と誌面どおりに言う',
    adv=['「待てば安くなる」は機器の値段の話。論点をずらさない',
         '突っ込まれたら、次ページの原材料の話まで出す'])

# ================================================================ 8. 「待ちたい」の正体は3つ
s = new_slide('② 待っても安くならない')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('「待ちたい」の正体は3つ。聞き分けてから潰す', {'size': 23, 'bold': True})])
box, tf = tb(s, 0.45, 1.86, 9.95, 0.32)
lines(tf, [('お客様の「まだいい」は1種類ではありません。潰し方が違うので、まずどれなのかを聞きます',
            {'size': 14, 'color': GRAY})])
rows = [
    ['スマホで言うと', '太陽光で言うと', '潰し方', '使うもの'],
    [('① 今じゃなくていいや', {'bold': True}), 'なんとなく先送り',
     ('時期が空くと、導入しない言い訳を探してしまう', {'bold': True, 'size': 12}),
     ('期日を切る', {'bold': True, 'color': RED})],
    [('② 待てば安くなる', {'bold': True}), '値下がり・補助金待ち',
     ('物価は6年で12.7％上昇。原材料の構造も同じ', {'bold': True, 'size': 12}),
     ('p.24', {'bold': True, 'color': RED})],
    [('③ 型落ちを買おう', {'bold': True}), '旧モデルでいい',
     ('数年で性能は変わらない', {'bold': True}),
     ('−', {'color': GRAY})],
]
table(s, 0.45, 2.28, 9.95, [2.35, 2.15, 4.20, 1.25], rows, font_size=12.5, row_h=0.76, head_h=0.32)
_, tf = rect(s, 0.45, 5.00, 9.95, 1.00, fill=LTBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
lines(tf, [
    ('なぜスマホで例えるのか', {'size': 14, 'bold': True, 'color': BLUE, 'space': 5}),
    ('「最新のスマホが欲しいのに、まだ買っていない人」の言い分と、まったく同じ構造だからです。',
     {'size': 14, 'space': 3}),
    ('お客様自身に心当たりがあるので、否定されずに聞いてもらえます。', {'size': 14, 'bold': True}),
])
_, tf = rect(s, 0.45, 6.18, 9.95, 0.74, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('いちばん多い失敗：どれか分からないまま、全部まとめて潰そうとする', {'size': 15, 'bold': True,
                                               'color': RED, 'space': 4}),
    ('「何か気になる点はありますか？」と先に聞く。①と②では打つ手がまったく違います。', {'size': 14}),
])
notes(s, '・出典：創蓄セット販売 強化合宿（時期訴求パート）')

# ================================================================ 9. なぜ安くならないのか
s = new_slide('② 待っても安くならない')
box, tf = tb(s, 0.45, 1.36, 7.80, 0.45)
lines(tf, [('なぜ蓄電池は安くならないのか', {'size': 24, 'bold': True})])
_, tf = rect(s, 8.45, 1.38, 1.95, 0.42, fill=LTYEL, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '応用', size=12, bold=True, color=WHITE, hl='C55A11')
run(p, '　突っ込まれたら', size=12, bold=True, color=GRAY)
box, tf = tb(s, 0.45, 1.88, 9.95, 0.32)
lines(tf, [('「本当に安くならないの？」と聞かれたときの材料。原材料の供給構造が変わらないからです',
            {'size': 14, 'color': GRAY})])
_, tf = rect(s, 0.45, 2.30, 4.80, 2.20, fill=LTGRAY)
lines(tf, [
    ('炭酸リチウム', {'size': 17, 'bold': True, 'space': 8}),
    ('・世界最大の生産国はチリ', {'size': 14, 'space': 5}),
    ('・日本は100％輸入、うち80％がチリから', {'size': 14, 'space': 5}),
    ('・アタカマ塩湖でかん水を1年天日干し', {'size': 14, 'space': 5}),
    ('　→ 増産しようにも時間がかかる', {'size': 14, 'bold': True, 'color': RED}),
])
_, tf = rect(s, 5.60, 2.30, 4.80, 2.20, fill=LTGRAY)
lines(tf, [
    ('コバルト', {'size': 17, 'bold': True, 'space': 8}),
    ('・レアメタル。コンゴ民主共和国に集中', {'size': 14, 'space': 5}),
    ('・産出の20％が手掘り', {'size': 14, 'space': 5}),
    ('・マンガンはコバルトの約1/10の価格', {'size': 14, 'space': 5}),
    ('　→ 安い材料への置き換えは進むが限界', {'size': 14, 'bold': True, 'color': RED}),
])
_, tf = rect(s, 0.45, 4.66, 9.95, 1.30, fill=LTBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
lines(tf, [
    ('言い方（断定しないこと）', {'size': 14, 'bold': True, 'color': BLUE, 'space': 6}),
    ('「もちろん、将来の価格を正確に予測することは誰にもできません。', {'size': 14, 'space': 3}),
    ('ただ、現在のデータやトレンドから合理的に判断すると、大幅な価格低下を期待して好機を逃すよりも、',
     {'size': 14, 'space': 3}),
    ('確実なメリットを早く享受する方が賢明な選択と言えるのではないでしょうか」', {'size': 14, 'bold': True}),
])
_, tf = rect(s, 0.45, 6.12, 9.95, 0.78, fill='FCE4E4', line=RED)
lines(tf, [
    ('初めての方は、このページは覚えなくて構いません', {'size': 15, 'bold': True,
                                  'color': RED, 'space': 4}),
    ('「安くなりますか？」には、p.24の物価指数だけで十分です。ここは聞かれたときの保険。', {'size': 14}),
])

# ================================================================ 10. p.25
s = new_slide('③ 財産になる')
page_slide(s, '⑮', 'p.25　まとまった費用をかけずに導入する方法',
    'approachbook/ab25.jpg', 'アプローチブック p.25　実質負担額と、賃貸 vs 持ち家',
    ['長期的には、財産になるのですね'],
    [('上の式を指す ─ 設置費用 − メリット ＝ 実質支払額', {'bold': True}),
     '「設置費用は、太陽光が稼いでくれます」',
     ('下の図で「払い続ける」と「払い終わる」を指で比べる', {'bold': True, 'color': RED})],
    '持ち家と同じように、長期的な財産になるものなのです',
    '「節約」で終わらせない。このページのキーワードは「財産」です',
    adv=['「家賃を払いたくないから持ち家。電気代も同じです」',
         '「同じ払うなら、将来の支出を減らすものへ」'])

# ================================================================ 11. p.26
s = new_slide('③ 財産になる')
page_slide(s, '⑮', 'p.26　みなさんは、どちらの生活を選びますか？',
    'approachbook/ab26.jpg', 'アプローチブック p.26　30年後の生涯比較',
    ['早く設置したら、その分お得ですね'],
    [('左右の2つのグラフを並べて指す。数字は2つだけ', {'bold': True}),
     '「左は15年でローンが終わり、以降は光熱費0円」',
     ('「右は何もしない場合。30年で生涯光熱費700万円です」', {'bold': True, 'color': RED})],
    '費用対効果でみれば、早いうちにセット導入するのがお得です！',
    'ここが最後のページ。閉じる前に必ず次のスライドの「握り」をやります',
    adv=['見出しがそのまま二者択一話法。本当に声に出して聞く',
         'フラット35なら家が財産になるのは35年後。創蓄は15年'])

# ================================================================ 12. 「待ちたい」の切り返し
s = new_slide('切り返しと締め')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('「待ちたい」と言われたときの切り返し', {'size': 24, 'bold': True})])
box, tf = tb(s, 0.45, 1.86, 9.95, 0.32)
lines(tf, [('そのまま声に出して言えるようにしておく2本です。棒読みだと見抜かれます',
            {'size': 14, 'color': GRAY})])
_, tf = rect(s, 0.45, 2.28, 9.95, 0.46, fill=LTYEL, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True)
run(p, '①　「補助金を待ちたい」', size=16, bold=True)
run(p, '　← 待ちたい理由②（値下がり待ち）に見えて、実は抽選リスクで潰す', size=12.5, color=GRAY)
_, tf = rect(s, 0.45, 2.80, 9.95, 1.42, fill=LTGRAY)
lines(tf, [
    ('「確かに、補助金があれば嬉しいですよね。ただ実は今、', {'size': 13.5, 'space': 3}),
    ('『補助金待ちで損をする』方が非常に多いんです。', {'size': 13.5, 'bold': True, 'space': 3}),
    ('待っている間に払い続ける電気代は、おそらく年間で15〜20万円以上。しかも補助金は抽選式ですので、',
     {'size': 13.5, 'space': 3}),
    ('必ずしももらえるとは限りません。待っていると製品値上げや申請待ちの渋滞ということも多くて。',
     {'size': 13.5, 'space': 3}),
    ('そう考えると、最も安く導入できるタイミングは間違いなく『今』なんです」',
     {'size': 13.5, 'bold': True, 'color': RED}),
])
_, tf = rect(s, 0.45, 4.36, 9.95, 0.46, fill=LTYEL, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True)
run(p, '②　「また考えて連絡します」', size=16, bold=True)
run(p, '　← 待ちたい理由①（なんとなく）。期日を切る', size=12.5, color=GRAY)
_, tf = rect(s, 0.45, 4.88, 9.95, 1.42, fill=LTGRAY)
lines(tf, [
    ('「お気持ちは分かります。ただ、', {'size': 13.5, 'space': 3}),
    ('時期が空くと導入しない言い訳を探してしまいがちで、絶好のタイミングを逃してしまうかもしれません。',
     {'size': 13.5, 'bold': True, 'space': 3}),
    ('○日11時までにご検討いただけますでしょうか。〇日に私からご連絡しますので、そのときにお教えください。',
     {'size': 13.5, 'space': 3}),
    ('それまでは、このキャンペーン枠を私の責任で確保しておきますね」',
     {'size': 13.5, 'bold': True, 'color': RED}),
])
_, tf = rect(s, 0.45, 6.44, 9.95, 0.48, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, 'どちらも「否定しない → ただ、で受ける → 事実を置く」の順です',
    size=16, bold=True, hl=YELLOW)
notes(s, '・出典：docs/知識_営業ルール20.md §断り文句の切り返し')

# ================================================================ 13. 閉じる前に握る
s = new_slide('切り返しと締め')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('p.26 を閉じる前に、2つ握る', {'size': 24, 'bold': True})])
box, tf = tb(s, 0.45, 1.86, 9.95, 0.32)
lines(tf, [('時期訴求の出口は、そのまま金額訴求の入口です。YESを取ってからアプローチブックを閉じます',
            {'size': 14, 'color': GRAY})])
holds = [
    ('①', '時期について握る', LTGREEN,
     '「費用対効果でみれば、早いうちにセット導入するのがお得です。',
     'ここまではよろしいですか？」'),
    ('②', '予算について握る', LTBLUE,
     '「いま月々◯◯円お支払いですよね。その範囲内に収まれば、',
     '導入に問題はないですか？」'),
]
y = 2.30
for n, t, fill, l1, l2 in holds:
    _, tf = rect(s, 0.45, y, 0.60, 1.22, fill=fill, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [(n, {'size': 24, 'bold': True, 'align': PP_ALIGN.CENTER})])
    _, tf = rect(s, 1.15, y, 2.30, 1.22, fill=fill, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [(t, {'size': 16, 'bold': True, 'align': PP_ALIGN.CENTER})])
    _, tf = rect(s, 3.55, y, 6.85, 1.22, fill=LTGRAY, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [
        (l1, {'size': 14, 'space': 3}),
        (l2, {'size': 14, 'bold': True}),
    ])
    y += 1.34
_, tf = rect(s, 0.45, 4.98, 9.95, 0.94, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('②でYESが取れたら、アプローチブックを閉じて見積りへ', {'size': 15, 'bold': True, 'space': 4}),
    ('「その範囲内に収まれば問題ない」と言わせておくと、次の金額提示がぐっと楽になります。', {'size': 14}),
])
_, tf = rect(s, 0.45, 6.08, 9.95, 0.84, fill='FCE4E4', line=RED, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, 'ここでも総額は出しません。', size=16, bold=True, color=RED)
run(p, '　出すのは次のパート（金額訴求）です', size=15, bold=True)

# ================================================================ 14. まとめ
s = new_slide('まとめ')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.50)
lines(tf, [('このパートで持ち帰ってほしい3つ', {'size': 25, 'bold': True})])
items = [
    ('1', 'お客様の頭を「いつやるか」に切り替える',
     '「導入するか / しないか」で議論している限り、断る理由はいくらでも出てきます。'),
    ('2', '待つといくら損かを、数字で置く',
     '誌面は10年で240万円。お客様の実額が分かれば、その場で計算し直す。'),
    ('3', '「待ちたい」は3つある。聞き分けてから潰す',
     '補助金待ち／値下がり待ち／なんとなく。打つ手がまったく違います。'),
]
y = 1.98
for n, t, d in items:
    _, tf = rect(s, 0.45, y, 0.72, 1.28, fill=RED, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [(n, {'size': 28, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER})])
    _, tf = rect(s, 1.28, y, 9.12, 1.28, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [
        (t, {'size': 19, 'bold': True, 'space': 5}),
        (d, {'size': 14}),
    ])
    y += 1.40
_, tf = rect(s, 0.45, 6.22, 9.95, 0.70, fill=None, line=RED, lw=1.75)
lines(tf, [
    ('次のパートは「金額訴求」。ここで初めて見積りと総額を出します。', {'size': 15, 'bold': True,
                                              'space': 3}),
    ('今日の最後に取った「その範囲内なら問題ない」が、そのまま次の入口になります。', {'size': 14}),
])

# ================================================================ 15. ロープレ
s = new_slide('ロールプレイング')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('通し練習：p.23 から p.26 まで', {'size': 24, 'bold': True})])
_, tf = rect(s, 0.45, 1.90, 9.95, 1.18, fill=LTGRAY)
lines(tf, [('進め方（合計30分）', {'size': 14, 'bold': True, 'color': GRAY, 'space': 6})])
for it in [
    '2人1組。営業役とお客様役。1本8分 → フィードバック4分 → 交代してもう1本',
    'p.23 を開いた状態から始め、p.26 の「握り」2つまで到達する',
    '基本｜4ページを順に通す　応用｜お客様役が「待ちたい」を1つ入れる（3つから選ぶ）',
]:
    p = para(tf, space_after=4); run(p, '・' + it, size=14)
_, tf = rect(s, 0.45, 3.20, 4.80, 1.72, fill=LTBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
lines(tf, [
    ('顧客設定（田中様ご夫婦）', {'size': 14, 'bold': True, 'color': BLUE, 'space': 6}),
    ('・ご主人48歳／奥様45歳／高校生・中学生のお子様', {'size': 12.5, 'space': 3}),
    ('・築14年・南向き切妻屋根・オール電化ではない', {'size': 12.5, 'space': 3}),
    ('・太陽光・蓄電池ともに未導入', {'size': 12.5, 'space': 3}),
    ('・電気代 月20,000円　・住宅ローン返済中', {'size': 12.5, 'space': 3}),
    ('・お子様が来年受験／ご主人は在宅勤務が週2日', {'size': 12.5}),
])
_, tf = rect(s, 5.60, 3.20, 4.80, 1.72, fill=LTGREEN, line='70A040')
lines(tf, [
    ('お客様役・講師が見るところ', {'size': 14, 'bold': True, 'space': 6}),
    ('・「機会損失」という言葉を使えたか', {'size': 12.5, 'space': 3}),
    ('・数字を田中様の実額に置き換えたか', {'size': 12.5, 'space': 3}),
    ('・「待ちたい」がどれなのかを先に聞いたか', {'size': 12.5, 'space': 3}),
    ('・最後の握り2つを取れたか', {'size': 12.5, 'space': 3}),
    ('・総額を口にしていないか', {'size': 12.5}),
])
_, tf = rect(s, 0.45, 5.06, 9.95, 0.96, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('フィードバックの型', {'size': 15, 'bold': True, 'space': 5}),
    ('① 良かったページを1つ、ページ番号で言う　② 直すのは1人1つだけ　③ 人ではなくページを指摘する',
     {'size': 14, 'bold': True}),
])
_, tf = rect(s, 0.45, 6.18, 9.95, 0.72, fill=None, line=RED, lw=1.5,
             anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [('お客様役は「ローン返済中なんだよね」と必ず言う。「追加の出費」ではなく「置き換え」と返せるか見ます',
            {'size': 14, 'bold': True, 'align': PP_ALIGN.CENTER})])

# ================================================================ 16. チェックシート
s = new_slide('ロールプレイング')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('セルフチェックシート', {'size': 24, 'bold': True})])
box, tf = tb(s, 0.45, 1.86, 9.95, 0.32)
lines(tf, [('ロープレのあとに記入し、できていない項目を次回までの宿題にしてください',
            {'size': 14, 'color': GRAY})])
rows = [
    ['確認すること', '今日', '次回'],
    ['p.23で「10年間で240万円」を言えた', ' ', ' '],
    ['「機会損失」という言葉を使えた', ' ', ' '],
    ['数字をお客様の実額に置き換えた', ' ', ' '],
    ['p.24で「6年で12.7％」を言えた', ' ', ' '],
    ['p.26で「どちらの生活を選びますか？」と聞けた', ' ', ' '],
    ['最後の握り2つを取れた', ' ', ' '],
    ['総額（販売金額）を自分から出さなかった', ' ', ' '],
]
table(s, 0.45, 2.28, 6.35, [4.35, 1.00, 1.00], rows, font_size=13, row_h=0.50, head_h=0.34)
_, tf = rect(s, 7.05, 2.26, 3.35, 1.95, fill=LTGRAY, line='A6A6A6')
p = para(tf, first=True, space_after=5)
run(p, '基本', size=12, bold=True, color=WHITE, hl='4F6228')
run(p, '　合格ライン', size=12, bold=True, color=GRAY)
q = para(tf, space_after=4); run(q, '7項目中 5つ', size=22, bold=True, color=RED)
q = para(tf, space_after=0)
run(q, 'うち「総額を出さなかった」は必須', size=12)
_, tf = rect(s, 7.05, 4.34, 3.35, 1.20, fill=LTYEL, line='BF8F00')
p = para(tf, first=True, space_after=5)
run(p, '応用', size=12, bold=True, color=WHITE, hl='C55A11')
run(p, '　合格ライン', size=12, bold=True, color=GRAY)
q = para(tf, space_after=0)
run(q, '7項目すべて＋「待ちたい」の切り返しが出る', size=12.5, bold=True)
_, tf = rect(s, 7.05, 5.66, 3.35, 1.26, fill=LTGRAY)
lines(tf, [
    ('次回までの宿題', {'size': 14, 'bold': True, 'space': 5}),
    ('担当顧客1件について、', {'size': 12, 'space': 2}),
    ('10年待った場合の機会損失額を', {'size': 12, 'space': 2}),
    ('計算してくること。', {'size': 12, 'space': 5}),
    ('翌日には50％忘れます。', {'size': 12, 'bold': True, 'color': RED}),
])

# ---------------------------------------------------------------- 並べ替え
ids = list(sldIdLst)
colophon = [x for x in ids if x is colophon_id][0]
sldIdLst.remove(colophon)
sldIdLst.append(colophon)
page_no[0] += 1

os.makedirs(os.path.dirname(OUT), exist_ok=True)
prs.save(OUT)
print('saved:', OUT, '/ slides:', len(prs.slides.__iter__.__self__._sldIdLst))

# ---------------------------------------------------------------- content-type fix
def ensure_jpg_content_type(path):
    """テンプレートが jpg を image/png と誤宣言しているため、python-pptx が
    保存時に Default を落とすことがある。表紙画像が壊れるので明示的に直す。"""
    import zipfile, shutil, re, tempfile
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        ct = z.read('[Content_Types].xml').decode('utf-8')
        needs = any(n.lower().endswith('.jpg') for n in names)
        if not needs or 'Extension="jpg"' in ct:
            if 'Extension="jpg" ContentType="image/jpeg"' in ct or not needs:
                return False
        data = {n: z.read(n) for n in names}
    ct = re.sub(r'<Default Extension="jpg"[^/]*/>', '', ct)
    ct = ct.replace('<Default Extension="png"',
                    '<Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="png"', 1)
    data['[Content_Types].xml'] = ct.encode('utf-8')
    tmp = path + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, data[n])
    shutil.move(tmp, path)
    return True

if ensure_jpg_content_type(OUT):
    print('fixed [Content_Types].xml (jpg -> image/jpeg)')
