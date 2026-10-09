# -*- coding: utf-8 -*-
"""経済メリット訴求パート（シミュレーションの見せ方）研修資料ビルダー

「必要性訴求パート」（①電気代・②災害対策）の続き＝③経済メリット。
講義25分＋ロープレ30分＋まとめ5分を想定し、講義パートを18枚に収めている。

基本方針は「伝え方②（電気代の減収分のみ）」。
ローンの話は後に回し、販売金額は最後に出す（合宿資料 p.47「分割して伝える」）。

    TRAINING_ASSETS=assets python3 slides/build/build_econ.py

画像
  <ASSETS>/approachbook/ab20.jpg, ab21.jpg      アプローチブック誌面
  <ASSETS>/simulation/sim34_cmp.jpg             P3・P4 の電気利用の流れを並べたもの
  <ASSETS>/simulation/sim6_bar.jpg              P6 上段（棒グラフ＋光熱費表）
  <ASSETS>/simulation/sim6_break.jpg            P6 下段（節約額の内訳）
  <ASSETS>/simulation/sim7_sum.jpg              P7 下段（累計削減額＋表）
  <ASSETS>/simulation/sim8_calc.jpg             P8 下段（実質負担額の引き算）
  <ASSETS>/simulation/sim8_loan.jpg             P8 上段（資金計画＝最後に出す）
  <ASSETS>/simulation/trust754.jpg              信憑性を疑ったことがある人 75.4％

内容の出典：docs/知識_シミュレーション.md ／ docs/知識_営業ルール20.md ／
「創蓄セット販売 強化合宿」／「第3回 必要性訴求②」
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
OUT    = os.path.join(REPO, 'slides', '2026営業研修_経済メリットパート_シミュレーションの見せ方.pptx')

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
            'val="31859B"><a:alpha val="80000"/></a:srgbClr>'))
    if shp.has_text_frame and shp.name == 'テキスト ボックス 5':
        tf = shp.text_frame; tf.clear()
        lines(tf, [
            ('%s　御中' % CLIENT, {'size': 26, 'bold': True, 'color': WHITE, 'font': UD, 'space': 10}),
            ('経済メリット訴求パート', {'size': 44, 'bold': True, 'color': WHITE, 'font': UD, 'space': 8}),
            ('シミュレーションの見せ方', {'size': 32, 'bold': True, 'color': WHITE, 'font': UD, 'space': 10}),
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

# ================================================================ 2. ゴールと進め方
s = new_slide('はじめに')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('このパートのゴールと、今日の進め方', {'size': 24, 'bold': True})])
_, tf = rect(s, 0.45, 1.92, 9.95, 0.62, fill=LTGRAY, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '講義 25分', size=16, bold=True, color=WHITE, hl='31859B')
run(p, '　→　', size=15, color=GRAY)
run(p, 'ロールプレイング 30分', size=16, bold=True, color=WHITE, hl='31859B')
run(p, '　→　', size=15, color=GRAY)
run(p, 'まとめ 5分', size=16, bold=True, color=WHITE, hl='31859B')
_, tf = rect(s, 0.45, 2.72, 4.80, 2.20, fill=LTGRAY, line='A6A6A6')
p = para(tf, first=True, space_after=8)
run(p, '基本', size=14, bold=True, color=WHITE, hl='4F6228')
run(p, '　初めての方・数回の方', size=14, bold=True)
for it in ['① 見せるページを絞れる',
           '② 残す数字3つを言い切れる',
           '③ 削減額を言い切ってから支払いに入る']:
    q = para(tf, space_after=6); run(q, it, size=14)
_, tf = rect(s, 5.60, 2.72, 4.80, 2.20, fill=LTYEL, line='BF8F00')
p = para(tf, first=True, space_after=8)
run(p, '応用', size=14, bold=True, color=WHITE, hl='C55A11')
run(p, '　契約率を上げたい方', size=14, bold=True)
for it in ['① 前提条件を先に自分から言える',
           '② 「FIT後は？」「上がらなかったら？」に即答',
           '③ 反応を見て伝え方①に切り替えられる']:
    q = para(tf, space_after=6); run(q, it, size=14)
_, tf = rect(s, 0.45, 5.18, 9.95, 0.72, fill=None, line=RED, lw=1.5,
             anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [('扱うのは アプローチブック p.20・p.21 ＋ 診断レポート（エネがえる）',
            {'size': 19, 'bold': True, 'align': PP_ALIGN.CENTER})])
_, tf = rect(s, 0.45, 6.14, 9.95, 0.76, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '「基本」だけで商談は成立します。', size=16, bold=True, hl=YELLOW)
run(p, '　「応用」は余力のある人だけ。', size=14)
notes(s, '・講義は25分。残りは全部ロープレに使うと最初に宣言する')

# ================================================================ 3. 位置づけ
s = new_slide('はじめに')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('経済メリットは、商談のどこに効くのか', {'size': 24, 'bold': True})])
_, tf = rect(s, 0.45, 1.92, 9.95, 0.92, fill=LTBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
             anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '購入　＝　', size=24, bold=True)
run(p, '価値', size=26, bold=True, color='1F6391')
run(p, '　÷　', size=24, bold=True)
run(p, '価格', size=26, bold=True, color=RED)
run(p, '　　この値が1を超えたとき、人は購入に至る', size=14, color=GRAY)
_, tf = rect(s, 0.45, 3.02, 4.80, 1.30, fill=LTGRAY)
lines(tf, [
    ('価値を上げる', {'size': 16, 'bold': True, 'color': '1F6391', 'space': 6}),
    ('① 電気代が高騰している（前回）', {'size': 14, 'space': 4}),
    ('② 災害対策（前回）', {'size': 14, 'space': 4}),
    ('＋ 商品の価値（FAB）', {'size': 14}),
])
_, tf = rect(s, 5.60, 3.02, 4.80, 1.30, fill='FCE4E4', line=RED)
lines(tf, [
    ('価格を下げる', {'size': 16, 'bold': True, 'color': RED, 'space': 6}),
    ('③ 経済メリット（今回）', {'size': 14, 'bold': True, 'space': 4}),
    ('「価格そのもの」ではなく', {'size': 14, 'space': 4}),
    ('「実質いくらになるか」を下げる', {'size': 14}),
])
_, tf = rect(s, 0.45, 4.50, 9.95, 1.55, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('ただし ③は「決め手」であって「入り口」ではない', {'size': 18, 'bold': True,
                                     'color': RED, 'space': 7}),
    ('狭小住宅・北向き屋根・電気使用量が少ないお宅では、数字が出ません。', {'size': 14, 'space': 4}),
    ('「元が取れる／取れない」で判断させると、そういうお宅に提案できなくなります。', {'size': 14, 'space': 4}),
    ('①②と商品の価値を積んだうえで、最後に③で背中を押す。順番を入れ替えないこと。',
     {'size': 14, 'bold': True}),
])
_, tf = rect(s, 0.45, 6.22, 9.95, 0.62, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '③経済メリットは、価格を下げる側の武器。入り口にはしない', size=17, bold=True, hl=YELLOW)

# ================================================================ 4. 今日の基本方針
s = new_slide('はじめに')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('今日の基本方針：まず「いくら減るか」だけを言い切る', {'size': 23, 'bold': True})])
_, tf = rect(s, 0.45, 1.92, 4.80, 0.78, fill='FCE4E4', line=RED, anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [
    ('✕　いきなり「月々◯◯円のお支払いで」から入る', {'size': 14, 'bold': True,
                                     'color': RED, 'space': 3}),
    ('→ 支払額だけが記憶に残る', {'size': 13, 'color': GRAY}),
])
_, tf = rect(s, 5.60, 1.92, 4.80, 0.78, fill=LTGREEN, line='70A040', anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [
    ('○　まず「電気代がいくら減るか」だけを言い切る', {'size': 14, 'bold': True, 'space': 3}),
    ('→ ローンの話は、そのあと', {'size': 13, 'color': GRAY}),
])
rows = [
    ['', '伝え方② ＝ 今日の基本', '伝え方① ＝ 切り替え用'],
    ['見せるもの', ('電気代の減収分だけ。ローンは外す', {'bold': True}),
     'ローン返済額を含めたトータル金額'],
    ['向いている方', ('まず全員これ。現金一括／月々の支出増に抵抗がある方／数字が多いと混乱される方',
                 {'bold': True}),
     '長期で考えられる方／完済後の姿に価値を感じる方'],
    ['切り替える合図', '−', ('②で「で、いくら払うの？」と聞かれたとき', {'bold': True})],
]
table(s, 0.45, 2.90, 9.95, [1.60, 4.35, 4.00], rows, font_size=13, row_h=0.78, head_h=0.34)
_, tf = rect(s, 0.45, 5.62, 4.80, 1.00, fill=LTGRAY, line='A6A6A6')
p = para(tf, first=True, space_after=5)
run(p, '基本', size=12, bold=True, color=WHITE, hl='4F6228')
q = para(tf, space_after=0)
run(q, '迷ったら②。ローンの話を自分から持ち出さない', size=13.5, bold=True)
_, tf = rect(s, 5.60, 5.62, 4.80, 1.00, fill=LTYEL, line='BF8F00')
p = para(tf, first=True, space_after=5)
run(p, '応用', size=12, bold=True, color=WHITE, hl='C55A11')
q = para(tf, space_after=0)
run(q, '②で得の実感が薄いときだけ①に切り替える', size=13.5, bold=True)
_, tf = rect(s, 0.45, 6.78, 9.95, 0.42, fill=None, line=RED, lw=1.5,
             anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [('出典：創蓄セット販売 強化合宿 p.17・p.18',
            {'size': 12, 'color': GRAY, 'align': PP_ALIGN.CENTER})])

# ================================================================ 5. 使う資料と見せるページ
s = new_slide('はじめに')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('使う資料と、見せるページ', {'size': 24, 'bold': True})])
_, tf = rect(s, 0.45, 1.92, 9.95, 0.52, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [('アプローチブック p.20・p.21　→　診断レポート（9ページ）',
            {'size': 17, 'bold': True, 'align': PP_ALIGN.CENTER})])
rows = [
    ['診断レポート', 'ページ名', '商談での扱い'],
    ['P1', '診断条件', '見せない（事前に自分が確認）'],
    ['P2', '太陽光 月別発電量', '聞かれたら開く'],
    [('P3・P4', {'bold': True}), ('現在の／FIT中の使い方', {'bold': True}),
     ('対比用に一瞬だけ', {'bold': True})],
    ['P5', 'FIT終了後の使い方', '見せない（P4と同じ）'],
    [('P6', {'bold': True, 'fill': LTYEL}), ('1ヶ月のシミュレーション', {'bold': True, 'fill': LTYEL}),
     ('★ 主役① いくら減るか', {'bold': True, 'fill': LTYEL, 'color': RED})],
    [('P7', {'bold': True}), ('長期シミュレーション', {'bold': True}), ('裏づけとして一瞬', {'bold': True})],
    [('P8', {'bold': True, 'fill': LTYEL}), ('お支払いシミュレーション', {'bold': True, 'fill': LTYEL}),
     ('★ 主役② いくら払うか', {'bold': True, 'fill': LTYEL, 'color': RED})],
    ['P9', '自家消費優先（特定月）', '聞かれたら開く'],
]
table(s, 0.45, 2.58, 6.20, [1.30, 2.70, 2.20], rows, font_size=12, row_h=0.44, head_h=0.32)
_, tf = rect(s, 6.95, 2.58, 3.45, 2.10, fill=LTBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
lines(tf, [
    ('お客様に残す数字は3つだけ', {'size': 14, 'bold': True, 'color': BLUE, 'space': 8}),
    ('① 1ヶ月の実質削減額（P6）', {'size': 13, 'space': 6}),
    ('② 25年間の累計削減額（P7）', {'size': 13, 'space': 6}),
    ('③ 1日あたりの実質負担額（P8）', {'size': 13}),
])
_, tf = rect(s, 6.95, 4.84, 3.45, 1.52, fill=LTGRAY)
lines(tf, [
    ('全部めくらない', {'size': 14, 'bold': True, 'space': 7}),
    ('9ページのうち、見せるのは', {'size': 13, 'space': 3}),
    ('P3・P4・P6・P7・P8 の5枚。', {'size': 13, 'bold': True, 'space': 6}),
    ('残りは聞かれたときの裏づけ。', {'size': 13}),
])
_, tf = rect(s, 0.45, 6.48, 9.95, 0.46, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, 'ページを行き来しない。綴じ順に1回通すだけで話が終わるように作られています',
    size=15, bold=True, hl=YELLOW)

# ================================================================ 6. p.20
s = new_slide('アプローチブック')
page_slide(s, '①', 'p.20　電気を購入する生活から電気を自給自足する生活へ',
    'approachbook/ab20.jpg', 'アプローチブック p.20　未導入のお家と、太陽光＋蓄電池のお家',
    ['電気を買うより、自給自足した方が良いですね'],
    [('左右の家の絵を順に指す。文章は読まない', {'bold': True}),
     '「①の電気代と②の災害、この2つへの答えがこれです」',
     '「自分で作って、自分で使う。これが答えです」'],
    '上昇する電気代を削減するために、自給自足が求められます',
    'ここで金額の話をしない。「考え方」を1つ置くだけのページです',
    adv=['①②で作ったネガを、ここで初めて解決策に変える',
         '「賦課金も使った分だけ」とp.14に戻してもよい'])

# ================================================================ 7. p.21
s = new_slide('アプローチブック')
page_slide(s, '①', 'p.21　太陽光発電・蓄電池の使い方',
    'approachbook/ab21.jpg', 'アプローチブック p.21　1日の電気の流れ',
    ['太陽光・蓄電池の活用方法が分かる'],
    [('図を左から順に指でなぞる（朝 → 昼 → 夕 → 深夜）', {'bold': True}),
     '「朝は買う。昼は作って使い、余りを貯める」',
     '「夕方からは、貯めた電気を使う」',
     ('「売電15円より買う電気30円が高い。だから売らずに使う」',
      {'bold': True, 'color': RED})],
    '具体的なシミュレーションを見ていきましょう！',
    'ここが曖昧だと、次のシミュレーションが何を言っているか伝わりません',
    adv=['この図は、あとで診断レポート P4 の流れ図とそのまま対応する',
         '図を見ずに「朝・昼・夕」の3つで言えるようにしておく'])

# ================================================================ 8. P3・P4 対比
s = new_slide('診断レポート')
page_slide(s, '②', 'P3・P4　使う量は同じ。買う量が減る',
    'simulation/sim34_cmp.jpg', '診断レポート P3（現在）／P4（FIT期間中）　1ヶ月の電気利用の流れ',
    ['買う電気がこんなに減るんですね'],
    [('左が「現在」、右が「導入後」。同じ絵なので並べて見せる', {'bold': True}),
     '「1ヶ月に使う量は470kWhのまま、変えていません」',
     ('「自給率が0％から86.75％になります」', {'bold': True, 'color': RED})],
    '使う量は変わらない。買う量が減るだけです',
    'P5（FIT終了後）はP4とほぼ同じ絵。並べると「さっきと同じでは」と言われます',
    adv=['数字が苦手なお客様には、金額よりこの2枚のほうが速い',
         '「使用のピークは19時〜20時」＝発電しない時間、を押さえておく'],
    img_w=5.15)

# ================================================================ 9. P6-① 棒グラフ
s = new_slide('診断レポート')
page_slide(s, '②', 'P6-①　棒の長さは3本とも同じ。中身が割れるだけ',
    'simulation/sim6_bar.jpg', '診断レポート P6　1ヶ月の光熱費の比較',
    ['払う総額が、こう分かれるんですね'],
    [('まず「3本の棒、長さは同じですよね」と確認させる', {'bold': True}),
     '「灰色が、これからも払う電気代です」',
     '「緑が蓄電池、オレンジが太陽光で減った分です」'],
    '安くなるのではなく、払い先が電力会社から自分の設備に替わります',
    '棒の内訳を全部読み上げない。灰色・緑・オレンジの3色だけ',
    adv=['下の黄色い棒は「売電収入」。支出減と収入は別物',
         '図の形が、そのまま「払い先が替わる」の言い方になる'],
    img_w=5.15)

# ================================================================ 10. P6-② 節約額の内訳
s = new_slide('診断レポート')
page_slide(s, '②', 'P6-②　節約額の内訳が、そのままセット販売の根拠',
    'simulation/sim6_break.jpg', '診断レポート P6　節約額の内訳（左：FIT期間中／右：FIT終了後）',
    ['蓄電池を入れると、こんなに変わるんですね'],
    [('太陽光の行と、蓄電池の行を指で2回指す', {'bold': True}),
     '「太陽光だけだと、この行だけです」',
     ('「蓄電池を足すと、ここまで増えます」', {'bold': True, 'color': RED})],
    '太陽光だけでは届かない。セットで初めてこの数字になります',
    '「電気料金プラン変更」「オール電化」が0なのは、提案に入っていないという意味',
    adv=['P4の「ピークは19〜20時」と結ぶ＝発電しない時間に要る',
         'セット販売の根拠が数字で載っている唯一の場所'],
    img_w=5.15)

# ================================================================ 11. P6-③ FIT後
s = new_slide('診断レポート')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('P6-③　FIT後に減るのは、売電収入だけ', {'size': 23, 'bold': True})])
box, tf = tb(s, 0.45, 1.86, 9.95, 0.32)
lines(tf, [('「10年後はどうなるの？」はほぼ必ず聞かれる。答えは同じページの右側にある',
            {'size': 14, 'color': GRAY})])
rows = [
    ['節約額の内訳', 'FIT期間中', 'FIT期間終了後', ''],
    ['電気料金プラン変更', '0 円', '0 円', ''],
    [('太陽光（自家消費）', {'bold': True}), ('6,612 円', {'bold': True}),
     ('6,612 円', {'bold': True}), ('変わらない', {'color': GRAY})],
    ['オール電化', '0 円', '0 円', ''],
    [('蓄電池', {'bold': True}), ('4,634 円', {'bold': True}),
     ('4,634 円', {'bold': True}), ('変わらない', {'color': GRAY})],
    [('節約額合計', {'bold': True, 'fill': LTGREEN}), ('11,246 円', {'bold': True, 'fill': LTGREEN}),
     ('11,246 円', {'bold': True, 'fill': LTGREEN}),
     ('1円も変わらない', {'bold': True, 'fill': LTGREEN, 'color': RED})],
    [('＋ 売電収入', {'fill': 'FCE4E4'}), ('1,706 円', {'fill': 'FCE4E4'}),
     ('590 円', {'bold': True, 'fill': 'FCE4E4', 'color': RED}),
     ('ここだけ下がる', {'bold': True, 'fill': 'FCE4E4', 'color': RED})],
    [('＝ 実質削減額', {'bold': True}), ('12,952 円', {'bold': True}),
     ('11,836 円', {'bold': True}), ('差は 1,116 円', {'color': GRAY})],
]
table(s, 0.45, 2.26, 9.95, [2.80, 2.20, 2.30, 2.65], rows, font_size=13, row_h=0.46, head_h=0.34)
_, tf = rect(s, 0.45, 5.90, 4.80, 1.00, fill=LTGRAY, line='A6A6A6')
p = para(tf, first=True, space_after=5)
run(p, '基本', size=12, bold=True, color=WHITE, hl='4F6228')
q = para(tf, space_after=0)
run(q, '「減るのは売電収入だけ。自分の分は変わりません」', size=13.5, bold=True)
_, tf = rect(s, 5.60, 5.90, 4.80, 1.00, fill=LTYEL, line='BF8F00')
p = para(tf, first=True, space_after=5)
run(p, '応用', size=12, bold=True, color=WHITE, hl='C55A11')
q = para(tf, space_after=0)
run(q, '買取単価が24円→8.3円に落ちるため、と理由まで', size=13.5, bold=True)
_, tf = rect(s, 0.45, 7.02, 9.95, 0.20, fill=None)

# ================================================================ 12. P7 長期
s = new_slide('診断レポート')
page_slide(s, '②', 'P7　25年でいくら減るか。前提も同じページに載っている',
    'simulation/sim7_sum.jpg', '診断レポート P7　長期シミュレーション（25年間の実質削減額）',
    ['25年だと、そんなに違うんですね'],
    [('大きいピンクの数字を1つだけ指す', {'bold': True}),
     '「25年間で、累計 455万円の差になります」',
     ('先に「電気代が年2％上がる前提です」と言ってから見せる', {'bold': True, 'color': RED})],
    '25年間の実質削減額は 累計 4,553,031 円',
    'グラフは40年まで目盛りがありますが、計算は25年。聞かれたら素直に答える',
    adv=['「上がらなかったら？」には右下を指す。0％でも360万円',
         '実質光熱費の定義も、同じページに書いてある'],
    img_w=5.15)

# ================================================================ 13. 伝え方②（基本）
s = new_slide('金額の伝え方')
box, tf = tb(s, 0.45, 1.36, 7.80, 0.45)
lines(tf, [('伝え方②　電気代の減収分のみをお伝えする', {'size': 23, 'bold': True})])
_, tf = rect(s, 8.45, 1.38, 1.95, 0.42, fill=LTGRAY, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '基本', size=12, bold=True, color=WHITE, hl='4F6228')
run(p, '　全員これ', size=12, bold=True, color=GRAY)
box, tf = tb(s, 0.45, 1.88, 9.95, 0.32)
lines(tf, [('ローンの話をいったん外して、「電気代がいくら減るか」だけを見せます',
            {'size': 14, 'color': GRAY})])
rows = [
    ['見る単位', '言い方', '金額'],
    ['月で見ると', '「毎月の電気代が、これだけ下がります」',
     ('15,089円 → 3,844円', {'bold': True})],
    ['', '「減る額は、毎月これだけです」', ('12,952 円/月', {'bold': True, 'color': RED})],
    ['年で見ると', '「1年にすると、これだけ残ります」', ('約 15.5 万円/年', {'bold': True})],
    [('25年で見ると', {'bold': True, 'fill': LTYEL}),
     ('「25年だと、累計でこれだけです」', {'bold': True, 'fill': LTYEL}),
     ('4,553,031 円', {'bold': True, 'fill': LTYEL, 'color': RED})],
]
table(s, 0.45, 2.28, 9.95, [2.00, 5.00, 2.95], rows, font_size=13.5, row_h=0.56, head_h=0.34)
_, tf = rect(s, 0.45, 5.12, 4.80, 1.02, fill=LTGRAY, line='A6A6A6')
p = para(tf, first=True, space_after=5)
run(p, '基本', size=12, bold=True, color=WHITE, hl='4F6228')
q = para(tf, space_after=0)
run(q, 'ここでローンの話をしない。「いくら払うか」は次', size=13.5, bold=True)
_, tf = rect(s, 5.60, 5.12, 4.80, 1.02, fill=LTYEL, line='BF8F00')
p = para(tf, first=True, space_after=5)
run(p, '応用', size=12, bold=True, color=WHITE, hl='C55A11')
q = para(tf, space_after=0)
run(q, '25年の数字はP7を開いて見せる（0％も同じページ）', size=13.5, bold=True)
_, tf = rect(s, 0.45, 6.34, 9.95, 0.58, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '削減額は「大きい単位」で言う。', size=13, bold=True, color=GRAY)
run(p, '月 → 年 → 25年 と積み上げて、得の大きさを作りきる', size=16, bold=True, hl=YELLOW)
notes(s, '・出典：創蓄セット販売 強化合宿 p.18')

# ================================================================ 14. 伝え方①（切り替え）
s = new_slide('金額の伝え方')
box, tf = tb(s, 0.45, 1.36, 7.80, 0.45)
lines(tf, [('伝え方①　ローン返済額を含めたトータル金額', {'size': 23, 'bold': True})])
_, tf = rect(s, 8.45, 1.38, 1.95, 0.42, fill=LTYEL, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '応用', size=12, bold=True, color=WHITE, hl='C55A11')
run(p, '　切り替え用', size=12, bold=True, color=GRAY)
box, tf = tb(s, 0.45, 1.88, 9.95, 0.32)
lines(tf, [('②で「で、いくら払うの？」と聞かれたときだけ、こちらに切り替えます',
            {'size': 14, 'color': GRAY})])
rows = [
    ['（例）', 'ローン支払い中', '完済後'],
    ['もともとの電気代', '15,000 円', '15,000 円'],
    ['＋ ローンの分割額', ('22,000 円', {'color': RED}), ('0 円', {'bold': True, 'color': RED})],
    ['− 電気代削減額', '−14,000 円', '−14,000 円'],
    [('支出合計', {'bold': True, 'fill': LTYEL}), ('23,000 円', {'bold': True, 'fill': LTYEL}),
     ('1,000 円', {'bold': True, 'fill': LTYEL, 'color': RED})],
]
table(s, 0.45, 2.28, 6.10, [2.50, 1.90, 1.70], rows, font_size=13, row_h=0.50, head_h=0.32)
_, tf = rect(s, 6.90, 2.28, 3.50, 1.22, fill=LTGRAY)
lines(tf, [
    ('向いているお客様', {'size': 14, 'bold': True, 'space': 6}),
    ('・長期で考えられる方', {'size': 13, 'space': 4}),
    ('・完済後の姿に価値を感じる方', {'size': 13}),
])
_, tf = rect(s, 6.90, 3.64, 3.50, 1.18, fill='FCE4E4', line=RED)
lines(tf, [
    ('注意', {'size': 14, 'bold': True, 'color': RED, 'space': 6}),
    ('返済中は支出が増えます。', {'size': 13, 'space': 4}),
    ('ここを隠すと必ず不信になります。', {'size': 13, 'bold': True}),
])
_, tf = rect(s, 0.45, 4.92, 9.95, 1.22, fill=LTBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
lines(tf, [
    ('使い分けの原則', {'size': 15, 'bold': True, 'color': BLUE, 'space': 7}),
    ('①　まず伝え方②（電気代の減収分のみ）で話す', {'size': 14, 'space': 4}),
    ('②　「で、いくら払うの？」と聞かれたら、伝え方①または P8 に進む', {'size': 14, 'space': 4}),
    ('どちらも同じ事実を、別の枠組みで見せているだけです。', {'size': 14, 'bold': True}),
])
_, tf = rect(s, 0.45, 6.34, 9.95, 0.58, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '完済後に支出が「1,000円」まで下がる絵を見せる', size=17, bold=True, hl=YELLOW)
notes(s, '・出典：創蓄セット販売 強化合宿 p.17')

# ================================================================ 15. P8 ＋ 分割して伝える
s = new_slide('金額の伝え方')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('P8　金額は「下から」伝える。総額は最後', {'size': 23, 'bold': True})])
pic(s, os.path.join(ASSETS, 'simulation/sim8_calc.jpg'), 0.45, 1.90, 5.20)
box, tf = tb(s, 0.45, 3.78, 5.20, 0.28)
lines(tf, [('診断レポート P8　毎月の実質負担額', {'size': 11, 'color': GRAY})])
_, tf = rect(s, 0.45, 4.12, 5.20, 1.05, fill=LTGRAY)
lines(tf, [
    ('この式の読み方', {'size': 13, 'bold': True, 'color': GRAY, 'space': 5}),
    ('P6で言った「実質削減額」が、そのままここに入っています。', {'size': 13, 'space': 3}),
    ('「さっきのこの金額が、ここです」と指でつなぐ。', {'size': 13, 'bold': True}),
])
steps = [
    ('①', '1日あたりの金額にして表す', '約345円 ≒ コーヒー一杯分',
     '「1日500円くらいコンビニで買われてますよね？」（YES取り）', LTBLUE),
    ('②', '1月あたりの金額にして表す', '10,364円 ≒ いまの電気代より安い',
     '「奥さんも、飲み会2回くらい削ったっていいですよね？」（YES取り）／「皆さん停電保険って」', LTGREEN),
    ('③', '総支払額を伝える', '聞かれてから。自分から出さない',
     '合宿資料「総支払額は最後に伝えるのがベター」', LTYEL),
]
y = 1.90
for n, t, amt, talk, fill in steps:
    _, tf = rect(s, 5.95, y, 0.52, 1.08, fill=fill, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [(n, {'size': 20, 'bold': True, 'align': PP_ALIGN.CENTER})])
    _, tf = rect(s, 6.58, y, 3.82, 1.08, fill=fill, anchor=MSO_ANCHOR.MIDDLE)
    p = para(tf, first=True, space_after=3)
    run(p, t, size=14, bold=True)
    p = para(tf, space_after=3)
    run(p, amt, size=12.5, bold=True, color=RED)
    p = para(tf, space_after=0)
    run(p, talk, size=11, color=GRAY)
    y += 1.17
_, tf = rect(s, 5.95, 5.44, 4.45, 0.74, fill=None, line=RED, lw=1.5, anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [
    ('下から積む。①→②→③の順', {'size': 14, 'bold': True, 'color': RED,
                      'align': PP_ALIGN.CENTER, 'space': 3}),
    ('フレーミング効果（合宿 p.46・p.47）', {'size': 12, 'color': GRAY,
                             'align': PP_ALIGN.CENTER}),
])
_, tf = rect(s, 0.45, 6.34, 9.95, 0.58, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '支払額は「小さい単位」で言う。', size=13, bold=True, color=GRAY)
run(p, '1日あたりから話し、販売金額は最後まで自分から出さない', size=16, bold=True, hl=YELLOW)
notes(s, '・P8の上段（販売金額・ローン条件）は、聞かれるまで見せない')

# ================================================================ 16. 75.4％ → 安心感
s = new_slide('疑いに答える')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('前提：シミュレーションは疑われています', {'size': 23, 'bold': True, 'color': RED})])
pic(s, os.path.join(ASSETS, 'simulation/trust754.jpg'), 0.45, 1.90, 5.35)
box, tf = tb(s, 0.45, 5.42, 5.35, 0.28)
lines(tf, [('出典：エネがえる運営事務局調べ（国際航業株式会社）', {'size': 10, 'color': GRAY})])
_, tf = rect(s, 6.05, 1.90, 4.35, 0.92, fill='FCE4E4', line=RED, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '疑ったことがある人　', size=14, bold=True)
run(p, '75.4', size=28, bold=True, color=RED)
run(p, ' ％', size=16, bold=True, color=RED)
_, tf = rect(s, 6.05, 2.98, 4.35, 2.12, fill=LTGRAY)
lines(tf, [
    ('だから、こう伝えます', {'size': 14, 'bold': True, 'space': 7}),
    ('・金額の大きさで勝とうとしない', {'size': 13, 'space': 6}),
    ('・前提条件を先に口で言う', {'size': 13, 'space': 2}),
    ('　（使用量・単価・上昇率）', {'size': 12, 'color': GRAY, 'space': 6}),
    ('・お客様の検針票の実額から出発する', {'size': 13, 'space': 6}),
    ('・診断日時・世帯IDを指す（専用の診断）', {'size': 13}),
])
_, tf = rect(s, 6.05, 5.26, 4.35, 0.92, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '✕　シミュレーション金額', size=16, bold=True, color=GRAY)
p = para(tf, align=PP_ALIGN.CENTER)
run(p, '◎　安心感', size=22, bold=True, color=RED)
_, tf = rect(s, 0.45, 6.34, 9.95, 0.58, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '「盛っていない」ことが伝わる方が、金額より効きます', size=17, bold=True, hl=YELLOW)
notes(s, '・出典：創蓄セット販売 強化合宿 p.25・p.26')

# ================================================================ 17. 即答集
s = new_slide('疑いに答える')
box, tf = tb(s, 0.45, 1.36, 7.80, 0.45)
lines(tf, [('聞かれたら、ページで答える', {'size': 24, 'bold': True})])
_, tf = rect(s, 8.25, 1.38, 2.15, 0.42, fill=LTGRAY, anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [('ロープレ中の参照用', {'size': 12, 'bold': True, 'color': GRAY,
                        'align': PP_ALIGN.CENTER})])
box, tf = tb(s, 0.45, 1.86, 9.95, 0.32)
lines(tf, [('答えは全部レポートの中にあります。言葉で説得せず、ページを開いて指す',
            {'size': 14, 'color': GRAY})])
rows = [
    ['お客様の質問', '答え方', '開くページ'],
    ['10年後（FIT後）はどうなるの？',
     ('「減るのは売電収入だけ。自分で使う分は変わりません」', {'bold': True}),
     ('P6 右側', {'bold': True})],
    ['電気代が上がらなかったら？',
     ('「上がらない前提の数字も、同じページに出ています」', {'bold': True}),
     ('P7 右下', {'bold': True})],
    ['本当にこんなに減るの？',
     '「○○様の検針票から計算。使う量は変えていません」',
     ('P3 と P6', {'bold': True})],
    ['蓄電池まで要る？',
     '「太陽光だけだと、この行だけになります」', ('P6 内訳', {'bold': True})],
    ['もっと大きい蓄電池の方が得では？',
     '「サイズアップ分のメリットは30年で約3.6万円です」',
     ('−', {'color': GRAY})],
    ['結局いくら払うの？',
     ('「1日あたり約345円です」（総額から答えない）', {'bold': True, 'color': RED}),
     ('P8 を下から', {'bold': True})],
]
table(s, 0.45, 2.28, 9.95, [3.05, 5.20, 1.70], rows, font_size=12.5, row_h=0.62, head_h=0.32)
_, tf = rect(s, 0.45, 6.26, 9.95, 0.66, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('いちばん多い失敗：「結局いくら？」に総額で答えてしまう', {'size': 15, 'bold': True,
                                       'color': RED, 'space': 4}),
    ('1日あたり → 1月あたり → 総額の順。総額は自分から出さない。', {'size': 14}),
])

# ================================================================ 18. まとめ
s = new_slide('まとめ')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.50)
lines(tf, [('このパートで持ち帰ってほしい3つ', {'size': 25, 'bold': True})])
items = [
    ('1', 'まず「いくら減るか」を言い切る',
     'ローンの話は、そのあと。自分から持ち出さない（伝え方②）。'),
    ('2', '削減額は大きい単位、支払額は小さい単位',
     '25年で455万円。払うのは1日345円。総額は聞かれてから。'),
    ('3', '金額で勝とうとしない',
     '✕ シミュレーション金額　◎ 安心感。前提を先に言い、検針票の実額から出発する。'),
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
_, tf = rect(s, 0.45, 6.22, 9.95, 0.70, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('おまけ：金額の褒め → 体験の褒め', {'size': 14, 'bold': True, 'space': 3}),
    ('「将来的には手出し0円で災害対策もできちゃいますね」「浮いたお金で何ができますか、奥さん？」',
     {'size': 14, 'bold': True}),
])
notes(s, '・「元が取れる／取れない」で判断させると、狭小住宅への提案が困難になる（合宿 p.36）')

# ================================================================ 19. ロープレ
s = new_slide('ロールプレイング')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('通し練習：p.20 から P8 まで', {'size': 24, 'bold': True})])
_, tf = rect(s, 0.45, 1.90, 9.95, 1.18, fill=LTGRAY)
lines(tf, [('進め方（合計30分）', {'size': 14, 'bold': True, 'color': GRAY, 'space': 6})])
for it in [
    '2人1組。営業役とお客様役。1本8分 → フィードバック4分 → 交代してもう1本',
    'p.20 を開いた状態から始め、P8 の「1日あたり」まで到達する',
    '基本｜伝え方②で通す　応用｜お客様役が疑いを1つ入れる（即答集から選ぶ）',
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
    ('・ローンの話を自分から出していないか', {'size': 12.5, 'space': 3}),
    ('・総額（販売金額）を口にしていないか', {'size': 12.5, 'space': 3}),
    ('・1日あたりまで落とせたか', {'size': 12.5, 'space': 3}),
    ('・前提条件を先に言えたか', {'size': 12.5, 'space': 3}),
    ('・太陽光だけ／蓄電池だけの話になっていないか', {'size': 12.5}),
])
_, tf = rect(s, 0.45, 5.06, 9.95, 0.96, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('フィードバックの型', {'size': 15, 'bold': True, 'space': 5}),
    ('① 良かったページを1つ、ページ番号で言う　② 直すのは1人1つだけ　③ 人ではなくページを指摘する',
     {'size': 14, 'bold': True}),
])
_, tf = rect(s, 0.45, 6.18, 9.95, 0.72, fill=None, line=RED, lw=1.5,
             anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [('「結局いくらですか？」を1回は必ず投げてください。総額で答えたらそこで止めます',
            {'size': 16, 'bold': True, 'align': PP_ALIGN.CENTER})])

# ================================================================ 20. チェックシート
s = new_slide('ロールプレイング')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('セルフチェックシート', {'size': 24, 'bold': True})])
box, tf = tb(s, 0.45, 1.86, 9.95, 0.32)
lines(tf, [('ロープレのあとに記入し、できていない項目を次回までの宿題にしてください',
            {'size': 14, 'color': GRAY})])
rows = [
    ['確認すること', '今日', '次回'],
    ['p.21で「朝・昼・夕」の流れを図なしで言えた', ' ', ' '],
    ['P6で太陽光の行と蓄電池の行を指して説明できた', ' ', ' '],
    ['「FIT後は？」に、売電収入だけと即答できた', ' ', ' '],
    ['前提条件（上昇率2％）を自分から先に言えた', ' ', ' '],
    ['削減額を25年の累計まで積み上げられた', ' ', ' '],
    ['1日あたりの金額まで落とせた', ' ', ' '],
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
run(q, '7項目すべて＋疑いへの即答が出る', size=12.5, bold=True)
_, tf = rect(s, 7.05, 5.66, 3.35, 1.26, fill=LTGRAY)
lines(tf, [
    ('次回までの宿題', {'size': 14, 'bold': True, 'space': 5}),
    ('担当顧客1件のシミュレーションで、', {'size': 12, 'space': 2}),
    ('残す数字3つを書き出してくること。', {'size': 12, 'space': 5}),
    ('翌日には50％忘れます。', {'size': 12, 'bold': True, 'color': RED}),
])

# ================================================================ 21. 参考：相見積もり
s = new_slide('参考')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('参考：相見積もりになったときのポイント', {'size': 23, 'bold': True})])
_, tf = rect(s, 0.45, 1.92, 9.95, 0.86, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '信頼　×　比較', size=30, bold=True)
rows = [
    ['', 'やること', '言い方'],
    [('信頼', {'bold': True}), '前提条件を隠さない。数字の出どころを言う',
     ('「この数字は○○様の検針票から計算しています」', {'bold': True})],
    [('比較', {'bold': True}), '他社下げをせず、お客様に判断軸を与える',
     ('「設置後の点検体制で比べてみてください」', {'bold': True})],
    [('−', {'color': GRAY}), '他社の強みで戦わず、自社の強みで戦う',
     '会社の安定性／地域密着性／メーカーとしての価値'],
]
table(s, 0.45, 2.94, 9.95, [1.20, 4.25, 4.50], rows, font_size=13, row_h=0.68, head_h=0.32)
_, tf = rect(s, 0.45, 5.30, 9.95, 1.00, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('金額ばかり褒めてはいけない理由も同じです', {'size': 15, 'bold': True, 'space': 4}),
    ('「元が取れる／取れない」で判断させると、狭小住宅・北向き屋根への提案が困難になります。',
     {'size': 14}),
])
_, tf = rect(s, 0.45, 6.44, 9.95, 0.48, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '金額で勝とうとせず、安心感で選ばれる', size=17, bold=True, hl=YELLOW)
notes(s, '・出典：創蓄セット販売 強化合宿 p.24・p.36／第3回 27枚目')

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
