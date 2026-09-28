# -*- coding: utf-8 -*-
"""必要性訴求パート（アプローチブックの伝え方）研修資料ビルダー

第2回・第3回に分かれていた必要性訴求の内容を1本にまとめ、
営業テクニックではなく「アプローチブック p.9〜p.21 を、どうめくり、
何を言うか」を軸に組み直したもの。

使い方は slides/build/build_session2.py と同じ。以下の画像が必要。

  <ASSETS>/ab09.jpg   アプローチブック p.9   電気料金の推移
  <ASSETS>/ab14.jpg   同 p.14  再エネ賦課金の推移
  <ASSETS>/ab17.jpg   同 p.17  生涯に支払う電気代
  <ASSETS>/ab18.jpg   同 p.18  予測が困難な地震情報
  <ASSETS>/ab19.jpg   同 p.19  自然災害と停電被害
  <ASSETS>/ab21.jpg   同 p.21  太陽光発電・蓄電池の使い方
  <ASSETS>/sim_trust.jpg  シミュレーションの信憑性（75.4％の円グラフ）

    TRAINING_ASSETS=/path/to/assets python slides/build/build_need.py

内容の出典：docs/知識_アプローチブック.md ／ docs/知識_営業ルール20.md ／
docs/知識_数値ファクト集.md ＋「20のポイント解説」
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
ASSETS = os.environ.get('TRAINING_ASSETS', '/tmp/training-assets')
AB     = os.path.join(ASSETS, 'ab')
TPL    = os.path.join(REPO, 'templates', '研修資料_フォーマット見本.pptx')
OUT    = os.path.join(REPO, 'slides', '2026営業研修_必要性訴求パート_アプローチブックの伝え方.pptx')

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
            'val="C0504D"><a:alpha val="80000"/></a:srgbClr>'))
    if shp.has_text_frame and shp.name == 'テキスト ボックス 5':
        tf = shp.text_frame; tf.clear()
        lines(tf, [
            ('%s　御中' % CLIENT, {'size': 26, 'bold': True, 'color': WHITE, 'font': UD, 'space': 10}),
            ('必要性訴求パート', {'size': 48, 'bold': True, 'color': WHITE, 'font': UD, 'space': 8}),
            ('アプローチブックの伝え方', {'size': 32, 'bold': True, 'color': WHITE, 'font': UD, 'space': 10}),
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


# ================================================================ 2. このパートのゴール
s = new_slide('このパートのゴール')
box, tf = tb(s, 0.45, 1.42, 9.9, 0.60)
lines(tf, [('アプローチブック p.9〜p.21 を、自分の言葉でめくれるようになる',
            {'size': 21, 'bold': True, 'color': RED})])
_, tf = rect(s, 0.45, 2.25, 9.95, 2.95, fill=LTGREEN)
lines(tf, [
    ('この1時間が終わったときの「できる状態」', {'size': 18, 'bold': True, 'space': 10}),
    ('①　各ページで「何を言うか」が、資料を見なくても口から出る', {'size': 19, 'space': 9}),
    ('②　各ページの「結論の一文」を、そのまま言い切れる', {'size': 19, 'space': 9}),
    ('③　お客様から引き出したい反応を、質問で取りにいける', {'size': 19, 'space': 9}),
    ('④　p.9からp.21まで、15分で止まらずに通せる', {'size': 19}),
])
_, tf = rect(s, 0.45, 5.45, 9.95, 1.05, fill=None, line=RED, lw=1.5,
             anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [
    ('このパートで扱うもの', {'size': 14, 'bold': True, 'color': GRAY,
                    'align': PP_ALIGN.CENTER, 'space': 3}),
    ('アプローチブック p.9〜p.21　＋　シミュレーション', {'size': 21, 'bold': True,
                                    'align': PP_ALIGN.CENTER}),
])
notes(s, '・営業テクニックを覚える会ではなく、アプローチブックをめくれるようになる会だと宣言する\n'
         '・「資料に書いてあることを読む」のではなく「資料を使って言わせる」のが目的')

# ================================================================ 3. 必要性訴求とは
s = new_slide('必要性訴求とは')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('必要性を訴求する3つの切り口', {'size': 26, 'bold': True})])
cuts = [
    (1.95, '①', '電気代が高騰している ⇒ 払わなくていい', 'p.9〜p.17（7ページ）', LTGREEN, '約20分'),
    (3.32, '②', '災害対策', 'p.18・p.19（2ページ）', LTYEL, '3分以内'),
    (4.69, '③', '経済メリット', 'p.20・p.21 ＋ シミュレーション', LTBLUE, '約15分'),
]
for y, n, t, pages, fill, mins in cuts:
    _, tf = rect(s, 0.45, y, 0.62, 1.20, fill=fill, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [(n, {'size': 26, 'bold': True, 'align': PP_ALIGN.CENTER})])
    _, tf = rect(s, 1.20, y, 7.15, 1.20, fill=fill, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [
        (t, {'size': 20, 'bold': True, 'space': 5}),
        ('使うページ：%s' % pages, {'size': 14, 'color': GRAY}),
    ])
    _, tf = rect(s, 8.50, y + 0.30, 1.90, 0.60, fill=GRAY, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [(mins, {'size': 15, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER})])
_, tf = rect(s, 0.45, 6.10, 9.95, 0.80, fill=LTGRAY)
lines(tf, [
    ('このパートの基本設計は「下げてから上げる」', {'size': 18, 'bold': True, 'color': RED, 'space': 4}),
    ('①②でネガティブな現状を自分事にしてもらい、③で解決策を出します。順番は入れ替えません。',
     {'size': 15}),
])

# ================================================================ 4. ページの地図
s = new_slide('アプローチブックの地図')
box, tf = tb(s, 0.45, 1.32, 9.95, 0.38)
lines(tf, [('p.9〜p.21で、どのページで何を言うか', {'size': 22, 'bold': True})])
rows = [
    ['切り口', 'ページ', '見せるもの', '引き出す反応'],
    [('①電気代', {'fill': LTGREEN}), ('p.9', {'fill': LTGREEN, 'bold': True}),
     ('電気料金の推移（2010〜2025年）', {'fill': LTGREEN, 'bold': True}),
     ('10年で5.5万円も上がっているのですね', {'fill': LTGREEN})],
    [('', {'fill': LTGREEN}), ('p.10・p.12', {'fill': LTGREEN}),
     ('自由化しても下がらない／全国的な値上げ', {'fill': LTGREEN}),
     ('全国的に上がっているのですね', {'fill': LTGREEN})],
    [('', {'fill': LTGREEN}), ('p.11・p.13', {'fill': LTGREEN}),
     ('電気代の内訳と再エネ賦課金とは', {'fill': LTGREEN}),
     ('こんなものを払っていたのですね', {'fill': LTGREEN})],
    [('', {'fill': LTGREEN}), ('p.14', {'fill': LTGREEN, 'bold': True}),
     ('再エネ賦課金の推移（13年で18倍）', {'fill': LTGREEN, 'bold': True}),
     ('賦課金で電気代が上がるのですね', {'fill': LTGREEN})],
    [('', {'fill': LTGREEN}), ('p.15・p.16', {'fill': LTGREEN}), ('燃料調整費の上昇', {'fill': LTGREEN}),
     ('これからも上がりそうですね', {'fill': LTGREEN})],
    [('', {'fill': LTGREEN}), ('p.17', {'fill': LTGREEN, 'bold': True}),
     ('生涯に支払う電気代（早見表）', {'fill': LTGREEN, 'bold': True}),
     ('30年でそんなにも払うのですね', {'fill': LTGREEN})],
    [('②災害', {'fill': LTYEL}), ('p.18', {'fill': LTYEL, 'bold': True}),
     ('予測が困難な地震情報', {'fill': LTYEL, 'bold': True}),
     ('いつ災害が起こるか分からないですね', {'fill': LTYEL})],
    [('', {'fill': LTYEL}), ('p.19', {'fill': LTYEL, 'bold': True}),
     ('停電被害の実績', {'fill': LTYEL, 'bold': True}), ('備えておいた方が安心ですね', {'fill': LTYEL})],
    [('', {'fill': LTYEL}), ('（なし）', {'fill': LTYEL, 'color': RED}),
     ('自社・メーカーの事例で語る', {'fill': LTYEL}), ('停電時も電気が使えた方が安心です', {'fill': LTYEL})],
    [('③経済メリット', {'fill': LTBLUE}), ('p.20・p.21', {'fill': LTBLUE, 'bold': True}),
     ('自給自足への転換と、その使い方', {'fill': LTBLUE, 'bold': True}),
     ('自給自足した方が良いですね', {'fill': LTBLUE})],
    [('', {'fill': LTBLUE}), ('（資料の外）', {'fill': LTBLUE}), ('シミュレーション', {'fill': LTBLUE}),
     ('思ったより減るんですね', {'fill': LTBLUE})],
]
table(s, 0.30, 1.85, 10.25, [1.40, 1.15, 3.55, 4.15], rows, font_size=11, row_h=0.38, head_h=0.32)
_, tf = rect(s, 0.30, 6.42, 10.25, 0.48, fill=None, line=RED, lw=1.5, anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [('太字が「必ず時間をかけるページ」。それ以外は流して構いません',
            {'size': 15, 'bold': True, 'align': PP_ALIGN.CENTER})])

# ================================================================ 5. 共通の伝え方の型
s = new_slide('伝え方の型')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('全ページ共通：1ページを「4ステップ」でめくる', {'size': 24, 'bold': True})])
steps = [
    ('STEP 1', '見出しを読む', '約5秒',
     'ページ上部の見出しを、そのまま読み上げる。ここは変えない。'),
    ('STEP 2', 'リード文は読まない', '約15秒',
     '小さい文字のリード文は読み上げず、要点を自分の言葉で1文に縮めて言う。'),
    ('STEP 3', '図の中の数字を1つだけ指す', '約30秒',
     '図の全部を説明しない。指でさして、いちばん大きい数字を1つだけ口に出す。'),
    ('STEP 4', '緑帯の結論を言い切り、質問を返す', '約20秒',
     '下の緑帯をそのまま言い切る。そのあと質問を投げて、お客様に言わせる。'),
]
y = 1.98
for tag, title, sec, desc in steps:
    _, tf = rect(s, 0.45, y, 1.10, 1.05, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [(tag, {'size': 14, 'bold': True, 'align': PP_ALIGN.CENTER})])
    _, tf = rect(s, 1.65, y, 8.75, 1.05, fill=LTGRAY, anchor=MSO_ANCHOR.MIDDLE)
    p = para(tf, first=True, space_after=4)
    run(p, title, size=18, bold=True)
    run(p, '　（%s）' % sec, size=13, color=GRAY)
    p = para(tf, space_after=0)
    run(p, desc, size=14)
    y += 1.17
_, tf = rect(s, 0.45, 6.62, 9.95, 0.42, fill=LTYEL, line='BF8F00', anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [('1ページ約70秒。13ページ通して約15分。これが「めくるだけ」の標準ペースです',
            {'size': 15, 'bold': True, 'align': PP_ALIGN.CENTER})])
notes(s, '・「資料を読み上げる」と「資料を使って言わせる」の違いを、講師が実演して見せる')


# ---------------------------------------------------------------- page helper
def page_slide(s, chapter, title, img, caption, goal, says, conc, caution=None):
    box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
    lines(tf, [(title, {'size': 23, 'bold': True})])
    if img:
        pic(s, os.path.join(ASSETS, img), 0.45, 1.92, 4.85)
        box, tf = tb(s, 0.45, 5.34, 4.85, 0.28)
        lines(tf, [(caption, {'size': 11, 'color': GRAY})])
    _, tf = rect(s, 5.65, 1.92, 4.75, 1.12, fill=LTBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    lines(tf, [('引き出す反応（お客様の台詞）', {'size': 12, 'bold': True, 'color': BLUE, 'space': 5})])
    for q in goal:
        p = para(tf, space_after=3); run(p, '「%s」' % q, size=15, bold=True)
    _, tf = rect(s, 5.65, 3.20, 4.75, 2.40, fill=LTGRAY)
    lines(tf, [('このページで言うこと', {'size': 13, 'bold': True, 'color': GRAY, 'space': 6})])
    for it in says:
        opts = {}
        if isinstance(it, tuple): it, opts = it
        p = para(tf, space_after=opts.pop('space', 6))
        run(p, it, size=opts.pop('size', 13), bold=opts.pop('bold', False),
            color=opts.pop('color', BLACK))
    _, tf = rect(s, 0.45, 5.80, 9.95, 0.58, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
    p = para(tf, first=True, align=PP_ALIGN.CENTER)
    run(p, '緑帯の結論（そのまま言い切る）　', size=12, bold=True, color=GRAY)
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

# ================================================================ 6. ①章扉
s = new_slide('① 電気代')
chapter_slide(s, '①', '電気代が高騰している ⇒ 払わなくていい',
              'p.9 〜 p.17（7ページ）', '約20分', LTGREEN,
              '「このままではまずい」と自分事で思ってもらう',
              ['メインは p.9・p.14・p.17 の3ページ。ここに時間をかける',
               '残りの4ページ（p.10〜13・15・16）は流してよい。根拠の補強に使う',
               'いちばん最初に、お客様の検針票の実額を聞く。そこから全部の数字を置き換える',
               '「全国平均では」で終わらせない。必ず「〇〇様の場合は」に変換する'])

# ================================================================ 7. p.9
s = new_slide('① 電気代')
page_slide(s, '①', 'p.9　直近10年間、過去最高水準の電気料金になっています',
    'ab09.jpg', 'アプローチブック p.9　電気料金の推移',
    ['10年間で5.5万円も電気代は上がっているのですね'],
    [('グラフ全体は説明せず、右端の¥180,675だけを指す', {'bold': True}),
     '「2010年は約12万円でした。それが2025年には18万円です」',
     '「震災を境に、15年で約5.5万円上がりました」',
     ('先に「〇〇様は月いくらですか？」と聞いてから開く', {'bold': True, 'color': RED})],
    '電気料金は2011年以降45％、約5.5万円も上昇している',
    'グラフの棒を1本ずつ読み上げると、お客様は数字を追えなくなります')

# ================================================================ 8. p.10・p.12
s = new_slide('① 電気代')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('p.10・p.12　「自分の地域だけではない」を作る2ページ', {'size': 23, 'bold': True})])
box, tf = tb(s, 0.45, 1.86, 9.95, 0.32)
lines(tf, [('この2ページは流して構いません。p.9の裏づけとして、20秒ずつで抜けます',
            {'size': 14, 'color': GRAY})])
for x, pg, ttl, says, conc in (
    (0.45, 'p.10', '「電力自由化」でも電気代は下がらなかった',
     ['2016年4月に電力小売が自由化され、会社を選べるようになった',
      '単価は 2016年 21円/kWh → 2025年 32円/kWh。11円上がっている',
      '「政府が法改正までしたのに、下がらなかったんです」'],
     '制度のねらいに反して、電気代はむしろ高くなっている'),
    (5.60, 'p.12', '電気料金が上がる理由：全国的な値上げ',
     ['各電力会社の値上げ幅を見せる（北陸42％・九州38％・関西29％…）',
      '東北電力は2023年6月1日から平均25.47％の値上げ',
      '「〇〇様の地域だけの話ではありません」'],
     '新単価を確認すると、10円以上も値上げしている'),
):
    _, tf = rect(s, x, 2.26, 4.80, 0.50, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
    p = para(tf, first=True)
    run(p, pg + '　', size=16, bold=True, color=RED)
    run(p, ttl, size=13, bold=True)
    _, tf = rect(s, x, 2.86, 4.80, 2.35, fill=LTGRAY)
    lines(tf, [('このページで言うこと', {'size': 12, 'bold': True, 'color': GRAY, 'space': 6})])
    for it in says:
        p = para(tf, space_after=6); run(p, it, size=13)
    _, tf = rect(s, x, 5.32, 4.80, 0.72, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
    p = para(tf, first=True, align=PP_ALIGN.CENTER)
    run(p, conc, size=13, bold=True, hl=YELLOW)
_, tf = rect(s, 0.45, 6.14, 9.95, 0.76, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('引き出す反応', {'size': 12, 'bold': True, 'color': GRAY, 'space': 3}),
    ('「法改正をしても電気代は上がっているのですね」／「全国的に上昇しているのですね」',
     {'size': 15, 'bold': True}),
])

# ================================================================ 9. p.11・p.13
s = new_slide('① 電気代')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('p.11・p.13　「なぜ上がるのか」の答えを渡す2ページ', {'size': 23, 'bold': True})])
_, tf = rect(s, 0.45, 1.92, 9.95, 1.00, fill=LTBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
             anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [
    ('p.11の計算式（この式は暗記してください）', {'size': 13, 'bold': True, 'color': BLUE,
                              'align': PP_ALIGN.CENTER, 'space': 4}),
    ('電気料金 ＝ 基本料金 ＋ 電力量料金（単価 × 使用量 ± 燃料費調整）＋ 再エネ賦課金',
     {'size': 17, 'bold': True, 'align': PP_ALIGN.CENTER}),
])
for x, pg, says, conc in (
    (0.45, 'p.11　高騰の原因は2つの料金単価の上昇',
     ['3つの内訳を、上から順に指でなぞる',
      '「量を減らしても、単価が上がれば電気代は上がる」',
      '「請求額は見ても、内訳まで見る方はいません」'],
     '毎月の支払料金は見ても、内訳まで見る人はいません'),
    (5.60, 'p.13　再エネ賦課金とは何か',
     ['再エネで発電した電気を電力会社が買い取るため、消費者から集めるお金',
      '国が定める割高な価格での買い取りが義務づけられ、その原資を負担している',
      '「いま4.18円/kWh。使った分だけ自動で増えます」'],
     '全国民が電気代とともに再エネ賦課金を負担しています'),
):
    _, tf = rect(s, x, 3.08, 4.80, 0.50, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [(pg, {'size': 13, 'bold': True})])
    _, tf = rect(s, x, 3.68, 4.80, 1.95, fill=LTGRAY)
    for i, it in enumerate(says):
        p = para(tf, first=(i == 0), space_after=6); run(p, it, size=13)
    _, tf = rect(s, x, 5.74, 4.80, 0.72, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
    p = para(tf, first=True, align=PP_ALIGN.CENTER)
    run(p, conc, size=13, bold=True, hl=YELLOW)
_, tf = rect(s, 0.45, 6.62, 9.95, 0.42, fill=None, line=RED, lw=1.5, anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [('ここで「賦課金」という言葉を初めて出す。次のp.14がこのパートの山場です',
            {'size': 15, 'bold': True, 'align': PP_ALIGN.CENTER})])

# ================================================================ 10. p.14
s = new_slide('① 電気代')
page_slide(s, '①', 'p.14　再エネ賦課金は値上がりし続けています',
    'ab14.jpg', 'アプローチブック p.14　再エネ賦課金の推移',
    ['再エネ賦課金が上昇することで、電気代が上昇するのですね'],
    [('左端と右端の2本だけを指す。途中の棒は説明しない', {'bold': True}),
     '「2012年は年1,210円。2026年は22,990円です」',
     '「13年で18倍。使った分だけ自動で増えます」',
     ('言い切ったら、必ず緑帯の一文で裏返す', {'bold': True, 'color': RED})],
    '裏返せば、太陽光を導入している方は続々とお得になっています！',
    '2023年だけ下がって見えます。聞かれたら「燃料費高騰による一時的な措置」と即答する')

# ================================================================ 11. p.15・p.16
s = new_slide('① 電気代')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('p.15・p.16　「これからも上がる」を作る2ページ', {'size': 23, 'bold': True})])
box, tf = tb(s, 0.45, 1.86, 9.95, 0.32)
lines(tf, [('ここは「過去」ではなく「これから」の話。時間は合わせて1分で抜けます',
            {'size': 14, 'color': GRAY})])
flow = ['中東情勢の悪化', 'ホルムズ海峡封鎖', '原油・LNGが高騰',
        '3か月平均で算定', '2か月後に反映']
for i, t in enumerate(flow):
    x = 0.45 + i * 2.00
    _, tf = rect(s, x, 2.28, 1.80, 0.72, fill=LTYEL, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [(t, {'size': 12, 'bold': True, 'align': PP_ALIGN.CENTER})])
    if i < 4:
        a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x + 1.84), Inches(2.52),
                               Inches(0.12), Inches(0.24))
        a.fill.solid(); a.fill.fore_color.rgb = RGBColor.from_string('808080')
        a.line.fill.background(); a.shadow.inherit = False
_, tf = rect(s, 0.45, 3.22, 4.80, 2.10, fill=LTGRAY)
lines(tf, [
    ('p.15で言うこと', {'size': 14, 'bold': True, 'color': GRAY, 'space': 7}),
    ('「燃料の値段は、3か月平均で計算して2か月後の請求に乗ります」', {'size': 13, 'space': 6}),
    ('「つまり、今の値上がりはこれから請求に出てきます」', {'size': 13, 'space': 6}),
    ('政府の電気・ガス料金支援は2026年3月使用分で終了', {'size': 13, 'bold': True, 'color': RED}),
])
_, tf = rect(s, 5.60, 3.22, 4.80, 2.10, fill=LTGRAY)
lines(tf, [
    ('p.16で言うこと', {'size': 14, 'bold': True, 'color': GRAY, 'space': 7}),
    ('東北電力の燃料調整単価（2026年）を見せる', {'size': 13, 'space': 6}),
    ('2月 −13.26円 → 6月 −8.36円', {'size': 13, 'space': 6}),
    ('「マイナス調整が5.1円減った＝1kWhあたり5.1円上がったということです」',
     {'size': 13, 'bold': True}),
])
_, tf = rect(s, 0.45, 5.45, 9.95, 0.62, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '緑帯の結論　', size=12, bold=True, color=GRAY)
run(p, 'ホルムズ海峡封鎖は電気代高騰にも影響を及ぼします', size=15, bold=True, hl=YELLOW)
_, tf = rect(s, 0.45, 6.20, 9.95, 0.72, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('このページの役割', {'size': 13, 'bold': True, 'color': GRAY, 'space': 3}),
    ('p.9〜p.14は「過去の実績」。p.15・16で初めて「これからも上がる」に変わります。',
     {'size': 15, 'bold': True}),
])

# ================================================================ 12. p.17
s = new_slide('① 電気代')
page_slide(s, '①', 'p.17　生涯に支払うことになる電気代は400万円を超えます',
    'ab17.jpg', 'アプローチブック p.17　生涯に支払う電気代',
    ['30年間でそんなにも電気代を払わないといけないのですね'],
    [('表の全部を見せず、お客様の金額の列だけを指でなぞる', {'bold': True}),
     '「〇〇様は月18,000円でしたね。この列です」',
     '「1年21万円、10年216万円、30年648万円です」',
     ('ここまでが①のゴール。金額の解決策はまだ出さない', {'bold': True, 'color': RED})],
    '月々の電気代が2万円なら、30年間の負担額は700万円以上になります',
    '「だから太陽光を」とすぐ言わない。ここは危機感を置いて、一度止まります')

# ================================================================ 13. ①のまとめ
s = new_slide('① 電気代')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('①のまとめ：7ページで言い切ること', {'size': 24, 'bold': True})])
rows = [
    ['ページ', '一言でいうと', '言い切る数字'],
    [('p.9', {'bold': True}), ('過去10年で上がった', {'bold': True}), ('45％・約5.5万円', {'bold': True})],
    ['p.10', '自由化しても下がらなかった', '21円 → 32円'],
    ['p.11', '上がる理由は単価', '（計算式）'],
    ['p.12', '全国どこでも上がっている', '北陸42％・九州38％'],
    ['p.13', '賦課金というものがある', '4.18円/kWh'],
    [('p.14', {'bold': True}), ('その賦課金が18倍になった', {'bold': True}), ('1,210円 → 22,990円', {'bold': True})],
    ['p.15・16', 'これからも上がる', '2月比で5.1円の負担増'],
    [('p.17', {'bold': True}), ('生涯ではこれだけ払う', {'bold': True}), ('30年で700万円以上', {'bold': True})],
]
table(s, 0.45, 1.95, 9.95, [1.35, 4.60, 4.00], rows, font_size=13, row_h=0.46, head_h=0.34)
_, tf = rect(s, 0.45, 6.15, 9.95, 0.80, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('太字の3ページ（p.9・p.14・p.17）だけは、資料を見ずに言えるようにしてください。',
     {'size': 16, 'bold': True, 'space': 4}),
    ('残り4ページは、聞かれたときに開ければ十分です。', {'size': 15}),
])

# ================================================================ 14. ②章扉
s = new_slide('② 災害対策')
chapter_slide(s, '②', '災害対策',
              'p.18・p.19（2ページ）＋ 事例', '3分以内', LTYEL,
              '「いつ起きてもおかしくない」と思ってもらう',
              ['長く話すほど失敗します。商談の雰囲気が暗くなるためです',
               '主語は必ず「事実」と「第三者」。お客様の家族を想像させない',
               '2ページめくって、事例を1本話したら、すぐ次へ抜ける',
               '「うちの地域は災害が来ない」と言われたら、否定せず事実で包み込む'])

# ================================================================ 15. p.18
s = new_slide('② 災害対策')
page_slide(s, '②', 'p.18　政府機関にも予測するのが困難な地震情報',
    'ab18.jpg', 'アプローチブック p.18　予測が困難な地震情報',
    ['いつ災害が起こるか分からないですね'],
    [('表の数字は読まない。赤枠の「札幌市」だけを指す', {'bold': True}),
     '「日本で最も権威のある地震調査委員会のデータです」',
     '「2018年6月の発表で、札幌が全国で最下位でした」',
     ('「ところが3か月後に、胆振東部地震が起きています」', {'bold': True, 'color': RED})],
    '専門家ですら、いつ「もしものこと」が起きるのか分からない状況です',
    '最後は「どこで起きるかは分からないですよね？」と質問で終える。断定で終えない')

# ================================================================ 16. p.19
s = new_slide('② 災害対策')
page_slide(s, '②', 'p.19　自然災害とともに停電被害はたくさん出ている',
    'ab19.jpg', 'アプローチブック p.19　自然災害と停電被害',
    ['備えておいた方が安心ですね'],
    [('家のアイコンの数で「規模」を見せる。件数は1つだけ言う', {'bold': True}),
     '「北海道の地震では約300万戸。日本初のブラックアウト」',
     '「復旧まで2週間かかった地域もありました」',
     ('食料は備蓄できるが、電気だけは備蓄できないと添える', {'bold': True, 'color': RED})],
    '災害大国だからこそ、もしもの時には備える必要があるのです',
    '南海トラフなど「これから」の話を膨らませない。1行触れて終わり')

# ================================================================ 17. ページのない部分
s = new_slide('② 災害対策')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('このあとに「ページのない1本」を挟みます', {'size': 24, 'bold': True})])
_, tf = rect(s, 0.45, 1.92, 4.95, 1.95, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('アプローチブックには', {'size': 16, 'bold': True, 'align': PP_ALIGN.CENTER, 'space': 3}),
    ('「備えるとどうなるか」のページがありません', {'size': 16, 'bold': True, 'color': RED,
                                'align': PP_ALIGN.CENTER, 'space': 10}),
    ('印刷された一般論では「うちの場合は？」に答えられないからです。', {'size': 14}),
])
_, tf = rect(s, 5.60, 1.92, 4.80, 1.95, fill=LTGREEN)
lines(tf, [
    ('だから事例で話します', {'size': 16, 'bold': True, 'space': 8}),
    ('✕「停電しても安心ですよ」', {'size': 14, 'space': 3}),
    ('　← 自分の意見。刺さらない', {'size': 12, 'color': GRAY, 'space': 8}),
    ('○「実際に〇〇市のA様が、台風の停電で3日間、冷蔵庫と照明を動かせたそうです」',
     {'size': 14, 'bold': True}),
])
rows = [
    ['事例の材料', '中身', 'URL'],
    ['自社の施工事例', 'いちばん強い。地域が近いほど効く', '−'],
    ['Panasonic', '動画＋テキストの記事', 'https://sumai.panasonic.jp/chikuden/'],
    ['SmartStar', '災害時の活用に特化した動画', 'https://www.smartstar.jp/voice/'],
]
table(s, 0.45, 4.05, 9.95, [2.05, 3.55, 4.35], rows, font_size=12, row_h=0.48, head_h=0.32)
_, tf = rect(s, 0.45, 6.15, 9.95, 0.78, fill='FCE4E4', line=RED)
lines(tf, [
    ('②全体で3分。時計を見て測ってください。', {'size': 16, 'bold': True, 'color': RED, 'space': 4}),
    ('p.18が1分、p.19が1分、事例が1分。これを超えると、そのあとの話が入らなくなります。', {'size': 14}),
])

# ================================================================ 18. ③章扉
s = new_slide('③ 経済メリット')
chapter_slide(s, '③', '経済メリット',
              'p.20・p.21 ＋ シミュレーション', '約15分', LTBLUE,
              '「自分の家ならいくら減るか」を具体的に持ち帰ってもらう',
              ['①②で作ったネガを、ここで初めて解決策に変えます',
               'p.20・p.21は「仕組みの説明」。ここが曖昧だとシミュレーションが伝わりません',
               'シミュレーションは金額の大きさで勝とうとしない。前提条件を先に言う',
               '見積り（総額）はこのパートでは出しません。それは金額訴求の役割です'])

# ================================================================ 19. p.20・p.21
s = new_slide('③ 経済メリット')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('p.20・p.21　仕組みを説明して、シミュレーションに渡す', {'size': 23, 'bold': True})])
pic(s, os.path.join(ASSETS, 'ab21.jpg'), 0.45, 1.92, 4.85)
box, tf = tb(s, 0.45, 5.34, 4.85, 0.28)
lines(tf, [('アプローチブック p.21　太陽光発電・蓄電池の使い方', {'size': 11, 'color': GRAY})])
_, tf = rect(s, 5.65, 1.92, 4.75, 1.45, fill=LTGRAY)
lines(tf, [
    ('p.20で言うこと', {'size': 13, 'bold': True, 'color': GRAY, 'space': 6}),
    ('「電気代の上昇と災害。この2つへの答えが、自分で作って自分で使うことです」', {'size': 13, 'space': 5}),
    ('結論：電気代を削減するために、自給自足が求められます', {'size': 12, 'bold': True}),
])
_, tf = rect(s, 5.65, 3.52, 4.75, 2.08, fill=LTGRAY)
lines(tf, [
    ('p.21で言うこと', {'size': 13, 'bold': True, 'color': GRAY, 'space': 6}),
    ('図を左から順に指でなぞる：朝は買う → 昼は作って使う・貯める → 夜は貯めた電気を使う',
     {'size': 13, 'space': 5}),
    ('「売電15円より、買う電気30円の方が高い。だから売らずに使うんです」', {'size': 13, 'space': 5}),
    ('ここが曖昧だと、次のシミュレーションが伝わりません', {'size': 13, 'bold': True, 'color': RED}),
])
_, tf = rect(s, 0.45, 5.80, 9.95, 0.58, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '緑帯の結論（そのまま言い切る）　', size=12, bold=True, color=GRAY)
run(p, '具体的なシミュレーションを見ていきましょう！', size=15, bold=True, hl=YELLOW)
_, tf = rect(s, 0.45, 6.48, 9.95, 0.50, fill=LTYEL, line='BF8F00', anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [('この一文で資料を閉じ、シミュレーションを開きます。ここが切り替えの合図です',
            {'size': 15, 'bold': True, 'align': PP_ALIGN.CENTER})])

# ================================================================ 20. シミュレーションの読み方
s = new_slide('③ 経済メリット')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('シミュレーションは「どこを読むか」だけ教える', {'size': 23, 'bold': True})])
box, tf = tb(s, 0.45, 1.86, 9.95, 0.32)
lines(tf, [('数字が多い資料です。お客様に見せるのは2か所だけにしてください',
            {'size': 14, 'color': GRAY})])
for x, ttl, a, b, c, d in (
    (0.45, 'FIT期間', '24,002 円/月', '14,429 円/月', '9,573', '約 319 円'),
    (5.60, 'FIT期間終了後', '24,002 円/月', '12,686 円/月', '11,316', '約 377 円'),
):
    _, tf = rect(s, x, 2.26, 4.80, 0.40, fill='DAEEF3', anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [(ttl, {'size': 15, 'bold': True, 'align': PP_ALIGN.CENTER})])
    rows = [['毎月の分割支払額', a], ['− 実質光熱費削減額', b]]
    table(s, x, 2.72, 4.80, [2.95, 1.85], rows, font_size=13, row_h=0.42, head_h=0.42, header=False)
    _, tf = rect(s, x, 3.62, 4.80, 0.90, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
    p = para(tf, first=True, align=PP_ALIGN.CENTER)
    run(p, '＝ 毎月の実質負担額　', size=13, bold=True)
    run(p, c, size=24, bold=True, color=RED)
    run(p, ' 円/月', size=13, bold=True)
    p = para(tf, align=PP_ALIGN.CENTER)
    run(p, '（1日あたり %s）' % d, size=12)
_, tf = rect(s, 0.45, 4.72, 4.80, 1.55, fill='FCE4E4', line=RED)
lines(tf, [
    ('✕ 見せてはいけないところ', {'size': 15, 'bold': True, 'color': RED, 'space': 6}),
    ('いちばん上の「分割支払額 24,002円」を先に見せると、', {'size': 13, 'space': 4}),
    ('そこで思考が止まります。', {'size': 13, 'bold': True}),
])
_, tf = rect(s, 5.60, 4.72, 4.80, 1.55, fill=LTGREEN, line='70A040')
lines(tf, [
    ('○ 見せるところ', {'size': 15, 'bold': True, 'space': 6}),
    ('「毎月の実質負担額」と「1日あたり」の2か所だけ。', {'size': 13, 'space': 4}),
    ('「1日319円。缶コーヒー2本分です」', {'size': 13, 'bold': True}),
])
_, tf = rect(s, 0.45, 6.40, 9.95, 0.58, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '1日あたりの金額まで落として、初めて判断できる数字になる', size=16, bold=True, hl=YELLOW)

# ================================================================ 21. 伝え方①②
s = new_slide('③ 経済メリット')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('伝え方は2通り。お客様を見て選びます', {'size': 23, 'bold': True})])
_, tf = rect(s, 0.45, 1.90, 4.80, 0.48, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [('伝え方①　ローン返済額を含めたトータル金額', {'size': 14, 'bold': True})])
rows = [
    ['（例）', 'ローン中', '完済後'],
    ['もともとの電気代', '15,000', '15,000'],
    ['＋ ローンの分割額', '22,000', '0'],
    ['− 電気代削減額', '−14,000', '−14,000'],
    [('支出合計', {'bold': True}), ('23,000', {'bold': True, 'fill': LTYEL}),
     ('1,000', {'bold': True, 'fill': LTYEL, 'color': RED})],
]
table(s, 0.45, 2.44, 4.80, [1.90, 1.50, 1.40], rows, font_size=12, row_h=0.42, head_h=0.32)
_, tf = rect(s, 0.45, 4.62, 4.80, 1.40, fill=LTGRAY)
lines(tf, [
    ('向いているお客様', {'size': 14, 'bold': True, 'space': 5}),
    ('・長期で考えられる方', {'size': 13, 'space': 3}),
    ('・完済後の姿に価値を感じる方', {'size': 13, 'space': 5}),
    ('返済中は支出が増えます。隠さないこと。', {'size': 13, 'bold': True, 'color': RED}),
])
_, tf = rect(s, 5.60, 1.90, 4.80, 0.48, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [('伝え方②　電気代の減収分のみ', {'size': 14, 'bold': True})])
rows = [
    ['（例）', '金額', '削減額'],
    ['もともとの電気代', '15,000', '1年 ¥168,000'],
    ['− 電気代削減額', '14,000', '10年 ¥1,680,000'],
    [('支出合計', {'bold': True}), ('1,000', {'bold': True, 'fill': LTYEL}),
     ('30年 ¥5,040,000', {'bold': True, 'fill': LTYEL, 'color': RED})],
]
table(s, 5.60, 2.44, 4.80, [1.90, 1.20, 1.70], rows, font_size=12, row_h=0.42, head_h=0.32)
_, tf = rect(s, 5.60, 4.62, 4.80, 1.40, fill=LTGRAY)
lines(tf, [
    ('向いているお客様', {'size': 14, 'bold': True, 'space': 5}),
    ('・現金一括で検討される方', {'size': 13, 'space': 3}),
    ('・月々の支出が増えることに抵抗がある方', {'size': 13, 'space': 5}),
    ('①で反応が鈍かったら②に切り替える。', {'size': 13, 'bold': True, 'color': RED}),
])
_, tf = rect(s, 0.45, 6.15, 9.95, 0.78, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('どちらも同じ事実を、別の枠組みで見せているだけです。', {'size': 16, 'bold': True, 'space': 4}),
    ('お客様の反応を見て、途中で切り替えて構いません。', {'size': 15}),
])

# ================================================================ 22. 疑われている
s = new_slide('③ 経済メリット')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('前提：シミュレーションは疑われています', {'size': 23, 'bold': True, 'color': RED})])
pic(s, os.path.join(ASSETS, 'sim_trust.jpg'), 0.45, 1.92, 5.35)
box, tf = tb(s, 0.45, 5.30, 5.35, 0.28)
lines(tf, [('出典：エネがえる運営事務局調べ（国際航業株式会社）', {'size': 10, 'color': GRAY})])
_, tf = rect(s, 6.10, 1.92, 4.30, 1.45, fill='FCE4E4', line=RED, anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [
    ('経済効果シミュレーションの信憑性を', {'size': 14, 'bold': True, 'align': PP_ALIGN.CENTER, 'space': 2}),
    ('疑ったことがある人', {'size': 14, 'bold': True, 'align': PP_ALIGN.CENTER, 'space': 5}),
    ('75.4 ％', {'size': 32, 'bold': True, 'color': RED, 'align': PP_ALIGN.CENTER}),
])
_, tf = rect(s, 6.10, 3.55, 4.30, 2.30, fill=LTGRAY)
lines(tf, [
    ('だから、こう伝えます', {'size': 15, 'bold': True, 'space': 8}),
    ('・金額の大きさで勝とうとしない', {'size': 13, 'space': 6}),
    ('・前提条件を先に口で言う', {'size': 13, 'space': 2}),
    ('　（電気使用量・単価・設置容量）', {'size': 13, 'color': GRAY, 'space': 6}),
    ('・お客様の検針票の実額から出発する', {'size': 13, 'space': 6}),
    ('「盛っていない」ことが伝わる方が', {'size': 13, 'bold': True, 'color': RED, 'space': 2}),
    ('金額より効きます', {'size': 13, 'bold': True, 'color': RED}),
])
_, tf = rect(s, 0.45, 6.05, 9.95, 0.85, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [
    ('言い方の例', {'size': 12, 'bold': True, 'color': GRAY, 'align': PP_ALIGN.CENTER, 'space': 3}),
    ('「この数字は、〇〇様の検針票の18,000円から計算しています。盛った数字ではありません」',
     {'size': 16, 'bold': True, 'align': PP_ALIGN.CENTER}),
])

# ================================================================ 23. ③のまとめ
s = new_slide('③ 経済メリット')
box, tf = tb(s, 0.45, 1.45, 9.95, 0.55)
lines(tf, [('このパートでは「全体の金額」を出しません', {'size': 26, 'bold': True, 'color': RED})])
rows = [
    ['', 'このパート（必要性訴求）', '次のパート（金額訴求）'],
    ['目的', '「電気代がいくら減るか」を納得してもらう', '「毎月いくら払うか」を払える形にする'],
    ['出す数字', '削減額（月14,000円／30年504万円）', '月々の支払額・概算見積・総額'],
    ['使う道具', 'アプローチブック＋シミュレーション', 'シミュレーション＋概算見積＋カタログ'],
]
table(s, 0.45, 2.20, 9.95, [1.30, 4.30, 4.35], rows, font_size=13.5, row_h=0.78, head_h=0.36)
_, tf = rect(s, 0.45, 5.15, 9.95, 1.30, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('なぜ分けるのか', {'size': 16, 'bold': True, 'space': 6}),
    ('「いくら得か」が腹落ちしていない状態で「いくら払うか」を出すと、支払額だけが記憶に残ります。',
     {'size': 15, 'space': 4}),
    ('このパートは「得の大きさ」を作りきるところまでです。', {'size': 15, 'bold': True, 'color': RED}),
])

# ================================================================ 24. 通しの時間配分
s = new_slide('通しで使う')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('必要性訴求パートの時間配分（商談での目安）', {'size': 23, 'bold': True})])
box, tf = tb(s, 0.45, 1.86, 9.95, 0.32)
lines(tf, [('※ 研修の時間割ではありません。お客様との商談1時間のうち、この部分に使えるのは約40分です',
            {'size': 14, 'color': GRAY})])
bars = [
    ('導入', '検針票の実額を聞く', '2分', 0.78, LTGRAY),
    ('① 電気代', 'p.9 〜 p.17（7ページ）', '20分', 4.20, LTGREEN),
    ('② 災害対策', 'p.18・p.19 ＋ 事例1本', '3分', 0.98, LTYEL),
    ('③ 経済メリット', 'p.20・p.21 ＋ シミュレーション', '15分', 3.20, LTBLUE),
]
y = 2.32
for name, detail, mins, blen, fill in bars:
    _, tf = rect(s, 0.45, y, 1.95, 0.72, fill=fill, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [(name, {'size': 15, 'bold': True, 'align': PP_ALIGN.CENTER})])
    _, tf = rect(s, 2.50, y, blen, 0.72, fill=fill, line='808080', lw=0.75,
                 anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [(mins, {'size': 18, 'bold': True, 'color': RED,
                       'align': PP_ALIGN.CENTER})])
    box, tf2 = tb(s, 7.00, y, 3.40, 0.72, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf2, [(detail, {'size': 12})])
    y += 0.86
_, tf = rect(s, 0.45, 5.90, 4.80, 1.00, fill=LTGRAY)
lines(tf, [
    ('時間が押したときに削る順番', {'size': 14, 'bold': True, 'space': 5}),
    ('p.15・16 → p.10・p.12 → p.11・p.13', {'size': 15, 'bold': True, 'color': RED, 'space': 3}),
    ('p.9・p.14・p.17・p.21 は削らない', {'size': 13}),
])
_, tf = rect(s, 5.60, 5.90, 4.80, 1.00, fill='FCE4E4', line=RED)
lines(tf, [
    ('いちばん多い失敗', {'size': 14, 'bold': True, 'color': RED, 'space': 5}),
    ('①で30分以上かけてしまい、', {'size': 14, 'space': 3}),
    ('③のシミュレーションが駆け足になる', {'size': 14, 'bold': True}),
])

# ================================================================ 25. NG集
s = new_slide('通しで使う')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('めくり方のNG集', {'size': 24, 'bold': True})])
box, tf = tb(s, 0.45, 1.86, 9.95, 0.32)
lines(tf, [('ロープレで実際に多い5つです。自分に当てはまるものに印をつけてください',
            {'size': 14, 'color': GRAY})])
rows = [
    ['✕ やりがちなこと', 'なぜまずいか', '○ こうする'],
    [('リード文を全部読み上げる', {'bold': True}), '読める文字を読まれると、聞くのをやめる',
     ('要点を1文に縮めて、自分の言葉で言う', {'bold': True})],
    [('グラフの棒を端から順に説明する', {'bold': True}), '数字が多すぎて、どれが重要か分からない',
     ('指でさして、数字は1ページ1つだけ', {'bold': True})],
    [('全国平均のまま話す', {'bold': True}), '「うちは違う」と思われ、他人事になる',
     ('検針票の実額に置き換えて話す', {'bold': True})],
    [('緑帯を飛ばして次のページへ', {'bold': True}), '結論が無いまま進むので、何も残らない',
     ('緑帯は必ず声に出して言い切る', {'bold': True})],
    [('災害の話を5分以上する', {'bold': True}), '場が暗くなり、そのあとの提案が入らない',
     ('3分で抜ける。事実と第三者で話す', {'bold': True})],
]
table(s, 0.45, 2.30, 9.95, [2.80, 3.60, 3.55], rows, font_size=12.5, row_h=0.74, head_h=0.34)
_, tf = rect(s, 0.45, 6.30, 9.95, 0.62, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
p = para(tf, first=True, align=PP_ALIGN.CENTER)
run(p, '共通しているのは「資料を読んでいる」状態。資料は読むものではなく、指すものです',
    size=16, bold=True, hl=YELLOW)

# ================================================================ 26. 通し練習
s = new_slide('通しで使う')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('通し練習：p.9 から、5分で行けるところまで', {'size': 23, 'bold': True})])
_, tf = rect(s, 0.45, 1.92, 9.95, 1.45, fill=LTGRAY)
lines(tf, [('進め方（合計11分）', {'size': 14, 'bold': True, 'color': GRAY, 'space': 6})])
for it in [
    '2人1組。営業役とお客様役。営業役だけ1本（5分）やります',
    '営業役はアプローチブックを持ち、p.9を開いた状態から始めます',
    'お客様役は「月18,000円」という設定。最初に必ず聞かれるので答えてください',
    'そのあとフィードバック3分 → 役を交代してもう1本（時間があれば）',
]:
    p = para(tf, space_after=4); run(p, '・' + it, size=14)
_, tf = rect(s, 0.45, 3.48, 4.80, 1.55, fill=LTGREEN, line='70A040')
lines(tf, [
    ('お客様役が見るところ', {'size': 15, 'bold': True, 'space': 6}),
    ('・自分が何回しゃべったか', {'size': 13, 'space': 4}),
    ('・「読まれている」と感じた場面', {'size': 13, 'space': 4}),
    ('・数字が多すぎて追えなくなった場面', {'size': 13}),
])
_, tf = rect(s, 5.60, 3.48, 4.80, 1.55, fill=LTBLUE, line='4B8FC0')
lines(tf, [
    ('講師が机間で見るところ', {'size': 15, 'bold': True, 'space': 6}),
    ('・発言比率 6：4 になっているか', {'size': 13, 'space': 4}),
    ('・緑帯を言い切れているか', {'size': 13, 'space': 4}),
    ('・太陽光だけ／蓄電池だけの話になっていないか', {'size': 13}),
])
_, tf = rect(s, 0.45, 5.16, 9.95, 1.00, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('フィードバックの型', {'size': 15, 'bold': True, 'space': 5}),
    ('① 良かったページを1つ、ページ番号で言う　② 直すのは1人1つだけ　③ 人ではなくページを指摘する',
     {'size': 14, 'bold': True}),
])
_, tf = rect(s, 0.45, 6.30, 9.95, 0.62, fill=None, line=RED, lw=1.5,
             anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [('5分で13ページは回りません。回りきらなかったページが、そのまま宿題です',
            {'size': 16, 'bold': True, 'align': PP_ALIGN.CENTER})])

# ================================================================ 27. チェックシート
s = new_slide('通しで使う')
box, tf = tb(s, 0.45, 1.36, 9.95, 0.45)
lines(tf, [('セルフチェックシート', {'size': 24, 'bold': True})])
box, tf = tb(s, 0.45, 1.86, 9.95, 0.32)
lines(tf, [('ロープレのあとに記入し、できていない項目を次回までの宿題にしてください',
            {'size': 14, 'color': GRAY})])
rows = [
    ['確認すること', '今日', '次回'],
    ['資料を見ずに、p.9の結論一文が言えた', ' ', ' '],
    ['p.14の「1,210円 → 22,990円」が言えた', ' ', ' '],
    ['p.17でお客様の実額の列を指せた', ' ', ' '],
    ['②災害を3分以内で抜けられた', ' ', ' '],
    ['p.21で「売電15円・買電30円」を言えた', ' ', ' '],
    ['シミュレーションで前提条件を先に言えた', ' ', ' '],
    ['13ページを15分で通せた', ' ', ' '],
    ['お客様の発言が4割あった', ' ', ' '],
]
table(s, 0.45, 2.30, 6.35, [4.35, 1.00, 1.00], rows, font_size=13, row_h=0.50, head_h=0.34)
_, tf = rect(s, 7.05, 2.30, 3.35, 2.15, fill=LTGREEN, line='70A040')
lines(tf, [
    ('合格ライン', {'size': 15, 'bold': True, 'space': 6}),
    ('8項目中 6つ', {'size': 26, 'bold': True, 'color': RED, 'space': 6}),
    ('うち、上から3つは必須です', {'size': 13, 'bold': True, 'space': 2}),
    ('（p.9・p.14・p.17）', {'size': 12, 'color': GRAY}),
])
_, tf = rect(s, 7.05, 4.60, 3.35, 2.06, fill=LTGRAY)
lines(tf, [
    ('次回までの宿題', {'size': 15, 'bold': True, 'space': 7}),
    ('できなかった項目を1つ選び、', {'size': 13, 'space': 3}),
    ('そのページだけを10回、', {'size': 13, 'space': 3}),
    ('声に出して読んでください。', {'size': 13, 'space': 7}),
    ('翌日には50％忘れます。', {'size': 13, 'bold': True, 'color': RED}),
])

# ================================================================ 28. まとめ
s = new_slide('まとめ')
box, tf = tb(s, 0.45, 1.42, 9.95, 0.55)
lines(tf, [('このパートで持ち帰ってほしい3つ', {'size': 26, 'bold': True})])
items = [
    ('1', '資料は読むものではなく、指すもの',
     'リード文は読まない。数字は1ページ1つだけ指す。残りはお客様に聞かせない。'),
    ('2', '緑帯の結論は、必ず声に出して言い切る',
     '13ページ分の緑帯が、そのまま必要性訴求のトークスクリプトです。'),
    ('3', '全国平均ではなく、お客様の検針票の実額で話す',
     '最初に月額を聞く。以降の数字はすべてそこから計算し直す。'),
]
y = 2.15
for n, t, d in items:
    _, tf = rect(s, 0.45, y, 0.72, 1.35, fill=RED, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [(n, {'size': 30, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER})])
    _, tf = rect(s, 1.28, y, 9.12, 1.35, fill=LTGREEN, anchor=MSO_ANCHOR.MIDDLE)
    lines(tf, [
        (t, {'size': 20, 'bold': True, 'space': 6}),
        (d, {'size': 14}),
    ])
    y += 1.50
_, tf = rect(s, 0.45, 6.30, 9.95, 0.62, fill=None, line=RED, lw=1.75,
             anchor=MSO_ANCHOR.MIDDLE)
lines(tf, [('次のパートは「時期訴求・金額訴求」。ここで作った"得の大きさ"が土台になります',
            {'size': 16, 'bold': True, 'align': PP_ALIGN.CENTER})])
notes(s, '・3つとも「アプローチブックの扱い方」の話であって、話術の話ではないことを強調する')

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
