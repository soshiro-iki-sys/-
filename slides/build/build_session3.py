# -*- coding: utf-8 -*-
"""第3回「必要性訴求②＆ロープレ」研修資料ビルダー

使い方は slides/build/build_session2.py と同じ。
アプローチブックのページ画像（<ASSETS>/ab/ab-01.jpg 〜 ab-26.jpg）と
ピラミッド図（<ASSETS>/pyramid.png）を用意してから実行する。

    TRAINING_ASSETS=/path/to/assets python slides/build/build_session3.py

書式は docs/研修資料フォーマット仕様.md、内容は docs/知識_*.md を参照。
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
OUT    = os.path.join(REPO, 'slides', '2026営業研修_第3回_必要性訴求②＆ロープレ.pptx')

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

def rect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, lw=1.0):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill: s.fill.solid(); s.fill.fore_color.rgb = RGBColor.from_string(fill)
    else: s.fill.background()
    if line:
        s.line.color.rgb = RGBColor.from_string(line); s.line.width = Pt(lw)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    tf = s.text_frame; tf.word_wrap = True
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

def ab_img(n):
    return os.path.join(AB, 'ab-%02d.jpg' % n)

# ---------------------------------------------------------------- 表紙
cover = prs.slides[0]
for shp in cover.shapes:
    if shp.shape_type == MSO_SHAPE_TYPE_AUTO and shp.name == '正方形/長方形 4':
        sp = shp._element.spPr
        fill = sp.find(qn('a:solidFill'))
        for ch in list(fill): fill.remove(ch)
        fill.append(parse_xml(
            '<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'val="F79646"><a:alpha val="80000"/></a:srgbClr>'))
    if shp.has_text_frame and shp.name == 'テキスト ボックス 5':
        tf = shp.text_frame; tf.clear()
        lines(tf, [
            ('2026年度　営業研修', {'size': 32, 'bold': True, 'color': WHITE, 'font': UD, 'space': 8}),
            ('第3回　必要性訴求②', {'size': 52, 'bold': True, 'color': WHITE, 'font': UD, 'space': 2}),
            ('＆ロープレ', {'size': 52, 'bold': True, 'color': WHITE, 'font': UD, 'space': 10}),
            ('〜太陽光＋蓄電池セット販売〜', {'size': 24, 'bold': True, 'color': WHITE, 'font': UD}),
        ])
    if shp.has_text_frame and shp.name == 'テキスト ボックス 6':
        tf = shp.text_frame; tf.clear()
        lines(tf, [('2026年8月下旬　＠船井総研', {'size': 14, 'color': WHITE, 'font': UD})],
              align=PP_ALIGN.RIGHT)
page_no[0] = 1

# ---------------------------------------------------------------- rule helpers
def rule_head(s, no_title, point):
    box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
    lines(tf, [(no_title, {'size': 24, 'bold': True})])
    _, tf = rect(s, 0.45, 1.88, 9.95, 0.46, fill=None, line=RED, lw=1.25)
    p = para(tf, first=True, align=PP_ALIGN.LEFT)
    run(p, 'POINT　', size=15, bold=True, color=RED)
    run(p, point, size=16, bold=True)

def goal_box(s, x, y, w, h, quote):
    _, tf = rect(s, x, y, w, h, fill=LTBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    lines(tf, [
        ('ゴール（お客様の反応）', {'size': 13, 'bold': True, 'color': BLUE, 'space': 6}),
    ])
    for q in quote:
        p = para(tf, space_after=3)
        run(p, '「%s」' % q, size=16, bold=True)

def points_box(s, x, y, w, h, items, title='説明のポイント'):
    _, tf = rect(s, x, y, w, h, fill=LTGRAY)
    lines(tf, [(title, {'size': 13, 'bold': True, 'color': GRAY, 'space': 6})])
    for it in items:
        opts = {}
        if isinstance(it, tuple): it, opts = it
        p = para(tf, space_after=opts.pop('space', 6))
        run(p, it, size=opts.pop('size', 12.5), bold=opts.pop('bold', False),
            color=opts.pop('color', BLACK))

def conclusion(s, text, y=6.36):
    _, tf = rect(s, 0.45, y, 9.95, 0.66, fill=LTGREEN)
    p = para(tf, first=True, align=PP_ALIGN.CENTER)
    run(p, 'このページの結論　', size=12, bold=True, color=GRAY)
    run(p, text, size=15, bold=True, hl=YELLOW)

def talk_slide(s, no_title, talks, conc=None, note=None):
    box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
    lines(tf, [(no_title, {'size': 24, 'bold': True})])
    y = 1.92
    for label, body, h in talks:
        _, tf = rect(s, 0.45, y, 8.05, h, fill=LTGRAY)
        for i, ln in enumerate(body):
            p = para(tf, first=(i == 0), space_after=3)
            run(p, ln, size=14)
        _, tf2 = rect(s, 8.68, y, 1.72, 0.38, fill=ORANGE)
        lines(tf2, [(label, {'size': 12, 'bold': True, 'color': WHITE,
                             'align': PP_ALIGN.CENTER})])
        y += h + 0.16
    if conc: conclusion(s, conc)
    if note: notes(s, note)

# ================================================================ 2. 本日のゴール
s = new_slide('本日のゴール')
box, tf = tb(s, 0.45, 1.45, 9.9, 0.62)
lines(tf, [('災害と使い方のパートを、サラッと・確実に通せるようになる',
            {'size': 28, 'bold': True, 'color': RED})])
_, tf = rect(s, 0.45, 2.35, 9.95, 3.05, fill=LTGREEN)
lines(tf, [
    ('研修が終わったときの「できる状態」', {'size': 18, 'bold': True, 'space': 10}),
    ('①　「うちの地域は災害が来ない」を、その場で潰せる', {'size': 19, 'space': 9}),
    ('②　災害の話を第三者話法だけで、3分以内に終えられる', {'size': 19, 'space': 9}),
    ('③　太陽光で創る→蓄電池に貯める→夜に使う、を図なしで説明できる', {'size': 19, 'space': 9}),
    ('④　p.21からシミュレーションへ、自然に受け渡せる', {'size': 19}),
])
_, tf = rect(s, 0.45, 5.65, 9.95, 0.85, fill=None, line=RED, lw=1.5)
lines(tf, [('本日の範囲：営業ルール ⑩〜⑫　／　アプローチブック p.18〜p.21',
            {'size': 22, 'bold': True, 'align': PP_ALIGN.CENTER})])
notes(s, '・扱うルールは3つだけ。その分ロープレを2周まわすと最初に伝える\n'
         '・災害パートは「長く話すほど失敗する」ことを最初に握っておく')

# ================================================================ 3. アジェンダ
s = new_slide('本日のアジェンダ')
rows = [
    ['時間', '内容', '扱うもの'],
    ['0:00-0:10', '前回の振り返り／宿題の確認', '第2回：ルール①〜⑨'],
    ['0:10-0:20', '必要性訴求の原理（価値÷価格・FAB）', '−'],
    ['0:20-0:50', ('ルール⑩　非常時に備える意義', {'bold': True}), 'アプローチブック p.18〜p.19'],
    ['0:50-1:00', '休憩', '−'],
    ['1:00-1:20', ('ルール⑪　非常時対策の効果の確認', {'bold': True}), '（該当ページなし）メーカー事例'],
    ['1:20-1:40', ('ルール⑫　太陽光・蓄電池の使い方', {'bold': True}), 'アプローチブック p.20〜p.21'],
    ['1:40-1:55', 'ロールプレイング（2周）', 'p.18〜21　→　p.9〜21 通し'],
    ['1:55-2:00', 'チェック／まとめ／宿題', '−'],
]
table(s, 0.45, 1.62, 9.95, [1.55, 4.3, 4.1], rows, font_size=14, row_h=0.50, head_h=0.38)
notes(s, '・ロープレを2周まわすため、解説は巻き気味に進める')

# ================================================================ 4. 前回の振り返り
s = new_slide('前回の振り返り')
box, tf = tb(s, 0.45, 1.42, 9.9, 0.45)
lines(tf, [('第2回「聞く姿勢づくり＆必要性訴求①」で押さえたこと', {'size': 24, 'bold': True})])
_, tf = rect(s, 0.45, 2.00, 4.85, 2.30, fill=LTGRAY)
lines(tf, [
    ('聴く姿勢作り（①〜⑥）', {'size': 17, 'bold': True, 'space': 8}),
    ('・失注理由は「商談設定」か「商談内容」しかない', {'size': 14, 'space': 5}),
    ('・先に相手の話を聞くことが、自分の話を聞いてもらうことに直結する', {'size': 14, 'space': 5}),
    ('・包み込みの法則で「先生－生徒」の関係をつくる', {'size': 14, 'space': 5}),
    ('・太陽光と蓄電池は切り離さず、常にセットで話す', {'size': 14, 'bold': True, 'color': RED}),
])
_, tf = rect(s, 5.55, 2.00, 4.85, 2.30, fill=LTGRAY)
lines(tf, [
    ('必要性訴求①（⑦〜⑨）', {'size': 17, 'bold': True, 'space': 8}),
    ('・電気代は2011年以降45％、約5.5万円上昇', {'size': 14, 'space': 5}),
    ('・再エネ賦課金は13年で18倍（2026年 4.18円/kWh）', {'size': 14, 'space': 5}),
    ('・30年で432万円〜700万円以上を支払う', {'size': 14, 'space': 5}),
    ('・基本設計は「下げてから上げる」', {'size': 14, 'bold': True, 'color': RED}),
])
_, tf = rect(s, 0.45, 4.60, 9.95, 1.75, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('宿題の確認（10分）', {'size': 18, 'bold': True, 'space': 8}),
    ('① 自社ストーリー5段構成を、20秒で言ってみてください（2名指名）', {'size': 16, 'space': 6}),
    ('② アプローチブック p.9・p.14・p.17 の結論の一文を、資料を見ずに言えますか', {'size': 16, 'space': 6}),
    ('③ 実商談のチェックシートで、○が付かなかったルールはどれでしたか', {'size': 16}),
])
notes(s, '・宿題をやっていない人を責めない。やってきた人を先に指名して場を温める')

# ================================================================ 5. 全体像① 4ステップ
s = new_slide('全体像①：契約までの4ステップ')
pic(s, os.path.join(ASSETS, 'pyramid.png'), 0.55, 1.55, 6.0, border=False)
_, tf = rect(s, 6.95, 1.75, 3.45, 1.55, fill=None, line=RED, lw=2.0)
lines(tf, [
    ('本日の範囲', {'size': 20, 'bold': True, 'color': RED, 'align': PP_ALIGN.CENTER, 'space': 8}),
    ('STEP2　必要性訴求（後半）', {'size': 17, 'bold': True, 'space': 5}),
    ('ルール⑩〜⑫', {'size': 17, 'bold': True}),
])
_, tf = rect(s, 6.95, 3.55, 3.45, 2.55, fill=LTGRAY)
lines(tf, [
    ('必要性訴求の残り3つ', {'size': 17, 'bold': True, 'space': 8}),
    ('⑦⑧⑨で「電気代」という', {'size': 14, 'space': 0}),
    ('お金のネガは作れました。', {'size': 14, 'space': 8}),
    ('今日はそこに', {'size': 14, 'space': 0}),
    ('「もしもの備え」と', {'size': 14, 'bold': True, 'space': 0}),
    ('「使い方のイメージ」', {'size': 14, 'bold': True, 'space': 0}),
    ('を足して、STEP2を閉じます。', {'size': 14, 'space': 8}),
    ('ここまでで「欲しい」を', {'size': 14, 'space': 0}),
    ('作りきります。', {'size': 14}),
])

# ================================================================ 6. 全体像② 20のルール
s = new_slide('全体像②：太陽光＋蓄電池営業 20のルール')
DONE = {'fill': LTGRAY, 'color': GRAY}
HL = {'fill': LTYEL}
rows = [
    ['項目', 'ゴール', 'ルール', '詳細'],
    [('前準備', dict(DONE)), ('必要な情報を揃える', dict(DONE)), ('①商談テーブルの確認', dict(DONE)), ('第2回で実施済み', dict(DONE))],
    [('聴く姿勢作り', dict(DONE)), ('商談の趣旨を認識する', dict(DONE)), ('②〜⑥', dict(DONE)), ('第2回で実施済み', dict(DONE))],
    [('必要性訴求', dict(DONE)), ('現状を正確に把握する', dict(DONE)), ('⑦直近数か年の電気代推移の理解', dict(DONE)), ('第2回で実施済み', dict(DONE))],
    [('', dict(DONE)), ('', dict(DONE)), ('⑧再エネ賦課金の構造の理解', dict(DONE)), ('第2回で実施済み', dict(DONE))],
    [('', dict(DONE)), ('', dict(DONE)), ('⑨長期で支払う電気代総額の理解', dict(DONE)), ('第2回で実施済み', dict(DONE))],
    [('', dict(HL)), ('', dict(HL)), ('⑩非常時に備える意義の理解', dict(HL)),
     ('度重なる災害から、非常時に備える必要があることを理解してもらう', dict(HL))],
    [('', dict(HL)), ('', dict(HL)), ('⑪非常時対策の効果の確認', dict(HL)),
     ('太陽光・蓄電池があれば、向こう30年近くは停電時の保証が得られることを理解してもらう', dict(HL))],
    [('', dict(HL)), ('', dict(HL)), ('⑫太陽光・蓄電池の使い方の理解', dict(HL)),
     ('シミュレーションに移る前に太陽光・蓄電池の一般的な使い方を理解してもらう', dict(HL))],
    ['時期訴求', '早期対策を進める', '⑬〜⑮', '第4回'],
    ['金額訴求', '費用対効果に納得する', '⑯〜⑳', '第4回'],
]
table(s, 0.30, 1.58, 10.25, [1.15, 1.70, 2.70, 4.70], rows, font_size=11, row_h=0.42, head_h=0.30)
_, tf = rect(s, 0.30, 6.32, 10.25, 0.50, fill=None, line=RED, lw=1.5)
lines(tf, [('黄色が本日の範囲。⑩〜⑫で「欲しい」を作りきり、シミュレーションに渡します',
            {'size': 17, 'bold': True, 'align': PP_ALIGN.CENTER})])

# ================================================================ 7. 章見出し
s = new_slide('必要性訴求②')
box, tf = tb(s, 0.45, 1.38, 9.9, 0.42)
lines(tf, [('ゴール：現状を正確に把握してもらう（後半）', {'size': 22, 'bold': True})])
rows = [
    ['項目', 'ゴール', 'ルール', 'POINT'],
    ['必要性訴求', '現状を正確に把握する', '⑩非常時に備える意義の理解',
     '「いつ身の回りで起きてもおかしくない」に腹落ちさせる'],
    ['', '', '⑪非常時対策の効果の確認', '災害トークは「第三者話法」でしか伝えない'],
    ['', '', '⑫太陽光・蓄電池の使い方の理解', 'シミュレーションに移る前に基本概念をおさらいする'],
]
for r in rows[1:]:
    for i in range(len(r)):
        r[i] = (r[i], {'fill': LTGREEN})
table(s, 0.30, 1.95, 10.25, [1.25, 1.75, 2.85, 4.40], rows, font_size=12, row_h=0.55, head_h=0.32)
_, tf = rect(s, 0.30, 4.10, 4.95, 2.30, fill=LTGRAY)
lines(tf, [
    ('使うアプローチブック', {'size': 17, 'bold': True, 'space': 8}),
    ('p.18　政府機関にも予測が困難な地震情報', {'size': 15, 'space': 6}),
    ('p.19　自然災害とともに停電被害はたくさん出ている', {'size': 15, 'space': 6}),
    ('p.20　電気を購入する生活から自給自足する生活へ', {'size': 15, 'space': 6}),
    ('p.21　太陽光発電・蓄電池の使い方', {'size': 15, 'space': 10}),
    ('※ ⑪に対応するページはありません（理由は後述）', {'size': 14, 'bold': True, 'color': RED}),
])
_, tf = rect(s, 5.60, 4.10, 4.95, 2.30, fill=LTYEL)
lines(tf, [
    ('このパートの勝ち筋', {'size': 17, 'bold': True, 'space': 8}),
    ('災害の話は、長く話すほど失敗します。', {'size': 15, 'space': 6}),
    ('商談の雰囲気が暗くなるからです。', {'size': 15, 'space': 10}),
    ('「事実 → 事例 → サラッと次へ」', {'size': 16, 'bold': True, 'color': RED, 'space': 6}),
    ('が鉄則。3分で抜けてください。', {'size': 15}),
])

# ================================================================ 8. 購入＝価値÷価格
s = new_slide('必要性訴求の原理')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('なぜ「必要性」を訴求するのか', {'size': 24, 'bold': True})])
_, tf = rect(s, 0.90, 1.98, 4.15, 1.70, fill=LTBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
lines(tf, [
    ('購入 ＝ 価値 ÷ 価格', {'size': 22, 'bold': True, 'align': PP_ALIGN.CENTER, 'space': 8}),
    ('この値が1を超えたとき、', {'size': 14, 'align': PP_ALIGN.CENTER, 'space': 2}),
    ('人は購入に至る', {'size': 14, 'align': PP_ALIGN.CENTER}),
])
_, tf = rect(s, 5.35, 1.98, 5.05, 1.70, fill=LTGRAY)
lines(tf, [
    ('100均で考えると', {'size': 15, 'bold': True, 'space': 6}),
    ('100円の商品を買うのは、100円以上の価値を感じているから。', {'size': 14, 'space': 5}),
    ('つまり X > 100円 ÷ 100円 の状態です。', {'size': 14}),
])
_, tf = rect(s, 0.45, 3.90, 4.95, 2.35, fill=LTGREEN)
lines(tf, [
    ('やり方①　価値を上げる', {'size': 18, 'bold': True, 'space': 8}),
    ('電気代のネガ（⑦⑧⑨）', {'size': 15, 'space': 4}),
    ('＋ もしもの備え（⑩⑪）', {'size': 15, 'space': 4}),
    ('＋ 使い方のイメージ（⑫）', {'size': 15, 'space': 10}),
    ('本日はここを積み上げます', {'size': 16, 'bold': True, 'color': RED}),
])
_, tf = rect(s, 5.60, 3.90, 4.95, 2.35, fill=LTGRAY)
lines(tf, [
    ('やり方②　価格を下げる', {'size': 18, 'bold': True, 'space': 8}),
    ('価格 −（経済メリット）を下げる', {'size': 15, 'space': 4}),
    ('＝ 時期訴求・金額訴求（⑬〜⑳）', {'size': 15, 'space': 10}),
    ('第4回で扱います', {'size': 16, 'bold': True, 'color': GRAY}),
])
conclusion(s, '必要性訴求とは「価値」を積み上げる作業のこと', y=6.40)

# ================================================================ 9. FAB
s = new_slide('必要性訴求の原理')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('価値の伝え方：FAB', {'size': 24, 'bold': True})])
rows = [
    ['', '要素', '意味', 'できる営業の割合'],
    ['F', 'Feature（特徴）', 'その製品の特長・長所　※固定', 'ほとんどの営業マンができる'],
    ['A', 'Advantage（メリット）', 'その特徴が「全ての人に」もたらす利点　※固定', '約8割ができる'],
    ['B', ('Benefit（ベネフィット）', {'bold': True}), ('だからこそ「目の前のお客様が」受ける利益　※変動', {'bold': True}),
     ('トップ2割しかできない', {'bold': True, 'color': RED})],
]
table(s, 0.45, 1.95, 9.95, [0.50, 2.55, 4.60, 2.30], rows, font_size=13, row_h=0.62, head_h=0.32)
_, tf = rect(s, 0.45, 4.30, 9.95, 1.95, fill=LTGRAY)
lines(tf, [
    ('例）コーヒーマシン', {'size': 16, 'bold': True, 'space': 7}),
    ('F：専用カプセルを入れてボタン一つで本格的なコーヒーが淹れられる', {'size': 15, 'space': 5}),
    ('A：誰でも安定した味を再現でき、手間も汚れも少ない', {'size': 15, 'space': 5}),
    ('B：忙しい朝でもボタン1つで一息つける／カプセルが毎月届くので買い忘れ・豆挽き・ボトル運びの負担がない',
     {'size': 15, 'bold': True, 'color': RED}),
])
conclusion(s, 'お客様に対してのB（利益）を、いかに具体的に話せるかが最重要', y=6.40)
notes(s, '・「太陽光＋蓄電池のBは？」と受講者に投げる。次回以降のワークにつなぐ')

# ================================================================ 10. ルール⑩-1
s = new_slide('必要性訴求②')
rule_head(s, 'ルール⑩：非常時に備える意義の理解', '「いつ身の回りで起きてもおかしくない」に腹落ちさせる')
pic(s, ab_img(18), 0.45, 2.50, 4.55)
box, tf = tb(s, 0.45, 5.85, 4.55, 0.3)
lines(tf, [('アプローチブック p.18　予測が困難な地震情報', {'size': 12, 'color': GRAY})])
goal_box(s, 5.35, 2.50, 5.05, 1.15, ['いつ災害が起こるか分からないですね'])
points_box(s, 5.35, 3.78, 5.05, 2.50, [
    ('地震リスクが最低とされた札幌で、大規模停電が起きた', {'bold': True, 'size': 13}),
    '地震調査委員会が2018年6月に発表した「今後30年以内に震度6弱以上」の確率で札幌が最下位',
    'その3か月後に北海道胆振東部地震が発生。気象庁も「確度の高い地震の予測は難しい」と明言',
    ('「うちの地域は地震来ないから」を潰すためのページ', {'bold': True, 'color': RED, 'size': 13}),
])
conclusion(s, '専門家ですら、いつ「もしものこと」が起きるのか分からない')
notes(s, '・全国展開の某社では「災害時に使える」トークは評判が良くないとされる。'
         '地域によって災害への考え方が大きく異なるため、事実で外堀を埋める')

# ================================================================ 11. ルール⑩-2
s = new_slide('必要性訴求②')
talk_slide(s, 'ルール⑩：3段構成で逃げ道を無くす', [
    ('①権威', [
        '「実は、日本で最も権威のある『地震調査委員会』という機関が、定期的に地震のデータを出しているんです。」',
    ], 0.60),
    ('②事実', [
        '「そのデータの2018年6月版では、札幌市の地震発生確率が全国で最も低かったんですね。」',
    ], 0.60),
    ('③事実', [
        '「ところが、その3か月後に胆振東部地震が起きてしまいました。気象庁も『確度の高い地震の予測は難しい』と明言しています。」',
    ], 0.75),
    ('着地', [
        '「どこでどのくらい大きな地震が来るかなんて、分からないですよね？」',
        '（＝だからこそ、今の生活を守るために備えが必要ですよね？）',
    ], 0.75),
], conc='最終着地は「分からないですよね？」の一言に落とす',
   note='・3段はテンポよく。1段ごとに間を空けると重くなる')

# ================================================================ 12. ルール⑩-3
s = new_slide('必要性訴求②')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('ルール⑩：停電被害の実績で「規模」を見せる', {'size': 24, 'bold': True})])
pic(s, ab_img(19), 0.45, 1.95, 4.55)
box, tf = tb(s, 0.45, 5.30, 4.55, 0.3)
lines(tf, [('アプローチブック p.19　自然災害と停電被害', {'size': 12, 'color': GRAY})])
rows = [
    ['災害', '停電戸数', '備考'],
    ['2016年4月　熊本地震', '約50万戸', '−'],
    ['2018年9月　台風21号（関東）', '240万戸', '−'],
    ['2018年9月　台風24号', '約180万戸', '−'],
    ['2018年9月　北海道胆振東部地震', '約300万戸', ('日本初のブラックアウト', {'bold': True, 'color': RED})],
    ['2019年9月　台風15号（関東）', '93万戸', '5日連続で停電'],
    ['今後10年以内　南海トラフ地震？', '？', '−'],
]
table(s, 5.35, 1.95, 5.05, [2.35, 1.15, 1.55], rows, font_size=10.5, row_h=0.48, head_h=0.30)
_, tf = rect(s, 5.35, 5.30, 5.05, 1.00, fill=LTGRAY)
lines(tf, [
    ('胆振東部地震は平成最大規模。', {'size': 14, 'space': 4}),
    ('復旧まで2週間かかった地域もありました。', {'size': 14}),
])
conclusion(s, '災害大国だからこそ、もしもの時には備える必要があるのです')

# ================================================================ 13. ルール⑩-4
s = new_slide('必要性訴求②')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('ルール⑩：「うちの地域は災害が来ない」への対処', {'size': 24, 'bold': True})])
box, tf = tb(s, 0.45, 1.88, 9.95, 0.35)
lines(tf, [('ここでも逃げ道を作らないことが目的です。真正面から否定せず、事実で包み込みます',
            {'size': 15, 'color': GRAY})])
rows = [
    ['よくあるネガ', '切り返し方'],
    ['この地域では災害なんて、ここ何十年起きていないから',
     'そうですよね。実は札幌の方も、まったく同じことをおっしゃっていたそうです（→p.18へ）'],
    ['備蓄はしてあるから大丈夫',
     '食料は備えられますが、電気だけは備蓄できないんですよね。冷蔵庫もスマホも止まってしまいます'],
    ['停電なんてすぐ復旧するでしょう',
     '胆振東部地震では、復旧まで2週間かかった地域もありました'],
    ['保険に入っているから',
     '保険はあとから戻ってくるお金です。停電したその日の生活は、保険では戻せないんですよね'],
]
table(s, 0.45, 2.32, 9.95, [3.65, 6.30], rows, font_size=14, row_h=0.80, head_h=0.34)
conclusion(s, '否定せず、事実と事例で包み込む。議論にしない', y=6.35)

# ================================================================ 14. 災害トークの落とし穴
s = new_slide('必要性訴求②')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('災害トークの落とし穴：長く話すほど失注する', {'size': 24, 'bold': True, 'color': RED})])
_, tf = rect(s, 0.45, 1.98, 4.95, 3.55, fill='FCE4E4', line=RED)
lines(tf, [
    ('✕　やりすぎ例', {'size': 20, 'bold': True, 'color': RED, 'space': 10}),
    ('「南海トラフが来たら、この地域は震度7です。', {'size': 14, 'space': 3}),
    ('　ライフラインは1か月止まると言われています。', {'size': 14, 'space': 3}),
    ('　お子さんもいらっしゃいますよね。', {'size': 14, 'space': 3}),
    ('　真っ暗な中で、水も出ない状態で……」', {'size': 14, 'space': 12}),
    ('何が起きるか', {'size': 15, 'bold': True, 'space': 5}),
    ('・商談の雰囲気が暗くなる', {'size': 14, 'space': 4}),
    ('・「不安を煽られた」と警戒される', {'size': 14, 'space': 4}),
    ('・お客様が黙り、ラリーが止まる', {'size': 14, 'space': 4}),
    ('・そのあとの金額の話が入らなくなる', {'size': 14}),
])
_, tf = rect(s, 5.60, 1.98, 4.95, 3.55, fill=LTGREEN, line='70A040')
lines(tf, [
    ('○　適切な例', {'size': 20, 'bold': True, 'space': 10}),
    ('「札幌が地震の確率が最も低いと出ていたのに、', {'size': 14, 'space': 3}),
    ('　3か月後に大きな地震が来てしまったんですね。', {'size': 14, 'space': 3}),
    ('　気象庁も予測は難しいと言っています。', {'size': 14, 'space': 3}),
    ('　どこで起きるかは分からないですよね？」', {'size': 14, 'space': 12}),
    ('何が違うか', {'size': 15, 'bold': True, 'space': 5}),
    ('・主語が「事実」と「第三者」', {'size': 14, 'space': 4}),
    ('・お客様の家族を想像させていない', {'size': 14, 'space': 4}),
    ('・最後が質問で終わり、ラリーが続く', {'size': 14, 'space': 4}),
    ('・3分で次のページへ抜けられる', {'size': 14}),
])
conclusion(s, '災害は「事実 → 事例 → サラッと次へ」。3分で抜ける', y=5.70)
notes(s, '・講師はまず「やりすぎ例」を実演して見せる。受講者に違和感を体験させてから正解を示す')

# ================================================================ 15. ルール⑪-1
s = new_slide('必要性訴求②')
rule_head(s, 'ルール⑪：非常時対策の効果の確認', '災害トークは「第三者話法」でしか伝えない')
_, tf = rect(s, 0.45, 2.50, 4.95, 3.40, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('このルールに対応するページは', {'size': 18, 'bold': True, 'align': PP_ALIGN.CENTER, 'space': 3}),
    ('アプローチブックにありません', {'size': 18, 'bold': True, 'color': RED, 'align': PP_ALIGN.CENTER, 'space': 14}),
    ('なぜか？', {'size': 16, 'bold': True, 'space': 6}),
    ('「事例」でないと意味が無いからです。', {'size': 15, 'space': 8}),
    ('印刷された一般論を見せても、', {'size': 15, 'space': 3}),
    ('「うちの場合はどうなの？」に答えられません。', {'size': 15, 'space': 8}),
    ('だからこのパートは、', {'size': 15, 'space': 3}),
    ('営業マンが自分で用意した事例で語ります。', {'size': 15, 'bold': True}),
])
points_box(s, 5.60, 2.50, 4.95, 3.40, [
    ('⑩で「いつ起きてもおかしくない」を作った。次は「備えると、どうなるか」', {'bold': True, 'size': 13}),
    '被災経験のある知人に話を聞いてみる、メーカーHPの事例を読み込むなどして、上滑りしないように練習しておく',
    'なるべく地域の近い事例や、お客様との境遇が重なる資料を見せる工夫が必要',
    ('太陽光＋蓄電池があれば、向こう30年近く停電時の保証が得られる', {'bold': True, 'color': RED, 'size': 13}),
], title='考え方')
conclusion(s, '一般論ではなく「事例」。だからページが無い', y=6.10)

# ================================================================ 16. ルール⑪-2
s = new_slide('必要性訴求②')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('ルール⑪：使えるメーカー事例（第三者話法の材料）', {'size': 24, 'bold': True})])
box, tf = tb(s, 0.45, 1.88, 9.95, 0.35)
lines(tf, [('自社の施工事例が一番強いですが、無い場合は下記を読み込んでおきましょう',
            {'size': 15, 'color': GRAY})])
rows = [
    ['メーカー', '特徴', 'URL'],
    ['Panasonic', '動画＋テキストベースでの記事があり、使いやすい', 'https://sumai.panasonic.jp/chikuden/'],
    ['SmartStar', '災害時の活用に特化して動画が上がっている', 'https://www.smartstar.jp/voice/'],
    ['ニチコン', '販売店にフィットしたコンテンツ作りがされている',
     'https://www.nichicon.co.jp/products/ess/about/voice.html'],
]
table(s, 0.45, 2.32, 9.95, [1.75, 4.20, 4.00], rows, font_size=13, row_h=0.62, head_h=0.34)
_, tf = rect(s, 0.45, 4.50, 9.95, 1.75, fill=LTGREEN)
lines(tf, [
    ('第三者話法の型', {'size': 17, 'bold': True, 'space': 8}),
    ('✕　「停電しても安心ですよ」　← 自分の意見。刺さらない', {'size': 16, 'space': 6}),
    ('○　「実際に○○市のA様が、台風の停電で3日間、冷蔵庫と照明を動かせたそうです」',
     {'size': 16, 'bold': True, 'space': 4}),
    ('　　← 事例で返す', {'size': 14, 'color': GRAY}),
])
conclusion(s, '「私が〜」ではなく「実際に〜あります」で返す', y=6.40)

# ================================================================ 17. ルール⑪-3 ワーク
s = new_slide('必要性訴求②')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('ワーク：自分の「災害事例」を1本つくる（10分）', {'size': 24, 'bold': True})])
box, tf = tb(s, 0.45, 1.88, 9.95, 0.35)
lines(tf, [('商談で必ず使う1本を、いま書いてしまいましょう。地域が近いほど強くなります',
            {'size': 15, 'color': GRAY})])
rows = [
    ['', '項目', '記入してください'],
    ['①', '誰の事例か（地域・家族構成）', ''],
    ['②', 'いつ・どんな災害だったか', ''],
    ['③', '停電は何日続いたか', ''],
    ['④', '太陽光＋蓄電池で何ができたか', ''],
    ['⑤', 'その方は何とおっしゃっていたか', ''],
]
table(s, 0.45, 2.32, 9.95, [0.45, 3.15, 6.35], rows, font_size=13, row_h=0.68, head_h=0.32)
_, tf = rect(s, 0.45, 6.05, 9.95, 0.80, fill=LTYEL)
lines(tf, [
    ('書けたらペアで実演。聞き手は「3分以内に終わったか」「暗くならなかったか」を答える',
     {'size': 16, 'bold': True, 'align': PP_ALIGN.CENTER, 'space': 4}),
    ('自社に事例が無い人は、上記メーカーHPの事例を自分の言葉に置き換える', {'size': 14, 'align': PP_ALIGN.CENTER}),
])

# ================================================================ 18. 休憩
s = new_slide(None)
_, tf = rect(s, 2.40, 2.85, 6.05, 1.80, fill=LTGREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
lines(tf, [
    ('休憩　10分', {'size': 44, 'bold': True, 'align': PP_ALIGN.CENTER, 'space': 10}),
    ('後半はルール⑫「使い方」とロープレです', {'size': 18, 'align': PP_ALIGN.CENTER}),
])

# ================================================================ 19. ルール⑫-1
s = new_slide('必要性訴求②')
rule_head(s, 'ルール⑫：太陽光・蓄電池の使い方の理解', 'シミュレーションに移る前に基本概念をおさらいする')
pic(s, ab_img(20), 0.45, 2.50, 4.55)
box, tf = tb(s, 0.45, 5.85, 4.55, 0.3)
lines(tf, [('アプローチブック p.20　電気を自給自足する生活へ', {'size': 12, 'color': GRAY})])
goal_box(s, 5.35, 2.50, 5.05, 1.15, ['電気を買うより、自給自足した方が良いですね'])
points_box(s, 5.35, 3.78, 5.05, 2.50, [
    ('電気代上昇（⑦⑧⑨）と災害対策（⑩⑪）を受けての結論のページ', {'bold': True, 'size': 13}),
    'これからは買うよりも自給自足の生活をした方がお得で安全であること',
    ('太陽光でつくることだけでなく、蓄電池で貯めることも必要であること', {'bold': True, 'size': 13}),
    ('ここでセットである理由が経済合理性からも説明できる（売電15円＜買電30円）', {'bold': True, 'color': RED, 'size': 13}),
])
conclusion(s, '上昇する電気代を削減するために、自給自足が求められます')

# ================================================================ 20. ルール⑫-2 1日の流れ
s = new_slide('必要性訴求②')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('ルール⑫：1日の電気の流れ（これは図なしで言えるように）', {'size': 24, 'bold': True})])
blocks = [
    (0.45, '朝', '発電量が少ない時間帯', ['電力会社から電気を購入する', '（買電消費）'], LTBLUE),
    (3.75, '昼', '太陽光発電が活発な時間帯', ['① 太陽光の電気をまず家庭で使う（自家消費）',
                                     '② 余った電気は蓄電池に充電する', '③ さらに余った電気は売電する'], LTYEL),
    (7.05, '夕方〜夜間', '発電できない時間帯', ['蓄電池に貯めた電気を使う'], LTGREEN),
]
for x, t, sub, items, fill in blocks:
    _, tf = rect(s, x, 1.98, 3.10, 2.45, fill=fill)
    lines(tf, [
        (t, {'size': 22, 'bold': True, 'align': PP_ALIGN.CENTER, 'space': 4}),
        (sub, {'size': 12, 'color': GRAY, 'align': PP_ALIGN.CENTER, 'space': 10}),
    ])
    for it in items:
        p = para(tf, space_after=6)
        run(p, it, size=13)
for x in (3.58, 6.88):
    a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(2.95), Inches(0.30), Inches(0.35))
    a.fill.solid(); a.fill.fore_color.rgb = RGBColor.from_string('808080')
    a.line.fill.background(); a.shadow.inherit = False
_, tf = rect(s, 0.45, 4.60, 9.95, 1.65, fill=LTGRAY)
lines(tf, [
    ('このサイクルは、太陽光と蓄電池が「両方」あって初めて成立します',
     {'size': 18, 'bold': True, 'color': RED, 'space': 8}),
    ('・太陽光だけ　→　昼に余った電気は売るしかない。売電15円＜買電30円なので損', {'size': 15, 'space': 5}),
    ('・蓄電池だけ　→　貯める電気を電力会社から買うことになる', {'size': 15, 'space': 5}),
    ('・セット　　　→　自分で創った電気を、自分で使い切れる', {'size': 15, 'bold': True}),
])
conclusion(s, '創る（太陽光）→ 貯める（蓄電池）→ 夜に使う', y=6.40)
notes(s, '・受講者に図を見ずに説明させる。ここが言えないとシミュレーションの説明が崩れる')

# ================================================================ 21. ルール⑫-3
s = new_slide('必要性訴求②')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('ルール⑫：シミュレーションへの受け渡し', {'size': 24, 'bold': True})])
pic(s, ab_img(21), 0.45, 1.95, 4.55)
box, tf = tb(s, 0.45, 5.30, 4.55, 0.3)
lines(tf, [('アプローチブック p.21　太陽光発電・蓄電池の使い方', {'size': 12, 'color': GRAY})])
_, tf = rect(s, 5.35, 1.95, 5.05, 1.55, fill=LTGRAY)
lines(tf, [
    ('①　ライフスタイルと結びつける', {'size': 15, 'bold': True, 'space': 6}),
    ('「ご家庭によって電気を使う時間帯や量は異なりますので、このサイクルをいかにご家庭の生活に合わせて最適化するかがポイントになります」',
     {'size': 13}),
])
_, tf = rect(s, 5.35, 3.62, 5.05, 1.45, fill=LTGRAY)
lines(tf, [
    ('②　理解度を確認する', {'size': 15, 'bold': True, 'space': 6}),
    ('「ここまでで、何かご不明な点はございますか？」', {'size': 13, 'space': 4}),
    ('「電気の流れのイメージは掴んでいただけましたでしょうか？」', {'size': 13}),
])
_, tf = rect(s, 5.35, 5.20, 5.05, 1.05, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('③　そのまま次へ', {'size': 15, 'bold': True, 'space': 6}),
    ('「具体的なシミュレーションを見ていきましょう！」', {'size': 14, 'bold': True}),
])
conclusion(s, 'p.21下部の一文をそのまま使い、シミュレーションへ渡す')

# ================================================================ 22. Paragraph02 通しの地図
s = new_slide('Paragraph 02 の通し')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('p.9〜p.21をひと続きで見る（第2回＋第3回）', {'size': 24, 'bold': True})])
rows = [
    ['ページ', '内容', 'ルール', '狙う反応'],
    ['p.9・10・12', '電気料金の推移／電力自由化／全国的な値上げ', '⑦', '10年で5.5万円も上がっている'],
    ['p.11・13〜16', '電気代の内訳／再エネ賦課金／燃料調整費', '⑧', '賦課金で電気代が上がる'],
    ['p.17', '生涯に支払う電気代', '⑨', '30年でそんなにも払うのですね'],
    [('p.18・19', {'bold': True}), ('地震予測／停電被害の実績', {'bold': True}), ('⑩', {'bold': True}),
     ('いつ災害が起こるか分からない', {'bold': True})],
    [('（ページなし）', {'bold': True}), ('メーカー事例・自社事例', {'bold': True}), ('⑪', {'bold': True}),
     ('停電時も電気が使えた方が安心', {'bold': True})],
    [('p.20・21', {'bold': True}), ('自給自足へ／太陽光・蓄電池の使い方', {'bold': True}), ('⑫', {'bold': True}),
     ('買うより自給自足した方が良い', {'bold': True})],
]
for i in (4, 5, 6):
    for j in range(4):
        t, o = rows[i][j]; o['fill'] = LTYEL; rows[i][j] = (t, o)
table(s, 0.45, 1.95, 9.95, [1.60, 3.55, 0.80, 4.00], rows, font_size=12, row_h=0.62, head_h=0.32)
_, tf = rect(s, 0.45, 6.05, 9.95, 0.80, fill=LTGREEN)
lines(tf, [
    ('黄色が本日追加した3ルール。これでParagraph 02が最後まで通ります', {'size': 17, 'bold': True,
                                                     'align': PP_ALIGN.CENTER, 'space': 4}),
    ('次回はParagraph 03（p.23〜26）＝時期訴求・金額訴求です', {'size': 14, 'align': PP_ALIGN.CENTER}),
])

# ================================================================ 23. ロープレ①
s = new_slide('ロールプレイング①')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('1周目：p.18〜p.21だけを通す（7分）', {'size': 24, 'bold': True})])
_, tf = rect(s, 0.45, 1.95, 4.95, 2.35, fill=LTGREEN)
lines(tf, [
    ('やること', {'size': 17, 'bold': True, 'space': 8}),
    ('① 2人1組（営業／お客様役）', {'size': 15, 'space': 6}),
    ('② p.18から始めて、p.21の「シミュレーションを見ていきましょう」で終わる', {'size': 15, 'space': 6}),
    ('③ 3分で交代。全員が営業役をやる', {'size': 15}),
])
_, tf = rect(s, 5.60, 1.95, 4.95, 2.35, fill=LTYEL)
lines(tf, [
    ('この周で見るポイント', {'size': 17, 'bold': True, 'space': 8}),
    ('・災害パートが3分以内に終わったか', {'size': 15, 'bold': True, 'color': RED, 'space': 6}),
    ('・第三者話法だけで語れたか（自分の意見を混ぜていないか）', {'size': 15, 'space': 6}),
    ('・p.21の受け渡しの一言まで言い切れたか', {'size': 15}),
])
_, tf = rect(s, 0.45, 4.55, 9.95, 1.55, fill=LTGRAY)
lines(tf, [
    ('お客様役はこの断り文句を必ず1回入れてください', {'size': 17, 'bold': True, 'space': 8}),
    ('「うちの地域、災害なんてここ何十年も起きてないですけどね」', {'size': 17, 'bold': True, 'color': RED, 'space': 6}),
    ('→ 営業役はスライド13の切り返しで対応する', {'size': 15}),
])
conclusion(s, '長く話さない。事実 → 事例 → サラッと次へ', y=6.30)

# ================================================================ 24. 顧客設定
s = new_slide('ロールプレイング')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('顧客設定（第2回と同じ田中様ご夫婦）', {'size': 24, 'bold': True})])
rows = [
    ['項目', '設定'],
    ['ご家族', '田中様ご夫婦（ご主人48歳・会社員／奥様45歳・パート）＋ 高校生・中学生のお子様2人'],
    ['お住まい', '築14年の戸建て（4LDK）。南向き切妻屋根。オール電化ではない'],
    ['導入状況', ('太陽光・蓄電池ともに未導入', {'bold': True, 'color': RED})],
    ['電気代', '月平均 18,000円（夏場は24,000円）。検針票はご主人が保管している'],
    ['温度感', 'ご主人＝慎重。「元が取れるのか」が最大の関心。奥様＝前向き。停電と電気代が心配'],
    ['本日の追加設定', ('停電の経験は一度もない。防災用品は備蓄しているが、電気の備えは考えたことがない',
                  {'bold': True, 'color': RED})],
    ['断り文句', '「うちの地域、災害なんてここ何十年も起きてないですけどね」'],
]
table(s, 0.45, 1.92, 9.95, [1.85, 8.10], rows, font_size=12.5, row_h=0.54, head_h=0.32)
_, tf = rect(s, 0.45, 6.35, 9.95, 0.56, fill=LTYEL)
lines(tf, [('奥様は停電に前向き、ご主人は懐疑的。夫婦の温度差をどう埋めるかも見ています',
            {'size': 17, 'bold': True, 'align': PP_ALIGN.CENTER})])
notes(s, '・夫婦の温度差は天秤。一方が盛り上がりすぎるともう一方が冷める。弱い方をフォローさせる')

# ================================================================ 25. ロープレ②
s = new_slide('ロールプレイング②')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('2周目：Paragraph 02を通しで（8分）', {'size': 24, 'bold': True})])
_, tf = rect(s, 0.45, 1.95, 4.95, 2.20, fill=LTGREEN)
lines(tf, [
    ('やること', {'size': 17, 'bold': True, 'space': 8}),
    ('① 3人1組（営業／ご主人役／奥様役）', {'size': 15, 'space': 6}),
    ('② p.9から p.21まで、止まらずに通す', {'size': 15, 'space': 6}),
    ('③ 4分で交代', {'size': 15}),
])
_, tf = rect(s, 5.60, 1.95, 4.95, 2.20, fill=LTYEL)
lines(tf, [
    ('この周で見るポイント', {'size': 17, 'bold': True, 'space': 8}),
    ('・発言比率は 営業6：お客様4 か', {'size': 15, 'space': 6}),
    ('・お客様側の4は、ご主人2：奥様2 か', {'size': 15, 'space': 6}),
    ('・最後までセット提案を崩していないか', {'size': 15, 'bold': True, 'color': RED}),
])
rows = [
    ['配分', 'ページ', '時間の目安'],
    ['電気代のネガ（⑦⑧⑨）', 'p.9〜p.17', '2分'],
    [('災害（⑩⑪）', {'bold': True, 'color': RED}), ('p.18〜p.19＋事例', {'bold': True}),
     ('1分　※ここを伸ばさない', {'bold': True, 'color': RED})],
    ['使い方（⑫）', 'p.20〜p.21', '1分'],
]
table(s, 0.45, 4.40, 9.95, [3.55, 3.20, 3.20], rows, font_size=14, row_h=0.58, head_h=0.34)
conclusion(s, '4分のうち、災害に使ってよいのは1分だけ', y=6.35)

# ================================================================ 26. チェックシート
s = new_slide('チェックシート')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.42)
lines(tf, [('ロープレのあと、自分で○△×をつけてください', {'size': 22, 'bold': True})])
rows = [
    ['ルール', '該当ページ', '狙った台詞を引き出せたか', '結論を言えたか'],
    ['⑩非常時に備える意義', 'p.18', 'いつ災害が起こるか分からないですね', '専門家ですら予測できない'],
    ['', 'p.19', '（規模を実感した反応）', '災害大国だからこそ備える必要がある'],
    ['⑪非常時対策の効果', '（ページなし）', '停電時も電気が使えた方が安心です', '「私が」ではなく「実際に」で返せた'],
    ['⑫太陽光・蓄電池の使い方', 'p.20', '買うより自給自足した方が良いですね', '自給自足が求められます'],
    ['', 'p.21', '活用方法が分かる', 'シミュレーションを見ていきましょう'],
    [('時間管理', {'bold': True}), '−', ('災害パートが3分以内で終わった', {'bold': True}), '−'],
    [('話法', {'bold': True}), '−', ('第三者話法だけで語れた（自分の意見を混ぜない）', {'bold': True}), '−'],
    [('商談運営', {'bold': True}), '−', '発言比率6：4／ご主人2：奥様2', 'セット提案を崩さなかった'],
]
table(s, 0.30, 1.90, 10.25, [2.45, 1.45, 3.55, 2.80], rows, font_size=11.5, row_h=0.48, head_h=0.32)
_, tf = rect(s, 0.30, 6.20, 10.25, 0.62, fill=LTGREEN)
lines(tf, [('「結論を言えたか」の欄は、資料を閉じて声に出して確認します（暗唱チェック）',
            {'size': 17, 'bold': True, 'align': PP_ALIGN.CENTER})])

# ================================================================ 27. フィードバックの型
s = new_slide('フィードバックの型')
box, tf = tb(s, 0.45, 1.38, 9.95, 0.45)
lines(tf, [('お互いにフィードバックするときの言い方', {'size': 24, 'bold': True})])
box, tf = tb(s, 0.45, 1.88, 9.95, 0.35)
lines(tf, [('褒め方は商談でもそのまま使えます。「質の高い褒め」を「数」多く言うのが基本です',
            {'size': 15, 'color': GRAY})])
_, tf = rect(s, 0.45, 2.32, 4.95, 1.95, fill=LTGREEN)
lines(tf, [
    ('型①　比較して褒める', {'size': 18, 'bold': True, 'space': 8}),
    ('✕「良かったです」', {'size': 15, 'space': 5}),
    ('○「1周目は災害で5分使っていましたが、2周目は2分半に収まっていました」',
     {'size': 15, 'bold': True}),
])
_, tf = rect(s, 5.60, 2.32, 4.95, 1.95, fill=LTGREEN)
lines(tf, [
    ('型②　人を褒める', {'size': 18, 'bold': True, 'space': 8}),
    ('✕「事例が良かったです」', {'size': 15, 'space': 5}),
    ('○「あの事例を自分で用意してきた○○さんが凄いですね」', {'size': 15, 'bold': True}),
])
_, tf = rect(s, 0.45, 4.50, 9.95, 1.75, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('直す点は「1人1つだけ」', {'size': 18, 'bold': True, 'space': 8}),
    ('一度に3つも4つも指摘すると、どれも直りません。', {'size': 15, 'space': 5}),
    ('その人が次の商談で一番効く1つに絞って伝えてください。', {'size': 15, 'space': 5}),
    ('順番は必ず「良い点 → 直す点1つ」。逆にしない。', {'size': 16, 'bold': True, 'color': RED}),
])
conclusion(s, '“質”の高い褒めを“数”多く言うことが重要である', y=6.40)

# ================================================================ 28. まとめ・宿題
s = new_slide('本日のまとめ')
_, tf = rect(s, 0.45, 1.45, 9.95, 2.30, fill=LTGREEN)
lines(tf, [
    ('本日持ち帰ること', {'size': 18, 'bold': True, 'space': 10}),
    ('① 災害は「事実 → 事例 → サラッと次へ」。3分で抜ける', {'size': 17, 'space': 8}),
    ('② ⑪にページが無いのは、一般論では意味が無いから。事例は自分で用意する', {'size': 17, 'space': 8}),
    ('③ 創る → 貯める → 夜に使う。このサイクルはセットでしか成立しない', {'size': 17, 'space': 8}),
    ('④ p.21の一言でシミュレーションに受け渡す', {'size': 17, 'bold': True, 'color': RED}),
])
_, tf = rect(s, 0.45, 4.00, 9.95, 2.30, fill=LTYEL, line='BF8F00')
lines(tf, [
    ('宿題（次回までに）', {'size': 18, 'bold': True, 'space': 10}),
    ('① 本日つくった災害事例を清書し、90秒で言えるようにしてくること', {'size': 17, 'space': 8}),
    ('② アプローチブック p.9〜p.21 の結論の一文を、資料を見ずに通しで言えるようにしてくること', {'size': 17, 'space': 8}),
    ('③ 実商談を1件、災害パートに何分使ったか計測してくること', {'size': 17}),
])
_, tf = rect(s, 0.45, 6.55, 9.95, 0.42, fill=None, line=RED, lw=1.5)
lines(tf, [('次回・第4回は「時期訴求・料金訴求（ルール⑬〜⑳）」／アプローチブック p.23〜p.26',
            {'size': 16, 'bold': True, 'align': PP_ALIGN.CENTER})])
notes(s, '・その日のうちに復習すると定着する（翌日には50％忘れる）ことを添える')

# ---------------------------------------------------------------- 並べ替え
ids = list(sldIdLst)
colophon = [x for x in ids if x is colophon_id][0]
sldIdLst.remove(colophon)
sldIdLst.append(colophon)
page_no[0] += 1

os.makedirs(os.path.dirname(OUT), exist_ok=True)
prs.save(OUT)
print('saved:', OUT, '/ slides:', len(prs.slides.__iter__.__self__._sldIdLst))
