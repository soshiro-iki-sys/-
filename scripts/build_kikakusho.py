# -*- coding: utf-8 -*-
"""企画書 空パッケージ ジェネレータ
prompts/企画書_空パッケージ_生成プロンプト.md の仕様に完全準拠。
"""
import copy, os
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn

# ---------- パレット ----------
NAVY  = RGBColor(0x00, 0x20, 0x60)
INK   = RGBColor(0x1F, 0x1F, 0x1F)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED   = RGBColor(0xFF, 0x00, 0x00)
YEL   = RGBColor(0xFF, 0xFF, 0x00)
PALE  = RGBColor(0xE9, 0xEE, 0xF6)

FONT = 'Meiryo UI'
LOGO = 'templates/assets/funai_logo.png'
COPY = 'Copyright Funai Consulting Inc. All Rights Reserved.'

# ---------- キャンバス ----------
W, H = 27.52, 19.05
CT, CB = 3.90, 15.20          # 本文エリア上下端
LM, CW = 0.19, 27.09          # 左マージン／全幅

prs = Presentation()
prs.slide_width  = Emu(9906000)
prs.slide_height = Emu(6858000)
BLANK = prs.slide_layouts[6]


# ---------- 低レベルヘルパ ----------
def style(run, size=18, bold=True, color=INK, font=FONT):
    f = run.font
    f.size = Pt(size); f.bold = bold
    f.color.rgb = color
    f.name = font
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {}); rPr.append(e)
        e.set('typeface', font)


def para(tf, parts, align=PP_ALIGN.LEFT, spc=None, first=False):
    """parts: [(text, size, color)] / [(text, size, color, bold)] / [(text,size,color,bold,font)]"""
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    if spc:
        p.line_spacing = spc
    for t in parts:
        text, size, color = t[0], t[1], t[2]
        bold = t[3] if len(t) > 3 else True
        font = t[4] if len(t) > 4 else FONT
        r = p.add_run(); r.text = text
        style(r, size, bold, color, font)
    return p


def textbox(s, x, y, w, h, anchor=MSO_ANCHOR.MIDDLE):
    tb = s.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    tf.auto_size = MSO_AUTO_SIZE.NONE
    return tf


def shape(s, x, y, w, h, fill=None, line=None, lw=1.0, kind=MSO_SHAPE.RECTANGLE):
    sh = s.shapes.add_shape(kind, Cm(x), Cm(y), Cm(w), Cm(h))
    st = sh.element.find(qn('p:style'))
    if st is not None:
        sh.element.remove(st)            # テーマ既定の見た目を外す
    sh.shadow.inherit = False            # 影なし
    if fill is not None:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    else:
        sh.fill.background()
    if line is not None:
        sh.line.color.rgb = line; sh.line.width = Pt(lw)
    else:
        sh.line.fill.background()
    tf = sh.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Cm(0.13)
    tf.margin_top = tf.margin_bottom = Cm(0.05)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    return sh


def arrow(s, x, y, w, h, kind, fill=RED):
    return shape(s, x, y, w, h, fill=fill, line=None, kind=kind)


# ---------- 共通 chrome ----------
def footer(s, page, strip=True):
    if strip:
        shape(s, 0, 18.55, W, 0.50, fill=NAVY)
    tf = textbox(s, 10.55, 18.31, 6.42, 1.01)
    para(tf, [(str(page), 9, WHITE, False)], PP_ALIGN.CENTER, first=True)
    tf = textbox(s, 18.36, 18.47, 9.00, 0.60)
    para(tf, [(COPY, 8, WHITE, False, 'Arial')], PP_ALIGN.RIGHT, first=True)


def header(s, title, tag):
    tf = textbox(s, 0.56, 0.19, 12.78, 1.53)
    para(tf, [(title, 28, INK)], PP_ALIGN.LEFT, first=True)
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(0), Cm(1.72), Cm(W), Cm(1.72))
    c.line.color.rgb = NAVY; c.line.width = Pt(2.25)
    s.shapes.add_picture(LOGO, Cm(24.56), Cm(0.19), Cm(2.72), Cm(1.40))
    sh = shape(s, LM, 1.85, 7.20, 0.81, fill=NAVY)
    para(sh.text_frame, [(tag, 16, WHITE)], PP_ALIGN.CENTER, first=True)


def lead(s, parts, h=1.10):
    tf = textbox(s, 0.38, 2.78, 26.76, h, MSO_ANCHOR.TOP)
    para(tf, parts, PP_ALIGN.LEFT, 1.0, first=True)


def band(s, parts):
    shape(s, 0, 15.93, W, 2.40, fill=NAVY)
    tf = textbox(s, 0, 15.93, W, 2.40)
    para(tf, parts, PP_ALIGN.CENTER, 0.9, first=True)


def note(s, owner, pending, basis='自社実績 or 業界ベンチマーク（要記入）'):
    tf = textbox(s, 0.38, 15.32, 26.76, 0.56)
    para(tf, [('主担当：%s　／　確定待ち：%s　／　根拠：%s' % (owner, pending, basis), 8, INK)],
         PP_ALIGN.LEFT, first=True)


def notes(s, text):
    s.notes_slide.notes_text_frame.text = text


# ---------- 高レベル部品 ----------
def card(s, x, y, w, hh, bh, head, body, body_fill=PALE, head_size=20,
         body_size=16, body_align=PP_ALIGN.LEFT, body_spc=1.3, lw=1.5):
    """head: [(text,size,color)] / body: [[(text,size,color)], ...]"""
    sh = shape(s, x, y, w, hh, fill=NAVY, line=NAVY, lw=lw)
    para(sh.text_frame, head, PP_ALIGN.CENTER, first=True)
    sb = shape(s, x, y + hh, w, bh, fill=body_fill, line=NAVY, lw=lw)
    for i, line_parts in enumerate(body):
        para(sb.text_frame, line_parts, body_align, body_spc, first=(i == 0))
    return sh, sb


def plain(s, x, y, w, h, lines, fill=None, line=NAVY, lw=1.5,
          align=PP_ALIGN.CENTER, spc=1.25):
    sh = shape(s, x, y, w, h, fill=fill, line=line, lw=lw)
    for i, parts in enumerate(lines):
        para(sh.text_frame, parts, align, spc, first=(i == 0))
    return sh


# ---------- スライド種別 ----------
SLIDES = []
def slide_blank():
    s = prs.slides.add_slide(BLANK)
    SLIDES.append({'s': s, 'strip': True})
    return s


def cover():
    s = slide_blank()
    shape(s, 0, 0, W, H, fill=NAVY)
    tf = textbox(s, 1.96, 4.11, 7.75, 6.60)
    para(tf, [('01', 144, WHITE)], PP_ALIGN.CENTER, first=True)
    shape(s, 3.22, 10.00, 5.22, 0.17, fill=WHITE)
    tf = textbox(s, 8.89, 5.55, 17.50, 1.20)
    para(tf, [('[ 企画書 ]　', 28, WHITE)], PP_ALIGN.LEFT, first=True)
    tf = textbox(s, 8.89, 6.93, 17.64, 3.60)
    para(tf, [('B2B2C提携による', 36, WHITE)], PP_ALIGN.LEFT, 1.15, first=True)
    para(tf, [('名簿獲得戦略', 36, WHITE)], PP_ALIGN.LEFT, 1.15)
    shape(s, 0.33, 15.61, 26.85, 2.50, fill=WHITE)
    s.shapes.add_picture(LOGO, Cm(0.89), Cm(16.25), Cm(3.19), Cm(1.64))
    tf = textbox(s, 9.31, 16.07, 17.64, 1.58)
    para(tf, [('株式会社 船井総合研究所', 15, INK)], PP_ALIGN.RIGHT, 1.0, first=True)
    para(tf, [('◯◯部　◯◯チーム　◯◯　　', 15, INK), ('◯◯　◯◯', 22, INK)],
         PP_ALIGN.RIGHT, 1.0)
    SLIDES[-1]['strip'] = False
    notes(s, '表紙。作成者名・部署・作成日は確定後に差し替える。')
    return s


def toc(active):
    s = slide_blank()
    header(s, '目次', '本日のアジェンダ')
    items = [
        ('①', 'Ⅰ. 結論　～ 何を決めていただきたいか ～'),
        ('②', 'Ⅱ. 戦略　～ センターピンはどこか ～'),
        ('③', 'Ⅲ. 施策　～ どう獲り、どう育て、どう刈るか ～'),
        ('④', 'Ⅳ. 数値　～ いくら投じ、いくら返るか ～'),
        ('⑤', 'Ⅴ. 実行　～ 誰が、いつまでに ～'),
    ]
    y = 4.03
    for i, (num, label) in enumerate(items):
        col = RED if i == active else INK
        tf = textbox(s, 1.67, y, 1.94, 1.40)
        para(tf, [(num, 32, col)], PP_ALIGN.CENTER, first=True)
        tf = textbox(s, 3.89, y, 21.96, 1.40)
        para(tf, [(label, 32, col)], PP_ALIGN.LEFT, first=True)
        y += 1.94
    notes(s, 'これから話す章を赤で示す。章が進むたびに再掲する。')
    return s


def interlude(lines, size=60):
    s = slide_blank()
    shape(s, 0, 0, W, H, fill=NAVY)
    n = len(lines)
    lh = size * 0.0353 * 0.95 * 1.24
    h = lh * n + 0.60
    y = (H - h) / 2 - 0.7
    tf = textbox(s, 1.39, y, 24.74, h)
    for i, t in enumerate(lines):
        para(tf, [(t, size, WHITE)], PP_ALIGN.CENTER, 0.95, first=(i == 0))
    SLIDES[-1]['strip'] = False
    return s


# ================= 本編 =================
# --- 1 表紙 ---
cover()

# --- 2 目次 ---
toc(0)

# --- 3 中扉 Ⅰ ---
interlude(['まず、', '結論から。'], 66)

# --- 4 エグゼクティブサマリー ---
s = slide_blank(); header(s, 'Ⅰ. 結論', 'エグゼクティブサマリー')
lead(s, [('■　提携【　　】社 × 名簿【　　　】件 → 粗利', 20, INK),
         ('【　　　　】円', 24, RED)])
cols = [(LM, '① 提携社数', '【　　】社'),
        (9.47, '② 1社あたり名簿数', '【　　　】件'),
        (18.74, '③ 総名簿数（センターピン）', '【　　　　】件')]
for x, hd, val in cols:
    card(s, x, 4.05, 8.54, 1.40, 2.90, [(hd, 18, WHITE)], [[(val, 32, RED)]],
         body_align=PP_ALIGN.CENTER)
card(s, LM, 8.75, CW, 1.30, 2.40, [('KGI｜粗利額', 20, YEL)],
     [[('粗利 ＝ 総名簿数【　　　　】件 × アポ率【　】％ × 商談化率【　】％', 20, INK)],
      [('　　　 × 成約率【　】％ × 平均単価【　　　】万円 × 粗利率【　】％', 20, INK)]],
     body_fill=WHITE, body_align=PP_ALIGN.CENTER, body_spc=1.2)
plain(s, LM, 12.55, CW, 2.65,
      [[('この企画で狙うのは粗利', 22, INK), ('【　　　　】円', 22, RED), ('。', 22, INK)],
       [('その全ては“総名簿数”という1つの数字に依存する。', 22, INK)]], fill=PALE, spc=1.2)
note(s, '起案者', '全数値（他章の確定後に最後に執筆）', '—')
band(s, [('追うべき数字は、', 24, WHITE), ('ひとつ', 32, YEL), ('。総名簿数です', 24, WHITE)])
notes(s, 'KGI＝粗利額。全数値が確定してから最後に執筆する。主担当：起案者。')

# --- 5 決裁依頼事項 ---
s = slide_blank(); header(s, 'Ⅰ. 結論', '決裁依頼事項')
lead(s, [('■　ご承認いただきたいのは、次の', 20, INK), ('3点', 24, RED), ('です', 20, INK)])
cw5 = [(0.19, 6.00), (6.19, 11.50), (17.69, 5.50), (23.19, 4.09)]
hdr = ['承認項目', '内　容', '金額・数値', '決裁区分']
for (x, w), t in zip(cw5, hdr):
    plain(s, x, 4.10, w, 1.10, [[(t, 16, WHITE)]], fill=NAVY, line=NAVY, lw=1.0)
rows = [('① 予算総額', '広告費／ハガキ発送費／セミナー運営費／ツール費', '【　　　】万円'),
        ('② 人　員', '訪問営業 ◯名／マーケ ◯名／事務代行 ◯名', '【　】名'),
        ('③ メンテ手数料の還元率', '提携先へ還元する手数料の料率', '【　】％')]
y = 5.20
for a, b, c in rows:
    plain(s, 0.19, y, 6.00, 2.40, [[(a, 16, NAVY)]], fill=PALE, line=NAVY, lw=1.0)
    plain(s, 6.19, y, 11.50, 2.40, [[(b, 16, INK)]], fill=WHITE, line=NAVY, lw=1.0,
          align=PP_ALIGN.LEFT)
    plain(s, 17.69, y, 5.50, 2.40, [[(c, 20, RED)]], fill=WHITE, line=NAVY, lw=1.0)
    plain(s, 23.19, y, 4.09, 2.40, [[('要決裁', 16, INK)]], fill=WHITE, line=NAVY, lw=1.0)
    y += 2.40
plain(s, LM, 12.75, CW, 2.45,
      [[('投資総額', 24, INK), ('【　　　】万円', 24, RED), ('　／　回収目標', 24, INK),
        ('【　　】ヶ月', 24, RED)]], fill=PALE)
note(s, '役員', 'メンテ手数料の還元率、投資総額')
band(s, [('この', 24, WHITE), ('3点', 32, YEL), ('のご承認をお願いします', 24, WHITE)])
notes(s, '投資総額＝固定費＋変動費。第14章の損益分岐点と数字を必ず一致させる。主担当：役員。')

# --- 6 中扉 Ⅱ ---
interlude(['個人を追うのを、', 'やめる。'], 60)

# --- 7 背景と課題 ---
s = slide_blank(); header(s, 'Ⅱ. 戦略', '背景と課題')
lead(s, [('■　個人向け集客はもう伸びない。だから、', 20, INK),
         ('名簿を持つ企業と組む', 24, RED)])
card(s, 0.29, 4.10, 13.17, 1.60, 5.20, [('これまで｜個人向け集客', 24, WHITE)],
     [[('・獲得単価が高騰（直近CPA【　　　】円）', 18, INK)],
      [('・リードが枯渇し、母数が増えない', 18, INK)],
      [('・広告を止めれば、受注も止まる', 18, INK)]], body_size=18)
card(s, 14.06, 4.10, 13.17, 1.60, 5.20, [('これから｜B2B2C提携', 24, WHITE)],
     [[('・提携先が既に持つ名簿にアクセス', 18, INK)],
      [('・1社の提携で【　　　】件がまとめて増える', 18, INK)],
      [('・獲得単価は「提携コスト」に置き換わる', 18, INK)]], body_size=18)
plain(s, LM, 11.20, CW, 4.00,
      [[('1件ずつ獲るのをやめ、', 28, INK)],
       [('【　　　】件', 28, RED), ('をまとめて獲りにいく', 28, INK)]], fill=PALE)
note(s, '経営企画', '自社の直近CPA実績、1社あたり想定名簿数')
band(s, [('限界なのは努力ではなく、', 24, WHITE), ('打ち手の構造', 32, YEL), ('です', 24, WHITE)])
notes(s, '現状のリード獲得単価を自社実績から出す。CPA高騰の根拠データを添付する。主担当：経営企画。')

# --- 8 センターピンの定義 ---
s = slide_blank(); header(s, 'Ⅱ. 戦略', 'センターピンの定義')
lead(s, [('■　勝敗を決めるのは、たった1つの数字＝', 20, INK), ('総名簿数', 24, RED)])
sh = shape(s, LM, 4.20, CW, 1.50, fill=NAVY, line=NAVY, lw=1.5)
para(sh.text_frame, [('セ ン タ ー ピ ン', 24, YEL)], PP_ALIGN.CENTER, first=True)
plain(s, LM, 5.70, CW, 3.20,
      [[('総名簿数 ＝ 提携社数', 28, INK), ('【　　】社', 28, RED)],
       [('× 1社あたり名簿数', 28, INK), ('【　　　】件', 28, RED)]],
      fill=WHITE, lw=2.5, spc=1.15)
decl = [(LM, '① 目的を一本化', ['訪問もDLもメルマガも、', '目的は提携社数']),
        (9.47, '② 提携後は必ず回収', ['提携＝ゴールではない。', '名簿の回収までが1件']),
        (18.74, '③ 他のKPIは追わない', ['名簿数に効かない施策は、', 'やらないと決める'])]
for x, hd, lines in decl:
    card(s, x, 9.30, 8.54, 1.40, 3.10, [(hd, 20, WHITE)],
         [[(t, 16, INK)] for t in lines], body_align=PP_ALIGN.CENTER, body_spc=1.35)
plain(s, LM, 14.05, CW, 1.15,
      [[('この式の左辺を最大化すること以外は、やらない。', 20, INK)]], fill=PALE, lw=1.0)
note(s, '起案者', '目標提携社数、1社あたり想定名簿数')
band(s, [('全リソースを、', 24, WHITE), ('総名簿数', 32, YEL), ('に集中させます', 24, WHITE)])
notes(s, '総名簿数 = 提携社数 × 1社あたり名簿数。未確定：目標提携社数、1社あたり想定名簿数。主担当：起案者。')

# --- 9 KPIツリー ---
s = slide_blank(); header(s, 'Ⅱ. 戦略', 'KPIツリー')
lead(s, [('■　粗利（KGI）は、名簿数（中間KPI）と施策KPIに', 20, INK),
         ('分解できる', 24, RED)])
plain(s, LM, 4.05, CW, 1.50,
      [[('KGI｜粗利　', 24, WHITE), ('【　　　　】円', 24, YEL)]], fill=NAVY, line=NAVY, lw=1.5)
arrow(s, 13.26, 5.70, 1.00, 0.90, MSO_SHAPE.DOWN_ARROW, NAVY)
plain(s, LM, 6.75, CW, 1.50,
      [[('中間KPI｜総名簿数　', 24, WHITE), ('【　　　　】件', 24, YEL)]],
      fill=NAVY, line=NAVY, lw=1.5)
arrow(s, 13.26, 8.40, 1.00, 0.90, MSO_SHAPE.DOWN_ARROW, NAVY)
kpi = [(0.19, '施策KPI ①\n訪問', ['提携社数', '【　　】社']),
       (7.13, '施策KPI ②\nDLレポート', ['獲得リード数', '【　　　】件']),
       (14.07, '施策KPI ③\n育成・刈り取り', ['商談化率', '【　】％']),
       (21.01, '施策KPI ④\n名簿回収', ['名簿提出率', '【　】％'])]
for x, hd, lines in kpi:
    sh = shape(s, x, 9.45, 6.24, 1.55, fill=NAVY, line=NAVY, lw=1.0)
    for i, t in enumerate(hd.split('\n')):
        para(sh.text_frame, [(t, 14, WHITE)], PP_ALIGN.CENTER, 1.1, first=(i == 0))
    sb = shape(s, x, 11.00, 6.24, 2.60, fill=PALE, line=NAVY, lw=1.0)
    para(sb.text_frame, [(lines[0], 16, INK)], PP_ALIGN.CENTER, 1.25, first=True)
    para(sb.text_frame, [(lines[1], 20, RED)], PP_ALIGN.CENTER, 1.25)
plain(s, LM, 13.85, CW, 1.35,
      [[('下の4つを動かせば、上の2つが動く。', 20, INK)]], fill=PALE, lw=1.0)
note(s, '経営企画', '各率の仮説値（全変数）')
band(s, [('KGIは', 24, WHITE), ('施策KPIの掛け算', 32, YEL), ('でしかありません', 24, WHITE)])
notes(s, 'KGI（粗利）→ 中間KPI（総名簿数）→ 施策KPI の3階層。各率は第11章の前提条件表と一致させる。主担当：経営企画。')

# --- 10 全体構造図 ---
s = slide_blank(); header(s, 'Ⅱ. 戦略', '全体構造図')
lead(s, [('■　獲得 → 育成 → 刈り取り → 回収。この', 20, INK),
         ('4段', 24, RED), ('で名簿が積み上がる', 20, INK)])
flow = [(0.19, '① 獲　得', ['訪問・飛び込み', 'DLレポート']),
        (7.13, '② 育成ハブ', ['週1メルマガ', '（事例・DL・告知）']),
        (14.07, '③ 刈り取り', ['月1セミナー', '個別テレアポ']),
        (21.01, '④ 名簿回収', ['無料点検の提示', 'ハガキ発送代行'])]
for x, hd, lines in flow:
    card(s, x, 4.20, 6.24, 1.50, 4.20, [(hd, 20, WHITE)],
         [[(t, 16, INK)] for t in lines], body_align=PP_ALIGN.CENTER, body_spc=1.35)
for x in (6.58, 13.52, 20.46):
    arrow(s, x, 7.30, 0.55, 1.00, MSO_SHAPE.RIGHT_ARROW, RED)
plain(s, LM, 10.05, CW, 3.20,
      [[('・獲得 → 育成：メルマガ登録率', 18, INK), ('【　】％', 18, RED)],
       [('・育成 → 刈り取り：開封率', 18, INK), ('【　】％', 18, RED),
        (' × セミナー参加率', 18, INK), ('【　】％', 18, RED)],
       [('・刈り取り → 回収：提携率', 18, INK), ('【　】％', 18, RED),
        (' × 名簿提出率', 18, INK), ('【　】％', 18, RED)]],
      fill=PALE, align=PP_ALIGN.LEFT, spc=1.3, lw=1.0)
plain(s, LM, 13.45, CW, 1.75,
      [[('詰まっている段だけを見つけて、そこだけ直す。', 24, INK)]], fill=WHITE)
note(s, '起案者', 'ファネル各段の転換率', '—')
band(s, [('名簿は', 24, WHITE), ('この1本の導線', 32, YEL), ('からしか積み上がりません', 24, WHITE)])
notes(s, '獲得→育成ハブ→刈り取り→名簿回収。各段の転換率は第11章の前提条件表で管理。主担当：起案者。')

# --- 11 中扉 Ⅲ ---
interlude(['集めて、育てて、', '刈り取る。'], 60)

# --- 12 訪問・飛び込み営業 ---
s = slide_blank(); header(s, 'Ⅲ. 施策', '訪問・飛び込み営業')
lead(s, [('■　断られても名刺は必ず獲る。それがそのまま', 20, INK),
         ('メルマガの母数', 24, RED), ('になる', 20, INK)])
card(s, LM, 4.05, CW, 1.30, 2.40, [('提携社数の公式', 20, YEL)],
     [[('提携社数 ＝ 訪問数', 20, INK), ('【　　　】件', 20, RED),
       (' × 面会率', 20, INK), ('【　】％', 20, RED),
       (' × 商談率', 20, INK), ('【　】％', 20, RED),
       (' × 提携率', 20, INK), ('【　】％', 20, RED)]],
     body_fill=WHITE, body_align=PP_ALIGN.CENTER, lw=2.5)
card(s, 0.29, 7.95, 13.17, 1.40, 3.45, [('即効獲得｜経営層に直接提示', 20, WHITE)],
     [[('・販売利益（フロー収益）', 18, INK)],
      [('・メンテ手数料（ストック収益）', 18, INK)],
      [('・この2本立てで経営層に刺す', 18, INK)]], body_size=18)
card(s, 14.06, 7.95, 13.17, 1.40, 3.45, [('リスト補給｜断られた後の動線', 20, WHITE)],
     [[('・名刺は必ず獲得する', 18, INK)],
      [('・即日メルマガへ合流させる', 18, INK)],
      [('・DL・セミナーで再接触する', 18, INK)]], body_size=18)
plain(s, LM, 13.00, CW, 2.20,
      [[('訪問の成果はゼロにならない。', 24, INK), ('断り ＝ リード獲得', 24, RED),
        ('と定義する。', 24, INK)]], fill=PALE)
note(s, '営業', '面会率・商談率・提携率の実測')
band(s, [('断られても、', 24, WHITE), ('名刺は残る', 32, YEL), ('。それが次の母数です', 24, WHITE)])
notes(s, '提携社数 = 訪問数 × 面会率 × 商談率 × 提携率／未確定：面会率・商談率・提携率。主担当：営業。')

# --- 13 DLレポート ---
s = slide_blank(); header(s, 'Ⅲ. 施策', 'DLレポート')
lead(s, [('■　新規は', 20, INK), ('網羅', 24, RED), ('し、既存は', 20, INK),
         ('再点火', 24, RED), ('する', 20, INK)])
card(s, LM, 4.05, CW, 1.30, 2.40, [('獲得リード数の公式', 20, YEL)],
     [[('獲得リード数 ＝（広告露出', 20, INK), ('【　　　】imp', 20, RED),
       (' ＋ メルマガ配信', 20, INK), ('【　　　】件', 20, RED), ('）', 20, INK)],
      [('× クリック率', 20, INK), ('【　】％', 20, RED),
       (' × DL率', 20, INK), ('【　】％', 20, RED)]],
     body_fill=WHITE, body_align=PP_ALIGN.CENTER, body_spc=1.2, lw=2.5)
card(s, 0.29, 7.95, 13.17, 1.40, 3.75, [('新規網羅｜Web・FB広告', 20, WHITE)],
     [[('・「メンテでストック収入を作る協業モデル」', 16, INK)],
      [('・配信面：Web ／ Facebook', 16, INK)],
      [('・広告予算【　　　】万円/月', 16, INK)],
      [('・DL者はメルマガへ自動登録', 16, INK)]], body_size=16)
card(s, 14.06, 7.95, 13.17, 1.40, 3.75, [('リスト活性化｜既存へ告知', 20, WHITE)],
     [[('・既存リストにも同じDLを案内', 16, INK)],
      [('・開封・クリックでスコアが再上昇', 16, INK)],
      [('・反応者を優先的にテレアポへ回す', 16, INK)],
      [('・眠っていた名刺が再び動き出す', 16, INK)]], body_size=16)
plain(s, LM, 13.30, CW, 1.90,
      [[('1本のレポートで、', 24, INK), ('新規獲得と既存活性を同時に', 24, RED),
        ('取りにいく。', 24, INK)]], fill=PALE)
note(s, 'マーケ', '広告予算、CTR・DL率のベンチマーク')
band(s, [('作るのは', 24, WHITE), ('1本', 32, YEL), ('。効かせる先は2つです', 24, WHITE)])
notes(s, '獲得リード数 =（広告露出＋メルマガ配信）× クリック率 × DL率／未確定：広告予算、CTR、DL率。主担当：マーケ。')

# --- 14 メルマガ・セミナー・テレアポ ---
s = slide_blank(); header(s, 'Ⅲ. 施策', '育成ハブ＆刈り取り')
lead(s, [('■　関心が', 20, INK), ('最高潮の瞬間', 24, RED), ('を狙い撃つ', 20, INK)])
steps = [(LM, 'STEP 1｜週1メルマガ', ['事例・DL案内・', 'セミナー告知を', '交互に配信する']),
         (9.47, 'STEP 2｜月1セミナー', ['収益モデルを', 'その場で理解して', 'もらう']),
         (18.74, 'STEP 3｜テレアポ', ['個別の収益', 'シミュレーションを', '持って電話する'])]
for x, hd, lines in steps:
    card(s, x, 4.05, 8.54, 1.40, 3.40, [(hd, 18, WHITE)],
         [[(t, 16, INK)] for t in lines], body_align=PP_ALIGN.CENTER, body_spc=1.35)
for x in (8.83, 18.10):
    arrow(s, x, 6.60, 0.55, 1.10, MSO_SHAPE.RIGHT_ARROW, RED)
card(s, LM, 9.05, CW, 1.30, 2.40, [('提携社数の公式', 20, YEL)],
     [[('提携社数 ＝ 総リスト数', 20, INK), ('【　　　】件', 20, RED),
       (' × 反応率', 20, INK), ('【　】％', 20, RED),
       (' × 商談化率', 20, INK), ('【　】％', 20, RED),
       (' × 提携率', 20, INK), ('【　】％', 20, RED)]],
     body_fill=WHITE, body_align=PP_ALIGN.CENTER, lw=2.5)
plain(s, LM, 12.95, CW, 2.25,
      [[('セミナー直後', 24, INK), ('48時間以内', 24, RED), ('に、必ず電話する。', 24, INK)]],
      fill=PALE)  # 3ステップ4.05-8.85 / 公式9.05-12.75
note(s, 'マーケ／営業', '開封率、セミナー参加率、参加者からの提携率')
band(s, [('育てた温度が', 24, WHITE), ('冷める前に', 32, YEL), ('、刈り取ります', 24, WHITE)])
notes(s, '提携社数 = 総リスト数 × 反応率 × 商談化率 × 提携率／未確定：開封率、セミナー参加率、参加者からの提携率。主担当：マーケ／営業。')

# --- 15 提携後の名簿引き出し ---
s = slide_blank(); header(s, 'Ⅲ. 施策', '提携後の名簿引き出し')
lead(s, [('■　提携直後の', 20, INK), ('1ヶ月', 24, RED), ('で、名簿を一括回収する', 20, INK)])
card(s, LM, 4.05, CW, 1.30, 2.40, [('獲得名簿数の公式', 20, YEL)],  # 2行組み
     [[('獲得名簿数 ＝ 提携社数', 20, INK), ('【　　】社', 20, RED),
       (' × 1社あたり平均名簿数', 20, INK), ('【　　　】件', 20, RED)],
      [('× 名簿提出率', 20, INK), ('【　】％', 20, RED)]],
     body_fill=WHITE, body_align=PP_ALIGN.CENTER, lw=2.5)
draw = [(LM, '① 無料点検を提示', ['提携先の顧客に', '無料点検を案内する', '大義名分をつくる']),
        (9.47, '② 手数料を還元', ['点検から生まれた', '売上の手数料を', '提携先へ還元する']),
        (18.74, '③ 発送を代行', ['ハガキ発送は', '自社負担で代行し', '提携先の手間をゼロに'])]
for x, hd, lines in draw:
    card(s, x, 7.95, 8.54, 1.40, 3.35, [(hd, 20, WHITE)],
         [[(t, 16, INK)] for t in lines], body_align=PP_ALIGN.CENTER, body_spc=1.35)
plain(s, LM, 12.90, CW, 2.30,
      [[('提携から1ヶ月以内。', 24, INK), ('ここを逃すと名簿は出てこない。', 24, RED)]],
      fill=PALE)
note(s, '営業', '名簿提出率、ハガキ発送の実費単価')
band(s, [('提携はゴールではない。', 24, WHITE), ('名簿の回収', 32, YEL),
         ('までが1件です', 24, WHITE)])
notes(s, '獲得名簿数 = 提携社数 × 1社あたり平均名簿数 × 名簿提出率／未確定：名簿提出率、ハガキ実費単価。主担当：営業。')

# --- 16 中扉 Ⅳ ---
interlude(['数字で、', '確かめる。'], 60)

# --- 17 前提条件表 ---
s = slide_blank(); header(s, 'Ⅳ. 数値', '前提条件表')
lead(s, [('■　すべての数字に、', 20, INK), ('根拠区分', 24, RED), ('を明記する', 20, INK)])
c17 = [(0.19, 7.50), (7.69, 4.50), (12.19, 7.00), (19.19, 8.09)]
for (x, w), t in zip(c17, ['変　数', '仮説値', '根拠区分', '出典・備考']):
    plain(s, x, 4.05, w, 1.00, [[(t, 16, WHITE)]], fill=NAVY, line=NAVY, lw=1.0)
vars17 = ['訪問数（月）', '面会率', '商談率', '提携率', 'クリック率（広告）',
          'DL率', 'メルマガ反応率', 'セミナー参加率', '名簿提出率', '1社あたり平均名簿数']
y = 5.05
for v in vars17:
    plain(s, 0.19, y, 7.50, 1.00, [[(v, 14, NAVY)]], fill=PALE, line=NAVY, lw=1.0,
          align=PP_ALIGN.LEFT)
    plain(s, 7.69, y, 4.50, 1.00, [[('【　　】', 14, RED)]], fill=WHITE, line=NAVY, lw=1.0)
    plain(s, 12.19, y, 7.00, 1.00, [[('自社実績／業界BM', 13, INK)]], fill=WHITE,
          line=NAVY, lw=1.0)
    plain(s, 19.19, y, 8.09, 1.00, [[('—', 13, INK)]], fill=WHITE, line=NAVY, lw=1.0)
    y += 1.00
note(s, '経営企画', '上記すべての率および平均名簿数')
band(s, [('仮説値には必ず', 24, WHITE), ('根拠', 32, YEL), ('をひも付けます', 24, WHITE)])
notes(s, '全公式の変数一覧。各値は「自社実績」か「業界ベンチマーク」かを必ず明記する。主担当：経営企画。')

# --- 18 収支シミュレーション ---
s = slide_blank(); header(s, 'Ⅳ. 数値', '収支シミュレーション')
lead(s, [('■　役員が見るのは、', 20, INK), ('保守ケース', 24, RED), ('です', 20, INK)])
card(s, LM, 3.95, CW, 1.20, 1.90, [('売上の公式', 18, YEL)],
     [[('売上 ＝ 名簿数 × アポ率 × 商談化率 × 成約率 × 平均単価', 20, INK)]],
     body_fill=WHITE, body_align=PP_ALIGN.CENTER, head_size=18, lw=2.5)
c18 = [(0.19, 9.09), (9.28, 6.00), (15.28, 6.00), (21.28, 6.00)]
for (x, w), t in zip(c18, ['項　目', '保守ケース', '標準ケース', '強気ケース']):
    plain(s, x, 7.35, w, 1.00, [[(t, 16, WHITE)]], fill=NAVY, line=NAVY, lw=1.0)
rows18 = [('総名簿数', '【　　　】件'), ('アポ率', '【　】％'), ('成約率', '【　】％'),
          ('平均単価', '【　　　】万円'), ('売上', '【　　　】万円'),
          ('粗利（KGI）', '【　　　】万円')]
y = 8.35
for a, b in rows18:
    plain(s, 0.19, y, 9.09, 1.00, [[(a, 16, NAVY)]], fill=PALE, line=NAVY, lw=1.0,
          align=PP_ALIGN.LEFT)
    for x, w in c18[1:]:
        plain(s, x, y, w, 1.00, [[(b, 16, RED)]], fill=WHITE, line=NAVY, lw=1.0)
    y += 1.00
plain(s, LM, 14.50, CW, 0.68,
      [[('※ 5つ目の公式は要追加 ― 名簿の質の定義（築年数・エリア・オール電化の有無）を変数化すること', 13, RED)]],
      fill=WHITE, line=RED, lw=1.5, align=PP_ALIGN.LEFT, spc=1.0)
note(s, '経営企画', 'アポ率・成約率・平均単価、名簿の質の定義')
band(s, [('判断の基準は、', 24, WHITE), ('保守ケース', 32, YEL), ('に置きます', 24, WHITE)])
notes(s, '売上 = 名簿数 × アポ率 × 商談化率 × 成約率 × 平均単価。'
         '※5つ目の公式（名簿の質の定義：築年数・エリア・オール電化有無）が未定義。主担当：経営企画。')

# --- 19 パートナー側の収益試算 ---
s = slide_blank(); header(s, 'Ⅳ. 数値', 'パートナー収益試算')
lead(s, [('■　提携先が儲からなければ、', 20, INK), ('提携率の前提が崩れる', 24, RED)])
card(s, LM, 4.05, CW, 1.30, 2.40, [('提携先の年間手数料', 20, YEL)],
     [[('年間手数料 ＝ 名簿数', 20, INK), ('【　　　】件', 20, RED),
       (' × 点検実施率', 20, INK), ('【　】％', 20, RED),
       (' × 手数料単価', 20, INK), ('【　　】円', 20, RED)]],
     body_fill=WHITE, body_align=PP_ALIGN.CENTER, lw=2.5)
part = [(LM, '初年度', '【　　　】円'), (9.47, '2年目以降（年）', '【　　　】円'),
        (18.74, '5年累計', '【　　　】円')]
for x, hd, val in part:
    card(s, x, 7.95, 8.54, 1.40, 3.35, [(hd, 20, WHITE)], [[(val, 28, RED)]],
         body_align=PP_ALIGN.CENTER)
plain(s, LM, 12.75, CW, 2.45,
      [[('提携先の年間ストック収入', 22, INK), ('【　　　】円', 22, RED), ('。', 22, INK)],
       [('これを提示できて初めて、提携率の前提が立つ。', 22, INK)]], fill=PALE, spc=1.2)
note(s, '営業', 'メンテ手数料率、点検実施率、手数料単価')
band(s, [('相手の', 24, WHITE), ('取り分', 32, YEL), ('から先に設計します', 24, WHITE)])
notes(s, '年間手数料 = 名簿数 × 点検実施率 × 手数料単価／未確定：メンテ手数料率、点検実施率。'
         'ここが弱いと第7章・第9章の提携率の前提が崩れる。主担当：営業。')

# --- 20 損益分岐点・回収期間 ---
s = slide_blank(); header(s, 'Ⅳ. 数値', '損益分岐点・回収期間')
lead(s, [('■　何社で回収でき、', 20, INK), ('どこで撤退するか', 24, RED),
         ('を先に決める', 20, INK)])
card(s, 0.29, 4.10, 13.17, 1.50, 4.30, [('固 定 費', 24, WHITE)],
     [[('・人件費　【　　　】万円/月', 18, INK)],
      [('・ツール・システム　【　　】万円/月', 18, INK)],
      [('・セミナー運営　【　　】万円/回', 18, INK)]], body_size=18)
card(s, 14.06, 4.10, 13.17, 1.50, 4.30, [('変 動 費', 24, WHITE)],
     [[('・広告費　【　　】円/リード', 18, INK)],
      [('・ハガキ発送　【　　】円/通', 18, INK)],
      [('・提携先への手数料還元　【　】％', 18, INK)]], body_size=18)
bep = [(LM, 'BEP提携社数', '【　　】社'), (9.47, '投資回収期間', '【　　】ヶ月'),
       (18.74, '撤退基準', '【　　】ヶ月時点で\n提携【　】社未満')]
for x, hd, val in bep:
    sh, sb = card(s, x, 10.15, 8.54, 1.30, 2.30, [(hd, 20, WHITE)], [], body_fill=PALE)
    for i, t in enumerate(val.split('\n')):
        para(sb.text_frame, [(t, 20, RED)], PP_ALIGN.CENTER, 1.2, first=(i == 0))
plain(s, LM, 14.00, CW, 1.20,
      [[('撤退基準を決めてから、投資する。', 20, INK)]], fill=WHITE, lw=1.0)
note(s, '経営企画', '固定費・変動費の内訳、BEP提携社数、回収月数')
band(s, [('引き際を決めてから、', 24, WHITE), ('踏み込みます', 32, YEL)])
notes(s, 'BEP提携社数と回収月数を算出。撤退基準を先に定義する。未確定：固定費・変動費の内訳。主担当：経営企画。')

# --- 21 中扉 Ⅴ ---
interlude(['4ヶ月で、', '回し切る。'], 60)

# --- 22 ロードマップ ---
s = slide_blank(); header(s, 'Ⅴ. 実行', 'ロードマップ')
lead(s, [('■　', 20, INK), ('4ヶ月', 24, RED),
         ('で、名簿が回収され始める状態をつくる', 20, INK)])
road = [(0.19, '1ヶ月目', ['訪問開始', 'DL第1号を制作']),
        (7.13, '2ヶ月目', ['メルマガ配信開始', 'DL広告を出稿']),
        (14.07, '3ヶ月目', ['初回セミナー開催', '直後にテレアポ']),
        (21.01, '4ヶ月目〜', ['提携・名簿回収', 'ハガキ発送代行'])]
for x, hd, lines in road:
    card(s, x, 4.15, 6.24, 1.50, 5.00, [(hd, 20, WHITE)],
         [[(t, 16, INK)] for t in lines], body_align=PP_ALIGN.CENTER, body_spc=1.4)
for x in (6.58, 13.52, 20.46):
    arrow(s, x, 7.60, 0.55, 1.10, MSO_SHAPE.RIGHT_ARROW, RED)
kpis = [(0.19, '訪問数', '【　　】件'), (7.13, '獲得リード', '【　　　】件'),
        (14.07, '提携社数', '【　】社'), (21.01, '獲得名簿', '【　　　】件')]
for x, hd, val in kpis:
    plain(s, x, 10.90, 6.24, 1.05, [[(hd, 16, WHITE)]], fill=NAVY, line=NAVY, lw=1.0)
    plain(s, x, 11.95, 6.24, 1.30, [[(val, 20, RED)]], fill=PALE, line=NAVY, lw=1.0)
plain(s, LM, 13.45, CW, 1.75,
      [[('毎月、この', 24, INK), ('4つの数字', 24, RED), ('だけを見て軌道修正する。', 24, INK)]],
      fill=WHITE)
note(s, '起案者', '各月のKPI目標値')
band(s, [('4ヶ月目に、', 24, WHITE), ('最初の名簿', 32, YEL), ('が返ってきます', 24, WHITE)])
notes(s, '月次KPI目標は第11章の前提条件表から逆算して設定する。主担当：起案者。')

# --- 23 体制・役割分担 ---
s = slide_blank(); header(s, 'Ⅴ. 実行', '体制・役割分担')
lead(s, [('■　', 20, INK), ('増員なしで回るか', 24, RED), ('、ここで見極める', 20, INK)])
c23 = [(0.19, 5.50), (5.69, 4.50), (10.19, 12.00), (22.19, 5.09)]
for (x, w), t in zip(c23, ['機　能', '担　当', '主な業務', '工数（人月）']):
    plain(s, x, 4.10, w, 1.10, [[(t, 16, WHITE)]], fill=NAVY, line=NAVY, lw=1.0)
rows23 = [('訪問営業', '◯◯', '工務店・不動産への訪問／提携交渉／名簿回収'),
          ('マーケティング', '◯◯', 'DLレポート制作／広告運用／メルマガ／セミナー'),
          ('事務代行', '◯◯', 'ハガキ原稿作成・印刷・発送／名簿データ化'),
          ('統括・数値管理', '◯◯', 'KPI集計／月次レビュー／役員報告')]
y = 5.20
for a, b, c in rows23:
    plain(s, 0.19, y, 5.50, 2.00, [[(a, 16, NAVY)]], fill=PALE, line=NAVY, lw=1.0)
    plain(s, 5.69, y, 4.50, 2.00, [[(b, 16, INK)]], fill=WHITE, line=NAVY, lw=1.0)
    plain(s, 10.19, y, 12.00, 2.00, [[(c, 16, INK)]], fill=WHITE, line=NAVY, lw=1.0,
          align=PP_ALIGN.LEFT)
    plain(s, 22.19, y, 5.09, 2.00, [[('【　　】', 18, RED)]], fill=WHITE, line=NAVY, lw=1.0)
    y += 2.00
plain(s, LM, 13.40, CW, 1.80,
      [[('必要人員', 24, INK), ('【　】名', 24, RED), ('　／　増員要否', 24, INK),
        ('【　　　　】', 24, RED)]], fill=PALE)
note(s, '役員', '増員要否、各機能の工数')
band(s, [('回らないなら、', 24, WHITE), ('外に出す', 32, YEL), ('という判断もあります', 24, WHITE)])
notes(s, '訪問・マーケ・事務代行（ハガキ発送）の担当と工数。増員要否は役員判断。主担当：役員。')

# --- 24 リスクと対策 ---
s = slide_blank(); header(s, 'Ⅴ. 実行', 'リスクと対策')
lead(s, [('■　潰しておくべきリスクは、', 20, INK), ('2つ', 24, RED)])
card(s, 0.29, 4.10, 13.17, 1.60, 6.00, [('① 個人情報の取り扱い', 24, WHITE)],
     [[('・名簿の受領は第三者提供にあたり、', 18, INK)],
      [('　原則として本人同意が必要', 18, INK)],
      [('・「点検ハガキ発送の受託（委託構成）」', 18, INK)],
      [('　であれば同意不要となる余地あり', 18, INK)],
      [('・委託契約書と安全管理措置の整備が前提', 18, INK)]], body_size=18, body_spc=1.25)
card(s, 14.06, 4.10, 13.17, 1.60, 6.00, [('② 非競合の明文化', 24, WHITE)],
     [[('・リフォーム案件は提携先へ戻す', 18, INK)],
      [('・提携先の商圏では直接営業しない', 18, INK)],
      [('・上記を契約条項として明文化し、', 18, INK)],
      [('　「顧客を取られる」懸念を潰す', 18, INK)],
      [('・違反時の取り扱いも条項に含める', 18, INK)]], body_size=18, body_spc=1.25)
plain(s, LM, 11.95, CW, 3.25,
      [[('法務レビューの結論：', 24, INK), ('【　　　　　　　　】', 24, RED)],
       [('（未了 ― 本企画の実行はレビュー完了を前提とする）', 20, RED)]],
      fill=WHITE, line=RED, lw=2.5)
note(s, '法務', '法務レビューの結論（委託構成の可否）', '—')
band(s, [('この2つを潰さないと、', 24, WHITE), ('名簿は動きません', 32, YEL)])
notes(s, '①個人情報：第三者提供と委託構成の切り分け。②非競合の契約条項化。'
         '法務レビューの結論待ち。主担当：法務。')

# --- 25 中扉 Appendix ---
interlude(['Appendix'], 66)

# --- 26 補足資料 ---
s = slide_blank(); header(s, 'Appendix', '補足資料')
lead(s, [('■　本編に添付する資料は', 20, INK), ('4点', 24, RED)])
apx = [(0.19, '① DLレポート案', ['「メンテでストック', '収入を作る協業モデル」']),
       (7.13, '② 訪問トークスクリプト', ['経営層向け', '初回接触の型']),
       (14.07, '③ 提携契約書ドラフト', ['非競合条項・', '委託構成を含む']),
       (21.01, '④ ハガキ原稿案', ['無料点検の', '案内文面'])]
for x, hd, lines in apx:
    sh, sb = card(s, x, 4.30, 6.24, 1.60, 5.60, [(hd, 16, WHITE)],
                  [[(t, 14, INK)] for t in lines], body_align=PP_ALIGN.CENTER,
                  body_spc=1.35)
    para(sb.text_frame, [('', 14, INK)], PP_ALIGN.CENTER, 1.35)
    para(sb.text_frame, [('担当：◯◯', 14, NAVY)], PP_ALIGN.CENTER, 1.35)
    para(sb.text_frame, [('状態：作成中', 14, RED)], PP_ALIGN.CENTER, 1.35)
plain(s, LM, 11.75, CW, 3.45,
      [[('別紙として同送する。', 24, INK)],
       [('本編の決裁とは切り離し、実務着手後に順次確定させる。', 24, INK)]], fill=PALE)
note(s, '各担当', '各資料の完成（DL案・スクリプト・契約書・ハガキ原稿）', '—')
band(s, [('本編は', 24, WHITE), ('決裁', 32, YEL), ('、別紙は', 24, WHITE),
         ('実務', 32, YEL), ('です', 24, WHITE)])
notes(s, 'DLレポート案・訪問トークスクリプト・提携契約書ドラフト・ハガキ原稿案。主担当：各担当。')

# --- 27 ご決裁のお願い ---
s = slide_blank(); header(s, 'まとめ', 'ご決裁のお願い')
lead(s, [('■　あらためて、ご承認いただきたい', 20, INK), ('3点', 24, RED),
         ('です', 20, INK)])
fin = [(LM, '① 予算総額', '【　　　】万円'), (9.47, '② 人　員', '【　】名'),
       (18.74, '③ 手数料還元率', '【　】％')]
for x, hd, val in fin:
    card(s, x, 4.20, 8.54, 1.60, 3.60, [(hd, 20, WHITE)], [[(val, 32, RED)]],
         body_align=PP_ALIGN.CENTER)
plain(s, LM, 9.75, CW, 2.60,
      [[('提携', 28, INK), ('【　　】社', 28, RED), (' × 名簿', 28, INK),
        ('【　　　】件', 28, RED), (' → 粗利', 28, INK), ('【　　　　】円', 28, RED)]],
      fill=PALE, lw=2.5)
plain(s, LM, 12.60, CW, 2.60,
      [[('本企画の成否は、', 22, INK), ('“総名簿数”という1つの数字', 22, RED),
        ('に集約されます。', 22, INK)]], fill=WHITE, spc=1.2)
note(s, '起案者', '第1章・第2章と同じ数値を転記（最終確定後）', '—')
band(s, [('ご決裁を', 24, WHITE), ('お願いいたします', 32, YEL)])
notes(s, 'エグゼクティブサマリー（第1章）と決裁依頼事項（第2章）の数値をそのまま転記する。主担当：起案者。')


# ================= ページ番号の付与と保存 =================
for i, rec in enumerate(SLIDES, start=1):
    footer(rec['s'], i, strip=rec['strip'])

os.makedirs('output', exist_ok=True)
out = 'output/企画書_B2B2C提携戦略_空パッケージ.pptx'
prs.save(out)
print('saved:', out, '/ slides:', len(prs.slides.__iter__.__self__._sldIdLst))
