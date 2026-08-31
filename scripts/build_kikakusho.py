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
from pptx.enum.dml import MSO_LINE_DASH_STYLE
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


def keymsg(s, msg, size=26):
    """キーメッセージ帯。そのページで持ち帰ってほしい一文を置く。"""
    band(s, [(msg, size, YEL)])


def band(s, parts):
    shape(s, 0, 15.93, W, 2.40, fill=NAVY)
    tf = textbox(s, 0, 15.93, W, 2.40)
    para(tf, parts, PP_ALIGN.CENTER, 0.9, first=True)


PENDING = []


def rec(owner, pending):
    """確定待ち事項を記録する（スライドには描かない）。"""
    PENDING.append((len(SLIDES), owner, pending))


def note(s, owner, pending, basis='自社実績 or 業界ベンチマーク（要記入）'):
    rec(owner, pending)
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
    tf = textbox(s, 1.37, 4.64, 18.97, 2.00)
    para(tf, [('[ 企画書 ]　', 44, WHITE)], PP_ALIGN.LEFT, first=True)
    tf = textbox(s, 1.37, 6.53, 19.12, 5.40)
    para(tf, [('アライアンス提携による', 54, WHITE)], PP_ALIGN.LEFT, 1.15, first=True)
    para(tf, [('名簿獲得戦略', 54, WHITE)], PP_ALIGN.LEFT, 1.15)
    shape(s, 0.33, 15.61, 26.85, 2.50, fill=WHITE)
    s.shapes.add_picture(LOGO, Cm(0.89), Cm(16.25), Cm(3.19), Cm(1.64))
    tf = textbox(s, 9.31, 16.07, 17.64, 1.58)
    para(tf, [('株式会社 船井総合研究所', 16, INK)], PP_ALIGN.RIGHT, 1.0, first=True)
    para(tf, [('専門工事支援部　住宅エネルギーチーム　佐野、伊木、岸野', 16, INK)],
         PP_ALIGN.RIGHT, 1.0)
    SLIDES[-1]['strip'] = False
    notes(s, '表紙。会議名・開催日を入れる場合はタイトル下に追記する。')
    return s


def toc():
    s = slide_blank()
    header(s, '目次', '本日の検討範囲')
    items = [
        ('①', 'Ⅰ. 結論', '決めていただきたいこと', 'P3'),
        ('②', 'Ⅱ. 戦略', 'センターピンはどこか', 'P4'),
        ('③', 'Ⅲ. 施策', 'どう獲り、育て、刈るか', 'P8'),
        ('④', 'Ⅳ. 数値', 'いくら投じ、いくら返るか', 'P14'),
        ('⑤', 'Ⅴ. 実行', '誰が、いつまでに', 'P17'),
        ('⑥', 'Ⅵ. 別軸', '本編とは切り離して見る2件', 'P21'),
    ]
    y = 4.10
    for num, label, sub, pg in items:
        tf = textbox(s, 1.67, y, 1.94, 1.40)
        para(tf, [(num, 32, INK)], PP_ALIGN.CENTER, first=True)
        tf = textbox(s, 3.89, y, 7.50, 1.40)
        para(tf, [(label, 28, INK)], PP_ALIGN.LEFT, first=True)
        tf = textbox(s, 11.60, y, 11.00, 1.40)
        para(tf, [('～ ' + sub + ' ～', 18, NAVY)], PP_ALIGN.LEFT, first=True)
        tf = textbox(s, 23.00, y, 4.28, 1.40)
        para(tf, [(pg, 20, NAVY)], PP_ALIGN.RIGHT, first=True)
        y += 1.80
    notes(s, '各章の掲載ページを示す。')
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




GRAY = RGBColor(0xBF, 0xBF, 0xBF)
SOFT = RGBColor(0xF2, 0xF2, 0xF2)


def photo(s, x, y, w, h, label='写 真'):
    """後から写真を差し込むための枠。破線＋淡いグレー地で「ここに入る」と分かる形にする。"""
    sh = shape(s, x, y, w, h, fill=SOFT, line=NAVY, lw=1.0)
    sh.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    para(sh.text_frame, [(label, 16, GRAY)], PP_ALIGN.CENTER, first=True)
    return sh


def case_page(tag, lead_parts, results, owner):
    """施策ごとの事例紹介ページ。POINT欄と成果帯だけを置き、中央は写真・本文用に空けておく。"""
    s = slide_blank(); header(s, 'Ⅲ. 施策', tag)
    lead(s, lead_parts)
    sh = shape(s, LM, 4.05, 3.31, 1.28, fill=RED, line=RED, lw=1.0)
    para(sh.text_frame, [('P O I N T', 18, WHITE)], PP_ALIGN.CENTER, first=True)
    plain(s, 3.62, 4.05, 23.66, 1.28, [[('', 18, INK)]], fill=WHITE, line=NAVY,
          lw=1.0, align=PP_ALIGN.LEFT)
    band(s, [('成果｜', 24, WHITE), (results, 28, YEL)])
    rec(owner, '事例企業名、取り組み内容、写真、成果の実績値')
    notes(s, '事例紹介ページ。POINT欄の下は写真と本文を差し込むために空けてある。'
             '事例企業名は確定後に〇社＠◯◯市から差し替える。主担当：%s。' % owner)
    return s


def exec_page(tag, lead_parts, items, target, message, owner):
    """施策ごとの実行ページ。何を・誰が・いつから・どこまでやるかを1枚に落とす。"""
    s = slide_blank(); header(s, 'Ⅴ. 実行', tag)
    lead(s, lead_parts)
    cols = [(0.19, 1.80), (1.99, 13.80), (15.79, 4.00), (19.79, 3.60), (23.39, 3.89)]
    for (x, w), t in zip(cols, ['No.', '実行すること', '担　当', '開始時期', '月次目標']):
        plain(s, x, 4.05, w, 1.00, [[(t, 16, WHITE)]], fill=NAVY, line=NAVY, lw=1.0)
    y = 5.05
    for i, it in enumerate(items, start=1):
        plain(s, 0.19, y, 1.80, 1.60, [[(str(i), 16, NAVY)]], fill=PALE, line=NAVY, lw=1.0)
        plain(s, 1.99, y, 13.80, 1.60, [[(it, 16, INK)]], fill=WHITE, line=NAVY, lw=1.0,
              align=PP_ALIGN.LEFT)
        for x, w in cols[2:]:
            plain(s, x, y, w, 1.60, [[('', 16, INK)]], fill=WHITE, line=NAVY, lw=1.0)
        y += 1.60
    plain(s, LM, y + 0.20, CW, 15.70 - (y + 0.20),
          [[('この施策の月次目標｜', 22, INK), (target, 22, RED)]], fill=PALE, lw=1.5)
    rec(owner, '担当・開始時期・月次目標')
    keymsg(s, message)
    notes(s, '施策ごとの実行計画。担当・開始時期・月次目標は打合せで埋める。主担当：%s。' % owner)
    return s



# ================= 本編（打合せ資料・22ページ） =================

# --- 1 表紙 ---
cover()

# --- 2 目次 ---
toc()

# --- 3 エグゼクティブサマリー ---
s = slide_blank(); header(s, 'Ⅰ. 結論', 'エグゼクティブサマリー')
lead(s, [('■　提携26社 × 1社あたり名簿68件 → 粗利', 20, INK),
         ('3,320万円', 24, RED)])
for x, hd, val in [(LM, '① 提携社数', '26社'),
                   (9.47, '② 1社あたり名簿数', '68件'),
                   (18.74, '③ 総名簿数（センターピン）', '1,755件')]:
    card(s, x, 4.05, 8.54, 1.40, 2.90, [(hd, 18, WHITE)], [[(val, 32, RED)]],
         body_align=PP_ALIGN.CENTER)
card(s, LM, 8.75, CW, 1.30, 2.40, [('KGI｜粗利額', 20, YEL)],
     [[('粗利 ＝ 総名簿数 1,755件 × アポ率 12％ × 商談化率 70％', 20, INK)],
      [('× 成約率 30％ × 平均単価 250万円 × 粗利率 30％ ＝ 3,320万円', 20, INK)]],
     body_fill=WHITE, body_align=PP_ALIGN.CENTER, body_spc=1.2)
plain(s, LM, 12.55, CW, 2.65,
      [[('この企画で狙うのは粗利', 22, INK), ('3,320万円', 22, RED), ('。', 22, INK)],
       [('その全ては“総名簿数”という1つの数字に依存する。', 22, INK)]], fill=PALE, spc=1.2)
note(s, '起案者', '—（全数値を反映済み）', 'KPI前提と事例ページを参照')
keymsg(s, '初年度は提携26社・名簿1,755件で、粗利3,320万円を狙う。')
notes(s, 'KGI＝粗利額。全数値が確定してから最後に執筆する。主担当：起案者。')

# --- 4 背景と課題 ---
s = slide_blank(); header(s, 'Ⅱ. 戦略', '背景と課題')
lead(s, [('■　個人向け集客はもう伸びない。だから、', 20, INK),
         ('名簿を持つ企業と組む', 24, RED)])
card(s, 0.29, 4.10, 13.17, 1.60, 5.20, [('これまで｜個人向け集客', 24, WHITE)],
     [[('・獲得単価が高騰（直近CPA 40,000円）', 18, INK)],
      [('・リードが枯渇し、母数が増えない', 18, INK)],
      [('・広告を止めれば、受注も止まる', 18, INK)]], body_size=18)
card(s, 14.06, 4.10, 13.17, 1.60, 5.20, [('これから｜アライアンス提携', 22, WHITE)],
     [[('・提携先が既に持つ名簿にアクセス', 18, INK)],
      [('・1社の提携で平均68件がまとめて増える', 18, INK)],
      [('・獲得単価は「提携コスト」に置き換わる', 18, INK)]], body_size=18)
plain(s, LM, 11.20, CW, 4.00,
      [[('1件ずつ獲るのをやめ、', 28, INK)],
       [('1,755件', 28, RED), ('をまとめて獲りにいく', 28, INK)]], fill=PALE)
note(s, '経営企画', '直近CPA 40,000円の裏取り')
keymsg(s, '1件ずつ獲る時代は終わった。名簿ごと獲りにいく。')
notes(s, '現状のリード獲得単価を自社実績から出す。CPA高騰の根拠データを別紙添付。主担当：経営企画。')

# --- 5 センターピンの定義 ---
s = slide_blank(); header(s, 'Ⅱ. 戦略', 'センターピンの定義')
lead(s, [('■　勝敗を決めるのは、たった1つの数字＝', 20, INK), ('総名簿数', 24, RED)])
sh = shape(s, LM, 4.20, CW, 1.50, fill=NAVY, line=NAVY, lw=1.5)
para(sh.text_frame, [('セ ン タ ー ピ ン', 24, YEL)], PP_ALIGN.CENTER, first=True)
plain(s, LM, 5.70, CW, 3.20,
      [[('総名簿数 ＝ 提携社数', 28, INK), ('26社', 28, RED)],
       [('× 1社あたり名簿数', 28, INK), ('68件', 28, RED), ('　＝　', 28, INK),
        ('1,755件', 28, RED)]],
      fill=WHITE, lw=2.5, spc=1.15)
for x, hd, lines in [(LM, '① 目的を一本化', ['訪問もDLもメルマガも、', '目的は提携社数']),
                     (9.47, '② 提携後は必ず回収', ['提携＝ゴールではない。', '名簿の回収までが1件']),
                     (18.74, '③ 他のKPIは追わない', ['名簿数に効かない施策は、', 'やらないと決める'])]:
    card(s, x, 9.30, 8.54, 1.40, 3.10, [(hd, 20, WHITE)],
         [[(t, 16, INK)] for t in lines], body_align=PP_ALIGN.CENTER, body_spc=1.35)
plain(s, LM, 14.05, CW, 1.15,
      [[('この式の左辺を最大化すること以外は、やらない。', 20, INK)]], fill=PALE, lw=1.0)
note(s, '起案者', '1社あたり保有名簿300件・名簿共有率35％の妥当性')
keymsg(s, '追う数字は「総名簿数」ただ1つ。')
notes(s, '総名簿数 = 提携社数 × 1社あたり名簿数。未確定：目標提携社数、1社あたり想定名簿数。主担当：起案者。')

# --- 6 KPIツリー ---
s = slide_blank(); header(s, 'Ⅱ. 戦略', 'KPIツリー')
lead(s, [('■　粗利（KGI）は、名簿数（中間KPI）と施策KPIに', 20, INK),
         ('分解できる', 24, RED)])
plain(s, LM, 4.05, CW, 1.50,
      [[('KGI｜粗利　', 24, WHITE), ('3,320万円', 24, YEL)]], fill=NAVY, line=NAVY, lw=1.5)
arrow(s, 13.26, 5.70, 1.00, 0.90, MSO_SHAPE.DOWN_ARROW, NAVY)
plain(s, LM, 6.75, CW, 1.50,
      [[('中間KPI｜総名簿数　', 24, WHITE), ('1,755件', 24, YEL)]],
      fill=NAVY, line=NAVY, lw=1.5)
arrow(s, 13.26, 8.40, 1.00, 0.90, MSO_SHAPE.DOWN_ARROW, NAVY)
for x, hd, lines in [(0.19, '施策KPI ①\n訪問', ['提携社数', '26社']),
                     (7.13, '施策KPI ②\nフォーム・メルマガ・DL', ['獲得リード数', '460件']),
                     (14.07, '施策KPI ③\nセミナー', ['実参加社数', '28社']),
                     (21.01, '施策KPI ④\n名簿共有', ['名簿共有率', '35％'])]:
    sh = shape(s, x, 9.45, 6.24, 1.55, fill=NAVY, line=NAVY, lw=1.0)
    for i, t in enumerate(hd.split('\n')):
        para(sh.text_frame, [(t, 14, WHITE)], PP_ALIGN.CENTER, 1.1, first=(i == 0))
    sb = shape(s, x, 11.00, 6.24, 2.60, fill=PALE, line=NAVY, lw=1.0)
    para(sb.text_frame, [(lines[0], 16, INK)], PP_ALIGN.CENTER, 1.25, first=True)
    para(sb.text_frame, [(lines[1], 20, RED)], PP_ALIGN.CENTER, 1.25)
plain(s, LM, 13.85, CW, 1.35,
      [[('下の4つを動かせば、上の2つが動く。', 20, INK)]], fill=PALE, lw=1.0)
note(s, '経営企画', '—（KPI前提と事例ページに集約）')
keymsg(s, '粗利は、下段の施策KPIを動かさない限り動かない。')
notes(s, 'KGI（粗利）→ 中間KPI（総名簿数）→ 施策KPI の3階層。各率は「KPI前提と事例」ページと一致させる。主担当：経営企画。')

# --- 7 全体構造図 ---
s = slide_blank(); header(s, 'Ⅱ. 戦略', '全体構造図')
lead(s, [('■　獲得 → 育成 → 刈り取り → 回収。この', 20, INK),
         ('4段', 24, RED), ('で名簿が積み上がる', 20, INK)])
for x, hd, lines in [(0.19, '① 獲　得', ['訪問・飛び込み', 'DLレポート']),
                     (7.13, '② 育成ハブ', ['週1メルマガ', '（事例・DL・告知）']),
                     (14.07, '③ 刈り取り', ['月1セミナー', '個別テレアポ']),
                     (21.01, '④ 名簿回収', ['無料点検の提示', 'ハガキ発送代行'])]:
    card(s, x, 4.20, 6.24, 1.50, 4.20, [(hd, 20, WHITE)],
         [[(t, 16, INK)] for t in lines], body_align=PP_ALIGN.CENTER, body_spc=1.35)
for x in (6.58, 13.52, 20.46):
    arrow(s, x, 7.30, 0.55, 1.00, MSO_SHAPE.RIGHT_ARROW, RED)
plain(s, LM, 10.05, CW, 3.20,
      [[('・獲得 → 育成：メルマガ登録率', 18, INK), ('80％', 18, RED)],
       [('・育成 → 刈り取り：開封率', 18, INK), ('16.9％', 18, RED),
        (' × セミナー申込率', 18, INK), ('2％', 18, RED),
        (' × 参加率', 18, INK), ('70％', 18, RED)],
       [('・刈り取り → 回収：提携率', 18, INK), ('20％', 18, RED),
        (' × 名簿共有率', 18, INK), ('35％', 18, RED)]],
      fill=PALE, align=PP_ALIGN.LEFT, spc=1.3, lw=1.0)
plain(s, LM, 13.45, CW, 1.75,
      [[('詰まっている段だけを見つけて、そこだけ直す。', 24, INK)]], fill=WHITE)
note(s, '起案者', '—', '—')
keymsg(s, '名簿は「獲得→育成→刈り取り→回収」の4段で積み上がる。')
notes(s, '獲得→育成ハブ→刈り取り→名簿回収。各段の転換率は「KPI前提と事例」ページで管理。主担当：起案者。')

# --- 8 訪問・飛び込み営業 ---
s = slide_blank(); header(s, 'Ⅲ. 施策', '訪問・飛び込み営業')
lead(s, [('■　断られても名刺は必ず獲る。それがそのまま', 20, INK),
         ('メルマガの母数', 24, RED), ('になる', 20, INK)])
card(s, LM, 4.05, CW, 1.30, 2.40, [('提携社数の公式', 20, YEL)],
     [[('提携社数 ＝ 訪問数', 20, INK), ('1,080社', 20, RED),
       (' × 商談率', 20, INK), ('7.5％', 20, RED),
       (' × 提携率', 20, INK), ('20％', 20, RED), ('　＝　', 20, INK),
       ('16社', 20, RED)]],
     body_fill=WHITE, body_align=PP_ALIGN.CENTER, lw=2.5)
card(s, 0.29, 7.95, 13.17, 1.40, 3.45, [('即効獲得｜経営層に直接提示', 20, WHITE)],
     [[('・販売利益（フロー収益）', 18, INK)],
      [('・メンテ手数料（ストック収益）', 18, INK)],
      [('・3名 × 月30社 × 12ヶ月で1,080社', 18, INK)]], body_size=18)
card(s, 14.06, 7.95, 13.17, 1.40, 3.45, [('リスト補給｜断られた後の動線', 20, WHITE)],
     [[('・名刺は必ず獲得する', 18, INK)],
      [('・即日メルマガへ合流させる', 18, INK)],
      [('・DL・セミナーで再接触する', 18, INK)]], body_size=18)
plain(s, LM, 13.00, CW, 2.20,
      [[('訪問の成果はゼロにならない。', 24, INK), ('断り ＝ リード獲得', 24, RED),
        ('と定義する。', 24, INK)]], fill=PALE)
note(s, '営業', '訪問先リスト1,080社の確保')
keymsg(s, '訪問1,080社が、提携16社と名刺864件を同時に生む。')
notes(s, '提携社数 = 訪問数 × 商談率 × 提携率／未確定：訪問先リストの確保、商談率。主担当：営業。')

# --- 9 事例｜訪問 ---
case_page('事例｜訪問',
          [('■　訪問で成果を出している', 20, INK), ('T社＠山梨県', 24, RED),
           ('の取り組み', 20, INK)],
          'リスト120件 → 商談15件（12.5％）→ 提携3社（20％）', '営業')

# --- 10 フォーム・メルマガ・DL ---
s = slide_blank(); header(s, 'Ⅲ. 施策', 'フォーム・メルマガ・DL')
lead(s, [('■　商圏', 20, INK), ('3,000社', 24, RED),
         ('をツールで抽出し、訪問先も含めて全社に届ける', 20, INK)])
card(s, LM, 4.05, CW, 1.30, 2.40, [('獲得リード数の公式', 20, YEL)],
     [[('フォーム送信 3,000社 × 年6回（約17,500通）→ 累積DL率', 20, INK),
       ('7.0％', 20, RED), (' → ', 20, INK), ('210社', 20, RED)],
      [('＋ FB広告 月30万円 × 12ヶ月 → ', 20, INK), ('250件', 20, RED),
       ('　＝　獲得リード ', 20, INK), ('460件', 20, RED)]],
     body_fill=WHITE, body_align=PP_ALIGN.CENTER, body_spc=1.2, lw=2.5)
card(s, 0.29, 7.95, 13.17, 1.40, 3.75, [('新規網羅｜フォーム送信', 20, WHITE)],
     [[('・リストDLで商圏3,000社を抽出', 16, INK)],
      [('・訪問先も除外せず3,000社全社へ配信', 16, INK)],
      [('・2ヶ月に1本のDL資料を年6回', 16, INK)],
      [('・月平均1,460通（プラン上限3,000社の49％）', 16, INK)]], body_size=16)
card(s, 14.06, 7.95, 13.17, 1.40, 3.75, [('育成｜メルマガ ＋ FB広告', 20, WHITE)],
     [[('・DL者と訪問名刺をリストに積み上げる', 16, INK)],
      [('・期中平均662件へ週1配信', 16, INK)],
      [('・月2,900通（プラン上限15,000通の19％）', 16, INK)],
      [('・FB広告 月30万円 → 年250件を上乗せ', 16, INK)]], body_size=16)
plain(s, LM, 13.30, CW, 1.90,
      [[('訪問できる先も、できない先も、', 24, INK),
        ('まとめて3,000社に届ける。', 24, RED)]], fill=PALE)
note(s, 'マーケ', 'フォーム送信のDL率、ツール割引の適用期間')
keymsg(s, '人が回れない分は、ツールが年17,500通で埋める。')
notes(s, '獲得リード数 = フォーム営業210社 ＋ FB広告250件 ＝ 460件。'
         'フォーム営業の母数は商圏3,000社の全社。訪問先にもメルマガ・DLを配信してよいという'
         '判断により、従来の「訪問先1,080社を除外」をやめた。主担当：マーケ。')

# --- 11 事例｜フォーム・DL ---
case_page('事例｜フォーム・DL',
          [('■　メルマガで成果を出している', 20, INK), ('船井総研', 24, RED),
           ('の実績', 20, INK)],
          '配信7,775通 → 開封16.9％ → クリック13社 → DL 9件', 'マーケ')

# --- 12 セミナー ---
s = slide_blank(); header(s, 'Ⅲ. 施策', 'セミナー')
lead(s, [('■　年4回のセミナーで、', 20, INK), ('提携を決める', 24, RED)])
for x, hd, lines in [(LM, 'STEP 1｜告　知', ['メルマガとDMで', 'セミナーを案内し', '申込を集める']),
                     (9.47, 'STEP 2｜送　客', ['反応者へ架電し', '参加を後押しする', '（テレアポは上積み）']),
                     (18.74, 'STEP 3｜当　日', ['個別の収益試算を', '提示し、その場で', '提携を打診する'])]:
    card(s, x, 4.05, 8.54, 1.40, 3.40, [(hd, 18, WHITE)],
         [[(t, 16, INK)] for t in lines], body_align=PP_ALIGN.CENTER, body_spc=1.35)
for x in (8.83, 18.10):
    arrow(s, x, 6.60, 0.55, 1.10, MSO_SHAPE.RIGHT_ARROW, RED)
card(s, LM, 9.05, CW, 1.30, 2.40, [('提携社数の公式', 20, YEL)],
     [[('期中平均リスト', 20, INK), ('662件', 20, RED),
       (' × 申込率', 20, INK), ('2％', 20, RED),
       (' × 参加率', 20, INK), ('70％', 20, RED),
       (' × 年4回', 20, INK)],
      [('→ 実参加', 20, INK), ('28社', 20, RED),
       ('（重複25％を控除）× 提携率', 20, INK), ('20％', 20, RED),
       ('　＝　', 20, INK), ('6社', 20, RED)]],
     body_fill=WHITE, body_align=PP_ALIGN.CENTER, lw=2.5)
plain(s, LM, 12.85, CW, 2.35,
      [[('前提：セミナーは', 22, INK), ('メルマガ・DM単独でも9社集客できる企画', 22, RED),
        ('にする。', 22, INK)],
       [('テレアポはその上積み。集客をテレアポに依存させない。', 22, INK)]],
      fill=PALE, spc=1.2)
note(s, 'マーケ／営業', 'セミナー申込率2％の実測、開催頻度')
keymsg(s, '年4回のセミナーが、提携を決める場になる。')
notes(s, 'フローはメルマガ（育成）→ テレアポ（送客）→ セミナー（刈り取り）。'
         'セミナーはDM単独でも成立する企画設計が前提で、テレアポは集客の上積み。主担当：マーケ／営業。')

# --- 13 事例｜セミナー ---
case_page('事例｜セミナー',
          [('■　セミナーで成果を出している', 20, INK), ('〇社＠◯◯市', 24, RED),
           ('の取り組み', 20, INK)],
          '参加【　】社　→　提携【　】社（提携率【　】％）', 'マーケ／営業')

# --- 14 KPI前提と事例 ---
s = slide_blank(); header(s, 'Ⅳ. 数値', 'KPI前提と事例')
lead(s, [('■　各施策のKPIと、その', 20, INK), ('根拠となる事例', 24, RED),
         ('まとめ', 20, INK)])
c14 = [(0.19, 6.00), (6.19, 7.00), (13.19, 5.00), (18.19, 9.09)]
for (x, w), t in zip(c14, ['施　策', 'K P I', '仮説値', '事例｜実績値']):
    plain(s, x, 4.05, w, 1.00, [[(t, 16, WHITE)]], fill=NAVY, line=NAVY, lw=1.0)
groups = [
    ('① 訪問・飛び込み',
     [('訪問数（年）', '1,080社', '3名 × 月30社 × 12ヶ月'),
      ('商談率', '7.5％', '目標値（現状1.0％・T社2.5％の中間）'),
      ('提携率', '20％', 'T社＠山梨県：商談15件 → 提携3社')]),
    ('② フォーム・メルマガ・DL',
     [('フォーム配信母数', '3,000社', '商圏の全社（訪問先も除外しない）'),
      ('累積DL率（年6回）', '7.0％', '210社がDL。月1,460通＝プラン枠の49％'),
      ('メルマガ開封率', '16.9％', '船井総研：7,775通 → 1,313件'),
      ('FB広告リード', '250件', '月30万円 × 12ヶ月')]),
    ('③ セミナー',
     [('セミナー申込率', '2％', '期中平均リスト662件 → 参加9社/回'),
      ('参加者からの提携率', '20％', '訪問の提携率と同水準で設定')]),
]
y = 4.85
for name, kpis in groups:
    plain(s, 0.19, y, 6.00, 0.95 * len(kpis), [[(name, 15, NAVY)]],
          fill=PALE, line=NAVY, lw=1.0, spc=1.2)
    for kpi, val, basis in kpis:
        plain(s, 6.19, y, 7.00, 0.95, [[(kpi, 15, INK)]], fill=WHITE, line=NAVY,
              lw=1.0, align=PP_ALIGN.LEFT)
        plain(s, 13.19, y, 5.00, 0.95, [[(val, 16, RED)]], fill=WHITE, line=NAVY, lw=1.0)
        plain(s, 18.19, y, 9.09, 0.95, [[(basis, 12, NAVY)]],
              fill=WHITE, line=NAVY, lw=1.0, align=PP_ALIGN.LEFT)
        y += 0.95
rec('経営企画', '事例企業名（〇社＠◯◯市）の確定')
keymsg(s, '仮説値はすべて、事例の実績値から置いている。')
notes(s, '施策ごとにKPIと事例をひとまとめにした前提条件表。仮説値は事例の実績値から置く。'
         '事例企業名は確定後に差し替える（現状は〇社＠◯◯市）。主担当：経営企画。')

# --- 15 収支シミュレーション ---
s = slide_blank(); header(s, 'Ⅳ. 数値', '収支シミュレーション')
lead(s, [('■　この1年間での', 20, INK), ('目標数値', 24, RED),
         ('です。投資は年452万円', 20, INK)])
card(s, LM, 4.57, CW, 1.20, 1.90, [('売上の公式', 18, YEL)],
     [[('売上 ＝ 総名簿数 × アポ率 × 商談化率 × 成約率 × 平均単価', 20, INK)]],
     body_fill=WHITE, body_align=PP_ALIGN.CENTER, head_size=18, lw=2.5)
c15 = [(0.19, 9.09), (9.28, 6.00), (15.28, 6.00), (21.28, 6.00)]
for (x, w), t in zip(c15, ['項　目', '悲観値', '標準値', '楽観値']):
    plain(s, x, 7.97, w, 1.00, [[(t, 16, WHITE)]], fill=NAVY, line=NAVY, lw=1.0)
y = 8.97
for a, vals in [('提携社数', ('17社', '26社', '35社')),
                ('総名簿数', ('1,150件', '1,755件', '2,360件')),
                ('アポ率', ('9％', '12％', '15％')),
                ('成約率', ('22％', '30％', '35％')),
                ('契約件数', ('16件', '44件', '87件')),
                ('売上', ('4,000万円', '1億1,100万円', '2億1,700万円')),
                ('粗利（KGI）', ('1,200万円', '3,320万円', '6,500万円'))]:
    plain(s, 0.19, y, 9.09, 1.00, [[(a, 16, NAVY)]], fill=PALE, line=NAVY, lw=1.0,
          align=PP_ALIGN.LEFT)
    for (x, w), v in zip(c15[1:], vals):
        plain(s, x, y, w, 1.00, [[(v, 16, RED)]], fill=WHITE, line=NAVY, lw=1.0)
    y += 1.00
rec('経営企画', '採用ケースの決定（標準値を計画値とするか）')
keymsg(s, '投資452万円に対し、標準ケースで粗利3,320万円。')
notes(s, '売上 = 名簿数 × アポ率 × 商談化率 × 成約率 × 平均単価。'
         '悲観値・標準値・楽観値の3ケースで置く。主担当：経営企画。')

# --- 16 パートナー収益試算 ---
s = slide_blank(); header(s, 'Ⅳ. 数値', 'パートナー収益試算')
lead(s, [('■　提携先が儲からなければ、', 20, INK), ('提携率の前提が崩れる', 24, RED)])
card(s, LM, 4.05, CW, 1.30, 2.40, [('提携先の年間手数料', 20, YEL)],
     [[('年間手数料 ＝ 預けた名簿', 20, INK), ('195件', 20, RED),
       (' × 契約率', 20, INK), ('2.5％', 20, RED),
       (' × 平均単価', 20, INK), ('250万円', 20, RED),
       (' × 手数料率', 20, INK), ('5％', 20, RED)]],
     body_fill=WHITE, body_align=PP_ALIGN.CENTER, lw=2.5)
for x, hd, val in [(LM, '初年度', '61万円'), (9.47, '2年目以降（年）', '40万円'),
                   (18.74, '5年累計', '220万円')]:
    card(s, x, 7.95, 8.54, 1.40, 3.35, [(hd, 20, WHITE)], [[(val, 28, RED)]],
         body_align=PP_ALIGN.CENTER)
plain(s, LM, 12.75, CW, 2.45,
      [[('手数料率5％だと提携先の取り分は', 22, INK), ('年61万円', 22, RED), ('。', 22, INK)],
       [('提携の動機として十分か、料率の見直しを含めて要検討。', 22, INK)]],
      fill=PALE, spc=1.2)
rec('営業', '手数料率5％の是非（提携先の取り分は年61万円）')
keymsg(s, '提携先が儲かる設計でなければ、提携率20％は成立しない。')
notes(s, '年間手数料 = 名簿数 × 点検実施率 × 手数料単価。'
         'ここが弱いと「訪問」「育成ハブ＆刈り取り」ページの提携率の前提が崩れる。主担当：営業。')

# --- 17 ロードマップ ---
s = slide_blank(); header(s, 'Ⅴ. 実行', 'ロードマップ')
lead(s, [('■　', 20, INK), ('3ヶ月', 24, RED), ('で、提携が生まれる状態をつくる', 20, INK)])
for x, hd, lines in [(LM, '1ヶ月目', ['訪問開始', 'DL第1号を制作']),
                     (9.47, '2ヶ月目', ['メルマガ配信開始', 'DL広告を出稿']),
                     (18.74, '3ヶ月目', ['初回セミナー開催', '直後に個別提案'])]:
    card(s, x, 4.42, 8.54, 1.50, 5.00, [(hd, 20, WHITE)],
         [[(t, 16, INK)] for t in lines], body_align=PP_ALIGN.CENTER, body_spc=1.4)
for x in (8.83, 18.10):
    arrow(s, x, 7.87, 0.55, 1.10, MSO_SHAPE.RIGHT_ARROW, RED)
for x, hd, val in [(LM, '訪問数', '90社 /月'), (9.47, '獲得リード', '38件 /月'),
                   (18.74, '提携社数', '2社 /月')]:
    plain(s, x, 11.17, 8.54, 1.05, [[(hd, 16, WHITE)]], fill=NAVY, line=NAVY, lw=1.0)
    plain(s, x, 12.22, 8.54, 1.30, [[(val, 20, RED)]], fill=PALE, line=NAVY, lw=1.0)
plain(s, LM, 13.72, CW, 1.75,
      [[('毎月、この', 24, INK), ('3つの数字', 24, RED), ('だけを見て軌道修正する。', 24, INK)]],
      fill=WHITE)
rec('起案者', '月次KPIの各月への配分')
keymsg(s, '3ヶ月で、提携が生まれる状態をつくる。')
notes(s, '3ヶ月で提携が生まれる状態までを引く。月次KPIは「KPI前提と事例」ページから逆算する。主担当：起案者。')

# --- 18 実行｜訪問 ---
exec_page('実行｜訪問',
          [('■　訪問を、', 20, INK), ('いつ・誰が・何件やるか', 24, RED),
           ('まで落とす', 20, INK)],
          ['対象リストの作成（工務店・不動産会社を抽出）',
           '訪問トークスクリプトの作成',
           '訪問の実施（1名あたり 週8社 ／ 月30社）',
           '獲得名刺のメルマガ登録',
           '商談化した先への提携提案・契約締結'],
          '訪問 90社　→　商談 7件　→　提携 1.3社（年16社）',
          '訪問は月90社。ここが、すべての起点になる。', '営業')

# --- 19 実行｜フォーム・DL ---
exec_page('実行｜フォーム・DL',
          [('■　DL資料を、', 20, INK), ('何を作り、どう配るか', 24, RED),
           ('まで落とす', 20, INK)],
          ['リストDLで商圏3,000社を抽出（訪問先も除外しない）',
           'DL資料の企画・原稿作成（2ヶ月に1本・年6本）',
           'フォーム送信で3,000社へ配信（月平均1,460通）',
           'FB広告の出稿（月30万円）',
           'DL者のメルマガ自動登録と週1配信（月2,900通）'],
          'フォーム配信 1,460通　→　リード 38件 /月',
          'DLは2ヶ月に1本。作るほどリードが積み上がる。', 'マーケ')

# --- 20 実行｜セミナー ---
exec_page('実行｜セミナー',
          [('■　年4回のセミナーを、', 20, INK), ('実行単位', 24, RED),
           ('まで落とす', 20, INK)],
          ['年間の開催計画を立てる（四半期に1回・年4回）',
           'セミナーの企画と当日コンテンツの作成',
           'メルマガ・DMでの告知と申込受付',
           'テレアポによる参加の後押し',
           '当日の個別提案と提携打診'],
          'セミナー参加 9社/回　→　提携 6社 /年',
          'セミナーは四半期に1回。開催日から逆算して動く。', 'マーケ／営業')

# --- 21 別軸①｜TRENDE社 ---
s = slide_blank(); header(s, 'Ⅵ. 別軸', '別軸①｜TRENDE社')
lead(s, [('■　本編とは', 20, INK), ('別軸', 24, RED),
         ('。伊藤忠商事の持分法適用会社との業務提携', 20, INK)])
card(s, 0.29, 4.05, 13.17, 1.30, 3.85, [('会社概要', 20, WHITE)],
     [[('・TRENDE株式会社（2017年設立・資本金14億円）', 13, INK)],
      [('・東京電力グループ発。2023年に伊藤忠商事が', 13, INK)],
      [('　子会社化し、現在は同社の持分法適用会社', 13, INK)],
      [('・東京センチュリー・全農・東芝・東急不動産も出資', 13, INK)],
      [('・小売電気／太陽光・蓄電池のリース・PPA', 13, INK)]],
     body_size=13, body_spc=1.30)
plain(s, 14.06, 4.05, 13.17, 1.30, [[('顧客基盤（同社資料より）', 20, WHITE)]],
      fill=NAVY, line=NAVY, lw=1.5)
y9 = 5.35
for nm, cnt in [('蓄電池レンタルサービス', '約9,000件'),
                ('テラリス（蓄電池・PVリース）', '約1,500件'),
                ('ほっとでんき（PPA）', '約2,800件'),
                ('じぶん電力・ひだまりでんき', '約500件'),
                ('合　計（全国・家庭）', '約13,800件')]:
    last = (cnt == '約13,800件')
    plain(s, 14.06, y9, 8.17, 0.72, [[(nm, 12, NAVY if last else INK)]],
          fill=PALE if last else WHITE, line=NAVY, lw=1.0, align=PP_ALIGN.LEFT)
    plain(s, 22.23, y9, 5.00, 0.72, [[(cnt, 12, RED)]],
          fill=PALE if last else WHITE, line=NAVY, lw=1.0)
    y9 += 0.72
for x, hd, lines in [
        (LM, '案①｜名簿を受け取る',
         ['・NEC蓄電池リースの', '　設置先名簿を受領する', '・当社が点検を提案する',
          '・最も名簿数に直結する']),
        (9.47, '案②｜案件を受け取る',
         ['・名簿は受領しない', '・同社の直販部門と組み', '　リフォーム案件のAPを受領',
          '・名簿ではなく案件が入る']),
        (18.74, '案③｜電力会社と組む',
         ['・同社が提携する電力会社', '　との協業に乗る', '・伊藤忠グループ案件も対象',
          '・中長期の広がりが大きい'])]:
    card(s, x, 9.30, 8.54, 1.25, 4.40, [(hd, 18, WHITE)],
         [[(t, 14, INK)] for t in lines], body_align=PP_ALIGN.LEFT, body_spc=1.45,
         body_size=14)
rec('営業', 'TRENDEとの提携形態（案①／②／③）と、受領できる名簿の件数・範囲')
keymsg(s, '1社との提携で、名簿が桁で動く可能性がある。')
notes(s, '本編の施策（訪問／フォーム・メルマガ・DL／セミナー）とは切り離した別軸の案件。'
         'TRENDE社（伊藤忠商事の持分法適用会社）との業務提携の検討ページ。'
         '同社は2024年に買収した蓄電池レンタル（約9,000件）のリプレイス提案を推進中で、'
         'テラリス（リース）は2025年度実績約900件。提携形態は名簿受領・案件受領・電力会社協業の'
         '3案。案①なら総名簿数が一気に増えるが、名簿の範囲と件数は未確認。主担当：営業。')

# --- 22 別軸②｜イー・スマイル 高橋顧問ルート ---
s = slide_blank(); header(s, 'Ⅵ. 別軸', '別軸②｜イー・スマイル')
lead(s, [('■　こちらも', 20, INK), ('別軸', 24, RED),
         ('。高橋顧問が持つ東海圏 5,000社のパイプ', 20, INK)])
card(s, 0.29, 4.05, 13.17, 1.30, 3.85, [('経　緯', 20, WHITE)],
     [[('・窓口を石原顧問から高橋顧問（イー・スマイル）へ', 13, INK)],
      [('　変更。春日井市在住、43歳', 13, INK)],
      [('・東海圏の工務店・設備に 5,000社 のパイプを持つ', 13, INK)],
      [('・訪問は春日井市の事務所を拠点に、坂元・田中・島の', 13, INK)],
      [('　3名で、金曜を除く平日16:00〜稼働', 13, INK)]],
     body_size=13, body_spc=1.30)
card(s, 14.06, 4.05, 13.17, 1.30, 3.85, [('この軸で狙うこと', 20, WHITE)],
     [[('・5,000社から点検先を営業していく', 13, INK)],
      [('・提携ではなく「点検案件」が先に立つ流れ', 13, INK)],
      [('・本編の名簿獲得モデル（提携 → 名簿共有）とは', 13, INK)],
      [('　入口も出口も別のルート', 13, INK)],
      [('・進み方しだいで、本編とは別に数字が立つ', 13, INK)]],
     body_size=13, body_spc=1.30)
for x, hd, lines in [
        (LM, '本編との違い①｜母数',
         ['・本編の訪問1,080社は', '　自社リストから積んだ数字',
          '・5,000社はこことは別枠', '・重複の有無は未確認']),
        (9.47, '本編との違い②｜出口',
         ['・本編は「提携 → 名簿共有」', '・こちらは「点検案件の獲得」',
          '・KGIへの効き方が違うため', '　同じ表では管理しない']),
        (18.74, '本編との違い③｜管理',
         ['・本編の26社には含めない', '・件数・確度が見えた段階で',
          '　改めて数字を置く', '・当面は月次で進捗のみ共有'])]:
    card(s, x, 9.30, 8.54, 1.25, 4.40, [(hd, 18, WHITE)],
         [[(t, 14, INK)] for t in lines], body_align=PP_ALIGN.LEFT, body_spc=1.45,
         body_size=14)
rec('営業', '5,000社の内訳と、実際に訪問できる社数')
keymsg(s, '5,000社のパイプは、本編とは別に数字を立てる。')
notes(s, '本編の施策とは切り離した別軸。イー・スマイル高橋顧問（春日井市）が持つ'
         '東海圏の工務店・設備5,000社のパイプから点検先を営業していくルート。'
         '本編の訪問1,080社とは母数も出口も別なので、同じKPI表では管理しない。'
         '件数と確度が見えた段階で改めて数字を置く。主担当：営業。')



# ================= ページ番号の付与と保存 =================
for i, rec_ in enumerate(SLIDES, start=1):
    footer(rec_['s'], i, strip=rec_['strip'])

os.makedirs('output', exist_ok=True)
out = 'output/企画書_アライアンス提携_空パッケージ.pptx'
prs.save(out)

lines = ['# 確定待ち事項リスト（企画書 空パッケージ・打合せ資料版）', '',
         '本パッケージで空欄（【　】）のまま残している項目の一覧です。', '',
         '| P | 主担当 | 確定待ち事項 |', '|---|---|---|']
for pg, owner, pending in PENDING:
    lines.append('| %d | %s | %s |' % (pg, owner, pending))
open('output/確定待ち事項リスト.md', 'w', encoding='utf-8').write('\n'.join(lines) + '\n')

print('saved:', out, '/ slides:', len(prs.slides._sldIdLst))
