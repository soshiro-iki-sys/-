# -*- coding: utf-8 -*-
"""撮影指示書ビルダー（マルユウハウジー お片付けサイト）

書式は templates/撮影指示書_フォーマット見本.pptx（遺品整理専門店 撮影指示書）を踏襲。
仕様は docs/撮影指示書フォーマット仕様.md のとおり。研修資料の書式とは別系統。

  A 表紙              1枚
  （一覧）            1枚   ← C型レイアウトでカット一覧
  B カットページ      52枚  ← 本体
  C Before/After 指示 1枚
  D 当日チェックリスト 1枚
                    ─────
                     56枚

実行： python3.13 slides/build/build_photobook.py
      （このコンテナでは python3 が 3.11、python-pptx は 3.13 側）
"""
import os
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from photobook_cuts import CUTS, GROUPS  # noqa: E402

# ── 基本設定 ──────────────────────────────────────────────
PAGE_W, PAGE_H = 11.692708, 8.267361       # A4横（297×210mm）
M = 0.47                                    # 左右マージン（版面幅 10.76）
BODY_W = PAGE_W - 2 * M

NAVY, BLUE, SKY = '102D46', '23658D', 'B9D9E9'
TEXT, SUB, RULE = '263843', '62747E', 'CBD8DE'
TINT1, TINT2 = 'F3F7F9', 'E8F1F5'
NG_BG, NG_LINE, NG_TEXT = 'FBF4F1', 'D6AA9E', '884333'
FONT = 'Noto Sans CJK JP'

SPEC_ROWS = ['撮影場所', '登場人物', '服装', '使用する小物',
             '動作', 'カメラ位置・構図', '表情']

CLIENT = 'マルユウハウジー'
SITE = 'お片付けサイト'
YEAR = '2026'

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   '..', '2026撮影指示書_マルユウハウジー_お片付けサイト.pptx')


# ── 図形ヘルパー ───────────────────────────────────────────
def rect(s, x, y, w, h, fill=None, line=None, lw=0.75, rounded=False):
    shp = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = RGBColor.from_string(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = RGBColor.from_string(line)
        shp.line.width = Pt(lw)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    shp.text_frame.text = ''
    return shp


def tb(s, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    return box, tf


def put(tf, text, size=10.5, bold=False, color=TEXT, align=PP_ALIGN.LEFT,
        space_after=0, line_spacing=None):
    """テキストフレームに1段落を書く（複数回呼べば段落が増える）。"""
    p = tf.paragraphs[0] if not tf.paragraphs[0].runs and not tf.paragraphs[0].text \
        else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    if line_spacing:
        p.line_spacing = line_spacing
    r = p.add_run()
    r.text = text
    f = r.font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = RGBColor.from_string(color)
    f.name = FONT
    rPr = r._r.get_or_add_rPr()
    for tag in ('a:latin', 'a:ea', 'a:cs'):
        el = rPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}'
                      + tag.split(':')[1])
        if el is None:
            from lxml import etree
            el = etree.SubElement(
                rPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}'
                + tag.split(':')[1])
        el.set('typeface', FONT)
    return p


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def fit(text, width_in, size, max_lines):
    """全角換算で収まる文字サイズを返す（収まらなければ段階的に縮める）。"""
    for s in (size, size - 0.75, size - 1.5, size - 2.25):
        per_line = max(1, int(width_in / (s / 72.0)))
        if len(text) <= per_line * max_lines:
            return s
    return size - 2.25


# ── A. 表紙 ───────────────────────────────────────────────
def cover(prs):
    s = blank(prs)
    rect(s, 0, 0, 0.28, PAGE_H, fill=NAVY)

    _, tf = tb(s, 0.91, 1.23, 7.29, 0.31)
    put(tf, 'PHOTO DIRECTION BOOK　　　/　　　' + YEAR,
        size=12, bold=True, color=BLUE)

    _, tf = tb(s, 0.91, 2.08, 9.38, 0.57)
    put(tf, '%s　%s' % (CLIENT, SITE), size=32.25, color=NAVY)

    _, tf = tb(s, 0.91, 2.78, 9.38, 0.78)
    put(tf, '撮影指示書', size=42.75, bold=True, color=NAVY)

    _, tf = tb(s, 0.91, 3.92, 9.38, 0.28)
    put(tf, '全52カット（優先度 S 21／A 31）　撮影2〜3日想定',
        size=11.25, color=SUB)
    return s


# ── C型ページの共通ヘッダー ────────────────────────────────
def section_head(s, title, lead):
    rect(s, M, 0.19, BODY_W, 0.04, fill=NAVY)
    _, tf = tb(s, M, 0.72, 10.42, 0.62)
    put(tf, title, size=26.25, color=NAVY)
    _, tf = tb(s, M, 1.38, 10.42, 0.34)
    put(tf, lead, size=12, color=SUB)


def must_bar(s, y, label, body):
    rect(s, M, y, BODY_W, 0.58, fill=TINT2)
    rect(s, M, y, 0.04, 0.58, fill=BLUE)
    _, tf = tb(s, 0.67, y + 0.13, 0.73, 0.3)
    put(tf, label, size=13.5, bold=True, color=NAVY)
    _, tf = tb(s, 1.50, y + 0.11, 9.38, 0.34)
    put(tf, body, size=13.5, bold=True, color=NAVY)


def table(s, y0, cols, rows, pitch=0.655, head_h=0.47, row_h=None,
          tint_first=True):
    """矩形＋テキストだけで表を描く（表オブジェクトは使わない）。"""
    row_h = row_h or (pitch - 0.045)
    xs, acc = [], M
    for _, w in cols:
        xs.append(acc)
        acc += w
    rect(s, M, y0, BODY_W, head_h, fill=NAVY)
    for (label, w), x in zip(cols, xs):
        _, tf = tb(s, x + 0.13, y0 + 0.12, w - 0.2, 0.26)
        put(tf, label, size=10.5, bold=True, color='FFFFFF')
    y = y0 + head_h + 0.04
    for row in rows:
        if tint_first and row[0]:
            rect(s, xs[0], y, cols[0][1], row_h, fill=TINT1)
        for i, ((_, w), x) in enumerate(zip(cols, xs)):
            val = row[i]
            if val == '':
                continue
            sz = fit(val, w - 0.26, 10.5, 2)
            _, tf = tb(s, x + 0.13, y + 0.07, w - 0.26, row_h - 0.1)
            put(tf, val, size=sz, bold=(i == 0),
                color=NAVY if i == 0 else TEXT)
        rect(s, M, y + row_h, BODY_W, 0.01, fill=RULE)
        y += pitch
    return y


# ── 一覧ページ ────────────────────────────────────────────
def index_page(prs):
    s = blank(prs)
    section_head(s, '撮影カット一覧',
                 '優先度 S・A のみ。S＝これが無いとページが作れない／A＝説得力が出る。')

    desc = {
        'A': 'メインビジュアル候補と朝礼',
        'B': '見積・養生・電話・買取・ワンストップ・集合',
        'C': '6サービス × 象徴＋作業（清掃・供養は3枚）',
        'D': '見積書の提示と、トラック1台分の物量',
        'E': '6ステップを1枚ずつ。6枚そろって初めて成立',
        'F': 'BEFORE／AFTER／作業中／物量。まず1件をこの型で',
        'G': 'ロゴ入りトラック',
        'H': '社屋・スタッフ集合・代表',
        'I': '捜索・梱包・仏具・分別・階段搬出',
        'J': '冷蔵庫・洗濯機・ソファ・ベッド・テレビ・ピアノ',
        'L': '手元アップと、文字乗せ用の余白カット',
    }
    rows = []
    for g, place in GROUPS:
        n = len([c for c in CUTS if c['no'].split('-')[0] == g])
        rows.append([g, place, desc[g], '%d' % n])
    rows.append(['', '合計', '', '52'])

    cols = [('群', 0.80), ('使う場所', 2.70), ('主なカット', 6.26), ('枚数', 1.00)]
    y = table(s, 1.95, cols, rows, pitch=0.40, head_h=0.40, row_h=0.38)

    _, tf = tb(s, M, y + 0.22, 10.73, 0.5)
    put(tf, '1カット＝1ページ。各ページは「写真（完成イメージ）＋仕様7項目＋カード3枚」で完結する。',
        size=9.75, color=SUB)
    put(tf, '買取品の物撮り（K群）は優先度がすべてBのため、本書では扱わない。',
        size=9.75, color=SUB)
    return s


# ── B. カットページ ───────────────────────────────────────
def cut_page(prs, c):
    s = blank(prs)

    # 見出し
    _, tf = tb(s, M, 0.53, 1.05, 0.56, anchor=MSO_ANCHOR.MIDDLE)
    put(tf, c['no'], size=13.5, bold=True, color=BLUE)

    _, tf = tb(s, 1.60, 0.53, 7.92, 0.56, anchor=MSO_ANCHOR.MIDDLE)
    put(tf, c['name'], size=fit(c['name'], 7.92, 21, 1), color=NAVY)

    _, tf = tb(s, 9.20, 0.60, 2.03, 0.4, anchor=MSO_ANCHOR.MIDDLE)
    put(tf, 'S　必須' if c['pri'] == 'S' else 'A　推奨',
        size=10.5, bold=True, color=NAVY if c['pri'] == 'S' else SUB,
        align=PP_ALIGN.RIGHT)

    rect(s, M, 1.24, BODY_W, 0.02, fill=NAVY)

    # 写真（未撮影のため枠のみ）
    rect(s, 0.53, 1.45, 6.15, 4.22, fill=TINT1, line=RULE)
    _, tf = tb(s, 0.90, 3.26, 5.41, 0.6, anchor=MSO_ANCHOR.MIDDLE)
    put(tf, '撮影完成イメージ', size=13.5, bold=True, color=SUB,
        align=PP_ALIGN.CENTER)
    put(tf, c['name'], size=10.5, color=SUB, align=PP_ALIGN.CENTER)

    _, tf = tb(s, M, 5.77, 1.53, 0.22)
    put(tf, 'REFERENCE IMAGE', size=7.5, bold=True, color=BLUE)
    _, tf = tb(s, 2.03, 5.77, 2.19, 0.22)
    put(tf, '撮影完成イメージ', size=7.5, color=SUB)
    _, tf = tb(s, 5.10, 5.77, 1.59, 0.22)
    put(tf, '撮影前・参考枠', size=7.5, color=SUB, align=PP_ALIGN.RIGHT)

    # 仕様表（7行固定）
    for n, (label, val) in enumerate(zip(SPEC_ROWS, c['spec'])):
        y = 1.43 + 0.655 * n
        rect(s, 6.90, y, 4.33, 0.61, fill=TINT1)
        rect(s, 6.90, y, 0.03, 0.61, fill=SKY)
        _, tf = tb(s, 7.04, y + 0.08, 1.17, 0.46)
        put(tf, label, size=fit(label, 1.17, 9.75, 1), bold=True, color=NAVY)
        _, tf = tb(s, 8.22, y + 0.07, 2.86, 0.48)
        put(tf, val, size=fit(val, 2.86, 10.5, 2), color=TEXT)

    # 下部カード3枚
    cards = [
        ('01', 'HP使用場所', M, 2.84, 'FFFFFF', RULE, BLUE, NAVY, c['use']),
        ('02', 'この写真で伝えること', 3.45, 3.81, 'FFFFFF', RULE, BLUE, NAVY, c['tell']),
        ('03', '撮影時の注意', 7.40, 3.83, NG_BG, NG_LINE, NG_TEXT, NG_TEXT, c['note']),
    ]
    for num, head, x, w, bg, ln, numc, headc, body in cards:
        rect(s, x, 6.16, w, 1.50, fill=bg, line=ln)
        _, tf = tb(s, x + 0.17, 6.31, 0.25, 0.23)
        put(tf, num, size=9, bold=True, color=numc)
        _, tf = tb(s, x + 0.49, 6.30, w - 0.66, 0.24)
        put(tf, head, size=11.25, bold=True, color=headc)
        rect(s, x + 0.17, 6.57, w - 0.33, 0.01, fill=RULE)
        _, tf = tb(s, x + 0.17, 6.73, w - 0.34, 0.80)
        put(tf, body, size=fit(body, w - 0.34, 10.5, 4), color=TEXT,
            line_spacing=1.1)
    return s


# ── C. Before / After 撮影指示 ────────────────────────────
def before_after(prs):
    s = blank(prs)
    section_head(
        s, 'BEFORE / AFTER 撮影指示',
        '施工実績の核。同条件で撮れていない2枚は、並べて使えないため撮り直しになる。')
    must_bar(s, 2.03, '必須',
             'BEFORE を撮らずに作業を始めない。AFTER を撮らずに撤収しない。')

    cols = [('', 1.15), ('BEFORE（F-1）', 4.80), ('AFTER（F-2）', 4.81)]
    rows = [
        ['タイミング', '片付けに手をつける前', '撤収する前（作業の最後）'],
        ['三脚位置', '部屋の角・対角線方向。床にマーキングする', 'BEFOREのマーキング位置に据え直す'],
        ['高さ・画角', '記録しておく（例：高さ140cm／24mm）', 'BEFOREと同一にする'],
        ['縦横', '横位置', '横位置（BEFOREと同一）'],
        ['光・露出', '室内灯＋自然光。露出を固定する', '同じ時刻帯・同じ露出・同じWB'],
        ['部屋の状態', '物を1つも動かさない', '床・隅に何も残さない'],
        ['人物', '入れない（無人で撮る）', '入れない（無人で撮る）'],
    ]
    y = table(s, 2.90, cols, rows, pitch=0.50, head_h=0.47, row_h=0.46)

    _, tf = tb(s, M, y + 0.20, 10.73, 0.9)
    put(tf, 'あわせて、この7項目を現場で記録する（サイトの施工実績に数字として載せる）',
        size=10.5, bold=True, color=NAVY, space_after=4)
    put(tf, '間取り ／ 作業人数 ／ 作業日数（時間） ／ 使用車両 ／ 料金 ／ 買取額 ／ 所在地（市区まで）',
        size=11.25, color=TEXT, space_after=6)
    put(tf, '2件目以降は同じ型を繰り返すだけ。実案件が出たときに随時積み増していく。',
        size=9.75, color=SUB)
    return s


# ── D. 撮影当日チェックリスト ─────────────────────────────
def checklist(prs):
    s = blank(prs)
    section_head(s, '撮影当日チェックリスト',
                 '各ロケーションの撤収前に、撮影責任者と現場責任者で確認する。')

    cards = [
        ('人物・装い', [
            '作業着のロゴの向き・汚れを確認',
            '帽子・手袋の着脱を確認した',
            'お客様役・ご本人役の許諾を得た',
            '手元カットの爪・手の汚れを確認',
        ]),
        ('空間・小物', [
            '表札・書類・ナンバーが写っていない',
            '床・隅に不要な物が残っていない',
            '養生のめくれ・継ぎ目を直した',
            '見積書・札・ダミー類を用意した',
        ]),
        ('構図・納品', [
            '引き・寄り・手元の3枚を撮った',
            '縦トリミングでも成立する構図',
            'F-1とF-2が同一アングル',
            '施工実績の7項目を記録した',
        ]),
    ]
    for (head, items), x in zip(cards, (0.47, 4.05, 7.64)):
        rect(s, x, 1.98, 3.40, 4.14, fill='FFFFFF', line=RULE)
        rect(s, x, 1.98, 3.40, 0.59, fill=NAVY)
        _, tf = tb(s, x + 0.17, 2.11, 3.06, 0.34)
        put(tf, head, size=14.25, bold=True, color='FFFFFF')
        y = 2.72
        for i, it in enumerate(items):
            _, tf = tb(s, x + 0.15, y, 0.30, 0.35)
            put(tf, '□', size=17.25, color=BLUE)
            _, tf = tb(s, x + 0.51, y + 0.02, 2.70, 0.62)
            put(tf, it, size=fit(it, 2.70, 10.5, 2), color=TEXT,
                line_spacing=1.1)
            if i < len(items) - 1:
                rect(s, x, y + 0.72, 3.40, 0.01, fill=RULE)
            y += 0.885

    rect(s, M, 6.51, BODY_W, 0.02, fill=NAVY)
    for label, x in (('撮影日：________________', 0.47),
                     ('撮影責任者：________________', 4.08),
                     ('現場責任者：________________', 7.79)):
        _, tf = tb(s, x, 6.77, 3.40, 0.36)
        put(tf, label, size=11.25, color=TEXT)
    return s


# ── 組み立て ──────────────────────────────────────────────
def main():
    prs = Presentation()
    prs.slide_width = Emu(int(PAGE_W * 914400))
    prs.slide_height = Emu(int(PAGE_H * 914400))

    cover(prs)
    index_page(prs)
    for c in CUTS:
        cut_page(prs, c)
    before_after(prs)
    checklist(prs)

    out = os.path.normpath(OUT)
    prs.save(out)
    print('%s  (%d slides)' % (out, len(prs.slides.__iter__.__self__._sldIdLst)))


if __name__ == '__main__':
    main()
