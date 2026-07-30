# -*- coding: utf-8 -*-
"""イエオモイのロゴを画像版LPのヘッダー／フッターに差し込む。

    python3 tools/insert_logo.py assets/logo/イエオモイロゴ.jpg lp/イエオモイ_画像LP.html

元データは黒地なので、まず背景を抜いて透過PNGにする。単純な閾値だと縁に黒い
にじみが残るので、いちばん近い不透明画素の色を「本来の色」とみなし、
観測値との比からアルファを求めて色を戻す（黒地への合成の逆算）。

そのうえでマーク部とワードマーク部を切り分け、
  ヘッダー: マーク＋ワードマークの横組みロックアップ
  フッター: 縦組みのロゴ全体
として data URI で埋め込む。何度実行しても結果は同じ（既存の差し込みを置換）。

必要: pillow, numpy, scipy
"""
import base64
import io
import re
import sys

import numpy as np
from PIL import Image
from scipy.ndimage import distance_transform_edt

MARKER_CSS = '/* === ロゴ（ヘッダー／フッター） === */'
MARKER_HEADER = '<!-- ロゴヘッダー -->'
MARKER_FOOTER = '<!-- ロゴフッター -->'


def knockout_black(im, core_level=60, floor=6):
    """黒地に合成されたロゴを、ストレートアルファの RGBA に戻す。"""
    a = np.asarray(im.convert('RGB'), float)
    mx = a.max(-1)
    core = mx > core_level
    # 各画素にいちばん近い「確実に不透明」な画素の色＝本来の色
    idx = distance_transform_edt(~core, return_distances=False, return_indices=True)
    pure = a[idx[0], idx[1]]
    alpha = np.clip(mx / np.maximum(pure.max(-1), 1e-6), 0, 1)
    alpha[core] = 1.0
    alpha[mx < floor] = 0.0
    return Image.fromarray(np.dstack([pure, alpha * 255]).astype(np.uint8), 'RGBA')


def split_lockup(im, min_gap=8):
    """マーク部とワードマーク部を、間の空白行で切り分ける。"""
    opaque = np.asarray(im)[..., 3] > 90
    rows = opaque.sum(1)
    runs, start = [], None
    for y, n in enumerate(rows):
        if n == 0:
            start = y if start is None else start
        elif start is not None:
            runs.append((start, y - 1))
            start = None
    inner = [r for r in runs if r[1] - r[0] >= min_gap and 0 < r[0] and r[1] < len(rows) - 1]
    if not inner:
        raise SystemExit('マークとワードマークの区切りが見つかりませんでした')
    cut = max(inner, key=lambda r: r[1] - r[0])
    trim = lambda x: x.crop(x.getbbox())
    return trim(im.crop((0, 0, im.width, cut[0]))), trim(im.crop((0, cut[1] + 1, im.width, im.height)))


def data_uri(im, height):
    """指定の表示高さの2倍でエンコードする（Retina 対応）。"""
    scale = height * 2 / im.height
    im = im.resize((max(1, round(im.width * scale)), height * 2), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format='PNG', optimize=True)
    return 'data:image/png;base64,' + base64.b64encode(buf.getvalue()).decode()


CSS = MARKER_CSS + '''
  .lp-header {
    display: flex; align-items: center; justify-content: space-between;
    gap: 16px; padding: 13px 22px; background: #fff;
    border-bottom: 1px solid #ece2d8;
  }
  .lp-header .lock { display: flex; align-items: center; gap: 12px; }
  .lp-header .lock img { display: block; width: auto; }
  .lp-header .lock .mark { height: 46px; }
  .lp-header .lock .word { height: 21px; }
  .lp-header .tag {
    color: #5f3f33; font-size: 13px; font-weight: 700;
    letter-spacing: .06em; white-space: nowrap;
  }
  .lp-footer {
    padding: 30px 20px 34px; background: #fff;
    border-top: 1px solid #ece2d8; text-align: center;
  }
  .lp-footer img { display: inline-block; height: 76px; width: auto; }
  @media (max-width: 480px) {
    .lp-header { padding: 10px 14px; }
    .lp-header .lock { gap: 9px; }
    .lp-header .lock .mark { height: 36px; }
    .lp-header .lock .word { height: 16px; }
    .lp-header .tag { display: none; }
    .lp-footer img { height: 62px; }
  }
'''


def strip_existing(html):
    html = re.sub(re.escape(MARKER_CSS) + r'.*?\n(?=\s*/\* ===|\s*\.lp \{|</style>)', '',
                  html, flags=re.S)
    html = re.sub(r'[ \t]*' + re.escape(MARKER_HEADER) + r'.*?</header>\n', '', html, flags=re.S)
    html = re.sub(r'[ \t]*' + re.escape(MARKER_FOOTER) + r'.*?</footer>\n', '', html, flags=re.S)
    return html


def main(logo_path, lp_path):
    logo = knockout_black(Image.open(logo_path))
    logo = logo.crop(logo.getbbox())
    mark, word = split_lockup(logo)
    print('mark %s / word %s / full %s' % (mark.size, word.size, logo.size))

    html = strip_existing(open(lp_path, encoding='utf-8').read())
    html = html.replace('</style>', CSS + '</style>', 1)

    header = ('    ' + MARKER_HEADER + '\n'
              '    <header class="lp-header">\n'
              '      <div class="lock">\n'
              '        <img class="mark" src="%s" alt="イエオモイ">\n'
              '        <img class="word" src="%s" alt="">\n'
              '      </div>\n'
              '      <div class="tag">住まいの定額おまかせパック</div>\n'
              '    </header>\n' % (data_uri(mark, 46), data_uri(word, 21)))
    footer = ('    ' + MARKER_FOOTER + '\n'
              '    <footer class="lp-footer">\n'
              '      <img src="%s" alt="イエオモイ">\n'
              '    </footer>\n' % data_uri(logo, 76))

    anchor = '    <!-- セクション1 -->'
    assert html.count(anchor) == 1
    html = html.replace(anchor, header + anchor, 1)
    assert html.count('  </main>') == 1
    html = html.replace('  </main>', footer + '  </main>', 1)

    open(lp_path, 'w', encoding='utf-8').write(html)
    print('updated', lp_path)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
