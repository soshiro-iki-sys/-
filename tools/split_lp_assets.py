# -*- coding: utf-8 -*-
"""data URI で画像を埋め込んだ1ファイル版LPを、HTML＋画像フォルダに分解する。

    python3 tools/split_lp_assets.py lp/イエオモイ_画像LP.html site/lp-image

出力:
    <出力先>/index.html      … 13KB 程度。テキストとしてそのまま扱える
    <出力先>/images/*.png    … 切り出した画像

1ファイル版は受け渡しが楽な代わりに 6MB あり、テキストとして開けない。
編集・コピペ・CDN配信にはこちらの分解版を使う。中身（見た目）は同一。
"""
import base64
import io
import os
import re
import sys

from PIL import Image

PATTERN = r'data:image/png;base64,[A-Za-z0-9+/=]+'


def name_for(im, counters):
    """画像の寸法から用途を推測してファイル名を決める。"""
    if im.size == (941, 1672):
        counters['section'] += 1
        return 'section%d.png' % counters['section']
    if im.size == (76, 76):
        counters['icon'] += 1
        return 'icon%d.png' % counters['icon']
    counters['logo'] += 1
    return ['logo-mark.png', 'logo-word.png', 'logo-full.png'][counters['logo'] - 1]


def main(lp_path, out_dir):
    html = open(lp_path, encoding='utf-8').read()
    img_dir = os.path.join(out_dir, 'images')
    os.makedirs(img_dir, exist_ok=True)

    counters = {'section': 0, 'icon': 0, 'logo': 0}
    written = []

    def replace(m):
        raw = base64.b64decode(m.group(0).split(',', 1)[1])
        im = Image.open(io.BytesIO(raw))
        name = name_for(im, counters)
        with open(os.path.join(img_dir, name), 'wb') as f:
            f.write(raw)
        written.append((name, len(raw)))
        return 'images/' + name

    html = re.sub(PATTERN, replace, html)
    out_html = os.path.join(out_dir, 'index.html')
    open(out_html, 'w', encoding='utf-8').write(html)

    for name, size in written:
        print('  images/%-14s %8.1f KB' % (name, size / 1024))
    print('%s  %.1f KB' % (out_html, len(html.encode()) / 1024))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
