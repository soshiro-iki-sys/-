# -*- coding: utf-8 -*-
"""画像に焼き込まれた LP の配色を、イエオモイ ロゴ準拠のブランドカラーへ置換する。

    python3 tools/recolor_lp_images.py lp/イエオモイ_画像LP.html

HTML に data URI で埋め込まれた PNG を取り出し、ベクター由来の
バーミリオン赤／ネイビー／レモンイエロー／スレートだけを
テラコッタ／コーヒーブラウン／マスタードゴールド／トープへ写して埋め戻す。
写真は触らない（判定は graphic_core を参照）。

必要: pillow, numpy, scipy
"""
import base64
import io
import re
import sys

import numpy as np
from PIL import Image
from scipy.ndimage import (binary_closing, binary_dilation, binary_fill_holes,
                           label, uniform_filter)

# --- リポジトリ既定のブランドカラー（lp/イエオモイ_LP.html の :root と同一） ---
BRAND_RED = (0xd8, 0x4a, 0x2e)    # --red   テラコッタ
BRAND_BROWN = (0x5f, 0x3f, 0x33)  # --brown  コーヒーブラウン
BRAND_YELLOW = (0xf5, 0xb3, 0x1c)  # --yellow マスタードゴールド
SRC_RED = (0xe6, 0x00, 0x12)
SRC_NAVY = (0x1b, 0x2a, 0x5a)
BRAND_TAUPE = (0x8a, 0x68, 0x5c)   # --taupe  トープ
SRC_YELLOW = (0xfa, 0xe2, 0x0f)
SRC_SLATE = (0x7c, 0x87, 0xa7)


def rgb2hsv(a):
    r, g, b = a[..., 0], a[..., 1], a[..., 2]
    mx = a.max(-1)
    d = mx - a.min(-1)
    h = np.zeros_like(mx)
    m = d > 1e-6
    i = m & (mx == r); h[i] = ((g - b)[i] / d[i]) % 6
    i = m & (mx == g) & ~(mx == r); h[i] = ((b - r)[i] / d[i]) + 2
    i = m & (mx == b) & ~(mx == r) & ~(mx == g); h[i] = ((r - g)[i] / d[i]) + 4
    return h * 60.0, np.where(mx > 1e-6, d / np.maximum(mx, 1e-6), 0.0), mx


def hsv2rgb(h, s, v):
    h = np.mod(h, 360.0) / 60.0
    i = np.floor(h).astype(int) % 6
    f = h - np.floor(h)
    p, q, t = v * (1 - s), v * (1 - s * f), v * (1 - s * (1 - f))
    return np.stack([np.choose(i, [v, q, p, p, t, v]),
                     np.choose(i, [t, v, v, q, p, p]),
                     np.choose(i, [p, p, t, v, v, q])], -1)


def hsv_of(c):
    h, s, v = rgb2hsv(np.array(c, float).reshape(1, 1, 3) / 255.0)
    return float(h[0, 0]), float(s[0, 0]), float(v[0, 0])


def hue_dist(h, c):
    return np.abs(np.mod(h - c + 180.0, 360.0) - 180.0)


def graphic_core(a, tight, v, s, h, ref, min_size, max_v_std, max_h_std,
                 min_white=0.35, max_ref_dist=0.25, floor_white=0.10, big=1200):
    """置換してよい画素（＝ベクター由来）を連結成分ごとに決める。

    次のどちらかを満たす成分を「図形」とみなす。
      A) 一定サイズ以上で、成分内の色のばらつきが小さい（ベタ塗り・グラデーション）。
         写真中の紺の作業着や屋根は陰影があるため vstd が大きく、ここで落ちる。
      B) 周囲に白（＝カード面や余白）が十分あり、かつ基準色に近い。
         細い文字や小さなアイコンは成分が小さく A の統計が効かないため、こちらで拾う。
         写真の中の紺の作業着は周囲に純白がないので通らない（実測 0.02 対 0.6〜0.8）。
    """
    lab, n = label(tight, structure=np.ones((3, 3)))
    if n == 0:
        return np.zeros_like(tight)
    fl = lab.ravel()
    cnt = np.bincount(fl, minlength=n + 1)[1:]
    mean = lambda x: np.bincount(fl, weights=x.ravel(), minlength=n + 1)[1:] / cnt
    hh = np.mod(h + 180.0, 360.0)  # 赤の 0/360 またぎを避ける
    with np.errstate(invalid='ignore', divide='ignore'):
        vstd = np.sqrt(np.maximum(mean(v * v) - mean(v) ** 2, 0))
        hstd = np.sqrt(np.maximum(mean(hh * hh) - mean(hh) ** 2, 0))
        white = mean(uniform_filter(((v > 0.92) & (s < 0.10)).astype(float), 21))
        dist = mean(np.sqrt(((a - np.array(ref, float) / 255.0) ** 2).sum(-1)))

    # 中〜小サイズの成分は A だけでは写真の中の作業着の一部などを拾ってしまう。
    # 周囲にカード面の純白がどれだけあるか（実測：文字 0.4〜0.8／写真の中 0.00〜0.03）
    # で足切りする。大きなベタ塗り（ボタン・リボン・見出し帯）は内部に白が無いので
    # この足切りから外し、A のばらつき判定にまかせる。
    flat_ok = (cnt >= min_size) & (vstd <= max_v_std) & (hstd <= max_h_std)
    on_light = (white > min_white) & (dist < max_ref_dist)
    keep = np.zeros(n + 1, bool)
    keep[1:] = (flat_ok & ((cnt >= big) | (white >= floor_white))) | on_light
    core = keep[lab]
    # 採用した図形の内側に閉じ込められた小片（アイコン内の窓など）も一緒に置換する
    inside = binary_fill_holes(binary_closing(core, np.ones((7, 7))))
    return core | (tight & inside)


def remap(a, mask, src, tgt):
    h, s, v = rgb2hsv(a)
    _, ss, sv = hsv_of(src)
    th, ts, tv = hsv_of(tgt)
    out = hsv2rgb(np.full_like(h, th), np.clip(s * (ts / ss), 0, 1), np.clip(v * (tv / sv), 0, 1))
    w = mask[..., None].astype(float)
    return a * (1 - w) + out * w


def recolor(im):
    """RGB 画像 1 枚をブランドカラーへ写して返す。"""
    a = np.asarray(im.convert('RGB'), float) / 255.0
    h, s, v = rgb2hsv(a)

    # 判定のしきい値は実測にもとづく。写真内の紺（作業着・屋根）は vstd>=0.07 で
    # 落ち、ベクター図形は vstd<=0.055 に収まる。赤はグラデーションボタンがあるため緩め。
    tight_red = (hue_dist(h, 358.0) < 12) & (s > 0.78) & (v > 0.55)
    tight_navy = (hue_dist(h, 223.0) < 22) & (s > 0.58) & (v > 0.08) & (v < 0.55)
    tight_yellow = (h > 36) & (h < 62) & (s > 0.55) & (v > 0.72)
    # 罫線・区切り線などのくすんだ青（低彩度）はブランドのトープへ
    tight_slate = (h > 200) & (h < 250) & (s > 0.12) & (s < 0.58) & (v > 0.35) & (v < 0.92)

    core_red = graphic_core(a, tight_red, v, s, h, SRC_RED, 30, 0.100, 2.0)
    core_navy = graphic_core(a, tight_navy, v, s, h, SRC_NAVY, 30, 0.055, 5.0)
    core_yellow = graphic_core(a, tight_yellow, v, s, h, SRC_YELLOW, 120, 0.030, 3.0)
    # 罫線類は写真の中の青みと紛れやすいので、白地の上にあることを強く要求する
    core_slate = graphic_core(a, tight_slate, v, s, h, SRC_SLATE, 40, 0.120, 12.0,
                              min_white=1.1, floor_white=0.40, big=10 ** 9)

    # アンチエイリアスの縁まで含める（同系色相・低彩度側へ 2px 膨張）
    near_red = (hue_dist(h, 358.0) < 22) & (s > 0.06)
    near_navy = (hue_dist(h, 223.0) < 34) & (s > 0.06) & (v < 0.98)
    near_yellow = (h > 30) & (h < 68) & (s > 0.10)
    m_red = binary_dilation(core_red, np.ones((5, 5))) & near_red
    m_navy = binary_dilation(core_navy, np.ones((5, 5))) & near_navy & ~m_red
    m_yellow = binary_dilation(core_yellow, np.ones((5, 5))) & near_yellow & ~m_red & ~m_navy
    near_slate = (h > 190) & (h < 260) & (s > 0.05)
    m_slate = binary_dilation(core_slate, np.ones((3, 3))) & near_slate & ~m_red & ~m_navy & ~m_yellow

    a = remap(a, m_red, SRC_RED, BRAND_RED)
    a = remap(a, m_navy, SRC_NAVY, BRAND_BROWN)
    a = remap(a, m_yellow, SRC_YELLOW, BRAND_YELLOW)
    a = remap(a, m_slate, SRC_SLATE, BRAND_TAUPE)
    print('  red %6d  navy %6d  yellow %5d  slate %5d'
          % (m_red.sum(), m_navy.sum(), m_yellow.sum(), m_slate.sum()))
    return Image.fromarray(np.clip(a * 255 + 0.5, 0, 255).astype(np.uint8))


def recolor_flat_icon(im, src=(20, 32, 74), tgt=BRAND_BROWN, tol=90):
    """単色ベタのアイコン（RGBA）を、アルファを保ったまま塗り替える。"""
    a = np.asarray(im.convert('RGBA')).copy()
    rgb = a[..., :3]
    hit = (np.abs(rgb.astype(int) - np.array(src)).sum(-1) < tol) & (a[..., 3] > 0)
    rgb[hit] = tgt
    a[..., :3] = rgb
    return Image.fromarray(a)


def encode(im):
    buf = io.BytesIO()
    im.save(buf, format='PNG', optimize=True)
    return 'data:image/png;base64,' + base64.b64encode(buf.getvalue()).decode()


def main(path):
    html = open(path, encoding='utf-8').read()
    for i, uri in enumerate(re.findall(r'data:image/png;base64,[A-Za-z0-9+/=]+', html)):
        im = Image.open(io.BytesIO(base64.b64decode(uri.split(',', 1)[1])))
        print('image %d %s' % (i, im.size))
        new = recolor_flat_icon(im) if im.mode == 'RGBA' else recolor(im)
        html = html.replace(uri, encode(new))
    open(path, 'w', encoding='utf-8').write(html)
    print('updated', path)




if __name__ == '__main__':
    main(sys.argv[1])
