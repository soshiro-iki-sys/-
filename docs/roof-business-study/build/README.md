# ビルド手順

```
pip install python-pptx pillow
cd docs/roof-business-study/build
python3 build2.py   # 新人コンサル版（36枚）
python3 build.py    # 現場向け（35枚）
```

## 検証

| スクリプト | 内容 |
|---|---|
| `overlap.py` / `overlap1.py` | 図形どうしの重なり・スライド枠外を幾何的に検出 |
| `verify2.py` | テキストが自身の枠を縦にはみ出していないかを検出 |
| `audit2.py` / `audit_p.py` | サイズ・フォント・配色・フッターの規定順守を監査 |

目視確認は LibreOffice で PDF 化 → `pdftoppm` で画像化して行う。
`libreoffice-impress` と `poppler-utils` が必要。

```
soffice --headless --convert-to pdf <file>.pptx --outdir prev
pdftoppm -png -r 100 prev/<file>.pdf prev/p
```

## レイアウトの原則（deck_lib.py）

- テキストは実フォントで幅を**実測**し、枠に収まらなければ自動縮小する（`textbox` / `label` / `panel`）
- 表は行高をテキスト量から自動決定する。`TABLE_FIXED = True` にすると行高を固定し、
  代わりにフォントを縮小する（既存の固定レイアウトを崩さないため `build.py` はこちら）
- 「・」で始まる行は折り返しをぶら下げ揃えにする
- 注記は `note_box` で内容量ぴったりの高さにし、下端を揃える
