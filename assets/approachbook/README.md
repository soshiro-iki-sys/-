# アプローチブック誌面画像（2026年版 p.9〜p.19）

必要性訴求パートの研修資料（`slides/build/build_need.py`）が読み込む誌面画像です。
**下端のキーメッセージ帯（緑帯）まで含めた1ページ丸ごと**を貼っています。

    TRAINING_ASSETS=assets/approachbook python3 slides/build/build_need.py

## 作り方

1. `スマートハウスのある暮らし`（アプローチブック本体 .pptx）をコピーする
2. **はみ出すシェイプの文字サイズだけを縮める**（下記）
3. PDF に書き出し、`pdftoppm -r 150 -png ab2.pdf n` で1ページ＝1画像にする
4. p.9〜p.19 を `ab09.jpg` 〜 `ab19.jpg` として保存（トリミングはしない）

### なぜ文字サイズを縮めるのか

LibreOffice で PDF 化すると日本語フォントが置換され、本来より文字幅が広くなります。
そのため下端のキーメッセージ帯が版面からはみ出し、p.9 ではグラフのデータラベルが
折り返してしまいます。**該当シェイプの文字サイズだけを縮めて回避**します。
誌面の文言・レイアウトそのものは変えません。

```python
import re, shutil
from pptx import Presentation
from pptx.util import Pt

shutil.copy('ab.pptx', 'ab2.pptx')
prs = Presentation('ab2.pptx'); slides = list(prs.slides)
BAR_SCALE = {9: 0.86, 12: 0.76, 13: 0.86, 14: 0.84, 15: 0.80, 17: 0.80, 18: 0.80}

def scale_runs(shape, sc, default=28):
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            cur = r.font.size.pt if r.font.size else default
            r.font.size = Pt(round(cur * sc, 1))

for pageno, sc in BAR_SCALE.items():            # 下端のキーメッセージ帯
    for sh in slides[pageno - 1].shapes:
        if (sh.has_text_frame and sh.top and sh.width
                and sh.top > 5900000 and sh.width > 6000000
                and sh.text_frame.text.strip()):
            scale_runs(sh, sc)

for sh in slides[8].shapes:                     # p.9 のグラフのデータラベル
    if sh.has_text_frame and re.fullmatch(r'¥[\d,]+', sh.text_frame.text.strip()):
        scale_runs(sh, 0.80, default=14)

prs.save('ab2.pptx')
```

本物の PowerPoint で書き出せる環境があれば、この前処理は不要です。
