# アプローチブック誌面画像（2026年版 p.9〜p.19）

必要性訴求パートの研修資料（`slides/build/build_need.py`）が読み込む誌面画像です。

    TRAINING_ASSETS=assets/approachbook python3 slides/build/build_need.py

## 作り方

1. `スマートハウスのある暮らし`（アプローチブック本体 .pptx）を PDF に書き出す
2. `pdftoppm -r 150 -png ab.pdf ab` で 1ページ＝1画像にする
3. **下端13％（誌面の緑帯と著作権表示）を切り落とす**
   緑帯は研修資料の側で正しい文言を作り直すため、画像には含めない

```python
from PIL import Image
im = Image.open('ab-10.png').convert('RGB')
w, h = im.size
im.crop((0, 0, w, int(h * 0.868))).save('ab10.jpg', quality=88)
```

## 出どころ

| ページ | 出どころ |
|---|---|
| p.10・p.11・p.12・p.13・p.15・p.16・p.17 | 2026年版アプローチブック本体から書き出し |
| p.9・p.14・p.18・p.19 | 既存の研修資料に貼られていた画像（内容は2026年版と同一。書き出しが綺麗なためこちらを流用） |

LibreOffice で PDF 化すると日本語フォントが置換され、グラフのデータラベルが
折り返してしまうページがあります。綺麗な画像が既にあるページはそちらを使ってください。
