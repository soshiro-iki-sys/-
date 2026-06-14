# 営業資料作成用ワークスペース 📊

営業のプレゼン資料（ピッチ）を **かんたん・きれい・速く** 作るためのセットアップです。
Markdown や git の知識は不要。**チャットで私（Claude）に頼むだけ** で、資料の作成・修正・PDF化・保存まで進みます。

> 詳しいワークフローとお作法は [`CLAUDE.md`](./CLAUDE.md) にまとまっています。

---

## 🚀 クイックスタート

いちばんかんたんなのは、チャットでこう頼むことです：

> 「○○社向けの提案資料を作って。商品は△△、相手は□□業界の会社です」

すると私が、**構成案 → スライド → PDF → 保存** の順に進めます。

途中でこんな頼み方もできます：

> 「先に構成案だけ見せて」
> 「料金の表を書き換えて」
> 「PDFにして」 / 「PowerPointにして」
> 「保存して」

---

## 📁 何が入っているか

| 場所 | 役割 |
|---|---|
| `CLAUDE.md` | ワークフロー・スキル早見表・お作法（中心の取説） |
| `README.md` | このクイックスタート |
| `drafts/` | **案件ごとの作業フォルダ**をここに切る（実作業の場所） |
| `templates/standard-pitch.md` | Marpベースのピッチ雛形（Addnessブランド：黒×青） |
| `templates/outline-template.md` | 構成案を最初に書くテンプレ |
| `assets/{images,logos,icons}/` | 画像・ロゴ・アイコン |
| `reference/` | 過去資料への近道（symlink）。型を学ぶ用 |
| `archive/` | 完了・お蔵入り案件の置き場 |
| `themes/addness.css` | ブランドテーマ（黒×青） |

---

## ✍️ 自分で書く場合の流れ

1. `templates/outline-template.md` をコピーして、まず**構成案（骨子）**を書く
2. 骨子が固まったら `templates/standard-pitch.md` をコピーして `drafts/＜案件名＞/` に置く
3. `◯◯` や `△△` を実際の内容に書き換える
4. 「PDFにして」「保存して」と私に伝える

### Markdown ミニ知識
- `#` … 大見出し / `##` … 中見出し
- `-` … 箇条書き / `**文字**` … 強調（青色になる）
- `---`（ハイフン3つ）… スライドのページ区切り

---

## 💻 手動でPDF/PowerPointにする場合

```bash
npm install   # 最初の1回だけ
npx marp drafts/＜案件名＞/index.md --theme themes/addness.css --allow-local-files -o output/資料.pdf
npx marp drafts/＜案件名＞/index.md --theme themes/addness.css --allow-local-files --pptx -o output/資料.pptx
```

---

困ったら、いつでもチャットで聞いてください 🙌
