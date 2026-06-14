# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## このリポジトリの目的

営業用プレゼンスライドを Markdown から生成するキット（[Marp](https://marp.app/) ベース）。
非エンジニアのユーザーがチャットで依頼し、Claude がスライドの作成・編集・PDF/PPTX 変換・GitHub への保存までを代行する想定。README とユーザー向けのやり取りは日本語が基本。

## コマンド

ビルド系のスクリプトは `package.json` に定義されており、入力 Markdown と出力先は実行時に渡す。

```bash
npm install                                              # 初回のみ（marp-cli を取得）

# PDF に変換（themes/sales.css が自動適用される）
npx marp slides/<file>.md --theme themes/sales.css -o output/<name>.pdf
# または: npm run pdf -- slides/<file>.md -o output/<name>.pdf

# PowerPoint に変換
npx marp slides/<file>.md --theme themes/sales.css --pptx -o output/<name>.pptx

npm run watch -- slides/<file>.md                        # 編集プレビュー（ライブ更新）
```

`output/` は `.gitignore` 済みのため、生成物はコミットされない。テスト・Lint の仕組みは無い。

## 構成と編集時の約束ごと

- `templates/営業スライド_テンプレート.md` … 新規スライドの雛形。新しい資料はこれを `slides/` にコピーして作る。
- `slides/` … 実際の資料（記入済み）。ファイル名・見出しは日本語。
- `themes/sales.css` … 全スライド共通の見た目。`@theme sales` で、各 Markdown の front-matter の `theme: sales` と対応する。

各スライド Markdown は必ず以下の front-matter で始まる：

```yaml
---
marp: true
theme: sales
paginate: true
---
```

### Marp の記法（このテーマ固有の挙動）

- `---`（水平線）でスライドが分割される。
- `<!-- _class: title -->` … 表紙（青グラデーション背景）。通常 `<!-- _paginate: false -->` と併用。
- `<!-- _class: section -->` … 章扉（中央寄せ）。
- `<!-- _class: lead -->` … 強調メッセージ用（中央寄せ）。
- `**強調**` は `themes/sales.css` でアクセントカラー（オレンジ）になる。`blockquote` はお客様の声・メッセージ用にスタイル済み。

### デザイン変更

色やフォントの変更依頼（例：「テーマの色を緑に」）は `themes/sales.css` の `:root` 内の CSS 変数（`--brand`, `--accent` など）を編集する。個別 Markdown は触らない。

## 作業フロー

ユーザーは Markdown を意識せず、チャットで「資料を作って／書き換えて／PDF にして」と依頼する。完成・修正後は指定ブランチへ commit & push して GitHub に保存する（生成した PDF/PPTX 自体はコミットしない）。
