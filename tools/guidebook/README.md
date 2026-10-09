# 外装工事ガイドブック（勉強会参加者特典）

- `guidebook.html`：本文（A4縦・12ページ）。文言や写真枠はここを直す
- `build_pptx.js`：PowerPoint版（A4縦・12ページ、すべて編集可能）を作る
- `render.js`：HTML版をPDFに書き出す

フォント（Noto Sans JP）はサイズが大きいため同梱していません。書き出す前に `fonts/` に置いてください。

```bash
mkdir -p fonts
curl -sS -A "Mozilla/5.0" "https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700;900" \
  | grep -o 'https://fonts.gstatic.com[^)]*' | paste - <(printf "400\n700\n900") \
  | while read u w; do curl -sS -o fonts/NotoSansJP-$w.ttf "$u"; done
NODE_PATH=$(npm root -g) node render.js "$PWD" "$PWD/guidebook.pdf"
```

## PowerPoint版

```bash
NODE_PATH=<pptxgenjsのあるnode_modules> node build_pptx.js out.pptx
```

フォントはメイリオを指定しています。
