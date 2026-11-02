# 企画書の生成方法

`build_deck.js` から `proposals/遺品整理事業_企画書.pptx` を生成します（pptxgenjs）。

```bash
npm install pptxgenjs react react-dom react-icons sharp
PPTX_SKILL=<pptxスキルのパス> node build_deck.js ../遺品整理事業_企画書.pptx
```

※通常は生成済みの `.pptx` を PowerPoint で直接編集して使ってください。
