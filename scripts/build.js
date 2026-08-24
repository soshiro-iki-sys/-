const K = require("./build_nagaoka_deck.js");
require("./slides_part1.js")(K);
require("./slides_part2.js")(K);
require("./slides_part3.js")(K);

const fs = require("fs");
const out = "/home/user/-/output/長岡_ヤシロ_補助金活用講座.pptx";
K.pres.writeFile({ fileName: out }).then(() => {
  console.log("written:", out);
  // 差し替え一覧
  let md = "# 差し替え箇所一覧 ─ 長岡_ヤシロ_補助金活用講座.pptx\n\n";
  md += "`.pptx` の空欄（表の値セル・グラフのデータ点・背景写真）の一覧。\n";
  md += "表の項目名・グラフの軸ラベル／凡例はすべて記入済みのため、値を入れるだけで完成します。\n\n";
  md += "| ページ | 要素 | 埋める内容 |\n|---|---|---|\n";
  K.notes.forEach((n) => { md += `| P${n.page} | ${n.element} | ${n.item} |\n`; });
  md += "\n## 事実確認が必要な箇所\n\n";
  md += "| ページ | 内容 |\n|---|---|\n";
  md += "| P8 | 東北電力の料金単価。登壇前に最新値を確認 |\n";
  md += "| P9 | 国の電気・ガス料金負担軽減支援事業の実施状況 |\n";
  md += "| P10 | 柏崎刈羽6号機の再稼働時期（2026年は見込み） |\n";
  md += "| P11・P12 | 中越地震の停電件数（約30万件）・復旧日数（11日）の一次出典 |\n";
  md += "| P13 | 2022年12月の大雪による停電規模、佐渡の停電日数（約9日間）の一次出典 |\n";
  md += "| P25 | ZEH／ZEH＋補助の対象要件（新築のみか） |\n";
  md += "| P26 | 補助金の受付状況・残予算 |\n";
  md += "| P38 | 太陽光トラブルの相談件数の出典（国民生活センター等） |\n";
  fs.writeFileSync("/home/user/-/output/差し替え箇所一覧.md", md, "utf8");
  console.log("notes:", K.notes.length);
}).catch((e) => { console.error(e); process.exit(1); });
