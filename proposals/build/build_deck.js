// 遺品整理事業 参入のご提案（企画書）生成スクリプト
// 使い方: node build_deck.js <出力パス.pptx>
// デザイン：船井総研の提案書（76p）・企画書（96p）の型を踏襲（4:3、紺の見出し帯、紺×黄の中扉、赤字の結論）
const path = require("path");
const pptxgen = require("pptxgenjs");

const OUT = process.argv[2] || "遺品整理事業_企画書.pptx";

// ---------- パレット ----------
const NAVY = "1F2D50";
const NAVY2 = "2B3F6E";
const BLUE = "2E5FA8";
const LBLUE = "E9EFF8";
const RED = "E60012";
const YEL = "FFF000";
const GOLD = "C9A227";
const ORANGE = "E07B2E";
const GREEN = "2E7D4F";
const LGREEN = "EAF4EE";
const PURPLE = "5B3F8C";
const BROWN = "7A5230";
const GRAY = "666666";
const LGRAY = "F3F4F6";
const LINE = "C9D1DE";
const DARK = "222222";
const WHITE = "FFFFFF";

const FONT = "Meiryo";
const MINCHO = "Yu Mincho";

const THEME = {
  name: "Ihin Seiri Proposal",
  headFontFace: FONT,
  bodyFontFace: FONT,
  colors: {
    dk1: "000000", lt1: "FFFFFF", dk2: NAVY, lt2: "F3F4F6",
    accent1: NAVY, accent2: BLUE, accent3: RED, accent4: ORANGE, accent5: GREEN, accent6: PURPLE,
    hlink: BLUE, folHlink: PURPLE,
  },
};

const pres = new pptxgen();
pres.defineLayout({ name: "STD_43", width: 10, height: 7.5 });
pres.layout = "STD_43";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = "遺品整理事業 参入のご提案";
pres.author = "株式会社船井総合研究所";
pres.company = "株式会社船井総合研究所";

// ---------- レイアウト ----------
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: WHITE },
  objects: [
    { rect: { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY } } },
    { rect: { x: 6.2, y: 0, w: 3.8, h: 0.85, fill: { color: NAVY2 } } },
    { text: { text: "Funai Soken", options: { x: 7.6, y: 0.18, w: 2.2, h: 0.5, fontFace: FONT, fontSize: 18, color: WHITE, align: "right", valign: "middle" } } },
    { text: { text: "Confidential", options: { x: 0.2, y: 7.15, w: 2, h: 0.25, fontFace: FONT, fontSize: 8, color: "888888" } } },
    { rect: { x: 9.25, y: 7.12, w: 0.75, h: 0.38, fill: { color: NAVY } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.3, y: 0.1, w: 7.2, h: 0.65, fontFace: FONT, fontSize: 20, color: WHITE, align: "left", valign: "middle", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 9.3, y: 7.16, w: 0.65, h: 0.3, fontFace: FONT, fontSize: 10, color: WHITE, align: "center" },
});

pres.defineSlideMaster({
  title: "DIVIDER",
  background: { color: NAVY },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 2.5, w: 8.8, h: 2.5, fontFace: FONT, fontSize: 40, bold: true, color: YEL, align: "center", valign: "middle" }, text: "" } },
  ],
});

pres.defineSlideMaster({
  title: "DIVIDER_BLACK",
  background: { color: "000000" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 2.5, w: 8.8, h: 2.5, fontFace: FONT, fontSize: 40, bold: true, color: YEL, align: "center", valign: "middle" }, text: "" } },
  ],
});

pres.defineSlideMaster({ title: "BLANK", background: { color: WHITE }, objects: [] });

// ---------- ヘルパー ----------
let objId = 0;
const nm = (p) => `${p}_${++objId}`;

function content(section, title, notes) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  if (notes) s.addNotes(notes);
  return s;
}
function divider(section, title, notes, black) {
  const s = pres.addSlide({ masterName: black ? "DIVIDER_BLACK" : "DIVIDER", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  if (notes) s.addNotes(notes);
  return s;
}
function lead(s, t, opt = {}) {
  s.addText(t, { x: 0.4, y: opt.y || 1.0, w: opt.w || 9.2, h: opt.h || 0.55, fontFace: FONT, fontSize: opt.size || 17, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true, objectName: nm("lead") });
}
function concl(s, t, opt = {}) {
  s.addText(t, { x: 0.3, y: opt.y || 6.42, w: 9.4, h: 0.6, fontFace: FONT, fontSize: opt.size || 17, bold: true, color: RED, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: nm("conclusion") });
}
function note(s, t) {
  s.addText(t, { x: 2.3, y: 7.12, w: 6.85, h: 0.3, fontFace: FONT, fontSize: 8, color: GRAY, align: "right", valign: "middle", margin: 0, isTextBox: true, objectName: nm("note") });
}
const MODEL_NOTE = "※船井総研作成のモデル数値です。成果を保証するものではありません。";
function tag(s, t) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.55, y: 0.93, w: 3.25, h: 0.36, rectRadius: 0.08, fill: { color: YEL }, line: { color: GOLD, width: 1 }, objectName: nm("tag") });
  s.addText("要カスタマイズ｜" + t, { x: 6.55, y: 0.93, w: 3.25, h: 0.36, fontFace: FONT, fontSize: 10, bold: true, color: DARK, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: nm("tagText") });
}
function rect(s, x, y, w, h, fill, opt = {}) {
  s.addShape(opt.round ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE, {
    x, y, w, h, fill: { color: fill }, line: opt.line ? { color: opt.line, width: opt.lw || 1, dashType: opt.dash || "solid" } : { type: "none" },
    rectRadius: opt.round ? (opt.r || 0.08) : undefined, objectName: nm("box"),
    shadow: opt.shadow ? { type: "outer", color: "000000", opacity: 0.15, blur: 4, offset: 2, angle: 90 } : undefined,
  });
}
function txt(s, t, x, y, w, h, opt = {}) {
  s.addText(t, {
    x, y, w, h, fontFace: opt.font || FONT, fontSize: opt.size || 12, bold: !!opt.bold, color: opt.color || DARK,
    align: opt.align || "left", valign: opt.valign || "top", margin: opt.margin === undefined ? 0.06 : opt.margin,
    lineSpacingMultiple: opt.ls || 1.1, paraSpaceAfter: opt.psa || 0, isTextBox: true, objectName: nm(opt.name || "text"), fit: "none",
  });
}
function circleNum(s, n, x, y, d, fill, color) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: nm("num") });
  s.addText(String(n), { x, y, w: d, h: d, fontFace: FONT, fontSize: Math.round(d * 30), bold: true, color: color || WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: nm("numText") });
}
function arrowR(s, x, y, w, h, color) {
  s.addShape(pres.shapes.RIGHT_ARROW, { x, y, w, h, fill: { color: color || NAVY }, line: { type: "none" }, objectName: nm("arrow") });
}
function arrowD(s, x, y, w, h, color) {
  s.addShape(pres.shapes.DOWN_ARROW, { x, y, w, h, fill: { color: color || NAVY }, line: { type: "none" }, objectName: nm("arrow") });
}
// 見出し帯付きカード
function card(s, x, y, w, h, head, body, opt = {}) {
  const c = opt.color || NAVY;
  const hh = opt.hh || 0.5;
  rect(s, x, y, w, h, opt.bg || WHITE, { line: LINE, round: false, shadow: true });
  rect(s, x, y, w, hh, c);
  const num = opt.num;
  if (num !== undefined) {
    circleNum(s, num, x + 0.1, y + (hh - 0.36) / 2, 0.36, WHITE, c);
    txt(s, head, x + 0.5, y, w - 0.55, hh, { size: opt.hsize || 13, bold: true, color: WHITE, valign: "middle" });
  } else {
    txt(s, head, x + 0.08, y, w - 0.16, hh, { size: opt.hsize || 13, bold: true, color: WHITE, valign: "middle", align: opt.halign || "left" });
  }
  if (body) txt(s, body, x + 0.1, y + hh + 0.08, w - 0.2, h - hh - 0.14, { size: opt.size || 12, color: DARK, ls: 1.15 });
}
// 表
function table(s, rows, x, y, w, colW, opt = {}) {
  const fs = opt.size || 11;
  const data = rows.map((r, i) => r.map((cell, j) => {
    const isObj = typeof cell === "object" && cell !== null;
    const text = isObj ? cell.text : String(cell);
    const base = {
      fontFace: FONT, fontSize: fs, valign: "middle", margin: [3, 5, 3, 5],
      border: { type: "solid", pt: 0.75, color: LINE },
      align: j === 0 ? "left" : (opt.align || "center"),
    };
    if (i === 0) Object.assign(base, { bold: true, color: WHITE, fill: { color: opt.head || NAVY }, align: "center" });
    else Object.assign(base, { color: DARK, fill: { color: (i % 2 === 0) ? LGRAY : WHITE } });
    if (opt.firstColFill && j === 0 && i > 0) Object.assign(base, { bold: true, fill: { color: LBLUE }, color: NAVY });
    if (isObj) Object.assign(base, cell.options || {});
    return { text, options: base };
  }));
  s.addTable(data, { x, y, w, colW, rowH: opt.rowH, autoPage: false, objectName: nm("table") });
}
// 記入欄（会社ごとに差し替え）
function ph(s, x, y, w, h, label, hint) {
  rect(s, x, y, w, h, "FAFAFA", { line: "9AA5B8", dash: "dash", lw: 1.25 });
  txt(s, [
    { text: label, options: { bold: true, color: NAVY, fontSize: 12, breakLine: true } },
    { text: hint || "", options: { color: GRAY, fontSize: 10.5 } },
  ], x + 0.08, y + 0.06, w - 0.16, h - 0.12, { ls: 1.15 });
}
function steps(s, items, y, h, opt = {}) {
  const n = items.length, gap = 0.28, x0 = opt.x || 0.4, W = opt.w || 9.2;
  const w = (W - gap * (n - 1)) / n;
  items.forEach((it, i) => {
    const x = x0 + i * (w + gap);
    card(s, x, y, w, h, it[0], it[1], { color: opt.colors ? opt.colors[i] : NAVY, num: opt.num === false ? undefined : i + 1, hh: opt.hh || 0.55, size: opt.size || 11, hsize: opt.hsize || 12 });
    if (i < n - 1) arrowR(s, x + w + 0.03, y + h / 2 - 0.13, gap - 0.06, 0.26, opt.arrow || ORANGE);
  });
}


// ---------- 元の企画書の型に合わせた部品 ----------
// 中見出し（例：「事業モデル」「紙販促」）
function subhead(s, t, opt = {}) {
  txt(s, t, 0.4, opt.y || 0.98, opt.w || 6.0, 0.55, { size: opt.size || 20, bold: true, color: DARK, valign: "middle", margin: 0 });
}
// スライド内のバナー見出し（例：「── 3つの追い風が遺品整理市場を拡大 ──」）
function banner(s, t, y, opt = {}) {
  const size = opt.size || 22;
  const tw = Math.min(8.6, t.length * size / 72 * 1.02 + 0.2);
  const c = opt.color || NAVY;
  txt(s, t, 5 - tw / 2, y, tw, 0.55, { size, bold: true, color: c, align: "center", valign: "middle", margin: 0 });
  if (opt.lines !== false && tw < 7.8) {
    rect(s, 0.5, y + 0.27, 5 - tw / 2 - 0.65, 0.025, c);
    rect(s, 5 + tw / 2 + 0.15, y + 0.27, 5 - tw / 2 - 0.65, 0.025, c);
  }
  if (opt.sub) txt(s, opt.sub, 0.4, y + 0.55, 9.2, 0.38, { size: 12.5, bold: true, color: DARK, align: "center", valign: "middle", margin: 0 });
}
// 番号付き見出し帯（例：「01｜遺品整理営業の全体フロー」）
function numHead(s, num, title, desc) {
  rect(s, 0.4, 1.0, 9.2, 0.78, NAVY);
  txt(s, num, 0.5, 1.0, 0.85, 0.78, { size: 28, bold: true, color: WHITE, valign: "middle", margin: 0 });
  rect(s, 1.32, 1.15, 0.03, 0.48, WHITE);
  txt(s, title, 1.45, 1.0, 4.9, 0.78, { size: 20, bold: true, color: WHITE, valign: "middle", margin: 0 });
  if (desc) txt(s, desc, 6.35, 1.0, 3.2, 0.78, { size: 9.5, bold: true, color: WHITE, valign: "middle", margin: 0 });
}
// 赤字の締めの一文
function redMsg(s, t, opt = {}) {
  txt(s, t, 0.3, opt.y || 6.45, 9.4, opt.h || 0.6, { size: opt.size || 19, bold: true, color: RED, align: opt.align || "center", valign: "middle", margin: 0 });
}
// 紺のまとめ帯（「」内を黄色にする）
function band(s, t, y, h, opt = {}) {
  rect(s, 0.4, y, 9.2, h, NAVY);
  const runs = [];
  t.split(/(「[^」]*」)/).forEach((p) => { if (p) runs.push({ text: p, options: { color: p.startsWith("「") ? YEL : WHITE } }); });
  if (opt.label) {
    rect(s, 0.55, y + (h - 0.45) / 2, 1.1, 0.45, WHITE);
    txt(s, opt.label, 0.55, y + (h - 0.45) / 2, 1.1, 0.45, { size: 13, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    txt(s, runs, 1.8, y, 7.7, h, { size: opt.size || 13, bold: true, valign: "middle" });
  } else {
    txt(s, runs, 0.55, y, 8.9, h, { size: opt.size || 15, bold: true, valign: "middle", align: opt.align || "center" });
  }
}
// 見出し付きパネル（紺の見出し＋白地）
function panel(s, x, y, w, h, head, color) {
  rect(s, x, y, w, h, WHITE, { line: color || NAVY, lw: 1.25 });
  rect(s, x, y, w, 0.42, color || NAVY);
  txt(s, head, x, y, w, 0.42, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
}
// 小さなラベル（例：「ワンストップで受注！」）
function pill(s, x, y, w, t, color) {
  rect(s, x, y, w, 0.34, color || NAVY);
  txt(s, t, x, y, w, 0.34, { size: 10, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
}
function photo(s, x, y, w, h, label) {
  rect(s, x, y, w, h, "E6E9EF", { line: "B8C0CC" });
  txt(s, label, x, y, w, h, { size: 10, color: "7A8494", align: "center", valign: "middle" });
}
const GEN_NOTE = "※画像は差し替えてご使用ください。";

// =====================================================================
// 表紙
// =====================================================================
const S0 = "会社紹介";
pres.addSection({ title: S0 });
{
  const s = pres.addSlide({ masterName: "BLANK", sectionTitle: S0 });
  s.addShape(pres.shapes.OVAL, { x: -2.5, y: 5.7, w: 16, h: 5, fill: { color: NAVY }, line: { type: "none" }, objectName: "cover_wave1" });
  s.addShape(pres.shapes.OVAL, { x: 1.5, y: 6.25, w: 14, h: 4, fill: { color: BLUE }, line: { type: "none" }, objectName: "cover_wave2" });
  s.addShape(pres.shapes.OVAL, { x: 6.2, y: -2.6, w: 7, h: 3.4, fill: { color: NAVY }, line: { type: "none" }, objectName: "cover_wave3" });
  s.addShape(pres.shapes.OVAL, { x: 7.4, y: -2.9, w: 6, h: 3.2, fill: { color: BLUE }, line: { type: "none" }, objectName: "cover_wave4" });
  s.addShape(pres.shapes.OVAL, { x: -0.5, y: -0.4, w: 1.6, h: 1.6, fill: { color: BLUE }, line: { type: "none" }, objectName: "cover_circle" });
  for (let i = 0; i < 12; i++) {
    const d = 0.13 - i * 0.008;
    s.addShape(pres.shapes.OVAL, { x: 0.75 + i * 0.2, y: 1.95 - d / 2, w: d, h: d, fill: { color: NAVY, transparency: i * 6 }, line: { type: "none" }, objectName: nm("dot") });
  }
  txt(s, "株式会社○○ 様\n遺品整理事業\n業績アップコンサルティング\nご提案書", 0.7, 2.25, 8.8, 2.4, { size: 32, font: MINCHO, color: DARK, ls: 1.1, name: "cover_title" });
  txt(s, "株式会社 船井総合研究所\n専門工事支援部\n○○ ○○", 0.7, 4.75, 6, 0.95, { size: 12, color: DARK, ls: 1.2, name: "cover_meta" });
  s.addNotes("【要カスタマイズ】会社名・担当者名を差し替える。\n事前経営相談では「遺品整理とは」〜「各業態におけるシナジー」までを使い、事後経営相談では「セミナーでのご質問について」以降を使う。");
}
{
  const s = content(S0, "会社紹介");
  txt(s, [
    { text: "株式会社船井総合研究所", options: { fontSize: 15, bold: true, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "創設者", options: { bold: true, breakLine: true } }, { text: "船井　幸雄", options: { bold: true, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "代表取締役社長", options: { bold: true, breakLine: true } }, { text: "真貝　大介", options: { bold: true, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "連結従業員数", options: { bold: true, breakLine: true } }, { text: "約1,500名", options: { bold: true, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "資本金", options: { bold: true, breakLine: true } }, { text: "3,000百万円", options: { bold: true, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "創立", options: { bold: true, breakLine: true } }, { text: "1970年3月6日", options: { bold: true, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "会社プロフィール", options: { bold: true, breakLine: true } },
    { text: "経営コンサルティング業界では世界ではじめて株式公開を果たした日本最大級の経営コンサルティングファーム。お客様の業績を向上させることを最重要テーマとし、現場の即時業績アップ支援に強みを持つ。約1,000名のコンサルタントが、全国10,000社以上の支援先のサポートに当たっている。", options: { bold: true } },
  ], 0.4, 1.0, 5.6, 5.9, { size: 11.5, ls: 1.1 });
  photo(s, 6.3, 1.05, 3.3, 2.4, "大阪本社 写真");
  txt(s, "大阪本社 〒530-0001\n大阪府大阪市北区梅田三丁目2番123号\nイノゲート大阪 21階・22階", 6.3, 3.5, 3.3, 0.7, { size: 9.5, color: DARK });
  photo(s, 6.3, 4.3, 3.3, 2.0, "東京本社 写真");
  txt(s, "東京本社 〒104-0028\n東京都中央区八重洲二丁目2番1号\n八重洲セントラルタワー 35階", 6.3, 6.35, 3.3, 0.7, { size: 9.5, color: DARK });
}
{
  const s = content(S0, "遺品整理事業 専門コンサルタント", "【要カスタマイズ】メンバーの写真・役職・資格を差し替える。");
  const M = [["専門工事支援部\nチームリーダー\n古民家鑑定士", "阪口 和輝"], ["専門工事支援部\nコンサルタント\n遺品整理士", "堂野 宗一郎"], ["専門工事支援部\nコンサルタント", "廣瀬 聡太朗"], ["専門工事支援部\nコンサルタント", "伊木 颯志朗"]];
  M.forEach((m, i) => {
    const x = 0.5 + (i % 2) * 4.7, y = 1.3 + Math.floor(i / 2) * 2.9;
    photo(s, x, y, 2.0, 2.3, "顔写真");
    txt(s, m[0] + "\n\n" + m[1], x + 2.15, y + 0.05, 2.4, 2.2, { size: 13, bold: true, color: DARK, ls: 1.2 });
  });
}
{
  const s = content(S0, "コンサルタント紹介", "【要カスタマイズ】担当コンサルタントのプロフィールに差し替える。");
  rect(s, 0.4, 1.05, 7.0, 0.85, "F6C700");
  rect(s, 4.4, 1.05, 3.0, 0.85, "FBE38A");
  txt(s, "コンサルタント紹介", 0.55, 1.05, 6, 0.85, { size: 30, color: WHITE, valign: "middle", margin: 0 });
  photo(s, 0.4, 2.2, 4.4, 4.3, "担当コンサルタント 写真");
  txt(s, [
    { text: "専門工事支援部", options: { fontSize: 14, bold: true, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "○○ ○○", options: { fontSize: 18, bold: true, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "【出身】　○○県○○市", options: { fontSize: 16, bold: true, breakLine: true } },
    { text: "【血液型】○型", options: { fontSize: 16, bold: true, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "【経歴/プロフィール】", options: { bold: true, breakLine: true } },
    { text: "大学卒業後、新卒で株式会社船井総合研究所に入社。入社後は遺品整理業界を専門にコンサルティングを行う。現在は、新規事業立上げから既存事業活性化まで、ご支援先の業績アップコンサルティングに従事している。", options: { bold: true, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "【ミッション】", options: { bold: true, breakLine: true } },
    { text: "遺品整理を『片付け』から、安心できる住まいと暮らしの『再編集』へ", options: { bold: true } },
  ], 5.0, 2.2, 4.6, 4.5, { size: 11, ls: 1.1 });
}
// ヒアリング（元の企画書のメモ形式）
{
  const s = content(S0, "ご相談内容について", "事前経営相談ではこのページで伺い、手元でメモする。事後経営相談では伺った内容を書き込んだ版に差し替え、最初に確認する。他社名・外注先名は書かない。");
  tag(s, "ヒアリング内容");
  const H = ["・セミナー・資料DLのお申し込みのきっかけ", "・貴社の現状", "・課題", "・セミナーで知りたいこと／資料DLでの疑問点", "・遺品整理事業を会社全体でどんな位置づけにしたいか"];
  H.forEach((h, i) => {
    const y = 1.45 + i * 1.1;
    txt(s, h, 0.4, y, 9.2, 0.4, { size: 14, color: DARK, margin: 0 });
    txt(s, "↪", 0.4, y + 0.42, 0.4, 0.4, { size: 13, color: GRAY, margin: 0 });
  });
}

// =====================================================================
// 遺品整理とは
// =====================================================================
const S1 = "遺品整理とは";
pres.addSection({ title: S1 });
divider(S1, "遺品整理とは");
{
  const s = content(S1, "遺品整理ビジネスのミッション");
  rect(s, 0.4, 1.0, 9.2, 5.95, "F7F4EE");
  txt(s, [
    { text: "遺品整理", options: { color: DARK } }, { text: "は、", options: { color: DARK, fontSize: 20, breakLine: true } },
    { text: "地域", options: { color: GREEN } }, { text: "の", options: { color: DARK } }, { text: "福祉", options: { color: GREEN } }, { text: "と", options: { color: DARK } }, { text: "空き家問題", options: { color: GREEN } }, { text: "を", options: { color: DARK, breakLine: true } },
    { text: "解決する事業です", options: { color: DARK } },
  ], 0.6, 1.1, 5.4, 1.8, { size: 26, bold: true, font: MINCHO, ls: 1.15 });
  txt(s, "私たちは、故人とご遺族の想いに寄り添い、\n地域の未来を支える遺品整理を行っています。", 0.6, 2.95, 5.4, 0.7, { size: 12, color: DARK, ls: 1.3 });
  photo(s, 6.1, 1.1, 3.4, 2.5, "イメージ写真（スタッフとご遺族）");
  rect(s, 0.6, 3.85, 8.8, 0.02, "B9A98F");
  txt(s, "遺品整理をしたいお客様は、こんなお悩みを抱えています", 1.4, 3.65, 7.2, 0.4, { size: 14, bold: true, color: DARK, align: "center", valign: "middle" });
  const P = [["亡くなったご遺族の遺品を\n丁寧に扱ってくれる、\n処分してくれる、\n安心できる業者を探している"], ["今の時代、実家に\nわざわざ帰って\n片付ける時間もない"], ["ネットで調べて出てきた、\n安心できるかどうかも\n分からない会社に\n頼むしかない"]];
  P.forEach((p, i) => {
    const x = 0.6 + i * 2.95;
    s.addShape(pres.shapes.OVAL, { x, y: 4.2, w: 0.75, h: 0.75, fill: { color: "EFE6D6" }, line: { type: "none" }, objectName: nm("icon") });
    txt(s, String(i + 1), x, 4.2, 0.75, 0.75, { size: 18, bold: true, color: GREEN, align: "center", valign: "middle", margin: 0 });
    txt(s, p[0], x + 0.85, 4.1, 2.0, 1.2, { size: 11, bold: true, color: DARK, ls: 1.15 });
  });
  s.addShape(pres.shapes.OVAL, { x: 0.6, y: 5.55, w: 1.2, h: 1.2, fill: { color: GREEN }, line: { type: "none" }, objectName: "mission_circle" });
  txt(s, "私たちの\nミッション", 0.6, 5.55, 1.2, 1.2, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  txt(s, [
    { text: "そんな業界の問題を解決するべく、", options: { color: DARK, breakLine: true } },
    { text: "地域密着", options: { color: GREEN } }, { text: "で", options: { color: DARK } }, { text: "安心できる遺品整理業者", options: { color: GREEN } }, { text: "を増やし、", options: { color: DARK, breakLine: true } },
    { text: "地域の福祉と未来を支えていきます。", options: { color: DARK } },
  ], 2.0, 5.5, 7.4, 1.3, { size: 15, bold: true, font: MINCHO, ls: 1.2 });
}
{
  const s = content(S1, "遺品整理とは");
  banner(s, "遺品整理の流れ【3ステップ】", 1.0, { lines: false });
  const C = [["1", "整理（仕分け）", "遺品を丁寧に仕分けします", ORANGE, "一つ一つ確認しながら、「残す物」「貴重品」「処分する物」に仕分けします。"], ["2", "搬出", "仕分けした段ボールを運び出します", GREEN, "仕分けが終わった段ボールを一つずつ丁寧に運び出します。"], ["3", "トラックへの積込", "トラックに積み込みます", NAVY, "安全にトラックへ積み込み、適切に処理・リサイクルを行います。"]];
  C.forEach((c, i) => {
    const x = 0.4 + i * 3.13, w = 2.94;
    rect(s, x, 1.65, w, 5.2, WHITE, { line: c[3], lw: 2, round: true, r: 0.06 });
    s.addShape(pres.shapes.OVAL, { x: x + 0.12, y: 1.78, w: 0.5, h: 0.5, fill: { color: c[3] }, line: { type: "none" }, objectName: nm("num") });
    txt(s, c[0], x + 0.12, 1.78, 0.5, 0.5, { size: 18, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, c[1], x + 0.7, 1.75, w - 0.75, 0.42, { size: 16, bold: true, color: c[3], valign: "middle", margin: 0 });
    txt(s, c[2], x + 0.7, 2.15, w - 0.75, 0.3, { size: 10, bold: true, color: DARK, margin: 0 });
    photo(s, x + 0.12, 2.6, w - 0.24, 2.6, "作業写真");
    txt(s, c[4], x + 0.15, 5.35, w - 0.3, 1.35, { size: 11, bold: true, color: DARK, ls: 1.2 });
    if (i < 2) s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: x + w + 0.02, y: 3.7, w: 0.17, h: 0.3, rotate: 90, fill: { color: NAVY }, line: { type: "none" }, objectName: nm("tri") });
  });
}

// =====================================================================
// 市場性
// =====================================================================
const S2 = "市場性";
pres.addSection({ title: S2 });
divider(S2, "市場性");
{
  const s = content(S2, "遺品整理市場の成長性");
  banner(s, "3つの追い風が遺品整理市場を拡大", 0.98, { size: 24 });
  const W = [["① 市場拡大の波", "多死社会の到来", "・年間死亡者数は今後も増加\n・遺品整理・実家整理の需要が拡大\n・「空き家をどうするか」を考える人が増加", "市場ニーズが拡大", "FFF3DC"], ["② 社会課題の波", "空き家問題の深刻化", "・全国の空き家 約900万戸\n・放置空き家も増加\n・活用・売却・解体の需要が拡大", "空き家対策が社会課題に", LBLUE], ["③ 法改正の波", "相続登記の義務化", "・相続後3年以内に登記が必要\n・放置が難しくなる\n・実家の活用・売却・整理を検討する人が増加", "法改正が行動を後押し", LGREEN]];
  W.forEach((w, i) => {
    const x = 0.4 + i * 3.13, ww = 2.94;
    rect(s, x, 1.65, ww, 3.2, WHITE, { line: LINE });
    rect(s, x, 1.65, ww, 0.5, NAVY);
    txt(s, w[0], x, 1.65, ww, 0.5, { size: 15, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, w[1], x, 2.2, ww, 0.4, { size: 13, bold: true, color: BROWN, align: "center", valign: "middle", margin: 0 });
    txt(s, w[2], x + 0.15, 2.65, ww - 0.3, 1.4, { size: 11, color: DARK, ls: 1.3 });
    rect(s, x + 0.15, 4.15, ww - 0.3, 0.55, w[4], { line: LINE, round: true });
    txt(s, w[3], x + 0.15, 4.15, ww - 0.3, 0.55, { size: 12, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
  });
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: 4.6, y: 4.95, w: 0.8, h: 0.3, rotate: 180, fill: { color: NAVY }, line: { type: "none" }, objectName: "tri_down" });
  rect(s, 0.4, 5.3, 9.2, 0.55, NAVY);
  txt(s, [{ text: "3つの波が重なり、", options: { color: WHITE } }, { text: "「放置」", options: { color: "F6C700" } }, { text: "から", options: { color: WHITE } }, { text: "「活用・処分」", options: { color: "F6C700" } }, { text: "へ動く人が急増", options: { color: WHITE } }], 0.4, 5.3, 9.2, 0.55, { size: 17, bold: true, align: "center", valign: "middle", margin: 0 });
  rect(s, 0.4, 5.85, 9.2, 1.1, WHITE, { line: NAVY });
  ["遺品整理", "実家整理", "リフォーム", "解体", "売却", "賃貸活用"].forEach((t, i) => {
    txt(s, t, 0.4 + i * 1.533, 5.85, 1.533, 1.1, { size: 14, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    if (i) rect(s, 0.4 + i * 1.533, 6.05, 0.015, 0.7, LINE);
  });
  note(s, "※空き家数：総務省「住宅・土地統計調査」");
}
{
  const s = content(S2, "遺品整理市場の現状");
  banner(s, "業界の現状", 0.95, { size: 28, lines: false, sub: "遺品整理業者が抱える課題と、見えてきた市場のチャンス" });
  rect(s, 0.6, 1.95, 8.8, 4.95, "FBF7F1", { line: "D9C7B2", round: true, r: 0.05 });
  rect(s, 0.6, 1.95, 8.8, 0.6, BROWN);
  txt(s, "遺品整理業の現状", 0.6, 1.95, 8.8, 0.6, { size: 18, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  const I = [["不用品回収を中心とした事業者が多い", "遺品整理という名義でも、実際には不用品回収の片付けを中心とした\nサービスが多く、「想いの整理」まで対応できる事業者は限られている。"], ["グリーフケアや供養への対応に差がある", "ご遺族への寄り添い方や供養の提案など、精神的なケアまで\n提供できる事業者はまだ多いとはいえない。"], ["サービス品質にばらつきがある", "会社規模や基準は業者に差があり、対応・料金・作業品質などが\nお客様満足度に大きく影響している。"]];
  I.forEach((it, i) => {
    const y = 2.7 + i * 1.1;
    s.addShape(pres.shapes.OVAL, { x: 0.8, y: y + 0.05, w: 0.45, h: 0.45, fill: { color: BROWN }, line: { type: "none" }, objectName: nm("num") });
    txt(s, String(i + 1), 0.8, y + 0.05, 0.45, 0.45, { size: 15, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, it[0], 1.4, y, 7.8, 0.5, { size: 16, bold: true, color: BROWN, valign: "middle", margin: 0 });
    txt(s, it[1], 1.4, y + 0.48, 7.8, 0.55, { size: 11, bold: true, color: DARK, margin: 0, ls: 1.2 });
    if (i < 2) rect(s, 0.8, y + 1.05, 8.4, 0.01, "D9C7B2");
  });
  rect(s, 0.8, 6.05, 8.4, 0.65, "FCEFD9", { round: true });
  txt(s, "「遺品整理＝片付け・回収」だけではない、心に寄り添うサービスが求められています。", 0.9, 6.05, 8.2, 0.65, { size: 12.5, bold: true, color: BROWN, align: "center", valign: "middle", margin: 0 });
}
{
  const s = content(S2, "遺品整理市場の成長性", "ライフサイクルの考え方：成長期は一番店が決まっていない。安定期に入ると価格競争になり、一番店化は難しい。");
  subhead(s, "ライフサイクルから分析する遺品整理市場", { w: 9 });
  const phases = [["導入期", 0.9, "D5DCE8"], ["成長期", 2.6, RED], ["成熟期", 3.3, "8FA3C4"], ["安定期", 2.9, "8FA3C4"], ["衰退期", 2.0, "B8C3D6"]];
  const bx = 0.6, bw = 1.2, base = 5.3;
  phases.forEach((p, i) => {
    const x = bx + i * 1.45;
    rect(s, x, base - p[1], bw, p[1], p[2]);
    txt(s, p[0], x, base + 0.05, bw, 0.35, { size: 12, bold: true, color: i === 1 ? RED : NAVY, align: "center" });
  });
  txt(s, "遺品整理\n総合型", bx + 1.45, base - 2.6 + 0.1, bw, 0.7, { size: 12, bold: true, color: WHITE, align: "center" });
  txt(s, "リフォーム\n新築住宅", bx + 2 * 1.45, base - 3.3 + 0.1, bw, 0.7, { size: 11, bold: true, color: WHITE, align: "center" });
  rect(s, 7.95, 1.7, 1.65, 1.75, RED);
  txt(s, "成長期\n地域一番店が不在", 7.95, 1.7, 1.65, 1.75, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
  rect(s, 7.95, 3.6, 1.65, 1.75, "8FA3C4");
  txt(s, "安定期\n価格競争が激化し、一番店化は困難", 7.95, 3.6, 1.65, 1.75, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle" });
  redMsg(s, "地域一番店が不在の今が、参入のタイミングです。");
}

// =====================================================================
// ビジネスモデル概要
// =====================================================================
const S3 = "ビジネスモデル概要";
pres.addSection({ title: S3 });
divider(S3, "ビジネスモデル概要");
{
  const s = content(S3, "ビジネスモデルと売上目標");
  subhead(s, "事業モデル");
  txt(s, [
    { text: "商圏人口", options: { fontSize: 18 } }, { text: "20", options: { fontSize: 36 } }, { text: "万人・平均客単価", options: { fontSize: 18 } }, { text: "30", options: { fontSize: 36 } }, { text: "万円で　売上", options: { fontSize: 18 } }, { text: "1億円", options: { fontSize: 32 } }, { text: "構築モデル", options: { fontSize: 18 } },
  ], 0.6, 1.55, 9.0, 0.75, { bold: true, color: RED, align: "center", valign: "middle", margin: 0 });
  txt(s, "（※　但し、商圏人口や会社の認知度や実績で年商には変動が生じます。商圏人口20万人を想定しています。）\n（※　数字はあくまでモデル数値であり、確約するものではありません。）", 0.9, 2.3, 8.5, 0.55, { size: 10, color: GRAY, margin: 0, ls: 1.2 });
  s.addShape(pres.shapes.OVAL, { x: 0.5, y: 3.15, w: 4.6, h: 2.9, fill: { color: "EEF0F3" }, line: { color: "BFC5CE", width: 6 }, objectName: "model_ring" });
  rect(s, 1.75, 3.0, 2.1, 0.7, WHITE, { line: LINE, round: true, shadow: true });
  txt(s, "遺品整理", 1.75, 3.0, 2.1, 0.7, { size: 24, bold: true, color: RED, align: "center", valign: "middle", margin: 0 });
  photo(s, 1.6, 3.85, 2.4, 1.4, "戸建住宅の写真");
  rect(s, 0.4, 5.45, 1.9, 0.5, WHITE, { line: LINE, round: true });
  txt(s, "（不用品回収）", 0.4, 5.45, 1.9, 0.5, { size: 13, bold: true, color: DARK, align: "center", valign: "middle", margin: 0 });
  rect(s, 3.3, 5.45, 1.9, 0.5, WHITE, { line: LINE, round: true });
  txt(s, "（出張買取）", 3.3, 5.45, 1.9, 0.5, { size: 13, bold: true, color: DARK, align: "center", valign: "middle", margin: 0 });
  rect(s, 5.45, 3.05, 4.15, 2.95, WHITE, { line: BLUE, lw: 1.5 });
  txt(s, "１．40代〜70代以上の故人の遺品整理をしたい人がターゲット。\n\n２．不用品回収や出張買取も対応するが、単価が大きい「戸建の遺品整理」をメインで集客。\n\n３．バックエンド商品を重視。解体・リフォーム・不動産仲介・葬儀等へ繋げる工夫でLTVをUPさせる。", 5.6, 3.15, 3.9, 2.8, { size: 12.5, bold: true, color: DARK, ls: 1.1 });
}
{
  const s = content(S3, "ビジネスモデルと売上目標", "【要確認】1年目の営業利益：元資料の記載は106万円だが、経費の内訳を足すと586万円になる。差額480万円の項目を確定させる。");
  txt(s, [{ text: "数値計画", options: { fontSize: 24 } }, { text: "（3年間）", options: { fontSize: 16 } }], 0.4, 0.98, 3.2, 0.6, { bold: true, color: NAVY, valign: "middle", margin: 0 });
  txt(s, "3年目には売上1億円を達成し、営業利益率23.1%の高収益体制へ\n安定した成長と収益性の向上を実現します。", 3.6, 0.98, 4.2, 0.6, { size: 10.5, bold: true, color: DARK, valign: "middle", margin: 0 });
  const B = (t) => ({ text: t, options: { bold: true } });
  table(s, [
    ["項目", "1年目", "", "2年目", "", "3年目", ""],
    [B("売上"), B("¥36,000,000"), "", B("¥54,000,000"), "", B("¥100,000,000"), ""],
    ["処分費", "¥10,800,000", "30%", "¥16,200,000", "30%", "¥30,000,000", "30%"],
    ["その他消耗品費", "¥1,000,000", "資材等", "¥1,500,000", "件数増加に伴い", "¥2,000,000", "件数増加に伴い"],
    [B("粗利"), B("¥24,200,000"), "67%", B("¥36,300,000"), "67%", B("¥68,000,000"), "68%"],
    ["アルバイト事務", "¥1,200,000", "事務1名", "¥1,200,000", "事務1名", "¥1,200,000", "事務1名"],
    ["広告宣伝費・広報費", "¥4,840,000", "粗利の20%", "¥7,260,000", "粗利の20%", "¥13,600,000", "粗利の20%"],
    ["その他事業経費", "¥1,500,000", "", "¥2,000,000", "", "¥2,500,000", ""],
    ["人件費", "¥6,000,000", "専任1名", "¥12,000,000", "専任2名", "¥18,000,000", "専任3名"],
    ["作業員アルバイト", "¥4,800,000", "アルバイト2名", "¥4,800,000", "アルバイト2名", "¥9,600,000", "アルバイト4名"],
    [B("営業利益"), B("¥1,060,000"), "", B("¥9,040,000"), "", B("¥23,100,000"), ""],
    [B("営業利益率"), B("2.94%"), "", B("16.74%"), "", B("23.10%"), ""],
  ], 0.4, 1.65, 7.0, [1.45, 1.05, 0.8, 1.05, 0.8, 1.05, 0.8], { size: 8.5, rowH: 0.34, firstColFill: true });
  rect(s, 7.6, 1.65, 2.0, 0.45, NAVY);
  txt(s, "初期投資", 7.6, 1.65, 2.0, 0.45, { size: 14, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  rect(s, 7.6, 2.1, 2.0, 1.5, WHITE, { line: LINE });
  txt(s, [{ text: "HP制作費 ", options: { fontSize: 11 } }, { text: "150万円", options: { fontSize: 16, color: RED, breakLine: true } }, { text: " ", options: { fontSize: 6, breakLine: true } }, { text: "古物商許可取得費用\n", options: { fontSize: 11 } }, { text: "5〜8万円", options: { fontSize: 16, color: RED } }], 7.7, 2.15, 1.85, 1.4, { bold: true, color: DARK, valign: "middle" });
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: 8.35, y: 3.7, w: 0.5, h: 0.22, rotate: 180, fill: { color: NAVY }, line: { type: "none" }, objectName: "tri_inv" });
  rect(s, 7.6, 4.0, 2.0, 0.45, NAVY);
  txt(s, "投資回収", 7.6, 4.0, 2.0, 0.45, { size: 14, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  rect(s, 7.6, 4.45, 2.0, 1.25, WHITE, { line: LINE });
  txt(s, [{ text: "2年目", options: { fontSize: 22, color: RED } }, { text: "で回収", options: { fontSize: 13, color: RED, breakLine: true } }, { text: "2年目の営業利益904万円により、初期投資を回収見込み。", options: { fontSize: 9.5, color: DARK } }], 7.7, 4.5, 1.85, 1.15, { bold: true, valign: "middle" });
  txt(s, [
    { text: "【特徴】", options: { color: DARK, breakLine: true } },
    { text: "初期投資費用が低く、社員1名からスタート可能", options: { color: RED } }, { text: "であり、", options: { color: DARK, breakLine: true } },
    { text: "20%以上の利益率", options: { color: RED } }, { text: "を目指せるビジネスモデルである。", options: { color: DARK } },
  ], 0.4, 5.85, 9.2, 1.15, { size: 17, bold: true, margin: 0, ls: 1.1 });
  note(s, MODEL_NOTE);
}
{
  const s = content(S3, "ビジネスモデルと売上目標");
  banner(s, "遺品整理を中心としたシナジーと商流の流れ", 0.95, { lines: false, size: 20, sub: "さまざまな入口から遺品整理に繋がり、整理後の出口サービスまでワンストップで提供" });
  panel(s, 0.4, 1.95, 2.75, 2.95, "入口（案件の発生源）", NAVY);
  txt(s, "① 葬儀会社\n　葬儀後に遺品整理の相談が発生\n② 老人ホーム紹介\n　入居前の生前整理が必要\n③ 不動産仲介・不動産売却\n　売却前には家財撤去が必須", 0.5, 2.42, 2.6, 2.4, { size: 10.5, bold: true, color: DARK, ls: 1.25 });
  arrowR(s, 3.2, 3.25, 0.45, 0.35, NAVY);
  s.addShape(pres.shapes.OVAL, { x: 3.7, y: 2.0, w: 2.6, h: 2.6, fill: { color: NAVY }, line: { color: LINE, width: 3 }, objectName: "center_ihin" });
  txt(s, "遺品整理\n（元請けで受注）", 3.7, 2.45, 2.6, 0.95, { size: 17, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, "遺品の仕分け・整理\n貴重品探索・供養品の選定\n買取・処分・搬出", 3.7, 3.4, 2.6, 0.8, { size: 9.5, bold: true, color: WHITE, align: "center", valign: "middle" });
  arrowR(s, 6.35, 3.25, 0.45, 0.35, GREEN);
  panel(s, 6.85, 1.95, 2.75, 2.95, "出口（整理後のサービス）", GREEN);
  txt(s, "④ 解体業者\n⑤ 一般廃棄物業者\n⑥ リユース・買取\n⑦ 清掃業\n⑧ リフォーム会社\n⑨ 不動産会社（売却・活用）", 6.95, 2.42, 2.6, 2.4, { size: 11, bold: true, color: DARK, ls: 1.3 });
  rect(s, 0.4, 5.05, 4.5, 1.9, "FFF6EE", { line: ORANGE, dash: "dash" });
  txt(s, "パターンA：解体会社からの残置物撤去依頼（下請けの流れ）", 0.5, 5.1, 4.3, 0.4, { size: 10.5, bold: true, color: ORANGE });
  steps(s, [["解体業者", "残置物撤去を依頼"], ["遺品整理会社", "撤去・処分"], ["解体業者", "解体工事"]], 5.55, 1.3, { x: 0.5, w: 4.3, num: false, size: 9.5, hsize: 10.5, hh: 0.42, colors: [ORANGE, ORANGE, ORANGE] });
  rect(s, 5.1, 5.05, 4.5, 1.9, "EEF3FB", { line: BLUE, dash: "dash" });
  txt(s, "パターンB：遺品整理を元請けで受注し、解体まで一括対応（元請けの流れ）", 5.2, 5.1, 4.3, 0.4, { size: 10.5, bold: true, color: BLUE });
  steps(s, [["お客様", "遺品整理を依頼"], ["遺品整理会社", "遺品整理を実施"], ["解体・不動産", "解体・売却・活用"]], 5.55, 1.3, { x: 5.2, w: 4.3, num: false, size: 9.5, hsize: 10.5, hh: 0.42, colors: [BLUE, BLUE, BLUE] });
}
{
  const s = content(S3, "ビジネスモデルと売上目標");
  subhead(s, "遺品整理から広がる7つの収益機会", { w: 9 });
  txt(s, [{ text: "遺品整理を中心に、関連サービスを階段状にご提案していくことで、\nお客様の課題を", options: { color: DARK } }, { text: "ワンストップで解決", options: { color: ORANGE } }, { text: "し、", options: { color: DARK } }, { text: "粗利を最大化", options: { color: ORANGE } }, { text: "します。", options: { color: DARK } }], 0.4, 1.5, 9.2, 0.65, { size: 13, bold: true, margin: 0 });
  const st = [["遺品整理（基本サービス）", ORANGE], ["買取・リユース", GOLD], ["供養・お焚き上げ", "7FA650"], ["特殊清掃・消臭・除菌", GREEN], ["解体工事", BLUE], ["不動産売却・買取サポート", NAVY], ["リフォーム・リノベーション", PURPLE]];
  st.forEach((t, i) => {
    const x = 0.4 + i * 0.32, y = 6.3 - i * 0.6, w = 3.6;
    rect(s, x, y, w, 0.52, t[1]);
    txt(s, String(i + 1), x, y, 0.45, 0.52, { size: 16, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, t[0], x + 0.45, y, w - 0.5, 0.52, { size: 12, bold: true, color: WHITE, valign: "middle" });
  });
  panel(s, 6.0, 2.35, 3.6, 4.5, "高粗利化の具体例（1件あたり）", NAVY);
  table(s, [["項目", "売上"], ["遺品整理（基本）", "30万円"], ["買取・リユース", "▲5万円（値引き）"], ["特殊清掃", "20万円"], ["供養・お焚き上げ", "5万円"], ["解体工事", "150万円"], ["不動産売却サポート", "売却価格の3%"], ["リフォーム", "200万円"]], 6.1, 2.85, 3.4, [1.75, 1.65], { size: 9.5, rowH: 0.3 });
  txt(s, [{ text: "売上合計 ", options: { fontSize: 12, color: DARK } }, { text: "490万円", options: { fontSize: 20, color: RED, breakLine: true } }, { text: "粗利合計 ", options: { fontSize: 12, color: DARK } }, { text: "180万円", options: { fontSize: 20, color: RED } }], 6.1, 5.4, 3.4, 1.35, { bold: true, align: "center", valign: "middle" });
  note(s, MODEL_NOTE);
}

// =====================================================================
// 許可・資格について
// =====================================================================
const S4 = "許可・資格について";
pres.addSection({ title: S4 });
divider(S4, "許可・資格について");
{
  const s = content(S4, "資格・許可について", "前提条件（お客様と許可業者の直接契約、当社は運搬しない、現場で許可業者へ引き渡す）と自治体ごとのルールを確認したうえで説明する。");
  rect(s, 2.6, 0.98, 4.8, 0.42, NAVY, { round: true, r: 0.2 });
  txt(s, "遺品整理・不用品回収業の新しいスタンダード", 2.6, 0.98, 4.8, 0.42, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  txt(s, [{ text: "三社間契約", options: { color: NAVY } }, { text: "を活用した", options: { color: NAVY, fontSize: 18 } }, { text: "安心・合法・低リスク", options: { color: GOLD } }, { text: "のサービス提供", options: { color: NAVY, fontSize: 18 } }], 0.4, 1.45, 9.2, 0.6, { size: 24, bold: true, align: "center", valign: "middle", margin: 0 });
  txt(s, "当社は「一廃許可」を持たなくても、一般廃棄物業者との三社間契約により、すべての工程を法令に沿って提供できます。", 0.4, 2.05, 9.2, 0.35, { size: 10.5, bold: true, color: DARK, align: "center", margin: 0 });
  const D = [[0.4, "お客様と一般廃棄物業者が直接契約", "お客様", "一般廃棄物業者\n（許可業者）", "❶直接契約（回収・処分の契約）\n❷回収・処分の実施", GREEN], [5.1, "お客様と当社が契約", "お客様", "当社\n（遺品整理）", "❶サービス契約（仕分け・搬出の契約）\n❷仕分け・搬出などの作業実施", NAVY]];
  D.forEach(([x, h, a, b, mid, c]) => {
    rect(s, x, 2.5, 4.5, 1.85, WHITE, { line: NAVY });
    rect(s, x + 0.8, 2.42, 2.9, 0.34, NAVY, { round: true, r: 0.15 });
    txt(s, h, x + 0.8, 2.42, 2.9, 0.34, { size: 10.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    rect(s, x + 0.12, 2.95, 1.2, 1.25, LBLUE);
    txt(s, a, x + 0.12, 2.95, 1.2, 1.25, { size: 12, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    rect(s, x + 3.18, 2.95, 1.2, 1.25, c);
    txt(s, b, x + 3.18, 2.95, 1.2, 1.25, { size: 10.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.LEFT_RIGHT_ARROW, { x: x + 1.45, y: 3.35, w: 1.6, h: 0.3, fill: { color: c }, line: { type: "none" }, objectName: nm("lr") });
    txt(s, mid.replace("❶", "①").replace("❷", "②"), x + 1.35, 2.85, 1.8, 0.5, { size: 7.5, bold: true, color: DARK, align: "center", valign: "middle", margin: 0 });
  });
  txt(s, [{ text: "一般廃棄物業者は直接お客様と契約するため、当社は", options: { color: DARK } }, { text: "「一廃許可」を持っていなくても", options: { color: RED } }, { text: "事業が可能です。", options: { color: DARK } }], 0.4, 4.4, 9.2, 0.4, { size: 12, bold: true, align: "center", valign: "middle", margin: 0 });
  panel(s, 0.4, 4.85, 5.9, 2.15, "当社のサービス提供の流れ", NAVY);
  ["ご相談・\nお見積り", "仕分け・\n整理", "搬出・\n運搬", "許可業者へ\n引渡し", "回収・処分\n（許可業者）"].forEach((t, k) => {
    const x = 0.5 + k * 1.16;
    rect(s, x, 5.45, 1.0, 1.35, LBLUE, { line: LINE });
    txt(s, String(k + 1), x, 5.5, 1.0, 0.35, { size: 13, bold: true, color: NAVY, align: "center", margin: 0 });
    txt(s, t, x, 5.85, 1.0, 0.85, { size: 9.5, bold: true, color: DARK, align: "center", valign: "middle", margin: 0 });
    if (k < 4) arrowR(s, x + 1.01, 6.0, 0.14, 0.25, NAVY);
  });
  panel(s, 6.45, 4.85, 3.15, 2.15, "三社間契約を導入するメリット", NAVY);
  txt(s, "■ 法令遵守で安心\n■ 許可がなくても事業可能\n■ 信頼性の高いサービス提供\n■ 事業拡大に集中できる", 6.55, 5.32, 3.0, 1.6, { size: 11.5, bold: true, color: NAVY, ls: 1.35 });
}
{
  const s = content(S4, "資格・許可について");
  const C = [["① 一般廃棄物収集運搬許可", "許可がなくても事業は可能。\n許可を取得すれば、利益率をもっと高められます。", [["利益率の向上", "自社で回収・運搬が完結するため、外注コストを抑え、収益性が向上します。"], ["スケジュールの柔軟性アップ", "外注の調整が不要になり、スケジュールを柔軟にコントロールできます。"], ["サービスの幅が広がる", "スポット対応や即日対応など、多様なニーズに応えられます。"]], GREEN], ["② 古物商許可", "「片付け」だけなら不要ですが、\n「買取」を行う場合は古物商許可が必要になります。", [["取得までの流れ", "書類準備 → 警察署へ申請 → 審査（約40日）→ 許可証交付 → 買取サービス開始"], ["費用", "申請手数料：約19,000円\n各種証明書：約1,000〜3,000円"], ["取得後に必要なこと", "本人確認／古物台帳への記録（3年間保管）／古物商プレートの掲示"]], NAVY]];
  C.forEach(([h, sub, items, c], i) => {
    const x = 0.4 + i * 4.7, w = 4.5;
    rect(s, x, 1.0, w, 0.6, c);
    txt(s, h, x + 0.1, 1.0, w - 0.2, 0.6, { size: 16, bold: true, color: WHITE, valign: "middle" });
    rect(s, x, 1.6, w, 5.35, WHITE, { line: c });
    txt(s, sub, x + 0.15, 1.7, w - 0.3, 0.75, { size: 11.5, bold: true, color: c, ls: 1.2 });
    items.forEach((it, k) => {
      const y = 2.55 + k * 1.42;
      rect(s, x + 0.15, y, w - 0.3, 1.3, i ? LBLUE : LGREEN);
      txt(s, it[0], x + 0.25, y + 0.05, w - 0.5, 0.4, { size: 12.5, bold: true, color: c });
      txt(s, it[1], x + 0.25, y + 0.45, w - 0.5, 0.8, { size: 10.5, bold: true, color: DARK, ls: 1.2 });
    });
  });
}

// =====================================================================
// 各業態におけるシナジー
// =====================================================================
const S5 = "各業態におけるシナジー";
pres.addSection({ title: S5 });
divider(S5, "各業態における\nシナジー");
{
  const s = content(S5, "既存ビジネスとのシナジー");
  subhead(s, "遺品整理事業参入で得られるシナジー", { w: 9 });
  txt(s, "業種を問わず、遺品整理事業には本業を伸ばす4つのシナジーがあります。", 0.4, 1.5, 9.2, 0.4, { size: 13, bold: true, color: DARK, margin: 0 });
  rect(s, 0.4, 2.0, 9.2, 3.5, WHITE, { line: NAVY });
  rect(s, 0.4, 2.0, 9.2, 0.5, NAVY);
  txt(s, "遺品整理事業参入によるシナジー", 0.4, 2.0, 9.2, 0.5, { size: 15, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  const E = [["商流の川上の獲得", "売却・解体・リフォームを決める前のお客様と、最初に接点を持てる。", "相見積もりを回避！"], ["既存資産の活用", "既存顧客・車両・職人・店舗・会員組織をそのまま活かせる。", "追加投資を抑制！"], ["外注費の内製化", "外注していた残置物撤去や家財処分が、自社の粗利に変わる。", "自社施工で粗利UP！"], ["紹介の増加", "「困ったらあの会社」と地域で認知され、紹介が集まる。", "安定した受注へ！"]];
  E.forEach((e, i) => {
    const x = 0.55 + i * 2.27, w = 2.12;
    rect(s, x, 2.65, w, 2.7, LBLUE);
    s.addShape(pres.shapes.OVAL, { x: x + 0.1, y: 2.75, w: 0.42, h: 0.42, fill: { color: NAVY }, line: { type: "none" }, objectName: nm("num") });
    txt(s, String(i + 1), x + 0.1, 2.75, 0.42, 0.42, { size: 14, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, e[0], x + 0.58, 2.72, w - 0.62, 0.5, { size: 12.5, bold: true, color: NAVY, valign: "middle", margin: 0 });
    txt(s, e[1], x + 0.1, 3.35, w - 0.2, 1.3, { size: 10.5, bold: true, color: DARK, ls: 1.25 });
    pill(s, x + 0.1, 4.85, w - 0.2, e[2]);
    if (i < 3) s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: x + w + 0.02, y: 3.85, w: 0.2, h: 0.12, rotate: 90, fill: { color: NAVY }, line: { type: "none" }, objectName: nm("tri") });
  });
  band(s, "遺品整理事業の参入は、どの業種にとっても「本業の入口」をつくり、「客単価の向上」と「安定受注」を実現する戦略となります。", 5.75, 1.0, { label: "まとめ", size: 13 });
}
{
  const s = content(S5, "既存ビジネスとのシナジー", "貴社の業種の行を指しながら説明し、次ページ以降の業種別ページへ進む。");
  subhead(s, "業種別のシナジー", { w: 9 });
  table(s, [
    ["業種", "本業の課題", "遺品整理で得られること", "つながる本業"],
    ["不動産会社", "仕入れ競争・広告費の高騰、案件の長期化", "売却を考える前の所有者と出会える／残置物撤去の外注費が利益に", "売買仲介・買取再販・賃貸管理"],
    ["解体業者", "価格競争、紹介依存、職人の稼働の波", "解体前の案件を元請けで獲得／家財撤去の利益も自社に", "解体工事・土地活用"],
    ["住宅・リフォーム業", "新築縮小、大型案件不足、OB客との接点不足", "相続した実家のリノベ相談の入口になる", "リフォーム・リノベーション"],
    ["葬儀業", "家族葬化で単価下落、葬儀後の収益がない", "葬儀後の困りごとを自社で受け、会員満足度が上がる", "法要・供養・墓じまい・相続相談"],
    ["ケアマネ・包括支援", "片付けの相談先不足、業務負担の増加", "施設入居前の生前整理に対応でき、家族の信頼が上がる", "入居相談・地域連携"],
    ["便利業", "小口案件ばかりで利益率が低い、ブランド力が弱い", "高単価商品を持てる／専門店として信頼される", "不用品回収・ハウスクリーニング"],
    ["リサイクル業", "良質な仕入れが不安定、一般客との接点が少ない", "良質なリユース品を直接仕入れられる", "資源回収・リユース販売"],
    ["買取業", "高価値品の仕入れが不安定、買取できない物が多い", "家一軒分から仕入れられ「全部お願いできる」が強みに", "出張買取・店舗販売"],
  ], 0.4, 1.6, 9.2, [1.5, 2.5, 3.25, 1.95], { size: 9.5, align: "left", rowH: 0.56, firstColFill: true });
}

const INDUSTRIES = [
  { name: "不動産会社",
    issues: [["仕入れ競争の激化", "ポータルサイトや一括査定経由の案件が増え、競合との奪い合いが激しい。"], ["集客コストの高騰", "反響サイト・Web広告・チラシなどの集客費用が上昇している。"], ["案件の長期化", "売却前の片付けや相続手続きに時間がかかり、案件が長引く。"], ["残置物撤去の外注", "売買・仲介で発生する残置物撤去を、1現場30〜50万円で外注している。"]],
    syn: [["相続案件の入口獲得", "売却を考える前の段階で、所有者と接点を持てる。", "風上で情報獲得！"], ["物件仕入れの強化", "相続物件の情報を、競合よりも早く取得できる。", "仕入れが増える！"], ["残置物撤去の内製化", "外注していた撤去費用を、自社の粗利に転換できる。", "外注費を粗利へ！"], ["売却スピード向上", "片付けから売却まで一貫して進められる。", "案件の長期化を防止！"]],
    eff: [["売却相談の増加", "片付けをきっかけに売却相談が増える。"], ["仕入れの安定化", "ポータルに依存しない仕入れルートを確立。"], ["粗利の向上", "外注していた撤去費用が自社の利益に。"], ["地域での信頼向上", "相続の困りごとを解決する会社として認知。"], ["事業の拡大", "買取再販・リフォーム・賃貸管理まで展開。"]],
    flow: ["遺品整理\n（家財整理）", "不動産情報の\n取得", "売却・買取再販", "リフォーム・\n解体"],
    sum: "遺品整理事業の参入は、不動産会社にとって「相続案件の入口獲得」「物件仕入れの強化」「売却スピード向上」を実現し、持続的な成長につながる強力な戦略となります。" },
  { name: "解体業者",
    issues: [["利益率の低下", "解体単体の価格競争で利益率が下がっている。"], ["集客コストの増加", "ネット広告や紹介依存により、集客コストが年々上昇している。"], ["新たな収益源の確保", "解体工事以外で収益を生み出し、売上の柱を増やしたい。"], ["職人稼働の平準化", "工事案件に波があり、職人の稼働が不安定になっている。"]],
    syn: [["解体前案件の獲得", "遺品整理を起点に、解体につながる案件を自社で獲得。", "ワンストップで受注！"], ["客単価の向上", "解体工事のみよりも、受注単価が大きく向上。", "単価UPで売上拡大！"], ["粗利の向上", "外注していた家財撤去を自社で行うことで、中間マージンを削減。", "自社施工で粗利UP！"], ["紹介依存からの脱却", "エンドユーザーとの接点が増え、直接受注の割合を高められる。", "直接取引で安定受注！"]],
    eff: [["受注件数の増加", "遺品整理をきっかけに解体につながる案件が増える。"], ["売上の安定化", "解体工事の波をカバーする新たな収益源に。"], ["地域での信頼向上", "「片付けから解体まで任せられる」安心感。"], ["職人の稼働率向上", "工事の合間でも稼働でき、稼働率が向上。"], ["将来の事業拡大", "リフォーム・不動産・売却までの展開が期待できる。"]],
    flow: ["遺品整理\n（家財整理）", "家財撤去・処分\n（運搬・処分）", "解体工事", "不動産・売却\n（提携連携）"],
    sum: "遺品整理事業の参入は、解体業者にとって「解体前案件の獲得」「客単価の向上」「粗利の向上」「安定受注」を実現し、持続的な成長につながる強力な戦略となります。" },
  { name: "住宅・リフォーム業",
    issues: [["新築市場の縮小", "人口減少や世帯数の減少により、新築需要が年々減少している。"], ["大型リフォーム案件の不足", "単価の低い小規模工事が増え、利益が伸び悩んでいる。"], ["OB顧客との接点不足", "築年数が経過したOB顧客との関係が希薄化している。"], ["空き家問題の増加", "親の死去により発生する空き家が増え、活用提案が求められている。"]],
    syn: [["実家リノベ案件の獲得", "遺品整理をきっかけに実家の活用相談につなげる。", "「片付け」から「活用提案」へ！"], ["相続空き家市場の攻略", "片付けから活用までワンストップで対応できる体制を構築。", "成長市場を先取り！"], ["リフォーム受注の増加", "売却ではなく“活用”の提案ができ、受注につながる。", "「活かす」提案が強みに！"], ["他事業との連携強化", "不動産・解体・士業などと連携し、競合優位性を高める。", "ワンストップ体制！"]],
    eff: [["新たな顧客層の獲得", "相続や高齢化で発生する空き家所有者にリーチ。"], ["高単価案件の増加", "リノベーション・フルリフォームなど高単価案件を獲得。"], ["OB顧客との関係強化", "整理を通じて再接点を持ち、将来の提案につなげる。"], ["事業の安定化", "新築依存から脱却し、安定した事業ポートフォリオへ。"], ["地域での競争優位性", "片付けから活用まで対応できる企業としてブランド力向上。"]],
    flow: ["遺品整理\n（片付け・整理）", "不動産・相続相談\n（提携先連携）", "リノベ・リフォーム\n提案・施工", "アフターサポート\n（定期点検など）"],
    sum: "遺品整理事業の参入は、リフォーム・住宅会社にとって「実家リノベ案件の獲得」「相続空き家市場の攻略」「リフォーム受注の増加」を実現する重要な戦略です。" },
  { name: "葬儀業",
    issues: [["葬儀単価の下落", "価格競争により、葬儀単価が下落し、収益が減少している。"], ["集客コストの高騰", "紹介サイトや広告費の上昇により、新規集客コストが年々増加している。"], ["アフター収益不足", "葬儀後の顧客接点が少なく、継続的な収益機会が限られている。"], ["会員組織活用不足", "会員やOB顧客との接点を十分に活かせておらず、LTVが低下している。"]],
    syn: [["葬儀後サポート強化", "遺族の「その後の困りごと」を解決し、信頼度と満足度が向上。", "信頼を獲得！"], ["アフター売上創出", "遺品整理・生前整理・空き家整理などのニーズを獲得し、新たな収益源に。", "新たな収益を確保！"], ["会員組織の活用", "会員向け特典・優待サービスで、会員満足度と継続率が向上。", "会員の価値向上！"], ["関連事業への展開", "不動産・解体・墓じまい・相続相談など関連事業へ展開が可能に。", "事業を拡大！"]],
    eff: [["顧客満足度の向上", "葬儀後の不安や負担を解消し「頼んでよかった」と思われる。"], ["アフター収益の増加", "葬儀単価の下落を補う新たな収益の柱を確立。"], ["紹介・口コミの増加", "ご遺族や会員からの紹介が増え、新規顧客を獲得。"], ["LTVの向上", "リピートや関連サービスの利用が増加。"], ["競合との差別化", "葬儀＋アフターサポートのワンストップ体制で差別化。"]],
    flow: ["葬儀施行\n（本業）", "遺品整理\n（片付け・整理）", "関連サービス\n（生前整理など）", "関連事業連携\n（不動産・相続等）"],
    sum: "遺品整理事業の参入は、葬儀業者にとって「葬儀後サポート強化」「アフター売上創出」「会員組織の活用」「関連事業への展開」を実現し、安定した収益基盤と持続的な成長につながる強力な戦略となります。" },
  { name: "ケアマネジャー・包括支援センター",
    issues: [["相談先不足", "整理に関する相談先や、専門業者の紹介先が不足。"], ["高齢者の住環境問題", "空き家・独居高齢者の増加で、家の片付け・管理が課題に。"], ["家族からの相談増加", "施設入居や相続等の相談が増加し、専門的な対応が必要。"], ["業務負担の増加", "片付けや不用品処分の相談対応に時間を取られ、負担が大きい。"]],
    syn: [["地域課題の解決", "空き家や放置物件の問題を解決し、地域包括ケアの一翼を担う。", "地域の安心・安全に！"], ["生前整理ニーズに対応", "施設入居前や高齢者の生前整理を支援し、本人・家族の安心に。", "高齢者の暮らしをサポート！"], ["家族支援の強化", "遠方に住む家族の負担を軽減し、安心して相談できる体制を提供。", "感謝される支援で信頼向上！"], ["地域包括ケアとの連携強化", "多職種と連携し、地域包括ケアのハブとしての役割を強化。", "支援の幅が拡大！"]],
    eff: [["信頼・満足度の向上", "困りごとをワンストップで解決し、利用者や家族の信頼が高まる。"], ["相談件数の増加", "安心して紹介できる体制で、片付け・整理の相談が増加。"], ["業務負担の軽減", "専門業者と連携することで、コア業務に集中できる。"], ["地域包括ケアの推進", "地域課題の解決に貢献し、地域包括ケアシステムの推進に寄与。"], ["他事業者との連携強化", "不動産・葬儀・士業などとの連携が深まる。"]],
    flow: ["相談対応\n（ケアマネ・包括）", "専門業者の紹介\n（遺品整理事業）", "整理・片付け\n実施", "関連サービスへ連携\n（不動産・解体等）"],
    sum: "遺品整理事業の参入は、ケアマネジャー・包括支援センターにとって「地域課題の解決」「生前整理ニーズへの対応」「家族支援の強化」「地域包括ケアとの連携強化」を実現し、信頼される存在としての価値を高めます。" },
  { name: "便利業",
    issues: [["利益率が低い", "単価が低く、利益が残りにくい。"], ["案件数依存", "小口案件が多く、件数をこなしても十分な利益を確保しづらい。"], ["生産性が低い", "移動や見積りの時間が多く、効率が悪化している。"], ["ブランド力が弱い", "専門性や信頼性をアピールしづらい。"]],
    syn: [["高単価商品化", "高単価案件を獲得し、1件あたりの粗利が大幅にアップ！", "粗利が大幅UP！"], ["顧客単価の向上", "生前整理や空き家片付けなど関連サービス展開で、顧客単価が向上！", "関連サービスで単価UP！"], ["既存顧客の活用", "過去の小口依頼顧客へ提案し、リピート・追加受注が増加！", "リピートが増加！"], ["ブランド力の向上", "「遺品整理の専門店」として信頼性が高まり、選ばれる理由が明確に！", "選ばれる理由が明確に！"]],
    eff: [["収益性の向上", "高単価案件の増加により、売上・利益が安定。"], ["安定した案件確保", "季節や天候に左右されにくく、年間を通じて案件獲得が可能。"], ["顧客満足度の向上", "遺族や依頼者の負担を軽減し、満足度・信頼感が向上。"], ["紹介・口コミの増加", "満足度の高いサービスが口コミや紹介につながる。"], ["事業の専門特化", "専門サービスの確立により、価格競争から脱却。"]],
    flow: ["生前整理・\n片付け", "遺品整理", "空き家片付け", "各種オプション\n（リフォーム・買取等）"],
    sum: "遺品整理事業の参入は、便利屋業者にとって「高単価商品化」「顧客単価の向上」「ブランド力の向上」「安定収益の確保」を実現し、持続的な成長につながる強力な戦略となります。" },
  { name: "リサイクル業",
    issues: [["良質な仕入れが安定しない", "市場価格に左右されやすく、資源物の安定確保が難しい。"], ["利益率が低い", "回収・選別のみでは利益率が伸びにくい。"], ["一般のお客様との接点が少ない", "法人案件が中心となり、個人顧客からの直接依頼が少ない。"], ["回収できる物が限定される", "資源物以外は処分や廃棄となり、収益機会を逃している。"]],
    syn: [["良質なリユース品を直接確保", "遺品整理現場から家具・家電・骨董品・ブランド品などを直接回収。", "仕入れコストを抑えて利益率UP！"], ["お客様との直接取引が増える", "一般のお客様から遺品整理を直接受注でき、地域密着企業として信頼を獲得。", "新規顧客の獲得につながる！"], ["リユース・買取収益を最大化", "家具・家電・古物・金属などを適正にリユース・資源化。", "利益率の高い事業へ！"], ["周辺サービスへ展開", "不用品回収・生前整理・解体・リフォームなどへ展開可能。", "地域密着サービスへ発展！"]],
    eff: [["顧客数の増加", "一般家庭からの直接依頼が増加。"], ["利益率向上", "リユース・買取利益で収益性アップ。"], ["回収量増加", "資源物の回収量が増え、安定した仕入れを確保。"], ["地域ブランド向上", "「困ったら相談できる会社」として認知が拡大。"], ["継続収益化", "生前整理・空き家管理・不用品回収などへ展開。"]],
    flow: ["遺品整理", "選別・\n仕分け", "リユース・\n買取", "資源回収・\n処分"],
    sum: "遺品整理事業への参入は、リサイクル業者にとって「良質な仕入れの確保」「利益率向上」「一般顧客の獲得」「地域密着サービスの拡大」を実現する新たな成長戦略です。" },
  { name: "買取業",
    issues: [["高価値商品の仕入れが安定しない", "商品の仕入れはお客様の持込みや出張買取に依存し、安定した仕入れルートの確保が難しい。"], ["買取できない品物が多い", "買取対象外の商品はお客様自身で処分する必要があり、機会損失につながっている。"], ["一度の売上が小さい", "買取のみでは一案件あたりの売上が限られ、利益拡大に限界がある。"], ["相続世代との接点が少ない", "家の整理や相続相談など、高単価案件につながる入口が少ない。"]],
    syn: [["高価値商品の仕入れを強化", "遺品整理現場からブランド品・貴金属・時計・骨董品・家具・家電などを直接査定・買取。", "高利益商品の仕入れが増える！"], ["買取できない物も対応可能", "不要品の処分や搬出まで対応できるため、お客様の手間を大幅に軽減。", "「全部お願いできる」が強み！"], ["客単価・利益率アップ", "買取利益だけでなく、遺品整理・搬出・仕分け・不用品回収などの売上を追加。", "一案件あたりの売上が大幅UP！"], ["次のビジネスへ展開", "遺品整理を入口に、生前整理・空き家整理・解体・不動産売却などの紹介が可能。", "継続的な売上につながる！"]],
    eff: [["高利益商品の仕入れ増加", "ブランド品・貴金属・骨董品などの買取機会が増加。"], ["客単価向上", "買取だけでなく遺品整理費用も受注可能。"], ["顧客満足度向上", "整理から買取まで一括対応で安心感を提供。"], ["地域ブランド向上", "「相続・片付けの相談先」として認知が拡大。"], ["継続収益化", "生前整理・空き家整理・不動産紹介など、継続的な収益を確保。"]],
    flow: ["遺品整理\n相談", "仕分け・\n査定・買取", "販売", "不用品回収・\n空き家・不動産紹介"],
    sum: "遺品整理事業への参入は、買取業者にとって「高利益商品の安定仕入れ」「客単価向上」「顧客満足度向上」「継続的な収益機会の創出」を実現する新たな成長戦略です。" },
];
function industrySlide(d) {
  const s = content(S5, "既存ビジネスとのシナジー");
  subhead(s, d.name, { w: 9 });
  // 経営課題
  rect(s, 0.4, 1.55, 2.95, 3.4, WHITE, { line: NAVY });
  rect(s, 0.4, 1.55, 2.95, 0.42, NAVY);
  txt(s, "経営課題", 0.4, 1.55, 2.95, 0.42, { size: 14, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  d.issues.forEach((it, i) => {
    const y = 2.03 + i * 0.73;
    s.addShape(pres.shapes.OVAL, { x: 0.48, y: y + 0.05, w: 0.32, h: 0.32, fill: { color: NAVY }, line: { type: "none" }, objectName: nm("num") });
    txt(s, String(i + 1), 0.48, y + 0.05, 0.32, 0.32, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, it[0], 0.86, y, 2.45, 0.32, { size: 10.5, bold: true, color: NAVY, valign: "middle", margin: 0 });
    txt(s, it[1], 0.86, y + 0.3, 2.45, 0.42, { size: 7.5, color: DARK, margin: 0, ls: 1.1 });
    if (i < 3) rect(s, 0.5, y + 0.71, 2.75, 0.008, LINE);
  });
  // シナジー
  rect(s, 3.5, 1.55, 6.1, 3.4, WHITE, { line: NAVY });
  rect(s, 3.5, 1.55, 6.1, 0.42, NAVY);
  txt(s, "遺品整理事業参入によるシナジー", 3.5, 1.55, 6.1, 0.42, { size: 14, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  d.syn.forEach((sy, i) => {
    const w = 1.42, x = 3.6 + i * (w + 0.08);
    rect(s, x, 2.05, w, 2.8, LBLUE);
    s.addShape(pres.shapes.OVAL, { x: x + 0.05, y: 2.12, w: 0.3, h: 0.3, fill: { color: NAVY }, line: { type: "none" }, objectName: nm("num") });
    txt(s, String(i + 1), x + 0.05, 2.12, 0.3, 0.3, { size: 10, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, sy[0], x + 0.38, 2.08, w - 0.42, 0.55, { size: 9.5, bold: true, color: NAVY, valign: "middle", margin: 0 });
    txt(s, sy[1], x + 0.05, 2.7, w - 0.1, 1.6, { size: 8, color: DARK, ls: 1.15 });
    pill(s, x + 0.04, 4.43, w - 0.08, sy[2]);
  });
  // 期待できる効果
  rect(s, 0.4, 5.05, 5.5, 1.35, WHITE, { line: GREEN });
  rect(s, 0.5, 4.95, 1.65, 0.3, GREEN);
  txt(s, "期待できる効果", 0.5, 4.95, 1.65, 0.3, { size: 10.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  d.eff.forEach((e, i) => {
    const w = 1.08, x = 0.45 + i * w;
    txt(s, e[0], x, 5.3, w, 0.3, { size: 8, bold: true, color: GREEN, align: "center", valign: "middle", margin: 0 });
    txt(s, e[1], x + 0.03, 5.62, w - 0.06, 0.75, { size: 7, color: DARK, align: "center", margin: 0, ls: 1.1 });
    if (i) rect(s, x, 5.35, 0.008, 0.95, "B9D8C4");
  });
  // ビジネスモデルのイメージ
  rect(s, 6.05, 5.05, 3.55, 1.35, "F4F0FA", { line: PURPLE });
  rect(s, 6.05, 5.05, 3.55, 0.3, PURPLE);
  txt(s, "ビジネスモデルのイメージ", 6.05, 5.05, 3.55, 0.3, { size: 10, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  d.flow.forEach((f, i) => {
    const w = 0.8, x = 6.12 + i * 0.87;
    rect(s, x, 5.45, w, 0.85, WHITE, { line: "CBBDE3" });
    txt(s, f, x, 5.45, w, 0.85, { size: 7, bold: true, color: PURPLE, align: "center", valign: "middle", margin: 0 });
  });
  band(s, d.sum, 6.5, 0.55, { label: "まとめ", size: 9.5 });
  return s;
}
function realEstateExtra() {
  const s = content(S5, "既存ビジネスとのシナジー", "【要確認】法改正の出典（法務省の通達等）を確認し、出典欄に記載する。");
  rect(s, 0.4, 1.0, 3.0, 0.4, NAVY);
  txt(s, "不動産会社の経営者様へ", 0.4, 1.0, 3.0, 0.4, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  txt(s, [{ text: "不動産情報を「風上」で掴む方法の一つが、", options: { color: DARK, fontSize: 20, breakLine: true } }, { text: "使えなくなります。", options: { color: RED, fontSize: 30 } }], 0.4, 1.45, 9.2, 1.1, { bold: true, margin: 0 });
  panel(s, 0.4, 2.7, 2.95, 2.7, "これまで", BLUE);
  txt(s, "相続登記が申請される\n↓\n法務局の不動産登記受付帳を開示請求\n↓\n対象の所有者へDM・営業アプローチ\n↓\n売却相談を風上で獲得", 0.5, 3.15, 2.75, 2.2, { size: 10, bold: true, color: DARK, align: "center", ls: 1.15 });
  panel(s, 3.5, 2.7, 3.0, 2.7, "2026年10月1日から", RED);
  txt(s, "法改正により取得できる情報が変わります\n\n不動産登記受付帳の様式変更により、「登記の目的」「不動産の所在事項」が表記されなくなります。\n\n営業目的での情報取得が事実上不可能になります。", 3.6, 3.15, 2.8, 2.2, { size: 10, bold: true, color: DARK, ls: 1.15 });
  panel(s, 6.65, 2.7, 2.95, 2.7, "これから", NAVY);
  txt(s, "・相続登記を起点としたDM送付ができない\n・売却前の所有者へアプローチが難しくなる\n・ポータルサイト・反響広告・既存の人脈の競争がさらに激化", 6.75, 3.15, 2.75, 2.2, { size: 10, bold: true, color: DARK, ls: 1.2 });
  rect(s, 0.4, 5.55, 9.2, 1.45, NAVY);
  txt(s, [{ text: "だからこそ今、新しい「風上チャネル」を持つ必要があります！\n", options: { color: WHITE, fontSize: 13 } }, { text: "その新しい入口の一つが「空き家×遺品整理」です。", options: { color: YEL, fontSize: 20 } }], 0.6, 5.55, 8.8, 1.45, { bold: true, align: "center", valign: "middle" });
  note(s, "出典：（要記入）");
}
function realEstateSim() {
  const s = content(S5, "既存ビジネスとのシナジー");
  txt(s, "不動産仲介業者が遺品整理事業を始めることで得られるシナジー", 0.4, 0.98, 9.2, 0.5, { size: 17, bold: true, color: NAVY, valign: "middle", margin: 0 });
  txt(s, "遺品整理を「不動産案件の入口」にし、物件ランクに応じた最適な出口（買取再販・仲介・解体）で収益化。", 0.4, 1.45, 9.2, 0.35, { size: 10.5, bold: true, color: DARK, margin: 0 });
  steps(s, [["遺品整理の実施", "年間遺品整理案件数\n120件"], ["不動産に関するご相談", "不動産相談発生率 5〜10%\n＝年間6〜12件"], ["物件のランク分け", "状態・立地・築年数で最適な出口に振り分け"], ["シナジーによる収益化", "買取再販・仲介・解体それぞれで収益を確保"]], 1.9, 1.6, { size: 10, hsize: 10.5, hh: 0.45 });
  table(s, [
    ["区分", "年間件数", "1件あたり粗利（目安）", "年間粗利", "最適な出口"],
    ["S・Aランク（立地良・築浅）", "2〜3件", "100万〜300万円", "200万〜900万円", "買取再販"],
    ["Bランク（標準）", "3〜4件", "約100万円", "300万〜600万円", "売買仲介"],
    ["Cランク（立地悪・築古）", "2〜3件", "30万〜100万円", "60万〜300万円", "解体"],
    [{ text: "合計", options: { bold: true } }, { text: "7〜10件", options: { bold: true } }, "―", { text: "約560万〜1,800万円", options: { bold: true, color: RED } }, ""],
  ], 0.4, 3.7, 9.2, [2.6, 1.2, 2.0, 2.0, 1.4], { size: 10.5, rowH: 0.42 });
  band(s, "遺品整理からつながる不動産案件を、物件ランクに応じて最適な出口へ。「買取再販・仲介・解体」の連携で、売上アップ・利益拡大を実現しましょう！", 6.0, 0.95, { size: 12.5 });
  note(s, MODEL_NOTE);
}
INDUSTRIES.forEach((d, i) => {
  industrySlide(d);
  if (i === 0) { realEstateExtra(); realEstateSim(); }
});

// =====================================================================
// セミナーでのご質問について
// =====================================================================
const S6 = "セミナーでのご質問について";
pres.addSection({ title: S6 });
divider(S6, "セミナーでの\nご質問について", "事後経営相談はここから。最初にヒアリングページ（書き込み版）で前回の内容を確認する。");
{
  const s = content(S6, "セミナーでのご質問について", "セミナーのアンケートや前回いただいた質問を記入し、該当ページを開きながら回答する。");
  tag(s, "いただいた質問");
  table(s, [["いただいたご質問", "回答", "参照ページ"], ...Array.from({ length: 7 }, () => ["", "", ""])], 0.4, 1.45, 9.2, [3.3, 4.8, 1.1], { size: 12, align: "left", rowH: 0.68 });
}

// =====================================================================
// 貴社で進めるには
// =====================================================================
const S7 = "貴社で進めるには";
pres.addSection({ title: S7 });
divider(S7, "貴社で進めるには");
{
  const s = content(S7, "商圏分析", "地域データは自治体の統計・住宅土地統計調査などから取得する。");
  tag(s, "商圏データ");
  subhead(s, "貴社商圏の遺品整理市場");
  const K = [["商圏人口", "○○万人"], ["年間死亡者数", "○,○○○人"], ["65歳以上の単身世帯", "○,○○○世帯"], ["空き家数", "○,○○○戸"]];
  K.forEach((k, i) => {
    const x = 0.4 + i * 2.33;
    rect(s, x, 1.65, 2.15, 1.4, LBLUE, { line: LINE });
    txt(s, k[0], x, 1.72, 2.15, 0.4, { size: 12, bold: true, color: NAVY, align: "center" });
    txt(s, k[1], x, 2.15, 2.15, 0.7, { size: 22, bold: true, color: RED, align: "center", valign: "middle" });
  });
  ph(s, 0.4, 3.25, 5.0, 3.1, "商圏地図", "対象エリア（市区町村）と、競合業者の位置をプロットした地図を配置");
  ph(s, 5.6, 3.25, 4.0, 3.1, "競合状況", "商圏内の遺品整理事業者／商圏内の○○業者の動き\n（Googleマップ・HPで調査）");
  redMsg(s, "貴社商圏の遺品整理市場は、約○億円と試算されます。");
}
{
  const s = content(S7, "商品設計", "料金はパック料金で見せる（内訳をそのまま見せると不安・不信感につながる）。");
  tag(s, "料金・オプション");
  const L = [0.4, 4.7];
  rect(s, 0.4, 1.0, 4.1, 0.5, "5A5F6A");
  txt(s, "✕ 悪い見せ方（内訳をそのまま見せる）", 0.4, 1.0, 4.1, 0.5, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  rect(s, 0.4, 1.5, 4.1, 5.45, WHITE, { line: "5A5F6A" });
  txt(s, "内訳を細かく見せると、かえって不安や不信感につながります。", 0.5, 1.55, 3.9, 0.4, { size: 9.5, bold: true, color: DARK, align: "center", margin: 0 });
  table(s, [["【間取り：2LDKの場合の料金例】", ""], ["交通費", "20,000円"], ["作業費", "80,000円"], ["処理費", "60,000円"], ["基本清掃費", "10,000円"], ["運搬費", "30,000円"], [{ text: "合計", options: { bold: true } }, { text: "200,000円", options: { bold: true } }]], 0.55, 2.0, 2.4, [1.2, 1.2], { size: 9.5, rowH: 0.32, head: "5A5F6A" });
  txt(s, "何にいくらかかるのかよくわからない…\n高いのか安いのか判断できない…\n追加料金が発生しそうで不安…", 3.0, 2.0, 1.45, 2.2, { size: 8, color: DARK, ls: 1.2 });
  txt(s, "悪い見せ方のデメリット\n✕ 料金の内訳が多く、分かりにくい\n✕ 何が必要で何が不要かわからず不安になる\n✕ 他社比較がしにくい\n✕ 追加料金を疑われ、信頼を損なう", 0.55, 4.45, 3.8, 1.6, { size: 9.5, bold: true, color: DARK, ls: 1.3 });
  rect(s, 4.7, 1.0, 4.9, 0.5, "C8102E");
  txt(s, "○ 良い見せ方（パック料金で見せる）", 4.7, 1.0, 4.9, 0.5, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  rect(s, 4.7, 1.5, 4.9, 5.45, WHITE, { line: "C8102E" });
  table(s, [["間取り", "物量の目安", "人数", "料金（税込）"], ["1K", "段ボール20箱程度", "1名", "30,000円〜"], ["1LDK", "段ボール50箱程度", "2〜3名", "80,000円〜"], ["2LDK", "段ボール100箱程度", "3〜4名", "150,000円〜"], ["3LDK", "段ボール150箱程度", "4〜6名", "200,000円〜"], ["4LDK以上", "段ボール200箱以上", "5名〜", "300,000円〜"]], 4.8, 1.6, 4.7, [0.9, 1.75, 0.8, 1.25], { size: 9.5, rowH: 0.33, head: "C8102E" });
  txt(s, "仕分け・搬出・処分・簡易清掃・運搬まで含んだ安心のパック料金です！", 4.8, 3.65, 4.7, 0.35, { size: 9.5, bold: true, color: RED, align: "center", margin: 0 });
  txt(s, "さらに安心のオプション（必要な場合のみ追加）", 4.8, 4.05, 4.7, 0.3, { size: 10, bold: true, color: DARK, margin: 0 });
  const O = [["供養", "15,000円〜"], ["特殊清掃", "30,000円〜"], ["原状回復", "40,000円〜"], ["解体", "100,000円〜"], ["空き家管理", "5,000円〜/月"], ["買取", "お値引き"]];
  O.forEach((o, i) => {
    const x = 4.8 + i * 0.79;
    rect(s, x, 4.4, 0.74, 0.85, "FFF5F5", { line: "E8B4B8" });
    txt(s, o[0] + "\n" + o[1], x, 4.4, 0.74, 0.85, { size: 7.5, bold: true, color: DARK, align: "center", valign: "middle", margin: 0 });
  });
  txt(s, "良い見せ方のメリット\n✓ 総額がすぐにわかり、安心して検討できる\n✓ 他社比較がしやすく、選ばれやすい\n✓ 追加オプションは必要な方だけ提案できる\n✓ 信頼感が高まり、成約率が向上する", 4.85, 5.35, 4.6, 1.55, { size: 9.5, bold: true, color: DARK, ls: 1.25 });
}
{
  const s = content(S7, "ビジネスモデルと売上目標", "1年目は準備期間を含めた月ごとの立ち上がりで示す。兼任スタート案と専任採用案を並べる。");
  tag(s, "貴社の数値");
  subhead(s, "貴社版 数値計画");
  s.addChart(pres.charts.BAR, [{ name: "月間契約件数（目安）", labels: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"], values: [0, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10] }], {
    x: 0.4, y: 1.6, w: 4.6, h: 3.4, barDir: "col", chartColors: [NAVY], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 9, dataLabelColor: DARK,
    showTitle: true, title: "1年目の月間契約件数（例）", titleFontSize: 12, titleColor: NAVY, titleFontFace: "+mn-lt",
    catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, catAxisLabelColor: GRAY, valAxisLabelColor: GRAY, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
    valGridLine: { color: "E5E7EB", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false, objectName: "ramp_chart",
  });
  table(s, [["", "兼任スタート案", "専任採用案"], ["1年目の体制", "既存営業1名が兼任", "専任営業1名を採用"], ["人件費", "本業と按分", "年600万円"], ["1年目売上", "○○万円", "○○万円"], ["2年目売上", "○○万円", "○○万円"], ["3年目売上", "○○万円", "○○万円"]], 5.2, 1.6, 4.4, [1.4, 1.5, 1.5], { size: 10.5, rowH: 0.5 });
  ph(s, 5.2, 4.75, 4.4, 1.35, "前提条件", "商圏人口・平均単価・契約率・広告予算など");
  redMsg(s, "準備期間を含めて、無理のない立ち上がりを計画します。");
  note(s, MODEL_NOTE);
}
{
  const s = content(S7, "ビジネスモデルと売上目標", "仲介手数料は宅建業法の上限（売買価格400万円超：3%＋6万円、税別）で計算。発生率は貴社の実績に合わせて調整する。");
  tag(s, "発生率・単価");
  subhead(s, "本業を含めた収益シミュレーション", { w: 6 });
  table(s, [
    ["項目（10件受託した場合）", "計算", "粗利"],
    ["遺品整理", "30万円 × 10件 × 粗利率60%", "180万円"],
    ["不動産売買（10件中1件）", "売買価格2,000万円の仲介手数料（両手）", "132万円"],
    ["解体工事（10件中1件）", "工事費150万円 × 粗利率30%", "45万円"],
    [{ text: "合計", options: { bold: true } }, "", { text: "357万円", options: { bold: true, color: RED } }],
  ], 0.4, 1.65, 6.2, [2.2, 2.9, 1.1], { size: 11, rowH: 0.55 });
  rect(s, 6.8, 1.65, 2.8, 2.75, LBLUE, { line: LINE });
  txt(s, [{ text: "1件あたりの粗利\n", options: { fontSize: 12, color: NAVY } }, { text: "遺品整理のみ　", options: { fontSize: 11, color: DARK } }, { text: "18万円\n", options: { fontSize: 18, color: NAVY } }, { text: "本業を含めると　", options: { fontSize: 11, color: DARK } }, { text: "約36万円", options: { fontSize: 22, color: RED } }], 6.85, 1.7, 2.7, 2.65, { bold: true, align: "center", valign: "middle", ls: 1.3 });
  txt(s, "※仲介手数料は宅建業法の上限（3%＋6万円、税別）で計算しています。", 0.4, 4.55, 9.2, 0.3, { size: 9.5, color: GRAY });
  redMsg(s, "本業まで含めると、遺品整理単体の約2倍の利益が残ります。", { y: 5.4 });
  note(s, MODEL_NOTE);
}

// ---- 集客について ----
divider(S7, "集客について", "", true);
{
  const s = content(S7, "集客について");
  subhead(s, "遺品整理事業で業績を上げるための集客3ステップ", { w: 9.2 });
  const C = [["① 一丁目一番地", "専門店HPと屋号を決める", "・信頼される屋号（ブランド）を決定\n・専門店としてのコンセプトを明確化", "ここがすべてのスタートライン", NAVY], ["② HPの設計・コンテンツ整備", "わかりやすく、頼みやすい\n反響が取れるHPにする", "・わかりやすい動線設計（相談までスムーズ）\n・料金・サービスを明確に掲載\n・事例・お客様の声で信頼性を強化\n・スマホ最適化で見やすく使いやすく", "反響が取れるHPを構築", BLUE], ["③ 販促活動で集客を最大化", "多様なチャネルで\n幅広く集客し、相談件数を最大化", "・チラシ・ポスティング\n・WEB広告（検索・ディスプレイ）\n・SNS広告\n・地域情報サイト\n・紹介・提携先からのご紹介\n・会員・既存顧客へのDM", "集客の最大化を実現", "1E7A8A"]];
  C.forEach((c, i) => {
    const x = 0.4 + i * 3.13, w = 2.94;
    rect(s, x, 1.65, w, 5.3, WHITE, { line: c[4], lw: 1.5 });
    rect(s, x, 1.65, w, 0.55, c[4]);
    txt(s, c[0], x, 1.65, w, 0.55, { size: 12.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, c[1], x + 0.1, 2.3, w - 0.2, 0.8, { size: 12.5, bold: true, color: DARK, align: "center", valign: "middle" });
    photo(s, x + 0.2, 3.15, w - 0.4, 1.2, "イメージ");
    txt(s, c[2], x + 0.12, 4.45, w - 0.24, 1.75, { size: 9.5, bold: true, color: DARK, ls: 1.25 });
    rect(s, x + 0.15, 6.3, w - 0.3, 0.5, i === 0 ? "FBEFD5" : LBLUE);
    txt(s, c[3], x + 0.15, 6.3, w - 0.3, 0.5, { size: 11, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    if (i < 2) arrowR(s, x + w + 0.02, 3.9, 0.17, 0.4, NAVY);
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.36, y: 1.61, w: 3.02, h: 5.38, fill: { color: WHITE, transparency: 100 }, line: { color: RED, width: 2.5 }, objectName: "red_frame" });
}
{
  const s = content(S7, "集客について");
  rect(s, 0.4, 1.0, 9.2, 4.95, NAVY);
  banner(s, "ホームページ完成後の集客戦略", 1.1, { color: WHITE, size: 22 });
  txt(s, "5つの集客施策で、安定した反響を継続的に獲得する", 0.4, 1.65, 9.2, 0.4, { size: 13, bold: true, color: "F6C700", align: "center", margin: 0 });
  const F = [["01", "SEO対策", "検索で見つけてもらう", "お役立ち記事の作成\n地域ページの充実\n事例・FAQの掲載", "長期的に反響が増える\n資産型集客"], ["02", "MEO対策", "地域検索で選ばれる", "Googleビジネス最適化\n写真・投稿の定期更新\n口コミの獲得", "地域密着で信頼される\n集客力アップ"], ["03", "Web広告", "今すぐ客を集める", "Google広告・SNS広告\n地域・年齢を細かく設定\n短期間で問い合わせ増加", "即効性が高く\n短期間で反響を獲得"], ["04", "チラシ・販促", "地域に認知してもらう", "折込チラシ・ポスティング\nDM・冊子・チェックリスト\n相談会・セミナー開催", "潜在層にアプローチし\n認知と信頼を拡大"], ["05", "アライアンス提携", "紹介で安定的に獲得する", "葬儀会社・不動産会社\nケアマネ・士業・便利屋など\n紹介ネットワークを構築", "広告費に依存せず\n安定した紹介を獲得"]];
  F.forEach((f, i) => {
    const x = 0.55 + i * 1.8, w = 1.7;
    rect(s, x, 2.15, w, 3.65, WHITE, { round: true, r: 0.05 });
    s.addShape(pres.shapes.OVAL, { x: x + w / 2 - 0.25, y: 2.0, w: 0.5, h: 0.4, fill: { color: NAVY }, line: { color: WHITE, width: 1 }, objectName: nm("num") });
    txt(s, f[0], x + w / 2 - 0.25, 2.0, 0.5, 0.4, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, f[1], x, 2.45, w, 0.45, { size: 14, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    txt(s, f[2], x, 2.88, w, 0.3, { size: 8.5, bold: true, color: GOLD, align: "center", margin: 0 });
    txt(s, f[3], x + 0.08, 3.3, w - 0.16, 1.3, { size: 8.5, bold: true, color: DARK, ls: 1.35 });
    rect(s, x + 0.08, 4.85, w - 0.16, 0.8, NAVY);
    txt(s, f[4], x + 0.08, 4.85, w - 0.16, 0.8, { size: 8.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  });
  redMsg(s, "遺品整理専門店として、５つの施策で効率的に集客を行う", { y: 6.2, size: 21 });
}
{
  const s = content(S7, "紙販促について");
  subhead(s, "紙販促");
  photo(s, 0.4, 1.6, 4.3, 5.3, "チラシ（表面）の画像");
  txt(s, "1.「遺品整理」を大きく打ち出す\n一目で何のサービスかわかるよう、最も目立つ位置に大きく掲載します。\n\n2.「誰のためのサービスか」を明確にする\n実家の片付けや相続で悩む方など、対象者を具体的に伝えます。\n\n3.「何を解決できるか」を明確にする\n遺品整理・空き家整理・買取など対応内容を分かりやすく示します。\n\n4.「保存版」を強調する\n今すぐ必要でなくても保管してもらえるよう促します。\n\n5.「実績・安心材料」を掲載する\n口コミ・実績・資格などで信頼感を高めます。", 4.95, 1.6, 4.65, 5.3, { size: 12, bold: true, color: DARK, ls: 1.15 });
}
{
  const s = content(S7, "紙販促について");
  subhead(s, "紙販促");
  photo(s, 0.4, 1.6, 4.3, 5.3, "チラシ（裏面）の画像");
  txt(s, "1. 対応サービスを記載する\n整理後の悩みまでワンストップで対応できることを伝えます。\n\n2. 料金プランを掲載する\n費用の目安を示し、問い合わせの不安を軽減します。\n\n3. 料金をわかりやすくする\n間取り別・物量別で比較しやすく掲載します。\n\n4. 追加料金の考え方を明記する\n料金トラブルへの不安を解消します。\n\n5. 作業量の目安を掲載する\nどのプランが適切か判断しやすくします。", 4.95, 1.6, 4.65, 5.3, { size: 12, bold: true, color: DARK, ls: 1.15 });
}
{
  const s = content(S7, "アライアンス提携について");
  subhead(s, "アライアンス提携");
  txt(s, [{ text: "アライアンス提携とは、地域のさまざまな専門企業や団体と", options: { color: DARK } }, { text: "信頼関係", options: { color: RED } }, { text: "を築き、\nお互いにお客様を紹介し合うことで、", options: { color: DARK } }, { text: "継続的に高単価・相続関連案件を獲得する仕組み", options: { color: RED } }, { text: "です。", options: { color: DARK } }], 0.4, 1.5, 9.2, 0.65, { size: 12.5, bold: true, margin: 0 });
  const B = [["ポイント", "地域の相続・空き家案件の発生源を押さえ、継続的に高単価案件を獲得する仕組みです。"], ["具体的にすること", "・不動産会社　・司法書士　・行政書士　・税理士\n・葬儀社　・ケアマネジャー　・地域包括支援センター\n・解体会社　・リフォーム会社　・ガス会社　・役所 など\nと提携する。"], ["意味合い", "単なる紹介獲得ではなく、地域の相続・空き家案件が最初に集まる会社（地域の案件ハブ）になることが目的です。"]];
  B.forEach((b, i) => {
    const y = 2.3 + i * 1.5;
    rect(s, 0.4, y, 1.7, 0.36, NAVY, { round: true, r: 0.15 });
    txt(s, b[0], 0.4, y, 1.7, 0.36, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    rect(s, 0.4, y + 0.4, 4.2, 1.0, WHITE, { line: NAVY });
    txt(s, b[1], 0.5, y + 0.42, 4.0, 0.95, { size: 9.5, bold: true, color: DARK, valign: "middle", ls: 1.2 });
  });
  s.addShape(pres.shapes.OVAL, { x: 4.85, y: 2.2, w: 4.75, h: 4.0, fill: { color: "FFF8E8" }, line: { type: "none" }, objectName: "net_bg" });
  rect(s, 6.35, 3.75, 1.75, 0.9, NAVY, { round: true });
  txt(s, "当社\n（遺品整理業）", 6.35, 3.75, 1.75, 0.9, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  const N = [["不動産会社", 6.55, 2.3], ["葬儀社", 8.3, 2.85], ["解体会社", 8.4, 4.0], ["リフォーム会社", 8.1, 5.15], ["役所", 6.55, 5.55], ["ケアマネジャー", 4.95, 5.15], ["税理士", 4.9, 4.0], ["司法書士\n行政書士", 5.0, 2.85]];
  N.forEach(([t, x, y]) => {
    rect(s, x, y, 1.35, 0.55, WHITE, { line: LINE, shadow: true });
    txt(s, t, x, y, 1.35, 0.55, { size: 9, bold: true, color: DARK, align: "center", valign: "middle", margin: 0 });
  });
  band(s, "期待できる効果：「高単価案件の獲得」「相見積もりを回避」「案件の継続的な流入」「相互送客による関係強化」", 6.4, 0.6, { size: 11 });
}
{
  const s = content(S7, "アライアンス提携について");
  subhead(s, "アライアンス提携");
  txt(s, [{ text: "初回訪問の目的は「ご挨拶と情報提供」。", options: { color: DARK } }, { text: "相手に安心感を与え、資料を置かせてもらう", options: { color: RED } }, { text: "ことをゴールにします。", options: { color: DARK } }], 0.4, 1.5, 9.2, 0.4, { size: 12.5, bold: true, margin: 0 });
  panel(s, 0.4, 2.0, 5.3, 4.95, "基本トークの流れ（5ステップ）", NAVY);
  const T = [["挨拶・自己紹介", "「はじめまして、○○（会社名）の□□と申します。本日はお時間をいただき、ありがとうございます。」"], ["訪問の目的を伝える", "「本日は営業のご案内ではなく、御社のお役に立てるかもしれない情報をお届けに伺いました。」"], ["相手のメリットを伝える", "「遺品整理や生前整理でお困りのお客様がいらっしゃった際に、御社がお役に立てると思っております。」"], ["資料のお渡し", "「サービス内容や事例をまとめた資料です。ご参考になれば幸いです。」"], ["今後のつながりをつくる", "「何かお困りのお客様がいらっしゃいましたら、いつでもご相談ください。」"]];
  T.forEach((t, i) => {
    const y = 2.5 + i * 0.88;
    s.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.12, w: 0.38, h: 0.38, fill: { color: NAVY }, line: { type: "none" }, objectName: nm("num") });
    txt(s, String(i + 1), 0.5, y + 0.12, 0.38, 0.38, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    rect(s, 0.95, y + 0.08, 1.6, 0.48, LBLUE);
    txt(s, t[0], 0.95, y + 0.08, 1.6, 0.48, { size: 9.5, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    txt(s, t[1], 2.6, y, 3.05, 0.75, { size: 8, color: DARK, valign: "middle", ls: 1.15 });
  });
  panel(s, 5.85, 2.0, 3.75, 2.75, "トークのポイント", NAVY);
  txt(s, "■ 売り込まない\n　まずは「信頼関係の土台づくり」を最優先に。\n■ 相手の立場で話す\n　「御社のお役に立つかもしれない」というスタンスで。\n■ 資料を置くことをゴールにする\n■ 次回のきっかけをつくる", 5.95, 2.48, 3.55, 2.2, { size: 9.5, bold: true, color: DARK, ls: 1.2 });
  panel(s, 5.85, 4.9, 3.75, 2.05, "初回訪問のゴール", NAVY);
  steps(s, [["資料", "置かせてもらう"], ["安心感", "持ってもらう"], ["関係", "相談できる第一歩"]], 5.4, 1.4, { x: 5.95, w: 3.55, num: false, size: 8.5, hsize: 10, hh: 0.42 });
}

// ---- 営業について ----
divider(S7, "営業について", "", true);
{
  const s = content(S7, "営業について");
  banner(s, "遺品整理事業 営業KPI", 0.98, { lines: false, size: 26, sub: "契約率80%を基準にした3年間の営業目標" });
  table(s, [
    ["", "1年目", "2年目", "3年目"],
    ["営業人員", "1名", "2名", "3名"],
    ["売上", "¥36,000,000", "¥54,000,000", "¥100,000,000"],
    ["単価", "¥300,000", "¥300,000", "¥300,000"],
    ["年間契約件数", "120件", "180件", "333件"],
    ["月間契約件数", "10件", "15件", "28件"],
    ["月間集客件数", "12.5件", "18.75件", "35件"],
    ["一人当たり契約件数（月）", "10件", "7.5件", "約9.3件"],
  ], 0.8, 1.95, 8.4, [2.7, 1.9, 1.9, 1.9], { size: 13, rowH: 0.5, firstColFill: true });
  band(s, "ポイント：「契約率80%」を基準に設定　／　売上は3年で約2.8倍　／　「一人当たり生産性」を最大化", 6.2, 0.6, { size: 12 });
  note(s, "※契約率は全期間80%を基準として算出しています。");
}
{
  const s = content(S7, "営業について");
  numHead(s, "01", "遺品整理営業の全体フロー", "遺品整理営業は「グリーフケア×課題解決提案×スピード対応」で信頼を獲得し、契約・高粗利化につなげます。");
  steps(s, [["問い合わせ", ""], ["現地確認", ""], ["見積提出・契約", ""], ["作業実施", ""]], 1.95, 0.5, { x: 0.4, w: 6.1, num: false, hh: 0.5, size: 9, hsize: 11, colors: [NAVY, NAVY, NAVY, NAVY] });
  rect(s, 0.4, 2.6, 6.1, 4.35, WHITE, { line: NAVY });
  rect(s, 1.5, 2.6, 3.9, 0.38, NAVY);
  txt(s, "案件の規模で営業方法を使い分ける", 1.5, 2.6, 3.9, 0.38, { size: 11.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  const A = [["30万円未満の案件", "オンライン中心で効率化", ["0次接客（電話ヒアリング）", "アンケート送付（SMS・LINE）", "写真提出（お客様）", "オンライン見積提出", "電話クロージング・契約", "作業実施"], "・現地訪問を省略し、コスト削減→価格に反映\n・即日見積・即対応で契約率アップ", ORANGE], ["30万円以上の案件", "現地訪問中心で信頼と提案力を強化", ["0次接客（電話ヒアリング）", "事前情報収集・日程調整", "現地確認（訪問）", "見積提示・追加提案", "即決クロージング・契約", "作業実施"], "・信頼構築と正確な見積で安心感を提供\n・追加提案で単価アップを実現", BLUE]];
  A.forEach(([h, sub, st, pt, c], i) => {
    const x = 0.5 + i * 3.0, w = 2.9;
    rect(s, x + 0.3, 3.08, w - 0.6, 0.3, c, { round: true, r: 0.15 });
    txt(s, h, x + 0.3, 3.08, w - 0.6, 0.3, { size: 10, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, sub, x, 3.4, w, 0.3, { size: 10.5, bold: true, color: c, align: "center", margin: 0 });
    st.forEach((t, k) => {
      rect(s, x + 0.05, 3.75 + k * 0.38, w - 0.1, 0.34, i ? "EEF3FB" : "FFF4EA");
      txt(s, `STEP${k + 1}　${t}`, x + 0.12, 3.75 + k * 0.38, w - 0.2, 0.34, { size: 8.5, bold: true, color: DARK, valign: "middle", margin: 0 });
    });
    txt(s, pt, x + 0.05, 6.1, w - 0.1, 0.8, { size: 8.5, bold: true, color: DARK, ls: 1.2 });
  });
  panel(s, 6.65, 1.95, 2.95, 5.0, "3つの重要ポイント", NAVY);
  txt(s, [{ text: "1 スピード対応\n", options: { color: ORANGE, fontSize: 13 } }, { text: "問い合わせ後5分以内の架電が理想。スピードが信頼に直結します。\n\n", options: { color: DARK } }, { text: "2 ヒアリング力\n", options: { color: ORANGE, fontSize: 13 } }, { text: "お客様の状況・課題・希望を正確に把握し、最適な提案につなげます。\n\n", options: { color: DARK } }, { text: "3 信頼構築と提案力\n", options: { color: ORANGE, fontSize: 13 } }, { text: "遺族の気持ちに寄り添い、課題解決のプロとして信頼を獲得します。", options: { color: DARK } }], 6.75, 2.45, 2.75, 4.4, { size: 9.5, bold: true, ls: 1.15 });
}
{
  const s = content(S7, "営業について");
  numHead(s, "02", "0次接客の重要性", "お問い合わせ後、初めてお客様とお会いする前に行う電話でのヒアリングが、その後の成約を大きく左右します。");
  txt(s, [{ text: "お客様は", options: { color: DARK } }, { text: "電話の印象", options: { color: RED } }, { text: "で「この会社に任せたいか」を判断しています。", options: { color: DARK } }], 0.4, 1.9, 9.2, 0.45, { size: 15, bold: true, align: "center", valign: "middle", margin: 0 });
  panel(s, 0.4, 2.45, 2.9, 4.5, "お客様が知りたいこと", ORANGE);
  txt(s, [{ text: "この会社は信頼できる？\n", options: { color: ORANGE, fontSize: 12 } }, { text: "親身に話を聞いてくれるか、丁寧に対応してくれるかを見ています。\n\n", options: { color: DARK } }, { text: "高くない？\n", options: { color: ORANGE, fontSize: 12 } }, { text: "料金の目安や追加料金の有無など、費用に関する不安を解消したいと思っています。\n\n", options: { color: DARK } }, { text: "対応は丁寧？\n", options: { color: ORANGE, fontSize: 12 } }, { text: "不安な気持ちに寄り添ってくれる、安心して相談できる雰囲気を重視しています。", options: { color: DARK } }], 0.5, 2.92, 2.7, 3.95, { size: 9.5, bold: true, ls: 1.15 });
  panel(s, 3.45, 2.45, 2.9, 4.5, "0次接客の3つの目的", NAVY);
  txt(s, "① 要望の正確な把握\n現状やご要望を正確に把握し、最適な提案につなげる。\n\n② 単価アップにつながる情報収集\nオプションや付帯サービスの提案材料を収集する。\n\n③ 信頼の獲得\n「この会社なら安心して任せられる」と思ってもらう。", 3.55, 2.92, 2.7, 3.95, { size: 9.5, bold: true, color: DARK, ls: 1.15 });
  panel(s, 6.5, 2.45, 3.1, 4.5, "営業側が確認すること", NAVY);
  txt(s, "■ 基本情報\nご入電者様と故人様の関係／決定権者（キーマン）／ご住所／作業希望時期／予算感／他社への相談状況\n■ 現場の状況\n間取り・階数／エレベーター／駐車スペース／物量／重量物\n■ お客様の個別事情\n捜索物／買取希望／供養や特殊清掃の希望／立会い予定\n■ 遺品整理後のご予定\n売却／解体／リフォーム／賃貸の退去", 6.6, 2.92, 2.9, 3.95, { size: 8.5, bold: true, color: DARK, ls: 1.15 });
}
{
  const s = content(S7, "営業について");
  numHead(s, "03", "契約率を上げるクロージング", "見積説明とクロージングの質が契約率を左右します。お客様の不安を解消し、安心してご契約いただきましょう。");
  panel(s, 0.4, 1.95, 3.0, 4.0, "クロージングの基本ステップ", NAVY);
  txt(s, "① 見積内容のご説明\n明確で分かりやすく説明し、不安や疑問を解消する\n\n② 安心のご提供\n追加料金なし・作業品質・保証などをお伝えして安心感を醸成\n\n③ 日程のご提案\n具体的な日程を提示し、スピード感を持ってご案内\n\n④ 契約のご提案・締結\n契約方法を案内し、スムーズにご契約へ", 0.5, 2.42, 2.8, 3.5, { size: 9.5, bold: true, color: DARK, ls: 1.15 });
  panel(s, 3.55, 1.95, 6.05, 4.0, "よくあるお客様の不安と解消トーク例", NAVY);
  table(s, [["お客様の不安", "解消トーク例"], ["他社と比較したい", "お気持ちはよく分かります。ただ、ご希望の日程が埋まってしまう可能性がございますので、まずは日程だけでも仮押さえさせていただきますと安心です。"], ["金額が高いかも…", "ご予算に合わせて最適なプランをご提案しております。また、買取や供養などを活用して費用を抑えることも可能です。"], ["追加料金が不安", "お見積りはすべてコミコミ価格です。当日お荷物が増えない限り、追加料金は一切かかりません。"], ["作業スタッフが信頼できるか不安", "当社は自社スタッフが対応いたします。礼儀・マナー研修を徹底しておりますので、安心してお任せいただけます。"]], 3.65, 2.45, 5.85, [1.7, 4.15], { size: 8.5, align: "left", rowH: 0.68 });
  rect(s, 0.4, 6.05, 9.2, 0.9, "FFF4EA", { line: ORANGE });
  txt(s, [{ text: "契約を後押しする3つのキーフレーズ　", options: { color: ORANGE, fontSize: 11 } }, { text: "①共感「そのお気持ち、よく分かります。」　②安心「追加料金は一切ありません。」　③希少性「ご希望の日程は埋まりやすいので、早めのご決断がおすすめです。」", options: { color: DARK, fontSize: 9.5 } }], 0.5, 6.05, 9.0, 0.9, { bold: true, valign: "middle" });
}

// ---- 見積出しについて ----
divider(S7, "見積出しについて", "", true);
{
  const s = content(S7, "見積作成について");
  const C = [["ダンボールのサイズと立米換算", "140ダンボールとは？", "3辺の合計が140cmのダンボールです。", "10箱で約1立米（1㎥）", "140ダンボール10箱分の容量がおよそ1立米になります。"], ["2トントラックの積載目安", "最大積載量は約10〜12立米", "ダンボールに換算すると、約100〜120箱分を積むことができます。", "10〜12立米／100〜120箱分", "※トラックのサイズや荷物の形状により、積載量は前後します。"], ["作業量の目安（1人あたり）", "1人あたり1.2立米／日", "荷物の種類や現場の状況により、多少前後します。", "1.2立米／日", "※荷物の種類・搬出経路・作業環境により変動します。"]];
  C.forEach((c, i) => {
    const x = 0.4 + i * 3.13, w = 2.94;
    rect(s, x, 1.0, w, 0.5, PURPLE);
    txt(s, c[0], x, 1.0, w, 0.5, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    rect(s, x, 1.5, w, 5.45, WHITE, { line: "CBBDE3" });
    txt(s, c[1], x + 0.1, 1.6, w - 0.2, 0.45, { size: 13, bold: true, color: DARK, font: MINCHO });
    txt(s, c[2], x + 0.1, 2.05, w - 0.2, 0.7, { size: 10, color: DARK, ls: 1.2 });
    photo(s, x + 0.15, 2.8, w - 0.3, 1.8, "写真");
    rect(s, x + 0.15, 4.75, w - 0.3, 1.25, "F4F0FA");
    txt(s, c[3], x + 0.15, 4.75, w - 0.3, 1.25, { size: 16, bold: true, color: PURPLE, align: "center", valign: "middle", font: MINCHO });
    txt(s, c[4], x + 0.1, 6.05, w - 0.2, 0.8, { size: 9, color: DARK, ls: 1.2 });
  });
}
{
  const s = content(S7, "見積作成について", "立米単価は紹介案件の場合 7,000〜10,000円が目安。");
  photo(s, 0.4, 1.0, 5.0, 5.9, "遺品整理 現地調査・見積シート（段ボール換算方式）");
  txt(s, "140サイズ段ボール\n＝0.605m × 0.435m × 0.345m\n≒ 約0.09㎥\n≒ 0.1㎥\n\n1リューベ＝140サイズ10箱\n\n1リューベ単価＝12,000〜15,000円\n※紹介等の場合は7,000〜10,000円", 5.6, 1.1, 4.0, 3.6, { size: 15, color: DARK, ls: 1.2 });
  txt(s, "段ボール何箱分に該当するか確認。\n箱単価を決めることで、見積金額を算出。", 5.6, 5.0, 4.0, 1.4, { size: 18, bold: true, color: RED, valign: "middle" });
}

// ---- 現場について ----
divider(S7, "現場について", "", true);
{
  const s = content(S7, "営業の現場担当について");
  txt(s, "遺品整理事業　3ヶ年の人員体制イメージ", 0.4, 0.98, 9.2, 0.55, { size: 24, bold: true, color: NAVY, valign: "middle", margin: 0 });
  txt(s, "正社員は営業・管理機能を担い、現場作業は派遣・アルバイトを中心に運営", 0.4, 1.5, 9.2, 0.38, { size: 13, bold: true, color: DARK, margin: 0 });
  const Y = [["1年目｜立上げ期", "基盤をつくる\nまずは一歩、着実に", "正社員営業 1名", "問い合わせ対応／現地見積／契約／現場作業／協力スタッフへの指示", "派遣社員・アルバイト中心。正社員も大型家具の搬出など現場に入る", "まずは営業＋現場の両方を経験"], ["2年目｜拡大期", "案件を増やし\n仕組みを強化", "正社員営業 2名", "営業・見積対応を分担しながら、現場管理・スタッフ指示を強化", "派遣社員・アルバイトを活用。正社員は「作業員」から「現場責任者」へ移行", "案件数の拡大と管理体制の強化"], ["3年目｜組織化期", "安定した成長と\n持続できる組織へ", "正社員営業 3名", "営業／見積／案件管理／現場監督／スタッフ育成・品質管理", "実作業は派遣社員・アルバイト中心。正社員は現場監督・管理業務が中心", "営業・管理機能を中心とした組織へ"]];
  Y.forEach((y, i) => {
    const x = 0.4 + i * 3.13, w = 2.94;
    rect(s, x, 2.0, w, 4.0, "F5F7FA", { line: LINE });
    rect(s, x, 2.0, w, 0.6, NAVY);
    txt(s, y[0], x + 0.1, 2.0, 1.75, 0.6, { size: 13, bold: true, color: WHITE, valign: "middle", margin: 0 });
    txt(s, y[1], x + 1.8, 2.0, w - 1.85, 0.6, { size: 7.5, bold: true, color: WHITE, valign: "middle", margin: 0 });
    txt(s, "正社員体制", x + 0.1, 2.68, w - 0.2, 0.3, { size: 10.5, bold: true, color: NAVY, margin: 0 });
    rect(s, x + 0.15, 2.98, w - 0.3, 0.42, "DCE6F5");
    txt(s, y[2], x + 0.15, 2.98, w - 0.3, 0.42, { size: 13, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    txt(s, "主な役割", x + 0.1, 3.5, w - 0.2, 0.3, { size: 10.5, bold: true, color: NAVY, margin: 0 });
    txt(s, y[3], x + 0.1, 3.8, w - 0.2, 0.75, { size: 9, bold: true, color: DARK, ls: 1.15 });
    txt(s, "現場体制", x + 0.1, 4.55, w - 0.2, 0.3, { size: 10.5, bold: true, color: NAVY, margin: 0 });
    txt(s, y[4], x + 0.1, 4.85, w - 0.2, 0.75, { size: 9, bold: true, color: DARK, ls: 1.15 });
    rect(s, x + 0.1, 5.55, w - 0.2, 0.38, "E3ECF8");
    txt(s, y[5], x + 0.1, 5.55, w - 0.2, 0.38, { size: 9.5, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
  });
  txt(s, "役割の進化", 0.4, 6.1, 1.5, 0.3, { size: 11, bold: true, color: DARK, margin: 0 });
  steps(s, [["営業＋現場作業", ""], ["営業＋現場管理", ""], ["営業・案件管理＋現場監督", ""]], 6.45, 0.48, { num: false, hh: 0.48, hsize: 12, colors: ["7FA3D6", "4F7FC0", NAVY] });
}
{
  const s = content(S7, "現場について");
  subhead(s, "遺品整理の絶対遵守3大ルール");
  txt(s, "故人様とご遺族のお気持ちを第一に考え、後悔のない遺品整理を行うために、必ず守るルールです。", 0.4, 1.5, 9.2, 0.4, { size: 12.5, bold: true, color: DARK, margin: 0 });
  const R3 = [["捨てる前に必ず確認", "現金・通帳・カード・印鑑・権利書・写真・アルバム・手紙は必ず確認し、誤って処分しない。"], ["迷ったら捨てない\n→確認保留ケースへ", "判断に迷うものはその場で処分せず「確認保留ケース」で保管。後からご遺族様と一緒に確認する。"], ["全品手作業確認", "ポケット・封筒・引き出し・箱の中まで、すべての品を一つひとつ手作業で確認する。"]];
  R3.forEach((r, i) => {
    const x = 0.4 + i * 3.13, w = 2.94;
    rect(s, x, 2.05, w, 3.0, WHITE, { line: RED, lw: 1.5 });
    rect(s, x, 2.05, w, 0.85, RED);
    txt(s, `ルール${i + 1}\n` + r[0].split("\n")[0], x, 2.05, w, 0.85, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, r[1], x + 0.15, 3.0, w - 0.3, 1.95, { size: 11, bold: true, color: DARK, ls: 1.3 });
  });
  table(s, [["作業前", "作業中", "引き渡し", "完了報告"], ["近隣挨拶・養生・作業前の写真撮影", "時計回りの仕分け・貴重品の記録と保管", "全室をご一緒に確認・貴重品をお渡し", "写真付きの作業完了報告書にご署名"]], 0.4, 5.25, 9.2, [2.3, 2.3, 2.3, 2.3], { size: 10.5, rowH: 0.5 });
  redMsg(s, "3大ルールの徹底が、ご遺族の安心と信頼につながります。", { y: 6.4 });
}

// =====================================================================
// 成功事例
// =====================================================================
const S8C = "成功事例";
pres.addSection({ title: S8C });
divider(S8C, "成功事例企業が続出！", "事例企業名・数値は掲載許可を確認してから使う。");
function caseSlide(name, area, r1, r2, right) {
  const s = content(S8C, "成功事例紹介", "【要確認】掲載許可。");
  tag(s, "HP画像");
  subhead(s, `${name}（${area}）`, { w: 6 });
  photo(s, 0.4, 1.6, 7.5, 2.9, "HPのファーストビュー画像");
  txt(s, [{ text: "立上げ2年目", options: { color: RED, fontSize: 24 } }, { text: "で\n遺品整理事業\n", options: { color: DARK, fontSize: 16 } }, { text: r1, options: { color: RED, fontSize: 20 } }, { text: "\nを実現！", options: { color: DARK, fontSize: 16 } }], 0.4, 4.7, 5.0, 2.2, { bold: true, valign: "middle", ls: 1.05 });
  s.addShape(pres.shapes.LINE, { x: 5.5, y: 4.8, w: 0, h: 1.9, line: { color: DARK, width: 2.5 }, objectName: nm("vline") });
  txt(s, right, 5.7, 4.7, 3.9, 2.2, { size: 16, color: DARK, valign: "middle", ls: 1.2 });
  return s;
}
caseSlide("株式会社 蒼炎様", "東京都多摩市", "売上8000万円・営業利益：3000万円", "", "参入2年で初期投資を回収し、\n“不用品買取”\n“空き家買取”“買取”など\n事業の裾野を拡大しながら\n成長中の会社。\n販促費用50万円/月\n反響数：40件/月");
caseSlide("Mevius Moving Works様", "大阪府大阪市", "売上6000万円・営業利益率30%超", "", "不動産会社からの紹介を中心\nに集客を行い、成約率80%超\nを実現！引越し事業と連携を\n図り、遺品整理事業を強化！");

// =====================================================================
// 船井総研がご提供できること
// =====================================================================
const S9 = "船井総研がご提供できること";
pres.addSection({ title: S9 });
{
  const s = pres.addSlide({ masterName: "BLANK", sectionTitle: S9 });
  rect(s, 0, 0.3, 10, 0.08, NAVY);
  photo(s, 0, 0.38, 10, 3.3, "オフィス写真");
  rect(s, 0, 3.68, 10, 0.08, NAVY);
  txt(s, "船井総研がご提供できること", 0.9, 4.6, 7, 0.8, { size: 30, color: DARK, valign: "middle" });
  for (let r = 0; r < 4; r++) for (let i = 0; i < 12; i++) {
    const d = 0.07 - i * 0.004;
    s.addShape(pres.shapes.OVAL, { x: 7.75 + i * 0.13, y: 4.78 + r * 0.13, w: d, h: d, fill: { color: NAVY, transparency: i * 6 }, line: { type: "none" }, objectName: nm("dot") });
  }
  s.addShape(pres.shapes.OVAL, { x: 7.5, y: 6.2, w: 6, h: 3, fill: { color: NAVY }, line: { type: "none" }, objectName: "wave9" });
  s.addNotes("事後経営相談の最後に使う。");
}
{
  const s = content(S9, "船井総研がご提供できること");
  banner(s, "貴社の課題に合わせて、3つのサービスでトータルにご支援します。", 1.15, { size: 15 });
  const SV = [["01", "遺品整理ホームページ\nディレクション支援", "WEBSITE SUPPORT", "集客・成約につながるホームページの企画・制作を、専門チームがトータルでサポートします。", NAVY], ["02", "遺品整理ツール＆\nマニュアル導入支援", "TOOL & MANUAL", "業務マニュアル・各種チェックリスト・帳票など、現場でそのまま使えるツールの導入を支援します。", "1F4E3D"], ["03", "遺品整理事業立上げ\n月次コンサルティング", "CONSULTING", "事業の立上げから軌道化・成長まで、経験豊富なコンサルタントが月次で継続的にご支援します。", "8A3B12"]];
  SV.forEach((v, i) => {
    const x = 0.4 + i * 3.13, w = 2.94;
    rect(s, x, 1.85, w, 1.05, v[4]);
    txt(s, v[0], x + 0.08, 1.85, 0.75, 1.05, { size: 26, color: WHITE, valign: "middle", margin: 0 });
    txt(s, v[1], x + 0.8, 1.88, w - 0.85, 0.8, { size: 12, bold: true, color: WHITE, valign: "middle", margin: 0 });
    txt(s, v[2], x + 0.8, 2.62, w - 0.9, 0.25, { size: 7, color: WHITE, align: "right", margin: 0 });
    photo(s, x, 2.9, w, 2.4, "イメージ写真");
    rect(s, x, 5.3, w, 1.0, "EEF1F6");
    txt(s, v[3], x + 0.1, 5.3, w - 0.2, 1.0, { size: 10, bold: true, color: DARK, align: "center", valign: "middle" });
  });
  band(s, "ともに、地域に必要とされる遺品整理事業を。", 6.45, 0.55, { size: 15 });
}
{
  const s = content(S9, "船井総研がご提供できること");
  rect(s, 0.4, 1.0, 6.2, 0.6, NAVY);
  txt(s, "01　遺品整理ホームページディレクション支援", 0.55, 1.0, 6.0, 0.6, { size: 15, bold: true, color: WHITE, valign: "middle", margin: 0 });
  txt(s, "成果につながるホームページを、プロの視点でトータルサポート\n商圏の競合調査からコンセプト設計、キャッチコピー提案まで、貴社の強みを最大限に活かしたホームページづくりをサポートします。", 0.4, 1.7, 4.8, 1.0, { size: 10, bold: true, color: DARK, ls: 1.2 });
  const H = [["競合状況の調査", "・商圏内の遺品整理事業者\n・商圏内の○○業者の動き"], ["強みの整理と打出し検討", "貴社の強み・差別化ポイントを整理し、効果的な打ち出し方を検討します。"], ["コンセプト設計のご提案", "ターゲット・訴求軸・メッセージを設計し、ホームページの方向性をご提案します。"], ["キャッチコピーのご提案", "お客様の心に響くキャッチコピーを複数パターンご提案します。"]];
  H.forEach((h, i) => {
    const y = 2.8 + i * 0.82;
    s.addShape(pres.shapes.OVAL, { x: 0.4, y: y + 0.17, w: 0.4, h: 0.4, fill: { color: NAVY }, line: { type: "none" }, objectName: nm("num") });
    txt(s, String(i + 1), 0.4, y + 0.17, 0.4, 0.4, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    rect(s, 0.95, y, 4.25, 0.74, "F2F4F8", { round: true });
    txt(s, h[0], 1.05, y + 0.02, 4.1, 0.3, { size: 12, bold: true, color: DARK, margin: 0 });
    txt(s, h[1], 1.05, y + 0.32, 4.1, 0.42, { size: 8.5, color: DARK, margin: 0 });
  });
  rect(s, 0.95, 6.15, 4.25, 0.7, NAVY);
  txt(s, "デザイン制作のみ\nホームページ制作会社様へ依頼", 0.95, 6.15, 4.25, 0.7, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  rect(s, 5.4, 1.75, 4.2, 5.1, "F2F4F8");
  txt(s, "ご提案イメージ", 5.4, 1.8, 4.2, 0.4, { size: 13, bold: true, color: NAVY, align: "center", margin: 0 });
  photo(s, 5.55, 2.25, 3.9, 4.45, "競合調査・コンセプト提案書の事例");
}
{
  const s = content(S9, "船井総研がご提供できること");
  rect(s, 0.4, 1.0, 6.2, 0.6, "1F4E3D");
  txt(s, "02　遺品整理ツール＆マニュアル導入支援", 0.55, 1.0, 6.0, 0.6, { size: 15, bold: true, color: WHITE, valign: "middle", margin: 0 });
  const MN = [["事業概要マニュアル", "正しい知識と仕組みで、安定した事業運営を実現。", "事業の全体像から収益モデル、運営体制までを体系化"], ["営業マニュアル", "成果を出す営業の型が、すべてここに。", "信頼構築から提案・クロージングまで、営業の「型」を標準化"], ["遺品整理現場マニュアル", "現場で迷わず、誰でも高品質。", "現場の「判断・作業・安全・品質」を標準化（全60ページ以上）"], ["遺品整理集客マニュアル", "現場で迷わず、誰でも成果を出せる。", "反響から問い合わせ、成約までの集客の流れを標準化（全60ページ以上）"]];
  MN.forEach((m, i) => {
    const x = 0.4 + (i % 2) * 4.7, y = 1.75 + Math.floor(i / 2) * 2.6, w = 4.5;
    txt(s, m[1], x, y, w, 0.3, { size: 10, color: DARK, font: MINCHO, margin: 0 });
    txt(s, m[0], x, y + 0.3, w, 0.6, { size: 22, color: DARK, font: MINCHO, valign: "middle", margin: 0 });
    txt(s, m[2], x, y + 0.9, w, 0.35, { size: 9, color: DARK, margin: 0 });
    photo(s, x, y + 1.3, w, 1.15, "マニュアルのページ画像");
  });
}
{
  const s = content(S9, "船井総研がご提供できること");
  rect(s, 0.4, 1.0, 6.2, 0.6, "8A3B12");
  txt(s, "03　遺品整理事業立上げ月次コンサルティング", 0.55, 1.0, 6.0, 0.6, { size: 15, bold: true, color: WHITE, valign: "middle", margin: 0 });
  s.addShape(pres.shapes.OVAL, { x: 3.5, y: 2.3, w: 3.0, h: 3.0, fill: { color: WHITE }, line: { color: NAVY, width: 1 }, objectName: "sol_ring" });
  s.addShape(pres.shapes.OVAL, { x: 4.05, y: 2.85, w: 1.9, h: 1.9, fill: { color: "2F3A4F" }, line: { type: "none" }, objectName: "sol_center" });
  txt(s, "遺品整理\nソリューション", 4.05, 2.85, 1.9, 1.9, { size: 15, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  const D = [["①事業戦略策定", "・3ヵ年経営計画策定\n・販促計画策定\n・財務分析", 0.4, 1.8, BLUE], ["②商品設計提案", "遺品整理・不用品回収・買取\n商品設計", 0.4, 4.2, "F2B705"], ["③集客戦略策定", "・ホームページ制作ディレクション\n・チラシ制作代行\n・WEBマーケティング分析", 6.4, 1.8, BLUE], ["④営業戦略策定", "・顧客管理システム導入\n・アプローチブック制作", 6.4, 3.45, "F2B705"], ["⑤採用戦略策定", "・採用戦略策定\n・作業員アルバイト用求人作成", 5.6, 5.0, BLUE]];
  D.forEach(([h, b, x, y, c]) => {
    rect(s, x, y, 3.2, 0.6, c, { round: true, r: 0.3 });
    txt(s, h, x, y, 3.2, 0.6, { size: 17, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, b, x + 0.3, y + 0.65, 3.0, 0.9, { size: 10.5, color: DARK, ls: 1.2 });
  });
  redMsg(s, "事業計画から商品設計・集客・営業まで一気通貫でサポートします。", { y: 6.45, size: 18 });
}
{
  const s = content(S9, "船井総研がご提供できること");
  banner(s, "1日の打ち合わせの例", 0.98, { size: 26 });
  panel(s, 0.4, 1.7, 3.0, 4.6, "支援内容", NAVY);
  txt(s, "① 事業戦略立案\n　事業計画、数値計画、商圏分析\n② 現状数値分析\n　売上・粗利・営業利益分析\n③ 集客力強化\n　定番チラシ・HP集客打合せ、リスティング運用、HPレポート分析\n④ 営業力強化\n　プロセス導入、営業ツール・トーク開発、研修\n⑤ 組織・管理力強化\n　案件管理のクラウド化、リーダー研修\n⑥ 店舗力強化\n⑦ 採用強化", 0.5, 2.15, 2.85, 4.1, { size: 8.5, bold: true, color: DARK, ls: 1.1 });
  arrowR(s, 3.45, 3.8, 0.3, 0.35, "F2B705");
  panel(s, 3.8, 1.7, 3.0, 4.6, "4時間基本パッケージ", NAVY);
  txt(s, [{ text: "★13：00〜13：30\n", options: { color: RED } }, { text: "①事業全体数値確認\n・事業数値進捗確認／営業KPIの確認\n", options: { color: DARK } }, { text: "★13：30〜14：30\n", options: { color: RED } }, { text: "②営業\n・営業進捗確認／追客案件の確認／営業ツール修正\n", options: { color: DARK } }, { text: "★14：40〜16：30\n", options: { color: RED } }, { text: "③集客\n・WEB・HP改善提案／集客スケジュールの策定／販促強化策\n", options: { color: DARK } }, { text: "★16：30〜17：00\n", options: { color: RED } }, { text: "④採用\n・採用進捗確認・人員体制確認", options: { color: DARK } }], 3.9, 2.15, 2.85, 4.1, { size: 9, bold: true, ls: 1.1 });
  arrowR(s, 6.85, 3.8, 0.3, 0.35, "F2B705");
  panel(s, 7.2, 1.7, 2.4, 4.6, "実行管理サポート", NAVY);
  txt(s, "★LINEサポート\n★TELサポート\n★メールサポート\n★ZOHO Cliqサポート\n★ハングアウト\n★ツール共有\n\n・進捗確認\n・課題相談解決\n\n※TELは9：00〜20：00\n※日曜は出れない場合もございます", 7.3, 2.15, 2.25, 4.1, { size: 9.5, bold: true, color: DARK, ls: 1.2 });
  rect(s, 1.2, 6.4, 7.6, 0.6, NAVY);
  txt(s, [{ text: "30日間で ", options: { color: WHITE } }, { text: "実行を最大化させる", options: { color: "F6C700" } }], 1.2, 6.4, 7.6, 0.6, { size: 20, bold: true, align: "center", valign: "middle", margin: 0 });
}
{
  const s = content(S9, "新規立上げスケジュール", "開始月を実際の月名に差し替える。");
  tag(s, "開始月");
  const X0 = 2.2, MW = 1.22, Y0 = 1.45;
  rect(s, 0.4, Y0, 9.2, 0.75, "000000");
  for (let m = 0; m < 6; m++) txt(s, `${m + 1}ヶ月目`, X0 + m * MW, Y0, MW, 0.75, { size: 14, color: WHITE, align: "center", valign: "middle" });
  const G = [["業者探し", 0, 2, "協力業者探し・原価確認", "5B9BD5"], ["商品設計", 0, 2, "基本パック・オプション・商品設計", "70AD47"], ["ブランド\n設計", 0, 2, "屋号・ロゴ・キャッチコピー\n・ブランドコンセプト", "4472C4"], ["HP制作", 0.45, 2.3, "TOPページ公開→全ページ公開", "ED7D31"], ["営業体制\n構築", 1.65, 4.35, "営業ツール制作・営業事例紹介・営業研修", "FFC000"], ["集客施策", 1.65, 4.35, "チラシ制作・WEB広告配信・SEO・MEO・AIO反響分析", "44546A"], ["アライアンス\n提携", 1.1, 4.9, "リストアップ・提携資料作成→提携に向けた営業活動", "F0605D"]];
  G.forEach((g, i) => {
    const y = Y0 + 0.75 + i * 0.66;
    rect(s, 0.4, y, 9.2, 0.66, WHITE, { line: "000000", lw: 0.75 });
    txt(s, g[0], 0.5, y, 1.65, 0.66, { size: 13, color: DARK, align: "center", valign: "middle" });
    rect(s, X0 + g[1] * MW, y + 0.1, g[2] * MW, 0.46, g[4], { shadow: true });
    txt(s, g[3], X0 + g[1] * MW, y + 0.1, g[2] * MW, 0.46, { size: 9, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  });
}

// =====================================================================
// コンサルティング形式・費用について
// =====================================================================
const S10 = "コンサルティング形式・費用について";
pres.addSection({ title: S10 });
{
  const s = pres.addSlide({ masterName: "BLANK", sectionTitle: S10 });
  photo(s, 5.0, 0, 5.0, 7.5, "握手の写真（青のオーバーレイ）");
  rect(s, 5.0, 0, 5.0, 7.5, BLUE);
  txt(s, "コンサルティング\n形式・費用について", 0.4, 2.8, 4.4, 1.6, { size: 28, bold: true, color: "1C4FA0", align: "center", valign: "middle" });
}
{
  const s = content(S10, "船井総研がご提供できること", "【要確認】プラン2（ツールのみ）の価格と研究会の年会費。最新の企画書（96p）の価格を記載。");
  const PL = [["❶【HP制作ディレクションのみ】", [["遺品整理　商品設計＋ホームページ制作ディレクション", "700,000円（税抜）"]], "➢ 制作費用別途　ホームページ制作会社様へデザイン制作費用はお支払いが必要です。", "DCE3F0"], ["❷【マニュアル・オリジナルツールのみ】", [["遺品整理　マニュアル・オリジナルツール導入支援", "1,500,000円（税抜）"]], "", "E2EFDA"], ["❸【マニュアル・オリジナルツール＋HP制作ディレクション＋月次支援", [["遺品整理　HP制作ディレクション", "700,000円（税抜）"], ["遺品整理　マニュアル・オリジナルツール導入支援　※②より50万円お値引き", "1,000,000円（税抜）"], ["遺品整理　業績アップ月次支援コンサルティング", "200,000円（税抜）/月"]], "", "FBE5D6"]];
  let y = 1.05;
  PL.forEach((p) => {
    txt(s, p[0].replace("❶", "①").replace("❷", "②").replace("❸", "③").replace("＋月次支援", "＋月次支援】"), 0.4, y, 9.2, 0.42, { size: 13.5, bold: true, color: DARK, margin: 0 });
    y += 0.45;
    p[1].forEach((r) => {
      rect(s, 0.4, y, 6.2, 0.42, p[3], { line: "000000", lw: 1.5 });
      txt(s, r[0], 0.5, y, 6.05, 0.42, { size: 10.5, bold: true, color: DARK, valign: "middle", margin: 0 });
      rect(s, 6.6, y, 3.0, 0.42, p[3], { line: "000000", lw: 1.5 });
      txt(s, r[1], 6.65, y, 2.9, 0.42, { size: 13, bold: true, color: DARK, valign: "middle", align: "right", margin: 0 });
      y += 0.42;
    });
    if (p[2]) { txt(s, p[2], 0.4, y + 0.03, 9.2, 0.28, { size: 8.5, bold: true, color: RED, align: "right", margin: 0 }); y += 0.3; }
    y += 0.18;
  });
  txt(s, "➢ 月次支援はオンライン＋必要に応じてご訪問となります。\n➢ コンサルティングは月次支援と研究会の両方を含むサービスとなります。\n➢ 月次支援のご契約は「クラウドサイン」、研究会のご契約は「船井総研マイページからのお手続き」と別になりますのでご了承ください。\n➢ 月次支援、交通費のお支払いは、毎月のご支援終了後、請求書を送付させていただき、銀行振り込みとなります。\n➢ PPC運用代行サービスは運用費用に対して、20%の手数料が発生します。（月ごとの広告費用の実額に対して計算されますので、毎月の請求金額が変動いたします。）\n➢ コンサルティングフィーは更新時に、事前確認の上、変動いたしますのでご了承ください。", 0.4, y, 9.2, 7.0 - y, { size: 8.5, bold: true, color: RED, ls: 1.15 });
}
{
  const s = content(S10, "コンサルティング形式・費用について");
  rect(s, 0.4, 1.05, 9.2, 0.6, "F2F4F8");
  txt(s, "コンサルティングに伴う契約概要", 0.7, 1.05, 8, 0.6, { size: 20, bold: true, color: "1C4FA0", valign: "middle", margin: 0 });
  const items = [
    "月次支援は原則1年契約とします。（調査期間は除く）",
    "月次支援フィーとは別に交通費は毎月固定額をご請求させていただきます。",
    "月次支援2年目以降のコンサルフィーはコンサルタントの役職に応じて金額が変動する場合がございます。",
    "月次支援フィーのご入金が滞りますと、コンサルティングが中止となりますので、ご入金は必ずお守り下さい。",
    "月次支援のご契約を締結する場合は、研究会のご入会が必須となります。",
    "月次支援でご提案した施策に、別途費用がかかる場合がございます。（購入費、印刷費、サイト制作費等）",
    "月次支援のご契約は一定の成果を保証するものではありません。成果がないことを理由に返金は、一切行いませんのでご了承ください。",
    "当企画書で記載している数値シミュレーションや事例はあくまでも一例であり、諸条件などにより数値が大きく変わりこの通りになることを保証したものではないことをご了承ください。",
    "当企画書で記載している数値や提案内容は、税理士や会計士、弁護士などの法律で定められた資格者による判断を保証するものではありません。",
    "月次支援で他社をご紹介することがございますが、他社とのご契約に関しては、直接契約となり、弊社は一切の責任を負いませんので、ご了承ください。",
    "月次支援のお時間は内容によって変動がありますが、1回あたり約2〜4時間を目安としております。",
    "弊社でお手伝いさせていただいたWEBサイトを活用し、リスティング広告を実施する場合、リスティング広告費用の20%〜が運用手数料として発生いたします。",
    "月次支援で、顧客情報などの個人情報を含む資料には原則取り扱うことができませんので、ご了承ください。",
    "月次支援で、弊社が制作したツールは、弊社に著作権が発生します。",
    "弊社の営業時間は平日（9時30分〜18時）ですので、お打ち合わせ、連絡について事前にご了承をお願いいたします。",
  ];
  txt(s, items.map((t, i) => ({ text: "■ " + t, options: { breakLine: i < items.length - 1, color: i === 6 || i === 14 ? RED : DARK } })), 0.4, 1.8, 9.2, 5.2, { size: 8.5, bold: true, ls: 1.15 });
  note(s, "※最新の契約条件に合わせて更新してください。");
}

// ---------- 出力 ----------
(async () => {
  await pres.writeFile({ fileName: OUT });
  const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
  await applyTheme(OUT, THEME);
  console.log("written:", OUT);
})();
