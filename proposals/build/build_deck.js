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
    { placeholder: { options: { name: "body", type: "body", x: 1, y: 2.35, w: 8, h: 0.6, fontFace: FONT, fontSize: 18, color: WHITE, align: "center", valign: "middle" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 3.0, w: 8.8, h: 1.5, fontFace: FONT, fontSize: 40, bold: true, color: YEL, align: "center", valign: "middle" }, text: "" } },
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
function divider(section, label, title, notes) {
  const s = pres.addSlide({ masterName: "DIVIDER", sectionTitle: section });
  s.addText(label, { placeholder: "body" });
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

// =====================================================================
// 表紙
// =====================================================================
const S0 = "1. はじめに";
pres.addSection({ title: S0 });
{
  const s = pres.addSlide({ masterName: "BLANK", sectionTitle: S0 });
  // 波形の装飾（紺〜青の楕円を画面外にはみ出させて配置）
  s.addShape(pres.shapes.OVAL, { x: -2.5, y: 5.7, w: 16, h: 5, fill: { color: NAVY }, line: { type: "none" }, objectName: "cover_wave1" });
  s.addShape(pres.shapes.OVAL, { x: 1.5, y: 6.25, w: 14, h: 4, fill: { color: BLUE }, line: { type: "none" }, objectName: "cover_wave2" });
  s.addShape(pres.shapes.OVAL, { x: 6.2, y: -2.6, w: 7, h: 3.4, fill: { color: NAVY }, line: { type: "none" }, objectName: "cover_wave3" });
  s.addShape(pres.shapes.OVAL, { x: 7.4, y: -2.9, w: 6, h: 3.2, fill: { color: BLUE }, line: { type: "none" }, objectName: "cover_wave4" });
  s.addShape(pres.shapes.OVAL, { x: -0.5, y: -0.4, w: 1.6, h: 1.6, fill: { color: BLUE }, line: { type: "none" }, objectName: "cover_circle" });
  // ドット
  for (let i = 0; i < 12; i++) {
    const d = 0.13 - i * 0.008;
    s.addShape(pres.shapes.OVAL, { x: 0.75 + i * 0.2, y: 1.85 - d / 2, w: d, h: d, fill: { color: NAVY, transparency: i * 6 }, line: { type: "none" }, objectName: nm("dot") });
  }
  txt(s, "株式会社○○ 様", 0.7, 2.1, 8, 0.5, { size: 20, font: MINCHO, color: DARK, name: "cover_client" });
  txt(s, "遺品整理事業\n参入のご提案", 0.7, 2.65, 8.6, 1.7, { size: 38, font: MINCHO, color: DARK, ls: 1.05, name: "cover_title" });
  txt(s, "―“片付け”を入口に、地域の相続・空き家案件の最初の相談先になる―", 0.7, 4.35, 8.8, 0.45, { size: 14, font: MINCHO, color: NAVY, name: "cover_sub" });
  txt(s, "20XX年○月○日\n株式会社船井総合研究所　専門工事支援部\n担当：○○ ○○", 0.7, 4.95, 6, 0.95, { size: 12, color: DARK, ls: 1.2, name: "cover_meta" });
  s.addNotes("【要カスタマイズ】会社名・日付・担当者名を差し替える。\n【使い方】事前経営相談（1回目）は1〜4章、事後経営相談（2回目）は2章（振り返り）＋5〜8章を使う。");
}

// 1-2 目次と進め方
{
  const s = content(S0, "目次と本日の進め方", "1回目はここで「本日は1〜4章、5〜8章は次回、貴社に当てはめてご説明します」と伝え、売り込みの場ではないことを明確にする。2回目は「前回の振り返り→5章から」と伝える。");
  lead(s, "まずは事業の全体像をご理解いただき、次回、貴社に当てはめた進め方をお示しします", { size: 15 });
  const left = [
    ["1", "はじめに", "本資料の目的・自己紹介"],
    ["2", "貴社のお話を伺う", "現状・目標・課題"],
    ["3", "遺品整理の事業概要", "市場・事業モデル・シナジー"],
    ["4", "よくあるご質問", "疑問への簡単な回答"],
  ];
  const right = [
    ["5", "疑問の解消", "セミナーでの疑問にお答えします"],
    ["6", "貴社で進めるには", "貴社版の計画と進め方"],
    ["7", "成功事例", "進め方の根拠"],
    ["8", "船井総研のご支援とお見積り", "支援内容・費用・次のステップ"],
  ];
  rect(s, 0.4, 1.7, 4.45, 0.5, BLUE, { round: true });
  txt(s, "事前経営相談（本日）", 0.4, 1.7, 4.45, 0.5, { size: 14, bold: true, color: WHITE, align: "center", valign: "middle" });
  rect(s, 5.15, 1.7, 4.45, 0.5, NAVY, { round: true });
  txt(s, "事後経営相談（次回）", 5.15, 1.7, 4.45, 0.5, { size: 14, bold: true, color: WHITE, align: "center", valign: "middle" });
  [[left, 0.4, BLUE], [right, 5.15, NAVY]].forEach(([list, x, c]) => {
    list.forEach((r, i) => {
      const y = 2.35 + i * 0.95;
      rect(s, x, y, 4.45, 0.82, i % 2 ? WHITE : LBLUE, { line: LINE });
      circleNum(s, r[0], x + 0.12, y + 0.16, 0.5, c);
      txt(s, r[1], x + 0.75, y + 0.06, 3.6, 0.4, { size: 14, bold: true, color: NAVY, valign: "middle" });
      txt(s, r[2], x + 0.75, y + 0.44, 3.6, 0.32, { size: 11, color: GRAY, valign: "middle" });
    });
  });
  concl(s, "本日は「ご説明とヒアリングの場」です。ご提案は次回、貴社に合わせて行います");
}

// 1-3 会社・担当コンサルタント紹介
{
  const s = content(S0, "会社・担当コンサルタント紹介", "会社紹介は手短に（1分）。担当者の遺品整理での支援実績を中心に話す。");
  tag(s, "担当者の写真・経歴");
  lead(s, "株式会社船井総合研究所", { w: 6 });
  table(s, [
    ["項目", "内容"],
    ["創立", "1970年3月6日"],
    ["代表取締役社長", "真貝 大介"],
    ["連結従業員数", "約1,500名"],
    ["特徴", "経営コンサルティング業界で初めて株式公開。約1,000名のコンサルタントが全国10,000社以上を支援"],
    ["所在地", "大阪本社（イノゲート大阪）／東京本社（八重洲セントラルタワー）"],
  ], 0.4, 1.65, 4.6, [1.4, 3.2], { size: 10.5, align: "left", firstColFill: true });
  rect(s, 5.3, 1.45, 4.3, 4.75, LBLUE, { line: LINE });
  ph(s, 5.45, 1.6, 1.6, 2.0, "顔写真", "担当コンサルタントの写真を配置");
  txt(s, [
    { text: "専門工事支援部", options: { fontSize: 11, color: GRAY, breakLine: true } },
    { text: "○○ ○○", options: { fontSize: 18, bold: true, color: NAVY, breakLine: true } },
    { text: "遺品整理士 など保有資格", options: { fontSize: 11, color: GRAY } },
  ], 7.15, 1.65, 2.4, 1.6, { ls: 1.25 });
  ph(s, 5.45, 3.75, 4.0, 2.3, "経歴・支援実績・ミッション", "例）遺品整理・リフォーム業界を専門に、新規事業立上げから業績アップまでを支援。\nミッション：遺品整理を『片付け』から、安心できる住まいと暮らしの『再編集』へ");
  txt(s, "遺品整理事業の専門チーム", 0.4, 4.45, 4.6, 0.4, { size: 13, bold: true, color: NAVY });
  ph(s, 0.4, 4.85, 4.6, 1.35, "チームメンバー", "顔写真・氏名・資格（例：古民家鑑定士、遺品整理士）を並べる");
  concl(s, "遺品整理の立上げを専門に支援するチームが、貴社を担当します");
}

// =====================================================================
// 2章 貴社のお話を伺う
// =====================================================================
const S2 = "2. 貴社のお話を伺う";
pres.addSection({ title: S2 });
divider(S2, "第2章", "貴社のお話を伺う", "1回目：質問だけのヒアリングシートとして使う（答えは手元でメモ）。2回目：伺った内容を書き込んだ版に差し替え、「前回伺ったこと」として確認してから5章に入る。他社名・外注先名は資料に書かない。");

function hearing(no, title, leadText, qs, tagText, conclText) {
  const s = content(S2, title, "1回目は質問を順に聞く。2回目は各記入欄に伺った内容を書き込み、間違いがないか確認する。");
  tag(s, tagText);
  lead(s, leadText, { w: 6 });
  qs.forEach((q, i) => {
    const y = 1.65 + i * (4.65 / qs.length);
    const h = 4.65 / qs.length - 0.12;
    rect(s, 0.4, y, 4.1, h, LBLUE, { line: LINE });
    circleNum(s, i + 1, 0.5, y + 0.1, 0.4, NAVY);
    txt(s, q[0], 1.0, y + 0.05, 3.45, 0.5, { size: 12, bold: true, color: NAVY, valign: "middle" });
    txt(s, q[1], 1.0, y + 0.52, 3.45, h - 0.55, { size: 10.5, color: GRAY });
    ph(s, 4.65, y, 4.95, h, "記入欄", "");
  });
  concl(s, conclText);
  return s;
}
hearing(1, "貴社のお話を伺う①　ご相談のきっかけ", "今回、遺品整理事業に関心を持たれたきっかけ", [
  ["お申し込み・ダウンロードのきっかけ", "セミナーへのお申し込み／資料ダウンロードの一番の理由"],
  ["気になったフレーズ・数値", "例）空き家の仕入れ、粗利改善、葬儀後の収益化"],
  ["社内の検討段階", "情報収集／具体案の比較検討／早期立上げ"],
  ["セミナー（資料）で特に知りたいこと", "商品・集客・営業・現場・許可・収支など"],
], "ヒアリング内容", "貴社の関心に合わせて、本日のご説明の重点を決めます");
hearing(2, "貴社のお話を伺う②　貴社の現状と目標", "貴社の経営基盤と、これから目指す姿", [
  ["事業規模と売上構成", "年間売上、主力事業の比率（仲介・買取・工事・施行数など）"],
  ["人員・組織", "社員数、営業人数、新規事業を兼任できる方の有無"],
  ["許認可・設備", "宅建・古物商・一般廃棄物の許可、トラック、倉庫"],
  ["3年後・5年後の目標", "売上目標、新会社設立、事業承継など"],
], "ヒアリング内容", "貴社の強みと目標から、遺品整理事業の位置づけを考えます");
hearing(3, "貴社のお話を伺う③　課題と遺品整理への期待", "現状と目標のギャップと、遺品整理に期待すること", [
  ["本業の課題", "例）広告費の高騰、単価の下落、仕入れの不足、紹介依存"],
  ["片付け・残置物の相談", "月に何件あるか／今はどこに依頼しているか"],
  ["遺品整理事業の位置づけ", "単独の収益事業か、本業の入口（川上）か"],
], "ヒアリング内容", "課題と遺品整理事業をつなげて、次回のご提案をつくります");

// =====================================================================
// 3章 遺品整理の事業概要
// =====================================================================
const S3 = "3. 遺品整理の事業概要";
pres.addSection({ title: S3 });
divider(S3, "第3章", "遺品整理の事業概要", "セミナーの全体像を先に渡すパート。詳しい数字・手法は5〜6章（2回目）に回し、ここでは全体像の理解に絞る。");

// 3-2 遺品整理とは
{
  const s = content(S3, "遺品整理とは", "不用品回収との違い（想いの整理・貴重品の捜索・供養）を強調する。");
  lead(s, "遺品整理の流れ【3ステップ】");
  steps(s, [
    ["整理（仕分け）", "一つひとつ確認しながら「残す物」「貴重品」「処分する物」に仕分けます"],
    ["搬出", "仕分けが終わった荷物を、建物を傷つけないよう丁寧に運び出します"],
    ["トラックへの積込", "安全に積み込み、許可業者による適正な処理・リサイクルにつなげます"],
  ], 1.65, 1.65, { colors: [ORANGE, GREEN, NAVY], size: 12.5, hsize: 14 });
  txt(s, "不用品回収との違い", 0.4, 3.5, 9.2, 0.4, { size: 14, bold: true, color: NAVY });
  table(s, [
    ["", "不用品回収", "遺品整理（専門店）"],
    ["目的", "不要な物を早く安く回収する", "故人とご遺族の想いを整理する"],
    ["貴重品・思い出の品", "対応しないことが多い", "1点ずつ手で確認し、見つけてお渡しする"],
    ["供養・心のケア", "対応しない", "供養・お焚き上げ、ご遺族への寄り添い"],
    ["その後の相談", "回収で終わり", "相続・売却・解体・リフォームまで相談できる"],
  ], 0.4, 3.9, 9.2, [2.2, 3.0, 4.0], { size: 12, rowH: 0.48 });
  concl(s, "遺品整理は「片付け」ではなく、心に寄り添う専門サービスです");
}

// 3-3 ミッション
{
  const s = content(S3, "遺品整理ビジネスのミッション", "");
  rect(s, 0.4, 1.0, 9.2, 1.55, LGREEN, { line: "B9D8C4" });
  txt(s, "遺品整理は、地域の福祉と空き家問題を解決する事業です", 0.6, 1.1, 8.8, 0.65, { size: 21, bold: true, color: GREEN, font: MINCHO, valign: "middle" });
  txt(s, "私たちは、故人とご遺族の想いに寄り添い、地域の未来を支える遺品整理を行っています。", 0.6, 1.8, 8.8, 0.6, { size: 13, color: DARK, valign: "middle" });
  txt(s, "遺品整理をしたいお客様は、こんなお悩みを抱えています", 0.4, 2.75, 9.2, 0.4, { size: 14, bold: true, color: NAVY, align: "center" });
  const pains = [
    ["安心できる業者を探している", "故人の遺品を丁寧に扱い、きちんと処分してくれる会社に頼みたい"],
    ["実家に帰る時間がない", "遠方に住んでいて、何度も実家に帰って片付ける時間がない"],
    ["どの会社が安心か分からない", "ネットで出てきた会社が信頼できるのか判断できない"],
  ];
  pains.forEach((p, i) => card(s, 0.4 + i * 3.13, 3.25, 2.94, 1.75, p[0], p[1], { color: GREEN, num: i + 1, size: 12, hsize: 12 }));
  rect(s, 0.4, 5.2, 9.2, 1.05, GREEN, { round: true });
  txt(s, "私たちのミッション：地域密着で安心できる遺品整理業者を増やし、地域の福祉と未来を支えていきます", 0.6, 5.2, 8.8, 1.05, { size: 15, bold: true, color: WHITE, valign: "middle", align: "center" });
  concl(s, "お客様の不安を解消できる「地域の専門店」が求められています");
}

// 3-4 3つの追い風
{
  const s = content(S3, "遺品整理市場の成長性", "");
  lead(s, "3つの追い風が遺品整理市場を拡大しています");
  const winds = [
    ["市場拡大の波：多死社会", "・年間死亡者数は今後も増加\n・遺品整理・実家整理の需要が拡大\n・「空き家をどうするか」を考える人が増加", ORANGE],
    ["社会課題の波：空き家問題", "・全国の空き家は約900万戸\n・放置空き家も増加\n・活用・売却・解体の需要が拡大", BLUE],
    ["法改正の波：相続登記の義務化", "・相続後3年以内に登記が必要\n・放置が難しくなる\n・実家の活用・売却・整理を検討する人が増加", GREEN],
  ];
  winds.forEach((w, i) => card(s, 0.4 + i * 3.13, 1.7, 2.94, 2.55, w[0], w[1], { color: w[2], num: i + 1, size: 12, hsize: 12 }));
  arrowD(s, 4.6, 4.35, 0.8, 0.4, NAVY);
  rect(s, 0.4, 4.85, 9.2, 1.35, NAVY, { round: true });
  txt(s, "3つの波が重なり、「放置」から「活用・処分」へ動く人が急増", 0.6, 4.9, 8.8, 0.55, { size: 16, bold: true, color: YEL, align: "center", valign: "middle" });
  txt(s, "遺品整理　／　実家整理　／　リフォーム　／　解体　／　売却　／　賃貸活用", 0.6, 5.5, 8.8, 0.55, { size: 14, color: WHITE, align: "center", valign: "middle" });
  concl(s, "遺品整理は、相続・空き家の需要の「最初の入口」です");
  note(s, "※空き家数：総務省「住宅・土地統計調査」。最新の公表値に合わせて更新してください。");
}

// 3-5 業界の現状
{
  const s = content(S3, "遺品整理業界の現状", "");
  lead(s, "遺品整理業者が抱える課題と、見えてきた市場のチャンス");
  const issues = [
    ["不用品回収を中心とした事業者が多い", "遺品整理という名前でも、実際は不用品回収の片付けが中心。「想いの整理」まで対応できる事業者は限られている。"],
    ["グリーフケアや供養への対応に差がある", "ご遺族への寄り添い方や供養の提案など、心のケアまで提供できる事業者はまだ多くない。"],
    ["サービス品質にばらつきがある", "対応・料金・作業品質が業者によって大きく違い、お客様の満足度に影響している。"],
  ];
  issues.forEach((it, i) => {
    const y = 1.7 + i * 1.45;
    rect(s, 0.4, y, 9.2, 1.3, i % 2 ? WHITE : "F7F1EA", { line: "D9C7B2" });
    circleNum(s, i + 1, 0.6, y + 0.35, 0.6, BROWN);
    txt(s, it[0], 1.4, y + 0.1, 8.0, 0.45, { size: 15, bold: true, color: BROWN, valign: "middle" });
    txt(s, it[1], 1.4, y + 0.55, 8.0, 0.7, { size: 12, color: DARK });
  });
  concl(s, "「片付け・回収」ではなく、心に寄り添う専門店が求められています");
}

// 3-6 今が参入のタイミング
{
  const s = content(S3, "今が参入のタイミング", "ライフサイクルの考え方：成長期は一番店が決まっていない。安定期に入ると価格競争になり、一番店化は難しい。");
  lead(s, "商品のライフサイクルで見ると、遺品整理は「成長期」にあります");
  // 曲線の代わりに段階ごとの柱で表現
  const phases = [["導入期", 0.9, "D5DCE8"], ["成長期", 2.6, RED], ["成熟期", 3.4, "8FA3C4"], ["安定期", 3.0, "8FA3C4"], ["衰退期", 2.0, "B8C3D6"]];
  const bx = 0.6, bw = 1.2, base = 5.0;
  phases.forEach((p, i) => {
    const x = bx + i * 1.45;
    rect(s, x, base - p[1], bw, p[1], p[2]);
    txt(s, p[0], x, base + 0.05, bw, 0.35, { size: 12, bold: true, color: i === 1 ? RED : NAVY, align: "center" });
  });
  txt(s, "遺品整理\n（いまここ）", bx + 1.45, base - 2.6 - 0.75, bw, 0.7, { size: 12, bold: true, color: RED, align: "center", valign: "bottom" });
  txt(s, "リフォーム\n新築など", bx + 2 * 1.45, base - 3.3, bw, 0.7, { size: 11, bold: true, color: WHITE, align: "center" });
  card(s, 7.95, 1.7, 1.65, 1.6, "成長期", "地域一番店が不在。今なら一番店をとれる", { color: RED, size: 11 });
  card(s, 7.95, 3.45, 1.65, 1.6, "安定期", "大手参入で価格競争。一番店化は困難", { color: "8FA3C4", size: 11 });
  txt(s, "※商品のライフサイクル理論（船井流）に基づく考え方です。", 0.4, 5.5, 7.4, 0.35, { size: 10, color: GRAY });
  concl(s, "地域一番店をとれるのは、専門店がまだ少ない「今」です");
}

// 3-7 事業モデル
{
  const s = content(S3, "事業モデルの全体像", "");
  txt(s, [
    { text: "商圏人口", options: { fontSize: 18, bold: true, color: RED } },
    { text: "20万人", options: { fontSize: 30, bold: true, color: RED } },
    { text: "・平均客単価", options: { fontSize: 18, bold: true, color: RED } },
    { text: "30万円", options: { fontSize: 30, bold: true, color: RED } },
    { text: "で 売上", options: { fontSize: 18, bold: true, color: RED } },
    { text: "1億円", options: { fontSize: 30, bold: true, color: RED } },
    { text: "を目指すモデル", options: { fontSize: 18, bold: true, color: RED } },
  ], 0.3, 1.0, 9.4, 0.8, { align: "center", valign: "middle" });
  txt(s, "※商圏人口や会社の認知度・実績により年商は変動します。数字はモデル数値であり、確約するものではありません。", 0.4, 1.8, 9.2, 0.35, { size: 10, color: GRAY, align: "center" });
  // 左：円の図
  s.addShape(pres.shapes.OVAL, { x: 0.6, y: 2.45, w: 4.0, h: 3.5, fill: { color: LBLUE }, line: { color: LINE, width: 1 }, objectName: "model_circle" });
  rect(s, 1.55, 2.75, 2.1, 0.7, WHITE, { line: NAVY, round: true });
  txt(s, "遺品整理", 1.55, 2.75, 2.1, 0.7, { size: 22, bold: true, color: RED, align: "center", valign: "middle" });
  txt(s, "戸建の遺品整理が中心", 1.2, 3.6, 2.8, 0.4, { size: 12, bold: true, color: NAVY, align: "center" });
  rect(s, 0.75, 4.55, 1.7, 0.55, WHITE, { line: LINE, round: true });
  txt(s, "不用品回収", 0.75, 4.55, 1.7, 0.55, { size: 12, bold: true, color: NAVY, align: "center", valign: "middle" });
  rect(s, 2.75, 4.55, 1.7, 0.55, WHITE, { line: LINE, round: true });
  txt(s, "出張買取", 2.75, 4.55, 1.7, 0.55, { size: 12, bold: true, color: NAVY, align: "center", valign: "middle" });
  // 右：3つのポイント
  const pts = [
    ["ターゲット", "親の家を相続する40〜70代。遠方に住むご家族も多い"],
    ["主力商品", "単価の大きい「戸建の遺品整理」を中心に集客する"],
    ["バックエンド", "解体・リフォーム・不動産・供養へつなげ、1顧客あたりの利益を高める"],
  ];
  pts.forEach((p, i) => card(s, 5.0, 2.4 + i * 1.22, 4.6, 1.08, p[0], p[1], { color: NAVY, num: i + 1, hh: 0.42, size: 11.5, hsize: 12.5 }));
  concl(s, "初期投資が低く、社員1名からスタートできるビジネスモデルです");
}

// 3-8 数値計画（概要）
{
  const s = content(S3, "3年間の数値計画（概要）", "ここでは概要のみ。経費の内訳は5章（2回目）で説明する。");
  lead(s, "3年目に売上1億円・営業利益率20%超を目指します");
  const cols = [["1年目", "3,600万円", "月10件", "2.9%", "立上げ期"], ["2年目", "5,400万円", "月15件", "16.7%", "拡大期"], ["3年目", "1億円", "月28件", "23.1%", "組織化期"]];
  cols.forEach((c, i) => {
    const x = 0.4 + i * 2.35, y = 1.7;
    rect(s, x, y, 2.15, 3.3, i === 2 ? NAVY : LBLUE, { line: LINE });
    txt(s, c[0] + "｜" + c[4], x, y + 0.1, 2.15, 0.4, { size: 13, bold: true, color: i === 2 ? YEL : NAVY, align: "center" });
    txt(s, "売上", x, y + 0.6, 2.15, 0.3, { size: 11, color: i === 2 ? WHITE : GRAY, align: "center" });
    txt(s, c[1], x, y + 0.9, 2.15, 0.6, { size: 24, bold: true, color: i === 2 ? WHITE : RED, align: "center", valign: "middle" });
    txt(s, "契約 " + c[2], x, y + 1.55, 2.15, 0.35, { size: 12, color: i === 2 ? WHITE : DARK, align: "center" });
    txt(s, "営業利益率", x, y + 2.1, 2.15, 0.3, { size: 11, color: i === 2 ? WHITE : GRAY, align: "center" });
    txt(s, c[3], x, y + 2.4, 2.15, 0.6, { size: 24, bold: true, color: i === 2 ? YEL : NAVY, align: "center", valign: "middle" });
  });
  card(s, 7.55, 1.7, 2.05, 1.55, "初期投資", "HP制作費 約150万円\n古物商許可 5〜8万円", { color: NAVY, size: 11.5 });
  arrowD(s, 8.35, 3.32, 0.45, 0.3, NAVY);
  card(s, 7.55, 3.7, 2.05, 1.3, "投資回収", "2年目の営業利益で回収見込み", { color: RED, size: 11.5 });
  txt(s, "前提：商圏人口20万人／平均客単価30万円／契約率80%", 0.4, 5.15, 9.2, 0.35, { size: 11, color: GRAY });
  concl(s, "初期投資が低く、社員1名から高い利益率を目指せます");
  note(s, MODEL_NOTE);
}

// 3-9 商流
{
  const s = content(S3, "仕事の入口と出口（商流）", "");
  lead(s, "さまざまな入口から遺品整理につながり、整理後の出口サービスまでワンストップで提供");
  card(s, 0.4, 1.7, 2.6, 2.7, "入口（案件の発生源）", "①葬儀会社\n　葬儀後に片付けの相談\n②介護・老人ホーム\n　入居前の生前整理\n③不動産会社\n　売却前の家財撤去", { color: NAVY, size: 11.5 });
  arrowR(s, 3.1, 2.85, 0.45, 0.4, NAVY);
  s.addShape(pres.shapes.OVAL, { x: 3.65, y: 1.75, w: 2.7, h: 2.6, fill: { color: NAVY }, line: { type: "none" }, objectName: "center_ihin" });
  txt(s, "遺品整理\n（元請けで受注）", 3.65, 2.2, 2.7, 1.0, { size: 16, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, "仕分け・貴重品捜索\n買取・搬出", 3.65, 3.2, 2.7, 0.7, { size: 10.5, color: WHITE, align: "center", valign: "middle" });
  arrowR(s, 6.45, 2.85, 0.45, 0.4, GREEN);
  card(s, 7.0, 1.7, 2.6, 2.7, "出口（整理後のサービス）", "④買取・リユース\n⑤清掃・特殊清掃\n⑥解体工事\n⑦リフォーム\n⑧不動産売却・活用\n⑨一般廃棄物業者（処分）", { color: GREEN, size: 11.5 });
  card(s, 0.4, 4.6, 4.5, 1.65, "下請けの流れ（よくある形）", "解体会社から残置物撤去だけを請け負う。\n単価が低く、後工程の利益は他社へ。", { color: "8A8F99", size: 11.5 });
  card(s, 5.1, 4.6, 4.5, 1.65, "元請けの流れ（目指す形）", "お客様から遺品整理を直接受注し、解体・売却・リフォームまで一括でご提案。", { color: RED, size: 11.5 });
  concl(s, "元請けで受けると、後工程の大きな利益が自社に残ります");
}

// 3-10 7つの収益機会
{
  const s = content(S3, "遺品整理から広がる7つの収益機会", "");
  lead(s, "関連サービスを段階的にご提案し、お客様の課題をワンストップで解決します");
  const st = [["遺品整理（基本サービス）", "すべての入口となる基盤サービス", ORANGE], ["買取・リユース", "値引きで契約率UP、販売で利益確保", GOLD], ["供養・お焚き上げ", "仏壇・人形・遺品の供養", "7FA650"], ["特殊清掃・消臭・除菌", "孤独死・ゴミ屋敷などで高利益", GREEN], ["解体工事", "老朽化した建物の解体", BLUE], ["不動産売却・買取サポート", "仲介手数料や買取利益", NAVY], ["リフォーム・リノベーション", "活用のための改修", PURPLE]];
  st.forEach((t, i) => {
    const x = 0.4 + i * 0.32, y = 5.55 - i * 0.6, w = 3.6;
    rect(s, x, y, w, 0.52, t[2]);
    txt(s, String(i + 1), x, y, 0.45, 0.52, { size: 16, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, t[0], x + 0.45, y, w - 0.5, 0.52, { size: 12, bold: true, color: WHITE, valign: "middle" });
  });
  card(s, 6.0, 1.6, 3.6, 1.35, "遺品整理のみ", "売上 30万円／粗利 10万円（粗利率 約33%）", { color: "8A8F99", size: 12 });
  arrowD(s, 7.55, 3.0, 0.5, 0.35, RED);
  card(s, 6.0, 3.4, 3.6, 2.15, "関連サービスを提案した場合", "買取（値引き）▲5万円／特殊清掃 20万円／供養 5万円／解体 150万円／不動産売却サポート 90万円／リフォーム 200万円\n→ 売上 490万円／粗利 180万円", { color: RED, size: 11 });
  concl(s, "遺品整理は、関連サービスにつながる「入口商品」です");
  note(s, MODEL_NOTE);
}

// 3-11 許可がなくても始められる
{
  const s = content(S3, "許可がなくても始められる（三社間契約）", "一般廃棄物の許可がなくても、お客様と許可業者が直接契約する形にすれば始められる。詳しい前提条件は5章で説明する。");
  lead(s, "お客様・当社・一般廃棄物の許可業者の三者で役割を分けます");
  const boxes = [["当社（遺品整理）", 0.4, NAVY], ["お客様", 3.9, LBLUE], ["一般廃棄物\n許可業者", 7.4, GREEN]];
  boxes.forEach(([t, x, c]) => {
    rect(s, x, 1.8, 2.2, 1.0, c, { round: true });
    txt(s, t, x, 1.8, 2.2, 1.0, { size: 14, bold: true, color: c === LBLUE ? NAVY : WHITE, align: "center", valign: "middle" });
  });
  s.addShape(pres.shapes.LEFT_RIGHT_ARROW, { x: 2.7, y: 2.1, w: 1.1, h: 0.4, fill: { color: NAVY }, line: { type: "none" }, objectName: "arrow_lr1" });
  s.addShape(pres.shapes.LEFT_RIGHT_ARROW, { x: 6.2, y: 2.1, w: 1.1, h: 0.4, fill: { color: GREEN }, line: { type: "none" }, objectName: "arrow_lr2" });
  txt(s, "①サービス契約\n仕分け・搬出・買取", 1.8, 2.9, 2.9, 0.65, { size: 11.5, color: NAVY, align: "center", bold: true });
  txt(s, "②直接契約\n収集運搬・処分", 5.3, 2.9, 2.9, 0.65, { size: 11.5, color: GREEN, align: "center", bold: true });
  table(s, [
    ["", "許可なし（三社間契約）", "一般廃棄物の許可あり", "古物商許可あり"],
    ["始めやすさ", "すぐに始められる", "取得に時間がかかる", "約40日で取得可能"],
    ["利益", "標準", "処分を自社で行い利益率UP", "買取で売上・利益UP"],
    ["位置づけ", "スタート時の基本形", "将来の選択肢", "早めの取得がおすすめ"],
  ], 0.4, 3.85, 9.2, [1.6, 2.6, 2.5, 2.5], { size: 12, rowH: 0.55 });
  concl(s, "法令に沿った方法で、スモールスタートできます");
}

// 3-12 4つの柱
{
  const s = content(S3, "事業を動かす4つの柱", "それぞれの詳しい内容はセミナー・次回（6章）で扱うと伝える。");
  lead(s, "集客・営業・現場・提携の4つを仕組みにすることで、未経験からでも立ち上げられます");
  const P = [
    ["集客", "専門店の屋号とHPをつくり、WEB・チラシで反響をとる", "・専門HP（10ページ構成）\n・SEO／MEO／Web広告\n・戸建向けのチラシ", BLUE],
    ["営業", "最初の電話対応と、現地での即決で契約率80%を目指す", "・問い合わせ後5分以内の電話\n・30万円未満はオンライン完結\n・30万円以上は現地で即決", ORANGE],
    ["現場", "誰がやっても同じ品質になる、作業のルールをつくる", "・捨てる前に必ず確認\n・貴重品の徹底捜索\n・作業完了報告", GREEN],
    ["提携", "不動産・士業・葬儀社などから紹介が集まる仕組みをつくる", "・地域の案件ハブになる\n・提携先の信用を守る\n・紹介手数料の設計", PURPLE],
  ];
  P.forEach((p, i) => {
    const x = 0.4 + (i % 2) * 4.7, y = 1.7 + Math.floor(i / 2) * 2.35;
    card(s, x, y, 4.5, 2.2, p[0], null, { color: p[3], num: i + 1, hsize: 15 });
    txt(s, p[1], x + 0.12, y + 0.58, 4.25, 0.6, { size: 12, bold: true, color: p[3] });
    txt(s, p[2], x + 0.12, y + 1.2, 4.25, 0.95, { size: 11.5, color: DARK, ls: 1.2 });
  });
  concl(s, "4つの柱を仕組みにすれば、未経験でも立ち上げられます");
}

// 3-13 全業種共通のシナジー
{
  const s = content(S3, "遺品整理がもたらすシナジー（全業種共通）", "");
  lead(s, "どの業種にとっても、遺品整理には4つの共通する効果があります");
  const E = [
    ["商流の川上がとれる", "売却・解体・リフォームを決める前のお客様と、最初に出会える。相見積もりになりにくい"],
    ["既存のお客様・資産が活きる", "既存顧客、車両、職人、店舗、会員組織などをそのまま活用できる"],
    ["外注していた作業が利益に変わる", "残置物撤去や家財処分を外注している場合、その費用が自社の粗利になる"],
    ["地域の信頼が上がり、紹介が増える", "「困ったらあの会社」という存在になり、提携先・お客様からの紹介が増える"],
  ];
  E.forEach((e, i) => {
    const x = 0.4 + (i % 2) * 4.7, y = 1.7 + Math.floor(i / 2) * 1.75;
    card(s, x, y, 4.5, 1.6, e[0], e[1], { color: [NAVY, BLUE, GREEN, ORANGE][i], num: i + 1, size: 12, hsize: 13.5 });
  });
  rect(s, 0.4, 5.3, 9.2, 0.95, NAVY, { round: true });
  txt(s, "遺品整理 ＝ 本業の「入口」をつくる事業", 0.4, 5.3, 9.2, 0.95, { size: 18, bold: true, color: YEL, align: "center", valign: "middle" });
  concl(s, "遺品整理は、どの業種にとっても「本業の入口」になります");
}

// 3-14 業種別シナジー一覧
{
  const s = content(S3, "業種別シナジー一覧", "貴社の業種の行を指しながら説明する。詳しくは付録の業種別ページ（2回目に6章で使用）。");
  lead(s, "8つの業種で、遺品整理が本業にもたらす効果");
  table(s, [
    ["業種", "本業の課題", "遺品整理で得られること", "つながる本業"],
    ["不動産", "仕入れ競争・広告費の高騰、案件の長期化", "売却を考える前の所有者と出会える／残置物撤去の外注費が利益に", "売買仲介・買取再販・賃貸管理"],
    ["解体", "価格競争、紹介依存、職人の稼働の波", "解体前の案件を元請けでとれる／家財撤去の利益も自社に", "解体工事・土地活用"],
    ["住宅・リフォーム", "新築縮小、大型案件不足、OB客との接点不足", "相続した実家のリノベ相談の入口になる", "リフォーム・リノベーション"],
    ["葬儀", "家族葬化で単価下落、葬儀後の収益がない", "葬儀後の困りごとを自社で受け、会員満足度が上がる", "法要・供養・墓じまい・相続相談"],
    ["介護（ケアマネ・包括）", "片付けの相談先不足、業務負担の増加", "施設入居前の生前整理に対応でき、家族の信頼が上がる", "入居相談・地域連携"],
    ["便利業", "小口案件ばかりで利益率が低い、ブランド力が弱い", "高単価商品を持てる／専門店として信頼される", "不用品回収・ハウスクリーニング"],
    ["リサイクル", "良質な仕入れが不安定、一般客との接点が少ない", "良質なリユース品を直接仕入れられる", "資源回収・リユース販売"],
    ["買取", "高価値品の仕入れが不安定、買取できない物が多い", "家一軒分から仕入れられ「全部お願いできる」が強みに", "出張買取・店舗販売"],
  ], 0.3, 1.6, 9.4, [1.45, 2.55, 3.25, 2.15], { size: 9.5, align: "left", firstColFill: true });
  concl(s, "どの業種にも、遺品整理と本業をつなぐ接点があります");
}

// 3-15 章のまとめ
{
  const s = content(S3, "第3章のまとめ", "");
  lead(s, "遺品整理事業のポイント");
  const M = [["需要", "多死社会・空き家・相続登記の義務化で、需要は伸び続ける", ORANGE], ["競合", "「想いの整理」までできる専門店は少なく、今なら地域一番店をとれる", RED], ["収益", "社員1名・低投資から、3年で売上1億円・営業利益率20%超を目指せる", NAVY], ["本業への効果", "売却・解体・リフォームの前のお客様と出会える「本業の入口」になる", GREEN]];
  M.forEach((m, i) => {
    const y = 1.7 + i * 1.1;
    rect(s, 0.4, y, 2.2, 0.95, m[2]);
    txt(s, m[0], 0.4, y, 2.2, 0.95, { size: 17, bold: true, color: WHITE, align: "center", valign: "middle" });
    rect(s, 2.6, y, 7.0, 0.95, i % 2 ? WHITE : LBLUE, { line: LINE });
    txt(s, m[1], 2.8, y, 6.7, 0.95, { size: 14, color: DARK, valign: "middle" });
  });
  concl(s, "伸びる市場で、本業の入口をつくる事業です");
}

// =====================================================================
// 4章 よくあるご質問
// =====================================================================
const S4 = "4. よくあるご質問";
pres.addSection({ title: S4 });
divider(S4, "第4章", "よくあるご質問", "");
{
  const s = content(S4, "よくいただくご質問", "1回目は1〜2行の簡単な回答にとどめ、「詳しくは○章／セミナーで」と次につなぐ。");
  lead(s, "事前によくいただくご質問に、簡単にお答えします");
  table(s, [
    ["ご質問", "回答", "詳しくは"],
    ["未経験でもできますか？", "できます。集客・営業・現場をマニュアルとツールで仕組みにして始めます", "6章"],
    ["どのくらい儲かりますか？", "商圏20万人のモデルで、3年目に売上1億円・営業利益率20%超を目指します", "5章"],
    ["許可は必要ですか？", "一般廃棄物の許可がなくても、三社間契約で始められます。古物商許可はあると有利です", "5章"],
    ["人はどのくらい必要ですか？", "1年目は社員1名（兼任も可）＋派遣・アルバイトで始められます", "5章"],
    ["トラックや倉庫は必要ですか？", "最初から持つ必要はありません。処分は許可業者に現場で引き渡します", "5章"],
    ["何から始めればいいですか？", "協力業者探し・料金設計・屋号とHPづくりから始めます（約6ヶ月で立上げ）", "8章"],
  ], 0.4, 1.65, 9.2, [2.6, 5.6, 1.0], { size: 11, align: "left", rowH: 0.62 });
  concl(s, "疑問の多くは、仕組みと進め方で解消できます");
}
{
  const s = content(S4, "セミナーで注目していただきたいポイント", "セミナー経由：本日の質問がセミナーのどこで解消されるかを示す。DL経由：資料で気になった箇所の補足と、次回の説明内容に差し替える。");
  tag(s, "セミナーの章立て／DL資料");
  lead(s, "本日のご質問が、どこで解消されるかの対応表", { w: 6 });
  table(s, [
    ["本日いただいたご質問", "セミナー（資料）で注目するポイント", "該当パート"],
    ["（例）本当に儲かるのか", "3年間の数値計画と、本業込みの利益", "（記入）"],
    ["（例）集客はどうするのか", "専門HPとWEB集客、チラシ、提携", "（記入）"],
    ["（例）現場を回せるか", "人員体制と見積の仕組み", "（記入）"],
    ["（記入）", "（記入）", "（記入）"],
    ["（記入）", "（記入）", "（記入）"],
  ], 0.4, 1.65, 9.2, [3.0, 4.6, 1.6], { size: 11.5, align: "left", rowH: 0.65 });
  concl(s, "この視点でお聞きいただくと、セミナーの理解が深まります");
}
{
  const s = content(S4, "次回のご案内", "日程は2択で提示する（例：「来週○日か○日はいかがでしょうか」）。DL経由は「情報収集の一環としてご覧いただくだけでも」とハードルを下げる。");
  tag(s, "日程候補");
  lead(s, "次回は、本日伺った内容をもとに「貴社に当てはめた進め方」をお持ちします", { w: 6, size: 15 });
  const N = [["貴社商圏の市場", "人口・死亡者数・空き家数・競合から市場規模を試算"], ["貴社版の数値計画", "貴社の商圏と体制に合わせた3年間の収支"], ["貴社業種でのシナジー", "本業にどんな利益が生まれるかを試算"], ["成功事例", "進め方の根拠となる参入企業の事例"]];
  N.forEach((n, i) => card(s, 0.4 + (i % 2) * 4.7, 1.75 + Math.floor(i / 2) * 1.55, 4.5, 1.4, n[0], n[1], { color: NAVY, num: i + 1, size: 12, hsize: 14 }));
  ph(s, 0.4, 4.95, 9.2, 1.3, "次回日程（候補）", "第1候補：　　月　　日（　）　　時〜　／　第2候補：　　月　　日（　）　　時〜\n場所：貴社／オンライン");
  concl(s, "次回、貴社専用の事業化イメージをお持ちします");
}

// =====================================================================
// 5章 疑問の解消
// =====================================================================
const S5 = "5. 疑問の解消";
pres.addSection({ title: S5 });
divider(S5, "第5章", "疑問の解消", "2回目の冒頭は2章（書き込み版）で前回の内容を確認してから入る。相手の疑問に合わせてページを選ぶ。");
{
  const s = content(S5, "いただいたご質問の整理", "セミナーのアンケートや1回目の質問を、7つの阻害要因に分類して記入する。最大のネックを特定し、そのページに時間をかける。");
  tag(s, "いただいた質問");
  lead(s, "ご質問を7つの観点で整理し、順にお答えします", { w: 6 });
  const R = [["お金", "儲かるのか・投資は回収できるか", "5-2"], ["時期", "今やるべきか", "5-5"], ["競合", "後発でも勝てるか", "5-5"], ["プラン", "現場や見積を回せるか", "5-3"], ["キーマン", "誰が進めるのか", "5-3"], ["会社", "法令・本業のブランドは大丈夫か", "5-4"], ["自分", "支援は信頼できるか", "7〜8章"]];
  table(s, [["観点", "よくある不安", "回答ページ", "貴社からいただいたご質問"], ...R.map(r => [r[0], r[1], r[2], "（記入）"])], 0.4, 1.65, 9.2, [1.2, 3.0, 1.2, 3.8], { size: 11.5, align: "left", rowH: 0.6, firstColFill: true });
  concl(s, "一番大きな不安から、根拠をもって解消します");
}
// 5-2 お金
{
  const s = content(S5, "お金：本当に儲かるのか", "【要確認】1年目の営業利益は、元資料の記載（106万円）と内訳の合計（586万円）が合わない。差額480万円の項目を確定させてから使う。");
  lead(s, "3年間の数値計画（経費の内訳）");
  const Y = (v) => v;
  table(s, [
    ["項目", "1年目", "2年目", "3年目", "備考"],
    [{ text: "売上", options: { bold: true } }, { text: "3,600万円", options: { bold: true } }, { text: "5,400万円", options: { bold: true } }, { text: "1億円", options: { bold: true } }, "月10件／15件／28件"],
    ["処分費", "1,080万円", "1,620万円", "3,000万円", "売上の30%"],
    ["その他消耗品費", "100万円", "150万円", "200万円", "資材など"],
    [{ text: "粗利", options: { bold: true } }, { text: "2,420万円", options: { bold: true } }, { text: "3,630万円", options: { bold: true } }, { text: "6,800万円", options: { bold: true } }, "粗利率 67〜68%"],
    ["アルバイト事務", "120万円", "120万円", "120万円", "事務1名"],
    ["広告宣伝費", "484万円", "726万円", "1,360万円", "粗利の20%"],
    ["その他事業経費", "150万円", "200万円", "250万円", ""],
    ["人件費（専任営業）", "600万円", "1,200万円", "1,800万円", "1名／2名／3名"],
    ["作業員アルバイト", "480万円", "480万円", "960万円", "2名／2名／4名"],
    [{ text: "立上げ期の追加費用 ※要確認", options: { color: RED } }, { text: "480万円", options: { color: RED } }, "―", "―", { text: "元資料との差額", options: { color: RED } }],
    [{ text: "営業利益", options: { bold: true, color: RED } }, { text: "106万円（2.9%）", options: { bold: true, color: RED } }, { text: "904万円（16.7%）", options: { bold: true, color: RED } }, { text: "2,310万円（23.1%）", options: { bold: true, color: RED } }, ""],
  ], 0.4, 1.55, 6.6, [1.9, 1.25, 1.25, 1.25, 0.95], { size: 9.5, rowH: 0.36 });
  card(s, 7.2, 1.55, 2.4, 1.7, "初期投資", "HP制作費 約150万円\n古物商許可 5〜8万円", { color: NAVY, size: 11.5 });
  arrowD(s, 8.15, 3.32, 0.5, 0.32, NAVY);
  card(s, 7.2, 3.72, 2.4, 1.55, "投資回収", "2年目の営業利益 904万円で回収見込み", { color: RED, size: 11.5 });
  concl(s, "社員1名から、3年目に営業利益率20%超を目指せます");
  note(s, MODEL_NOTE);
}
// 5-3 プラン・キーマン
{
  const s = content(S5, "プラン・キーマン：現場と人は回るのか", "");
  lead(s, "トラック・倉庫・専門の目利きがなくても始められる3つの仕組み");
  card(s, 0.4, 1.7, 3.0, 4.5, "人員体制", null, { color: NAVY, num: 1 });
  table(s, [["", "正社員", "役割"], ["1年目", "1名", "営業＋現場作業"], ["2年目", "2名", "営業＋現場管理"], ["3年目", "3名", "案件管理＋現場監督"]], 0.5, 2.3, 2.8, [0.75, 0.65, 1.4], { size: 10, rowH: 0.42 });
  txt(s, "作業は派遣・アルバイト中心。正社員は営業と管理に集中します", 0.5, 4.15, 2.8, 1.0, { size: 11.5, color: DARK });
  card(s, 3.5, 1.7, 3.0, 4.5, "見積の仕組み", null, { color: ORANGE, num: 2 });
  txt(s, [
    { text: "物量を段ボールの箱数に換算", options: { bold: true, color: ORANGE, breakLine: true } },
    { text: "140サイズ10箱 ＝ 1立米", options: { breakLine: true } },
    { text: "立米単価 1.2〜1.5万円", options: { breakLine: true } },
    { text: "2tトラック ＝ 10〜12立米", options: { breakLine: true } },
    { text: "作業量 1人1日1.2立米", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "見積 ＝ 物量×単価＋オプション−買取額", options: { bold: true, color: NAVY } },
  ], 3.6, 2.3, 2.8, 3.8, { size: 11.5, ls: 1.25 });
  card(s, 6.6, 1.7, 3.0, 4.5, "処分費を下げる仕組み", null, { color: GREEN, num: 3 });
  const route = ["国内で買取・販売", "海外へ輸出リユース", "まとめ売り（福袋）", "自社で再利用", "最後に廃棄"];
  route.forEach((r, i) => {
    rect(s, 6.75, 2.3 + i * 0.72, 2.7, 0.5, i === 4 ? "8A8F99" : GREEN, { round: true });
    txt(s, r, 6.75, 2.3 + i * 0.72, 2.7, 0.5, { size: 11.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    if (i < 4) arrowD(s, 7.95, 2.8 + i * 0.72, 0.3, 0.2, GREEN);
  });
  concl(s, "社員1名＋外部スタッフで、未経験から始められます");
}
// 5-4 会社
{
  const s = content(S5, "会社：法令と本業のブランドは大丈夫か", "「100%合法」とは言わず、前提条件を守れば法令に沿って運営できると説明する。自治体ごとのルールは事前に確認する。");
  lead(s, "三社間契約の前提条件と、本業のブランドを守る運営体制");
  card(s, 0.4, 1.7, 4.5, 3.0, "三社間契約で守ること", "①お客様と許可業者が、収集運搬・処分の契約を直接結ぶ\n②当社は仕分け・搬出までを担い、廃棄物を運搬しない\n③廃棄物は現場で許可業者のパッカー車へ引き渡す\n④自治体のルールを事前に確認する", { color: NAVY, size: 12 });
  card(s, 5.1, 1.7, 4.5, 3.0, "本業のブランドを守る仕組み", "①請負契約書・買取同意書を必ず交わす\n②古物台帳で買取の記録を3年間保管\n③損害賠償保険に加入\n④作業前の写真撮影・貴重品返却シートでトラブルを防ぐ", { color: GREEN, size: 12 });
  rect(s, 0.4, 4.9, 9.2, 1.3, LBLUE, { line: LINE });
  txt(s, "コンプライアンスを守った運用マニュアル・契約書のひな形をご提供します（8章）", 0.6, 4.9, 8.8, 1.3, { size: 14, bold: true, color: NAVY, valign: "middle", align: "center" });
  concl(s, "法令に沿った運営体制を、最初から整えて始めます");
}
// 5-5 競合・時期
{
  const s = content(S5, "競合・時期：後発でも勝てるか", "");
  tag(s, "商圏の競合状況");
  lead(s, "専門店がいない今、専門店として打ち出せば後発でも勝てます", { w: 6, size: 15 });
  ph(s, 0.4, 1.7, 4.5, 3.0, "貴社商圏の競合状況", "例）商圏内の遺品整理業者 ○社（うち専門店 ○社）\n不用品回収・便利屋 ○社\nGoogleマップ・HPで調査した結果を記入");
  card(s, 5.1, 1.7, 4.5, 3.0, "勝つための3つの打ち手", "①「遺品整理専門店」の屋号とHPで打ち出す\n②供養・買取・不動産相談までワンストップで対応\n③不動産・士業・葬儀社との提携で、相見積もりのない紹介を得る", { color: RED, size: 12 });
  rect(s, 0.4, 4.9, 9.2, 1.3, NAVY, { round: true });
  txt(s, "準備（提携先の開拓・HP制作・営業研修）に約6ヶ月かかるため、\n広告単価が上がる前に、早めに着手することが重要です", 0.6, 4.9, 8.8, 1.3, { size: 14, bold: true, color: WHITE, valign: "middle", align: "center" });
  concl(s, "専門店がいない今が、参入の時期です");
}

// =====================================================================
// 6章 貴社で進めるには
// =====================================================================
const S6 = "6. 貴社で進めるには";
pres.addSection({ title: S6 });
divider(S6, "第6章", "貴社で進めるには", "2回目の中心。「貴社の場合」に置き換えて示す。");
{
  const s = content(S6, "貴社商圏の市場", "地域データは自治体の統計・住宅土地統計調査などから取得する。");
  tag(s, "商圏データ");
  lead(s, "貴社の商圏で、遺品整理の市場規模を試算します", { w: 6 });
  const K = [["商圏人口", "○○万人"], ["年間死亡者数", "○,○○○人"], ["65歳以上の単身世帯", "○,○○○世帯"], ["空き家数", "○,○○○戸"]];
  K.forEach((k, i) => {
    const x = 0.4 + i * 2.33;
    rect(s, x, 1.7, 2.15, 1.5, LBLUE, { line: LINE });
    txt(s, k[0], x, 1.78, 2.15, 0.4, { size: 12, bold: true, color: NAVY, align: "center" });
    txt(s, k[1], x, 2.25, 2.15, 0.7, { size: 22, bold: true, color: RED, align: "center", valign: "middle" });
  });
  ph(s, 0.4, 3.4, 4.5, 2.8, "商圏地図", "対象エリア（市区町村）と、競合業者の位置をプロットした地図を配置");
  card(s, 5.1, 3.4, 4.5, 2.8, "市場規模の試算", "参考：商圏人口20万人で市場規模 約8.4億円\n\n貴社商圏の試算：約○億円\n目標シェア：3年目 約○%", { color: NAVY, size: 13 });
  concl(s, "貴社商圏の遺品整理市場は、約○億円です");
}
{
  const s = content(S6, "貴社版の事業コンセプト", "");
  tag(s, "コンセプト・屋号案");
  lead(s, "売り先・売り物・売り方を絞り込み、貴社の強みを活かした専門店にします", { w: 6, size: 14 });
  const C3 = [["売り先", "誰に売るか", "例）親の家を相続する50〜60代、賃貸オーナー"], ["売り物", "何を売るか", "例）遺品整理の標準パック＋供養・買取"], ["売り方", "どう売るか", "例）現地確認・即決型、提携先からの紹介"]];
  C3.forEach((c, i) => {
    const x = 0.4 + i * 3.13;
    card(s, x, 1.7, 2.94, 0.95, c[0], c[1], { color: NAVY, num: i + 1, size: 11, hh: 0.45 });
    ph(s, x, 2.75, 2.94, 1.5, "貴社の場合", c[2]);
  });
  ph(s, 0.4, 4.45, 4.5, 1.8, "貴社の強み", "例）地域での施工実績・既存顧客数・資格・車両など");
  ph(s, 5.1, 4.45, 4.5, 1.8, "屋号・コンセプト案（2〜3案）", "例）「○○の遺品整理」— 想いをつなぐ、地域の片付け相談室");
  concl(s, "貴社の強み○○を活かした専門店として打ち出します");
}
{
  const s = content(S6, "貴社版の料金設計", "料金はパック料金で見せる（内訳をそのまま見せると不安・不信感につながる）。箱数・立米を併記し、見積の計算と一致させる。平均単価30万円になるのは、戸建2LDK以上が中心になるため。");
  tag(s, "料金・オプション");
  lead(s, "お客様には分かりやすいパック料金、社内は立米単価で原価を管理します", { w: 6, size: 14 });
  table(s, [
    ["間取り", "物量の目安", "作業人数", "作業時間", "料金（税込）"],
    ["1K", "段ボール20箱（約2立米）", "1〜2名", "1〜3時間", "30,000円〜"],
    ["1LDK", "段ボール50箱（約5立米）", "2〜3名", "3〜6時間", "80,000円〜"],
    ["2LDK", "段ボール100箱（約10立米）", "3〜4名", "5〜8時間", "150,000円〜"],
    ["3LDK", "段ボール150箱（約15立米）", "4〜6名", "6〜9時間", "200,000円〜"],
    ["4LDK以上・戸建", "段ボール200箱以上（約20立米〜）", "5名〜", "1〜2日", "300,000円〜"],
  ], 0.4, 1.6, 9.2, [1.6, 3.2, 1.3, 1.4, 1.7], { size: 11, rowH: 0.45 });
  txt(s, "仕分け・搬出・処分・簡易清掃・運搬まで含んだ安心のパック料金（追加料金なし）", 0.4, 4.4, 9.2, 0.35, { size: 12, bold: true, color: RED, align: "center" });
  const O = [["供養", "15,000円〜"], ["特殊清掃", "30,000円〜"], ["原状回復", "40,000円〜"], ["解体", "100,000円〜"], ["不動産売却", "ご相談"], ["空き家管理", "5,000円〜/月"], ["買取", "お値引き"]];
  O.forEach((o, i) => {
    const x = 0.4 + i * 1.32;
    rect(s, x, 4.85, 1.2, 1.3, LBLUE, { line: LINE });
    txt(s, o[0], x, 4.9, 1.2, 0.5, { size: 11, bold: true, color: NAVY, align: "center", valign: "middle" });
    txt(s, o[1], x, 5.45, 1.2, 0.5, { size: 10.5, color: DARK, align: "center", valign: "middle" });
  });
  concl(s, "分かりやすいパック料金で、相見積もりでも選ばれます");
}
{
  const s = content(S6, "貴社版の数値計画", "1年目は準備期間（約6ヶ月）を含めた月ごとの立ち上がりで示す。兼任スタート案と専任採用案の2つを並べる。");
  tag(s, "貴社の数値");
  lead(s, "貴社の商圏と体制に合わせた3年間の計画", { w: 6 });
  // 1年目の立ち上がりグラフ（ネイティブチャート）
  s.addChart(pres.charts.BAR, [{ name: "月間契約件数（目安）", labels: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"], values: [0, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10] }], {
    x: 0.4, y: 1.65, w: 4.6, h: 3.3, barDir: "col", chartColors: [NAVY], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 9, dataLabelColor: DARK,
    showTitle: true, title: "1年目の月間契約件数（例）", titleFontSize: 12, titleColor: NAVY, titleFontFace: "+mn-lt",
    catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, catAxisLabelColor: GRAY, valAxisLabelColor: GRAY, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
    valGridLine: { color: "E5E7EB", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false, objectName: "ramp_chart",
  });
  txt(s, "※月ごとの件数は例です。貴社の準備状況に合わせて調整します", 0.4, 4.95, 4.6, 0.3, { size: 9.5, color: GRAY });
  table(s, [
    ["", "兼任スタート案", "専任採用案"],
    ["1年目の体制", "既存営業1名が兼任", "専任営業1名を採用"],
    ["人件費", "本業と按分", "年600万円"],
    ["向いている会社", "片付け相談がすでにある", "早く立ち上げたい"],
    ["1年目売上", "○○万円", "○○万円"],
    ["3年目売上", "○○万円", "○○万円"],
  ], 5.2, 1.65, 4.4, [1.4, 1.5, 1.5], { size: 10.5, rowH: 0.5 });
  ph(s, 5.2, 4.75, 4.4, 1.4, "貴社の前提条件", "商圏人口・平均単価・契約率・広告予算など");
  concl(s, "準備期間を含めても、無理のない立ち上がりを計画します");
  note(s, MODEL_NOTE);
}
{
  const s = content(S6, "貴社の業種でのシナジー", "付録の業種別ページから、貴社の業種のページをこの位置に差し込む（複数業種にまたがる場合は複数枚）。このページは差し込み後に削除してよい。");
  tag(s, "業種別ページを差し込む");
  lead(s, "貴社の業種で、遺品整理が本業にもたらす効果", { w: 6 });
  const Q = [["経営課題", "貴社の本業の課題（ヒアリング②③より）"], ["遺品整理事業参入によるシナジー", "本業にもたらす効果を3〜4つ"], ["期待できる効果", "売上・利益・紹介・ブランドなど"], ["ビジネスモデルのイメージ", "遺品整理→本業へのつながりのフロー"]];
  Q.forEach((q, i) => ph(s, 0.4 + (i % 2) * 4.7, 1.7 + Math.floor(i / 2) * 2.15, 4.5, 2.0, q[0], q[1]));
  concl(s, "（付録の業種別ページを使用）");
}
{
  const s = content(S6, "本業込みの利益", "仲介手数料は宅建業法の上限（売買価格400万円超：3%＋6万円、税別）で計算。発生率は貴社の実績や業種に合わせて調整する。");
  tag(s, "発生率・単価");
  lead(s, "遺品整理の利益に、本業への波及効果を加えて考えます", { w: 6 });
  rect(s, 0.4, 1.65, 9.2, 0.8, NAVY, { round: true });
  txt(s, "本業込みの利益 ＝ 遺品整理の粗利 ＋ 不動産の発生率×仲介手数料 ＋ 工事の発生率×工事粗利", 0.5, 1.65, 9.0, 0.8, { size: 14, bold: true, color: WHITE, align: "center", valign: "middle" });
  table(s, [
    ["項目（10件受託した場合の例）", "計算", "粗利"],
    ["遺品整理", "30万円 × 10件 × 粗利率60%", "180万円"],
    ["不動産売買（10件中1件）", "売買価格2,000万円の仲介手数料（両手）\n（2,000万円×3%＋6万円）×2", "132万円"],
    ["解体工事（10件中1件）", "工事費150万円 × 粗利率30%", "45万円"],
    [{ text: "10件あたりの合計", options: { bold: true } }, "", { text: "357万円", options: { bold: true, color: RED } }],
  ], 0.4, 2.65, 6.2, [2.2, 2.8, 1.2], { size: 11, rowH: 0.55 });
  rect(s, 6.8, 2.65, 2.8, 3.4, LBLUE, { line: LINE });
  txt(s, "1件あたりの利益", 6.8, 2.8, 2.8, 0.4, { size: 13, bold: true, color: NAVY, align: "center" });
  txt(s, "遺品整理のみ", 6.8, 3.25, 2.8, 0.3, { size: 11, color: GRAY, align: "center" });
  txt(s, "18万円", 6.8, 3.55, 2.8, 0.55, { size: 22, bold: true, color: NAVY, align: "center", valign: "middle" });
  arrowD(s, 7.95, 4.15, 0.5, 0.3, RED);
  txt(s, "本業込み", 6.8, 4.5, 2.8, 0.3, { size: 11, color: GRAY, align: "center" });
  txt(s, "約36万円", 6.8, 4.8, 2.8, 0.6, { size: 26, bold: true, color: RED, align: "center", valign: "middle" });
  txt(s, "※税別", 6.8, 5.5, 2.8, 0.3, { size: 9.5, color: GRAY, align: "center" });
  concl(s, "本業まで含めると、遺品整理単体の約2倍の利益が残ります");
  note(s, MODEL_NOTE);
}
{
  const s = content(S6, "集客の進め方", "広告予算は、立上げ直後はテスト運用で反響の出るキーワード・エリアを見極め、効果の高い媒体に絞って拡大する。");
  tag(s, "広告予算");
  lead(s, "4つのチャネルで、安定して反響をとります", { w: 6 });
  table(s, [
    ["チャネル", "役割", "獲得単価の目安", "効果が出るまで"],
    ["WEB（専門HP・SEO・MEO・広告）", "今すぐ依頼したい人を集める", "約3万円／件", "広告はすぐ／SEO 3〜6ヶ月"],
    ["チラシ（戸建・B4・保存版）", "検索していない人に認知してもらう", "約5万円／件", "配布後すぐ"],
    ["提携先からの紹介", "相見積もりのない紹介を得る", "成約時のみ手数料", "3〜6ヶ月"],
    ["既存のお客様", "本業のお客様に案内する", "ほぼ0円", "すぐ"],
  ], 0.4, 1.65, 9.2, [2.8, 2.6, 1.7, 2.1], { size: 10.5, align: "left", rowH: 0.55 });
  steps(s, [["テスト期（立上げ〜3ヶ月）", "月20万円程度で反響の出るキーワード・エリアを検証"], ["拡大期（4ヶ月目以降）", "粗利の15〜20%を目安に、効果の高い媒体へ集中"], ["安定期", "紹介・SEOの比率を高め、広告費を下げていく"]], 4.65, 1.55, { size: 11, hsize: 11.5, colors: [BLUE, NAVY, GREEN] });
  concl(s, "広告ですぐ反響をとり、紹介とSEOで広告費を下げていきます");
}
{
  const s = content(S6, "営業の進め方", "");
  lead(s, "契約率80%を基準に、最初の電話対応と現地での即決で受注します");
  table(s, [
    ["", "1年目", "2年目", "3年目"],
    ["営業人員", "1名", "2名", "3名"],
    ["月間集客件数", "12.5件", "18.75件", "35件"],
    ["月間契約件数", "10件", "15件", "28件"],
    ["一人当たり契約件数（月）", "10件", "7.5件", "約9.3件"],
  ], 0.4, 1.65, 4.4, [1.9, 0.83, 0.83, 0.84], { size: 10.5, rowH: 0.46, firstColFill: true });
  txt(s, "※契約率80%を基準に算出", 0.4, 4.0, 4.4, 0.3, { size: 9.5, color: GRAY });
  card(s, 5.0, 1.65, 2.2, 2.6, "30万円未満", "オンライン完結\n電話→アンケート→写真→見積→電話で契約\n\n訪問を省いて安く・早く", { color: BLUE, size: 11 });
  card(s, 7.4, 1.65, 2.2, 2.6, "30万円以上", "現地確認＋即決\n電話→事前準備→現地確認→その場で見積→契約\n\n追加提案で単価UP", { color: NAVY, size: 11 });
  const K3 = [["スピード", "問い合わせ後5分以内に電話"], ["ヒアリング", "お金・時期・競合・プラン・キーマンを確認"], ["信頼と提案", "ご遺族に寄り添い、関連サービスも提案"]];
  K3.forEach((k, i) => card(s, 0.4 + i * 3.13, 4.5, 2.94, 1.75, k[0], k[1], { color: ORANGE, num: i + 1, size: 12, hsize: 13 }));
  concl(s, "最初の電話対応が、契約率を決めます");
}
{
  const s = content(S6, "現場品質の仕組み", "");
  lead(s, "誰が作業しても「頼んでよかった」と言われる品質を、ルールで標準化します");
  const R3 = [["捨てる前に必ず確認", "現金・通帳・印鑑・権利書・写真・手紙は必ず確認"], ["迷ったら捨てない", "判断に迷う物は「確認保留ケース」へ。ご遺族と一緒に確認"], ["全品手作業で確認", "ポケット・封筒・引き出し・箱の中まで一つずつ確認"]];
  R3.forEach((r, i) => card(s, 0.4 + i * 3.13, 1.7, 2.94, 1.75, r[0], r[1], { color: RED, num: i + 1, size: 12, hsize: 12.5 }));
  steps(s, [["作業前", "近隣挨拶・養生・作業前の写真撮影"], ["作業中", "時計回りの仕分け・貴重品の記録と保管"], ["引き渡し", "全室をご一緒に確認・貴重品をお渡し"], ["完了報告", "写真付きの報告書にサインをいただく"]], 3.7, 1.55, { size: 11, hsize: 12, colors: [NAVY, NAVY, NAVY, NAVY], arrow: GOLD });
  rect(s, 0.4, 5.4, 9.2, 0.85, LGREEN, { line: "B9D8C4" });
  txt(s, "現場マニュアル（60ページ以上）とチェックリストで、未経験のスタッフでも同じ品質に", 0.6, 5.4, 8.8, 0.85, { size: 13, bold: true, color: GREEN, align: "center", valign: "middle" });
  concl(s, "品質の高さが紹介を生み、広告費を下げます");
}

// =====================================================================
// 7章 成功事例
// =====================================================================
const S7 = "7. 成功事例";
pres.addSection({ title: S7 });
divider(S7, "第7章", "成功事例：進め方の根拠", "6章の進め方で、実際に成果が出ていることを示す。事例企業名・数値は掲載許可を確認してから使う。");
function caseSlide(name, area, result, detail, tie, notes) {
  const s = content(S7, "成功事例紹介", notes);
  tag(s, "HP画像・掲載許可");
  lead(s, `${name}（${area}）`, { w: 6 });
  ph(s, 0.4, 1.65, 9.2, 2.3, "HPのスクリーンショット", "事例企業のHPファーストビュー画像を配置");
  txt(s, result, 0.4, 4.15, 4.6, 1.9, { size: 20, bold: true, color: RED, valign: "middle" });
  s.addShape(pres.shapes.LINE, { x: 5.1, y: 4.2, w: 0, h: 1.8, line: { color: DARK, width: 2 }, objectName: nm("vline") });
  txt(s, detail, 5.3, 4.15, 4.3, 1.1, { size: 13, color: DARK, valign: "middle" });
  rect(s, 5.3, 5.3, 4.3, 0.85, LBLUE, { line: LINE });
  txt(s, tie, 5.4, 5.3, 4.1, 0.85, { size: 11, bold: true, color: NAVY, valign: "middle" });
  return s;
}
{
  const s = caseSlide("株式会社 蒼炎様", "東京都多摩市", "立上げ2年目で\n遺品整理事業\n売上8,000万円・\n営業利益3,000万円を実現！", "参入2年で初期投資を回収し、不用品買取・空き家買取など事業の裾野を拡大。販促費 月50万円／反響 月40件。", "6章との対応：専門店の屋号・HP（多摩エリア特化）とWEB集客", "【要確認】掲載許可。");
  concl(s, "地域特化の専門店として打ち出し、2年で投資を回収");
}
{
  const s = caseSlide("Mevius Moving Works様", "大阪府大阪市", "立上げ2年目で\n遺品整理事業\n売上6,000万円・\n営業利益率30%超を実現！", "不動産会社からの紹介を中心に集客し、成約率80%超。引越し事業と連携して遺品整理事業を強化。", "6章との対応：提携先からの紹介と、本業（引越し）とのシナジー", "【要確認】掲載許可。");
  concl(s, "提携からの紹介と本業との連携で、高い利益率を実現");
}
{
  const s = content(S7, "成功企業の共通点", "同業種の事例がある場合は、このページの前に追加する。");
  lead(s, "成功している企業は、共通して4つのことを実践しています");
  table(s, [
    ["成功の共通点", "蒼炎様", "Mevius様", "貴社の進め方（6章）"],
    ["専門店として打ち出す", "多摩エリア特化の専門店", "遺品整理の専門ブランド", "貴社版の事業コンセプト"],
    ["紹介の仕組み", "事業の裾野を広げ紹介を獲得", "不動産会社からの紹介が中心", "提携先からの紹介"],
    ["スピード対応", "反響月40件に素早く対応", "成約率80%超", "5分以内の電話・現地即決"],
    ["本業との連携", "買取・空き家買取へ拡大", "引越し事業と連携", "本業込みの利益"],
  ], 0.4, 1.7, 9.2, [2.2, 2.3, 2.3, 2.4], { size: 11.5, align: "left", rowH: 0.75, firstColFill: true });
  concl(s, "成功の型は、貴社でも再現できます");
}

// =====================================================================
// 8章 船井総研のご支援とお見積り
// =====================================================================
const S8 = "8. 船井総研のご支援とお見積り";
pres.addSection({ title: S8 });
{
  const s = pres.addSlide({ masterName: "BLANK", sectionTitle: S8 });
  rect(s, 5.4, 0, 4.6, 7.5, BLUE);
  rect(s, 5.4, 0, 4.6, 7.5, NAVY, {});
  s.addShape(pres.shapes.OVAL, { x: 6.2, y: 4.8, w: 6, h: 4, fill: { color: BLUE }, line: { type: "none" }, objectName: "div8_wave" });
  txt(s, "第8章", 0.6, 2.6, 4.6, 0.5, { size: 16, color: GRAY, align: "center" });
  txt(s, "船井総研の\nご支援とお見積り", 0.4, 3.1, 5.0, 1.4, { size: 30, bold: true, color: NAVY, align: "center", valign: "middle" });
  txt(s, "ともに、地域に必要とされる遺品整理事業を。", 5.6, 3.3, 4.2, 0.8, { size: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
  s.addNotes("2回目の最後にだけ提示する。ここまでで疑問が解消され、進め方に納得いただいていることを確認してから入る。");
}
{
  const s = content(S8, "3つの支援サービス", "");
  lead(s, "貴社の課題に合わせて、3つのサービスでトータルにご支援します");
  const SV = [
    ["01 HPディレクション支援", "・商圏の競合調査\n・強みの整理と打ち出し\n・コンセプト設計\n・キャッチコピーのご提案\n\n※デザイン制作はご希望のHP制作会社へ依頼", NAVY],
    ["02 ツール＆マニュアル導入支援", "・事業概要マニュアル\n・営業マニュアル\n・現場マニュアル（60ページ以上）\n・集客マニュアル（60ページ以上）\n・見積システム、契約書、チラシ、提携資料など", GREEN],
    ["03 月次コンサルティング", "・事業戦略（3ヵ年計画・財務分析）\n・商品設計\n・集客戦略（HP・チラシ・広告）\n・営業戦略（ツール・研修）\n・採用戦略", BROWN],
  ];
  SV.forEach((v, i) => card(s, 0.4 + i * 3.13, 1.7, 2.94, 4.5, v[0], v[1], { color: v[2], size: 11.5, hsize: 12.5, hh: 0.6 }));
  concl(s, "事業計画から商品・集客・営業・現場まで、一気通貫でご支援します");
}
{
  const s = content(S8, "6ヶ月の立ち上げスケジュール", "開始月を実際の月名に差し替える。各項目で貴社にお願いすること（担当者の決定、協力業者との面談、写真素材の準備など）も伝える。");
  tag(s, "開始月");
  lead(s, "約6ヶ月で、集客・営業・現場の仕組みを立ち上げます", { w: 6 });
  const X0 = 2.2, MW = 1.22, Y0 = 1.65;
  rect(s, 0.4, Y0, 9.2, 0.45, NAVY);
  txt(s, "項目", 0.4, Y0, 1.8, 0.45, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle" });
  for (let m = 0; m < 6; m++) txt(s, `${m + 1}ヶ月目`, X0 + m * MW, Y0, MW, 0.45, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle" });
  const G = [["業者探し", 0, 1.5, "協力業者探し・原価確認", BLUE], ["商品設計", 0, 1.5, "基本パック・オプション設計", GREEN], ["ブランド設計", 0, 1.5, "屋号・ロゴ・キャッチコピー", NAVY], ["HP制作", 0.5, 2.0, "TOP公開→全ページ公開", ORANGE], ["営業体制", 1.7, 4.3, "営業ツール制作・営業研修", GOLD], ["集客施策", 1.7, 4.3, "チラシ・WEB広告・SEO・MEO", "44546A"], ["提携", 1.2, 4.8, "リスト・資料作成→提携営業", RED]];
  G.forEach((g, i) => {
    const y = Y0 + 0.5 + i * 0.5;
    rect(s, 0.4, y, 9.2, 0.46, i % 2 ? WHITE : LGRAY, { line: LINE });
    txt(s, g[0], 0.45, y, 1.75, 0.46, { size: 11, bold: true, color: NAVY, valign: "middle" });
    rect(s, X0 + g[1] * MW, y + 0.06, g[2] * MW, 0.34, g[4], { round: true });
    txt(s, g[3], X0 + g[1] * MW, y + 0.06, g[2] * MW, 0.34, { size: 9.5, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
  txt(s, "月次支援：月1回（約4時間）の打ち合わせ＋LINE・電話・メールでの日常サポート", 0.4, 5.85, 9.2, 0.4, { size: 12, bold: true, color: NAVY });
  concl(s, "○月スタートで、○月に営業開始を目指します");
}
{
  const s = content(S8, "支援プランと費用", "【要確認】ツール＆マニュアル単体の価格（100万円か150万円か）、研究会の年会費。ここでは最新の企画書（96p）の価格を記載。");
  tag(s, "提案プラン");
  lead(s, "貴社の状況に合わせて、3つのプランからお選びいただけます", { w: 6, size: 15 });
  const PL = [
    ["プラン1｜HP制作ディレクションのみ", [["商品設計＋HP制作ディレクション", "700,000円"]], "※HP制作会社へのデザイン制作費は別途", LBLUE],
    ["プラン2｜マニュアル・オリジナルツールのみ", [["マニュアル・オリジナルツール導入支援", "1,500,000円"]], "", LGREEN],
    ["プラン3｜フルサポート（おすすめ）", [["HP制作ディレクション", "700,000円"], ["マニュアル・オリジナルツール導入支援（プラン2より50万円お値引き）", "1,000,000円"], ["業績アップ月次支援コンサルティング", "200,000円／月"]], "※月次支援はオンライン＋必要に応じてご訪問", "FBEDE3"],
  ];
  let y = 1.6;
  PL.forEach((p, i) => {
    txt(s, p[0], 0.4, y, 9.2, 0.38, { size: 13, bold: true, color: i === 2 ? RED : NAVY });
    y += 0.4;
    p[1].forEach((r) => {
      rect(s, 0.4, y, 6.6, 0.42, p[3], { line: DARK, lw: 1.25 });
      txt(s, r[0], 0.5, y, 6.4, 0.42, { size: 11.5, bold: true, color: DARK, valign: "middle" });
      rect(s, 7.0, y, 2.6, 0.42, p[3], { line: DARK, lw: 1.25 });
      txt(s, r[1] + "（税抜）", 7.05, y, 2.5, 0.42, { size: 11.5, bold: true, color: DARK, valign: "middle", align: "right" });
      y += 0.42;
    });
    if (p[2]) { txt(s, p[2], 0.4, y, 9.2, 0.3, { size: 9.5, color: RED, align: "right" }); y += 0.3; }
    y += 0.12;
  });
  txt(s, "・コンサルティングは月次支援と研究会の両方を含むサービスです　・PPC運用代行は広告運用費の20%の手数料が発生します", 0.4, 5.75, 9.2, 0.5, { size: 9.5, color: GRAY });
  concl(s, "おすすめは、立上げから定着まで伴走するプラン3「フルサポート」です");
}
{
  const s = content(S8, "投資総額と回収の見通し", "初期投資（HP・許可）だけでなく、支援費用も含めた総額で示す。後から「聞いていた金額と違う」とならないようにする。");
  tag(s, "総額・回収時期");
  lead(s, "支援費用を含めた投資総額と、回収の見通しをお示しします", { w: 6, size: 15 });
  table(s, [
    ["初年度の投資（プラン3の場合の例）", "金額（税抜）"],
    ["HP制作ディレクション", "70万円"],
    ["マニュアル・ツール導入支援", "100万円"],
    ["月次支援コンサルティング（20万円×12ヶ月）", "240万円"],
    ["HP制作費（制作会社へ）", "○○万円"],
    ["研究会 年会費", "○○万円"],
    ["古物商許可の取得", "約2〜8万円"],
    [{ text: "合計", options: { bold: true } }, { text: "○○○万円", options: { bold: true, color: RED } }],
  ], 0.4, 1.65, 4.9, [3.4, 1.5], { size: 11, align: "right", rowH: 0.5 });
  card(s, 5.5, 1.65, 4.1, 2.0, "回収の見通し", "遺品整理の営業利益：1年目 ○○万円／2年目 ○○万円\n本業への波及利益：年 ○○万円", { color: NAVY, size: 12 });
  arrowD(s, 7.3, 3.72, 0.5, 0.32, RED);
  rect(s, 5.5, 4.1, 4.1, 1.55, RED, { round: true });
  txt(s, "○年目に投資総額を回収\n（本業の利益を含めると ○ヶ月）", 5.5, 4.1, 4.1, 1.55, { size: 16, bold: true, color: WHITE, align: "center", valign: "middle" });
  concl(s, "支援費用を含めても、○年目に回収できる計画です");
  note(s, MODEL_NOTE);
}
{
  const s = content(S8, "次のステップ", "本日決めていただきたいことを3つに絞り、その場で決めていただく。");
  tag(s, "開始月・担当者");
  lead(s, "本日決めていただきたいこと", { w: 6 });
  const D = [["支援プラン", "プラン1〜3から選択"], ["推進担当者", "社内で事業を進める方（兼任可）"], ["開始月", "○月スタート"]];
  D.forEach((d, i) => card(s, 0.4 + i * 3.13, 1.65, 2.94, 1.35, d[0], d[1], { color: RED, num: i + 1, size: 12, hsize: 14 }));
  txt(s, "契約後90日の動き", 0.4, 3.2, 9.2, 0.4, { size: 14, bold: true, color: NAVY });
  steps(s, [["1ヶ月目", "競合調査・料金設計・協力業者探し・屋号の検討"], ["2ヶ月目", "屋号・ロゴ決定、HP制作、提携先リスト作成"], ["3ヶ月目", "HP公開・集客開始・営業研修・提携営業開始"]], 3.65, 1.9, { size: 12, hsize: 13, colors: [BLUE, NAVY, GREEN] });
  txt(s, "契約概要（要点）：月次支援は原則1年契約／研究会への入会が必要／成果を保証するものではありません（全文は付録）", 0.4, 5.7, 9.2, 0.5, { size: 10, color: GRAY });
  concl(s, "○月スタートで、○月の営業開始を目指しましょう");
}

// =====================================================================
// 付録
// =====================================================================
const SA = "付録";
pres.addSection({ title: SA });
divider(SA, "Appendix", "付録", "必要に応じて開く。業種別ページは6章「貴社の業種でのシナジー」に差し込んで使う。");

const INDUSTRIES = [
  ["不動産会社", ["仕入れ競争の激化：ポータル・一括査定経由で奪い合い", "集客コストの高騰：Web広告・反響サイトの費用が上昇", "案件の長期化：売却前の片付けや相続手続きで止まる"], ["相続案件の入口獲得", "物件仕入れの強化", "残置物撤去の内製化", "売却スピードの向上"], ["売却前の所有者と接点", "外注費が自社の粗利に", "相見積もりのない仕入れ", "地域での信頼向上"], ["遺品整理", "不動産相談", "売却・買取再販", "リフォーム・解体"], "「相続案件の入口獲得」「物件仕入れの強化」「売却スピード向上」を実現"],
  ["解体業者", ["利益率の低下：解体単体の価格競争", "集客コストの増加：ネット広告や紹介依存", "新たな収益源の確保", "職人稼働の平準化"], ["解体前案件の獲得", "客単価の向上", "家財撤去を自社で行い粗利UP", "紹介依存からの脱却"], ["受注件数の増加", "売上の安定化", "地域での信頼向上", "職人の稼働率向上"], ["遺品整理", "家財撤去・処分", "解体工事", "不動産・売却"], "「解体前案件の獲得」「客単価の向上」「粗利の向上」「安定受注」を実現"],
  ["住宅・リフォーム業", ["新築市場の縮小", "大型リフォーム案件の不足", "OB顧客との接点不足", "空き家問題の増加"], ["実家リノベ案件の獲得", "相続空き家市場の攻略", "リフォーム受注の増加", "他事業との連携強化"], ["新たな顧客層の獲得", "高単価案件の増加", "OB顧客との関係強化", "地域での競争優位性"], ["遺品整理", "不動産・相続相談", "リノベ・リフォーム", "アフターサポート"], "「実家リノベ案件の獲得」「相続空き家市場の攻略」「リフォーム受注の増加」を実現"],
  ["葬儀業", ["家族葬化で葬儀単価が下落", "集客コストの高騰", "葬儀後のアフター収益不足", "会員組織を活かしきれていない"], ["葬儀後サポートの強化", "アフター売上の創出", "会員組織の活用", "関連事業への展開"], ["顧客満足度の向上", "アフター収益の増加", "紹介・口コミの増加", "LTVの向上"], ["葬儀施行", "遺品整理", "供養・生前整理", "不動産・相続連携"], "「葬儀後サポート強化」「アフター売上創出」「会員組織の活用」を実現"],
  ["ケアマネジャー・包括支援センター", ["片付けの相談先が不足", "高齢者の住環境問題", "家族からの相談増加", "業務負担の増加"], ["地域課題の解決", "生前整理ニーズへの対応", "家族支援の強化", "地域包括ケアとの連携強化"], ["利用者・家族の信頼向上", "相談件数の増加", "業務負担の軽減", "他事業者との連携強化"], ["相談対応", "専門業者の紹介", "整理・片付け", "関連サービス連携"], "「地域課題の解決」「生前整理への対応」「家族支援の強化」を実現"],
  ["便利業", ["利益率が低い", "小口案件が多く件数頼み", "移動や見積で生産性が低い", "ブランド力が弱い"], ["高単価商品化", "顧客単価の向上", "既存顧客の活用", "ブランド力の向上"], ["収益性の向上", "安定した案件確保", "顧客満足度の向上", "紹介・口コミの増加"], ["生前整理・片付け", "遺品整理", "空き家片付け", "各種オプション"], "「高単価商品化」「顧客単価の向上」「ブランド力の向上」を実現"],
  ["リサイクル業", ["良質な仕入れが安定しない", "回収・選別だけでは利益率が低い", "一般のお客様との接点が少ない", "回収できる物が限られる"], ["良質なリユース品を直接確保", "お客様との直接取引が増える", "リユース・買取収益を最大化", "周辺サービスへ展開"], ["顧客数の増加", "利益率の向上", "回収量の増加", "地域ブランドの向上"], ["遺品整理", "選別・仕分け", "リユース・買取", "資源回収・処分"], "「良質な仕入れの確保」「利益率向上」「一般顧客の獲得」を実現"],
  ["買取業", ["高価値商品の仕入れが安定しない", "買取できない品物が多い", "一度の売上が小さい", "相続世代との接点が少ない"], ["高価値商品の仕入れ強化", "買取できない物も対応可能", "客単価・利益率アップ", "次のビジネスへ展開"], ["高利益商品の仕入れ増加", "客単価の向上", "顧客満足度の向上", "継続的な収益化"], ["遺品整理相談", "査定・買取", "販売", "不用品回収・不動産紹介"], "「高利益商品の安定仕入れ」「客単価向上」「継続的な収益機会の創出」を実現"],
];
INDUSTRIES.forEach(([name, issues, syn, eff, flow, sum]) => {
  const s = content(SA, "業種別シナジー｜" + name, "6章「貴社の業種でのシナジー」に差し込んで使う。");
  rect(s, 0.4, 1.0, 3.1, 0.45, NAVY);
  txt(s, "経営課題", 0.4, 1.0, 3.1, 0.45, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
  issues.forEach((t, i) => {
    const y = 1.5 + i * 0.75;
    rect(s, 0.4, y, 3.1, 0.68, i % 2 ? WHITE : LBLUE, { line: LINE });
    circleNum(s, i + 1, 0.48, y + 0.17, 0.34, NAVY);
    txt(s, t, 0.9, y, 2.55, 0.68, { size: 10.5, color: DARK, valign: "middle" });
  });
  rect(s, 3.7, 1.0, 5.9, 0.45, NAVY);
  txt(s, "遺品整理事業参入によるシナジー", 3.7, 1.0, 5.9, 0.45, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
  syn.forEach((t, i) => {
    const w = (5.9 - 0.15 * 3) / 4, x = 3.7 + i * (w + 0.15);
    rect(s, x, 1.5, w, 1.45, LBLUE, { line: LINE });
    circleNum(s, i + 1, x + w / 2 - 0.2, 1.6, 0.4, BLUE);
    txt(s, t, x + 0.05, 2.05, w - 0.1, 0.85, { size: 11, bold: true, color: NAVY, align: "center", valign: "middle" });
  });
  txt(s, "期待できる効果", 3.7, 3.05, 5.9, 0.35, { size: 12, bold: true, color: GREEN });
  eff.forEach((t, i) => {
    const w = (5.9 - 0.15 * 3) / 4, x = 3.7 + i * (w + 0.15);
    rect(s, x, 3.4, w, 0.85, LGREEN, { line: "B9D8C4" });
    txt(s, t, x + 0.05, 3.4, w - 0.1, 0.85, { size: 10.5, bold: true, color: GREEN, align: "center", valign: "middle" });
  });
  rect(s, 0.4, 4.6, 9.2, 0.4, PURPLE);
  txt(s, "ビジネスモデルのイメージ", 0.4, 4.6, 9.2, 0.4, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
  flow.forEach((t, i) => {
    const w = 1.9, x = 0.55 + i * 2.33;
    rect(s, x, 5.08, w, 0.62, "F1ECF8", { line: "CBBDE3", round: true });
    txt(s, t, x, 5.08, w, 0.62, { size: 11, bold: true, color: PURPLE, align: "center", valign: "middle" });
    if (i < flow.length - 1) arrowR(s, x + w + 0.07, 5.27, 0.28, 0.24, PURPLE);
  });
  rect(s, 0.4, 5.85, 9.2, 0.5, NAVY, { round: true });
  txt(s, "まとめ：" + sum, 0.5, 5.85, 9.0, 0.5, { size: 11.5, bold: true, color: YEL, valign: "middle" });
  concl(s, `遺品整理は、${name.replace(/・包括支援センター/, "")}にとって本業を伸ばす戦略になります`, { size: 15 });
});

// 不動産補足
{
  const s = content(SA, "補足｜不動産会社にとっての環境変化", "【要確認】法改正の出典（法務省の通達等）を確認し、出典欄に記載する。");
  lead(s, "不動産情報を「風上」で掴む方法の一つが、使えなくなります");
  card(s, 0.4, 1.7, 2.95, 2.4, "これまで", "相続登記をきっかけに、法務局の不動産登記受付帳を開示請求し、所有者へDM・営業でアプローチできた", { color: "8A8F99", size: 11.5 });
  arrowR(s, 3.42, 2.7, 0.35, 0.4, RED);
  card(s, 3.85, 1.7, 2.95, 2.4, "2026年10月から", "不動産登記受付帳の様式変更により、「登記の目的」「不動産の所在」が表記されなくなり、営業目的での情報取得が事実上できなくなる", { color: RED, size: 11.5 });
  arrowR(s, 6.87, 2.7, 0.35, 0.4, RED);
  card(s, 7.3, 1.7, 2.3, 2.4, "これから", "新しい「風上の入口」が必要\n＝「空き家×遺品整理」", { color: NAVY, size: 11.5 });
  table(s, [["遺品整理120件あたり", "件数の目安", "主な出口", "年間粗利の目安"], ["不動産相談の発生（5〜10%）", "6〜12件", "", ""], ["S・Aランク（立地良・築浅）", "2〜4件", "買取再販", "200〜900万円"], ["Bランク（標準）", "2〜5件", "売買仲介", "300〜600万円"], ["Cランク（立地悪・築古）", "2〜3件", "解体", "60〜300万円"]], 0.4, 4.3, 9.2, [3.0, 1.6, 2.0, 2.6], { size: 11, rowH: 0.42 });
  note(s, "出典：（要記入）　" + MODEL_NOTE);
}
// 契約概要
{
  const s = content(SA, "コンサルティングに伴う契約概要", "");
  const items = [
    "月次支援は原則1年契約です（調査期間を除く）。月次支援フィーとは別に、交通費を毎月固定額でご請求します。",
    "月次支援2年目以降のフィーは、コンサルタントの役職に応じて変動する場合があります。",
    "月次支援のご契約には研究会へのご入会が必要です。月次支援は「クラウドサイン」、研究会は「船井総研マイページ」でのお手続きです。",
    "月次支援・交通費は毎月のご支援終了後に請求書を送付し、銀行振込となります。",
    "ご提案した施策に別途費用がかかる場合があります（購入費・印刷費・サイト制作費など）。",
    "月次支援のご契約は一定の成果を保証するものではありません。成果がないことを理由とした返金は行いません。",
    "本資料の数値シミュレーションや事例は一例であり、諸条件により大きく変わります。",
    "リスティング広告を実施する場合、広告費用の20%〜が運用手数料として発生します。",
    "顧客情報などの個人情報を含む資料は原則お預かりしません。必要な場合は別途確認事項に合意いただきます。",
    "弊社が制作したツールの著作権は弊社に帰属します。",
    "弊社の営業時間は平日9:30〜18:00です。時間外・土日祝のご連絡への返信は翌営業日以降となります。",
  ];
  txt(s, items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1 } })), 0.4, 1.1, 9.2, 5.8, { size: 12.5, ls: 1.15, psa: 7 });
  note(s, "※最新の契約条件に合わせて更新してください。");
}

// ---------- 出力 ----------
(async () => {
  await pres.writeFile({ fileName: OUT });
  const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
  await applyTheme(OUT, THEME);
  console.log("written:", OUT);
})();
