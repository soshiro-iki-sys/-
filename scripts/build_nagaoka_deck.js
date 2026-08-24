/* 株式会社ヤシロ 第2講座 資料ジェネレータ
   フォーマットは docs/ヤマキシ_補助金活用講座_資料フォーマット仕様.md に準拠 */
const PptxGenJS = require("pptxgenjs");

const NAVY = "002060";
const RED = "C00000";
const HRED = "FF0000";
const YEL = "FFFF00";
const WHITE = "FFFFFF";
const BLACK = "000000";
const PALE = "EDF3FB";
const GRAY = "BFBFBF";
const F = "Meiryo UI";

const SW = 11.69, SH = 8.27;
const ML = 0.16;               // 左右マージン
const CW = SW - ML * 2;        // コンテンツ幅 11.37

const pres = new PptxGenJS();
pres.defineLayout({ name: "A4LAND", width: SW, height: SH });
pres.layout = "A4LAND";
pres.author = "株式会社ヤシロ";
pres.company = "株式会社ヤシロ";
pres.title = "太陽光・蓄電池 補助金活用講座";

const notes = [];   // 差し替え一覧
function note(page, element, item) { notes.push({ page, element, item }); }

/* ---------- 共通グリッド ---------- */
function chapterTitle(s, text) {
  s.addText(text, {
    x: ML, y: 0.03, w: CW, h: 0.61, fontFace: F, fontSize: 30, bold: true,
    color: BLACK, margin: 0, valign: "middle", align: "left",
  });
}

function chip(s, text) {
  const w = Math.min(text.length * 0.34 + 0.34, 7.0);
  s.addText(text, {
    x: ML, y: 0.73, w, h: 0.42, fontFace: F, fontSize: 20, bold: true,
    color: WHITE, fill: { color: NAVY }, align: "center", valign: "middle", margin: 0,
  });
}

function lead(s, text, lines) {
  const h = (lines || 1) === 1 ? 0.5 : 0.86;
  s.addText(text, {
    x: ML, y: 1.12, w: CW, h, fontFace: F, fontSize: 22, bold: true,
    color: BLACK, margin: 0, valign: "middle", align: "left", lineSpacingMultiple: 1.15,
  });
}

function source(s, text) {
  s.addText(text, {
    x: ML, y: 6.76, w: CW, h: 0.28, fontFace: F, fontSize: 11,
    color: "595959", margin: 0, valign: "middle", align: "right",
  });
}

/* 結論バー。text は文字列、または {t, hl} の配列（hl:true で黄色ハイライト） */
function footerBar(s, text, sizeOverride) {
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 7.12, w: SW, h: SH - 7.12, fill: { color: NAVY }, line: { color: NAVY, width: 0 },
  });
  let runs, plainLen;
  if (typeof text === "string") {
    plainLen = text.length;
    runs = [{ text, options: { color: WHITE } }];
  } else {
    plainLen = text.reduce((a, r) => a + r.t.length, 0);
    runs = text.map((r) => ({ text: r.t, options: { color: r.hl ? YEL : WHITE } }));
  }
  let sz = sizeOverride;
  if (!sz) sz = plainLen <= 24 ? 32 : plainLen <= 32 ? 28 : plainLen <= 46 ? 24 : 20;
  runs.forEach((r) => Object.assign(r.options, { fontFace: F, fontSize: sz, bold: true, breakLine: false }));
  s.addText(runs, {
    x: 0.3, y: 7.12, w: SW - 0.6, h: SH - 7.12, align: "center", valign: "middle",
    margin: 0, lineSpacingMultiple: 1.1,
  });
}

/* 本編スライドの雛形 */
function contentSlide(o) {
  const s = pres.addSlide();
  chapterTitle(s, o.chapter);
  if (o.chip) chip(s, o.chip);
  if (o.lead) lead(s, o.lead, o.leadLines);
  if (o.source) source(s, o.source);
  if (o.footer) footerBar(s, o.footer, o.footerSize);
  if (o.note) s.addNotes(o.note);
  return s;
}

/* ---------- 部品 ---------- */
/* 角丸カード */
function card(s, x, y, w, h, opts) {
  const o = opts || {};
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.09,
    fill: { color: o.fill || WHITE },
    line: { color: o.line || NAVY, width: o.lw === undefined ? 1.5 : o.lw },
  });
}

/* 見出しタブ付きカード（2カラム比較用） */
function tabCard(s, x, y, w, h, title, tabColor) {
  card(s, x, y + 0.28, w, h - 0.28, { fill: WHITE, line: NAVY, lw: 1.5 });
  s.addText(title, {
    x: x + 0.24, y, w: w - 0.48, h: 0.56, fontFace: F, fontSize: 20, bold: true,
    color: WHITE, fill: { color: tabColor }, align: "center", valign: "middle", margin: 0,
  });
}

/* 番号つきの項目行 */
function bulletRow(s, x, y, w, h, num, text, circleColor, fontSize) {
  s.addShape(pres.ShapeType.ellipse, {
    x, y: y + (h - 0.52) / 2, w: 0.52, h: 0.52,
    fill: { color: circleColor }, line: { color: circleColor, width: 0 },
  });
  s.addText(String(num), {
    x, y: y + (h - 0.52) / 2, w: 0.52, h: 0.52, fontFace: F, fontSize: 18, bold: true,
    color: WHITE, align: "center", valign: "middle", margin: 0,
  });
  s.addText(text, {
    x: x + 0.68, y, w: w - 0.68, h, fontFace: F, fontSize: fontSize || 17, bold: true,
    color: BLACK, align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.12,
  });
}

/* 数字を大きく見せる統計カード */
function statCard(s, x, y, w, h, label, value, unit, opts) {
  const o = opts || {};
  card(s, x, y, w, h, { fill: o.fill || WHITE, line: o.line || NAVY, lw: 2 });
  s.addText(label, {
    x: x + 0.12, y: y + 0.16, w: w - 0.24, h: 0.42, fontFace: F, fontSize: 15, bold: true,
    color: o.labelColor || NAVY, align: "center", valign: "middle", margin: 0,
  });
  s.addText(
    [
      { text: value, options: { fontFace: F, fontSize: o.valueSize || 44, bold: true, color: o.valueColor || NAVY } },
      { text: unit ? " " + unit : "", options: { fontFace: F, fontSize: 18, bold: true, color: o.valueColor || NAVY } },
    ],
    { x: x + 0.12, y: y + 0.56, w: w - 0.24, h: h - 0.72, align: "center", valign: "middle", margin: 0 }
  );
}

/* 表 */
function table(s, rows, x, y, w, colW, opts) {
  const o = opts || {};
  s.addTable(rows, {
    x, y, w, colW,
    border: { type: "solid", color: o.border || GRAY, pt: 1 },
    fontFace: F, fontSize: o.fontSize || 14, valign: "middle",
    rowH: o.rowH || 0.34, margin: o.margin === undefined ? 0.06 : o.margin,
    autoPage: false,
  });
}
function th(t) { return { text: t, options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "center" } }; }
function lbl(t) { return { text: t, options: { fill: { color: PALE }, color: BLACK, bold: true, align: "left" } }; }
function val(t, o) { return { text: t || "", options: Object.assign({ color: BLACK, align: "left" }, o || {}) }; }
function blank() { return { text: "", options: { color: BLACK } }; }

/* 目次 */
const AGENDA = [
  "長岡市のエネルギー事情と災害対策",
  "長岡市でも太陽光はお得なのか",
  "太陽光・蓄電池のデメリットと弊社の対策",
  "正しい業者の選び方",
  "まとめ",
];
function agendaSlide(active) {
  const s = pres.addSlide();
  chapterTitle(s, "本日の目次");
  const runs = [];
  AGENDA.forEach((t, i) => {
    runs.push({
      text: t,
      options: {
        fontFace: F, fontSize: 36, bold: true,
        color: i === active ? HRED : BLACK,
        breakLine: i !== AGENDA.length - 1,
        paraSpaceAfter: 14,
      },
    });
  });
  s.addText(runs, { x: 0.6, y: 1.5, w: SW - 1.2, h: 5.4, align: "left", valign: "middle", margin: 0 });
  return s;
}

/* 章扉（全面ネイビー） */
function dividerSlide(linesArr, size) {
  const s = pres.addSlide();
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: SW, h: SH, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  const runs = linesArr.map((t, i) => ({
    text: t,
    options: {
      fontFace: F, fontSize: size || 48, bold: true, color: t === "" ? NAVY : WHITE,
      breakLine: i !== linesArr.length - 1, paraSpaceAfter: 8,
    },
  }));
  s.addText(runs, { x: 0.8, y: 1.2, w: SW - 1.6, h: SH - 2.4, align: "center", valign: "middle", margin: 0 });
  return s;
}

/* 2カラム比較（デメリット / ヤシロの対応） */
function twoColSlide(o) {
  const s = contentSlide({
    chapter: o.chapter, chip: o.chip, lead: o.lead, leadLines: o.leadLines,
    footer: o.footer, footerSize: o.footerSize, note: o.note, source: o.source,
  });
  const cy = 1.86, ch = 5.02, cw = 5.36;
  const lx = 0.42, rx = SW - 0.42 - cw;
  tabCard(s, lx, cy, cw, ch, o.leftTitle, o.leftColor || RED);
  tabCard(s, rx, cy, cw, ch, o.rightTitle, o.rightColor || NAVY);
  const top = cy + 0.52, rowH = 1.36, gap = 0.14;
  o.left.forEach((t, i) => bulletRow(s, lx + 0.26, top + i * (rowH + gap), cw - 0.52, rowH, i + 1, t, o.leftColor || RED, o.fs));
  o.right.forEach((t, i) => bulletRow(s, rx + 0.26, top + i * (rowH + gap), cw - 0.52, rowH, i + 1, t, o.rightColor || NAVY, o.fs));
  return s;
}

module.exports = {
  pres, notes, note, NAVY, RED, HRED, YEL, WHITE, BLACK, PALE, GRAY, F, SW, SH, ML, CW,
  chapterTitle, chip, lead, source, footerBar, contentSlide, card, tabCard, bulletRow,
  statCard, table, th, lbl, val, blank, agendaSlide, dividerSlide, twoColSlide, AGENDA,
};
