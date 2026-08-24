/* 株式会社ヤシロ 第2講座 資料ジェネレータ（共通部品）
   フォーマットは docs/ヤマキシ_補助金活用講座_資料フォーマット仕様.md に準拠 */
const PptxGenJS = require("pptxgenjs");

const NAVY = "002060", NAVY2 = "12386E", RED = "C00000", HRED = "FF0000";
const YEL = "FFFF00", WHITE = "FFFFFF", BLACK = "000000";
const PALE = "EDF3FB", PINK = "FDE9E7", GRAY = "BFBFBF", GRAY2 = "808080";
const INK = "404040", MUTE = "595959", SOFT = "F2F2F2";
const F = "Meiryo UI";

const SW = 11.69, SH = 8.27;
const ML = 0.16;
const CW = SW - ML * 2;
const BAR_Y = 7.12, BAR_H = SH - BAR_Y;

const pres = new PptxGenJS();
pres.defineLayout({ name: "A4LAND", width: SW, height: SH });
pres.layout = "A4LAND";
pres.author = "株式会社ヤシロ";
pres.company = "株式会社ヤシロ";
pres.title = "太陽光・蓄電池 補助金活用講座";

const notes = [];
function note(page, element, item) { notes.push({ page, element, item }); }

/* ---------- 文字幅の推定（全角=1em、半角=0.55em） ---------- */
function wideLen(s) {
  let n = 0;
  for (const ch of s) {
    const c = ch.codePointAt(0);
    const half = (c >= 0x20 && c <= 0x7e) || (c >= 0xff61 && c <= 0xff9f);
    n += half ? 0.55 : 1;
  }
  return n;
}
/* 幅 w(inch) に1行で収まる最大サイズ */
function fitSize(text, w, max, min) {
  const we = Math.max(1, wideLen(text));
  return Math.max(min, Math.min(max, Math.floor((w * 72) / we)));
}

/* ---------- 共通グリッド ---------- */
function chapterTitle(s, text) {
  s.addText(text, {
    x: ML, y: 0.03, w: CW, h: 0.61, fontFace: F, fontSize: 30, bold: true,
    color: BLACK, margin: 0, valign: "middle", align: "left",
  });
}

function chip(s, text) {
  const w = Math.min(wideLen(text) * 0.31 + 0.36, 7.2);
  s.addText(text, {
    x: ML, y: 0.73, w, h: 0.42, fontFace: F, fontSize: 20, bold: true,
    color: WHITE, fill: { color: NAVY }, align: "center", valign: "middle", margin: 0,
  });
}

/* リード文。長ければ自動で2行枠にする */
function lead(s, text) {
  const oneLine = wideLen(text) * (22 / 72) <= CW - 0.1;
  s.addText(text, {
    x: ML, y: 1.12, w: CW, h: oneLine ? 0.5 : 0.86, fontFace: F, fontSize: 22, bold: true,
    color: BLACK, margin: 0, valign: "middle", align: "left", lineSpacingMultiple: 1.15,
  });
  return oneLine ? 1.68 : 2.04;   // 本文の推奨開始 y
}

function source(s, text) {
  s.addText(text, {
    x: ML, y: 6.78, w: CW, h: 0.26, fontFace: F, fontSize: 10.5,
    color: MUTE, margin: 0, valign: "middle", align: "right",
  });
}

/* 結論バー：必ず1行に収まるサイズで描く */
function footerBar(s, text) {
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: BAR_Y, w: SW, h: BAR_H, fill: { color: NAVY }, line: { color: NAVY, width: 0 },
  });
  const parts = typeof text === "string" ? [{ t: text }] : text;
  const plain = parts.map((p) => p.t).join("");
  const sz = fitSize(plain, SW - 0.9, 32, 20);
  const runs = parts.map((p) => ({
    text: p.t,
    options: { fontFace: F, fontSize: sz, bold: true, color: p.hl ? YEL : WHITE },
  }));
  s.addText(runs, {
    x: 0.35, y: BAR_Y, w: SW - 0.7, h: BAR_H, align: "center", valign: "middle", margin: 0,
  });
}

function contentSlide(o) {
  const s = pres.addSlide();
  chapterTitle(s, o.chapter);
  if (o.chip) chip(s, o.chip);
  s.bodyY = o.lead ? lead(s, o.lead) : 1.3;
  if (o.source) source(s, o.source);
  if (o.footer) footerBar(s, o.footer);
  if (o.note) s.addNotes(o.note);
  return s;
}

/* ---------- 部品 ---------- */
function card(s, x, y, w, h, o) {
  o = o || {};
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.09,
    fill: { color: o.fill || WHITE },
    line: { color: o.line || NAVY, width: o.lw === undefined ? 1.5 : o.lw },
  });
}

/* 見出しタブ付きカード */
function tabCard(s, x, y, w, h, title, tabColor) {
  card(s, x, y + 0.28, w, h - 0.28, { fill: WHITE, line: NAVY, lw: 1.5 });
  s.addText(title, {
    x: x + 0.24, y, w: w - 0.48, h: 0.56, fontFace: F,
    fontSize: fitSize(title, w - 0.7, 20, 14), bold: true,
    color: WHITE, fill: { color: tabColor }, align: "center", valign: "middle", margin: 0,
  });
}

/* 帯見出し */
function band(s, x, y, w, text, opts) {
  const o = opts || {};
  s.addText(text, {
    x, y, w, h: o.h || 0.48, fontFace: F, fontSize: o.size || fitSize(text, w - 0.4, 20, 13),
    bold: true, color: o.color || WHITE, fill: { color: o.fill || NAVY },
    align: "center", valign: "middle", margin: 0,
  });
}

function bulletRow(s, x, y, w, h, num, text, circleColor, fontSize) {
  const d = 0.5;
  s.addShape(pres.ShapeType.ellipse, {
    x, y: y + (h - d) / 2, w: d, h: d,
    fill: { color: circleColor }, line: { color: circleColor, width: 0 },
  });
  s.addText(String(num), {
    x, y: y + (h - d) / 2, w: d, h: d, fontFace: F, fontSize: 17, bold: true,
    color: WHITE, align: "center", valign: "middle", margin: 0,
  });
  s.addText(text, {
    x: x + d + 0.16, y, w: w - d - 0.16, h, fontFace: F, fontSize: fontSize || 16, bold: true,
    color: BLACK, align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.14,
  });
}

function statCard(s, x, y, w, h, label, value, unit, o) {
  o = o || {};
  card(s, x, y, w, h, { fill: o.fill || WHITE, line: o.line || NAVY, lw: 2 });
  s.addText(label, {
    x: x + 0.1, y: y + 0.14, w: w - 0.2, h: 0.4, fontFace: F,
    fontSize: fitSize(label, w - 0.3, 15, 11), bold: true,
    color: o.labelColor || NAVY, align: "center", valign: "middle", margin: 0,
  });
  const vs = o.valueSize || 40;
  s.addText([
    { text: value, options: { fontFace: F, fontSize: vs, bold: true, color: o.valueColor || NAVY } },
    { text: unit ? " " + unit : "", options: { fontFace: F, fontSize: Math.round(vs * 0.42), bold: true, color: o.valueColor || NAVY } },
  ], { x: x + 0.08, y: y + 0.52, w: w - 0.16, h: h - 0.64, align: "center", valign: "middle", margin: 0 });
}

function table(s, rows, x, y, w, colW, o) {
  o = o || {};
  s.addTable(rows, {
    x, y, w, colW,
    border: { type: "solid", color: o.border || GRAY, pt: 1 },
    fontFace: F, fontSize: o.fontSize || 14, valign: "middle",
    rowH: o.rowH || 0.34, margin: o.margin === undefined ? 0.06 : o.margin,
    autoPage: false,
  });
}
function th(t, o) { return { text: t, options: Object.assign({ fill: { color: NAVY }, color: WHITE, bold: true, align: "center" }, o || {}) }; }
function lbl(t) { return { text: t, options: { fill: { color: PALE }, color: BLACK, bold: true, align: "left" } }; }
function val(t, o) { return { text: t || "", options: Object.assign({ color: BLACK, align: "left" }, o || {}) }; }
function blank() { return { text: "", options: { color: BLACK } }; }

/* 3列の比較テーブル（項目 / A / B）。行ラベルの重なりを避けるための専用部品 */
function compareTable(s, x, y, w, titleA, titleB, rows, o) {
  o = o || {};
  const cLabel = o.labelW || 2.1;
  const cCol = (w - cLabel) / 2;
  const body = rows.map((r) => [
    { text: r[0], options: { fill: { color: PALE }, color: NAVY, bold: true, align: "center", fontSize: 14 } },
    { text: r[1], options: { fill: { color: SOFT }, color: INK, bold: true, align: "center", fontSize: 15 } },
    { text: r[2], options: { fill: { color: WHITE }, color: NAVY, bold: true, align: "center", fontSize: 15 } },
  ]);
  const head = [[
    { text: "", options: { fill: { color: WHITE }, border: [{ type: "none" }, { type: "none" }, { type: "solid", color: GRAY, pt: 1 }, { type: "none" }] } },
    th(titleA, { fill: { color: GRAY2 } }),
    th(titleB),
  ]];
  table(s, head.concat(body), x, y, w, [cLabel, cCol, cCol], { rowH: o.rowH || 0.66, fontSize: 15 });
}

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
  const y0 = 1.42, h = 0.94, gap = 0.2;
  AGENDA.forEach((t, i) => {
    const y = y0 + i * (h + gap);
    const on = i === active;
    card(s, 1.5, y, 8.7, h, { fill: on ? PINK : WHITE, line: on ? HRED : GRAY, lw: on ? 2.5 : 1 });
    s.addShape(pres.ShapeType.ellipse, {
      x: 1.76, y: y + (h - 0.56) / 2, w: 0.56, h: 0.56,
      fill: { color: on ? HRED : NAVY }, line: { color: on ? HRED : NAVY, width: 0 },
    });
    s.addText(String(i + 1), {
      x: 1.76, y: y + (h - 0.56) / 2, w: 0.56, h: 0.56, fontFace: F, fontSize: 20, bold: true,
      color: WHITE, align: "center", valign: "middle", margin: 0,
    });
    s.addText(t, {
      x: 2.54, y, w: 7.4, h, fontFace: F, fontSize: on ? 26 : 24, bold: true,
      color: on ? HRED : BLACK, align: "left", valign: "middle", margin: 0,
    });
  });
  return s;
}

/* 章扉 */
function dividerSlide(lines, size) {
  const s = pres.addSlide();
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: SW, h: SH, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  const runs = lines.map((t, i) => ({
    text: t || " ",
    options: {
      fontFace: F, fontSize: size || 48, bold: true, color: WHITE,
      breakLine: i !== lines.length - 1, paraSpaceAfter: 8,
    },
  }));
  s.addText(runs, { x: 0.8, y: 1.2, w: SW - 1.6, h: SH - 2.4, align: "center", valign: "middle", margin: 0 });
  return s;
}

/* 2カラム比較（デメリット / ヤシロの対応） */
function twoColSlide(o) {
  const s = contentSlide({
    chapter: o.chapter, chip: o.chip, lead: o.lead,
    footer: o.footer, note: o.note, source: o.source,
  });
  const cy = Math.max(s.bodyY + 0.18, 1.86), cw = 5.36;
  const ch = (o.source ? 6.64 : 6.94) - cy;
  const lx = 0.42, rx = SW - 0.42 - cw;
  tabCard(s, lx, cy, cw, ch, o.leftTitle, o.leftColor || RED);
  tabCard(s, rx, cy, cw, ch, o.rightTitle, o.rightColor || NAVY);
  const inner = ch - 0.62, n = o.left.length;
  const rowH = (inner - 0.24) / n, top = cy + 0.68;
  o.left.forEach((t, i) => bulletRow(s, lx + 0.24, top + i * rowH, cw - 0.5, rowH - 0.12, i + 1, t, o.leftColor || RED, o.fs));
  o.right.forEach((t, i) => bulletRow(s, rx + 0.24, top + i * rowH, cw - 0.5, rowH - 0.12, i + 1, t, o.rightColor || NAVY, o.fs));
  return s;
}

module.exports = {
  pres, notes, note, wideLen, fitSize,
  NAVY, NAVY2, RED, HRED, YEL, WHITE, BLACK, PALE, PINK, GRAY, GRAY2, INK, MUTE, SOFT, F,
  SW, SH, ML, CW, BAR_Y, BAR_H,
  chapterTitle, chip, lead, source, footerBar, contentSlide,
  card, tabCard, band, bulletRow, statCard, table, th, lbl, val, blank,
  compareTable, agendaSlide, dividerSlide, twoColSlide, AGENDA,
};
