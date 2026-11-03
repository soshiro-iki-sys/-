// 「遺品整理を中心としたシナジーと商流の流れ」スライド単体の生成スクリプト
// 使い方: PPTX_SKILL=<pptxスキルのパス> node build_slide_shoryu.js <出力パス.pptx>
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa");

const OUT = process.argv[2] || "商流の流れ.pptx";
const NAVY = "1F2D50", NAVY2 = "2B3F6E", BLUE = "2E5FA8", GREEN = "2E7D32", ORANGE = "E07B2E", DARK = "222222", WHITE = "FFFFFF";
const FONT = "Meiryo";

async function png(name, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(FA[name], { color: "#" + color, size: 256 }));
  return "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.defineLayout({ name: "STD_43", width: 10, height: 7.5 });
  pres.layout = "STD_43";
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.defineSlideMaster({
    title: "CONTENT", background: { color: WHITE },
    objects: [
      { rect: { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY } } },
      { rect: { x: 6.2, y: 0, w: 3.8, h: 0.85, fill: { color: NAVY2 } } },
      { text: { text: "Funai Soken", options: { x: 7.6, y: 0.18, w: 2.2, h: 0.5, fontFace: FONT, fontSize: 18, color: WHITE, align: "right", valign: "middle" } } },
      { text: { text: "Confidential", options: { x: 0.2, y: 7.15, w: 2, h: 0.25, fontFace: FONT, fontSize: 8, color: "888888" } } },
      { placeholder: { options: { name: "title", type: "title", x: 0.3, y: 0.1, w: 7.2, h: 0.65, fontFace: FONT, fontSize: 20, color: WHITE, align: "left", valign: "middle", margin: 0 }, text: "" } },
    ],
  });
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("ビジネスモデルと売上目標", { placeholder: "title" });

  let n = 0;
  const nm = (p) => `${p}_${++n}`;
  const txt = (t, x, y, w, h, o = {}) => s.addText(t, { x, y, w, h, fontFace: FONT, fontSize: o.size || 10, bold: !!o.bold, color: o.color || DARK, align: o.align || "left", valign: o.valign || "middle", margin: o.margin === undefined ? 0.02 : o.margin, lineSpacingMultiple: o.ls || 1.1, isTextBox: true, fit: "none", objectName: nm("text") });
  const box = (x, y, w, h, fill, line, o = {}) => s.addShape(o.round === false ? pres.shapes.RECTANGLE : pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: o.r || 0.06, fill: fill ? { color: fill } : { color: WHITE, transparency: 100 }, line: line ? { color: line, width: o.lw || 1, dashType: o.dash || "solid" } : { type: "none" }, shadow: o.shadow ? { type: "outer", color: "000000", opacity: 0.15, blur: 3, offset: 1.5, angle: 90 } : undefined, objectName: nm("box") });
  const ln = (x, y, w, h, color, o = {}) => s.addShape(pres.shapes.LINE, { x, y, w, h, flipV: !!o.flipV, flipH: !!o.flipH, line: { color, width: o.w || 1.5, dashType: o.dash || "solid", endArrowType: o.arrow ? "triangle" : undefined }, objectName: nm("line") });
  const img = async (name, color, x, y, d) => s.addImage({ data: await png(name, color), x, y, w: d, h: d, objectName: nm("icon") });
  const bullets = (arr, x, y, w, h, size) => txt(arr.map((t, i) => ({ text: "・" + t, options: { breakLine: i < arr.length - 1 } })), x, y, w, h, { size, valign: "top", ls: 1.15 });

  // タイトル
  txt("遺品整理を中心としたシナジーと商流の流れ", 0.25, 0.92, 9.5, 0.42, { size: 18, bold: true, color: NAVY, align: "center" });
  txt("さまざまな入口から遺品整理に繋がり、整理後の出口サービスまでワンストップで提供", 0.25, 1.32, 9.5, 0.3, { size: 11, bold: true, align: "center" });

  // ---------- 入口 ----------
  box(0.25, 1.72, 2.5, 0.3, NAVY, null, { r: 0.04 });
  txt("入口（案件の発生源）", 0.25, 1.72, 2.5, 0.3, { size: 10.5, bold: true, color: WHITE, align: "center" });
  const IN = [
    ["FaPrayingHands", "① 葬儀会社", ["葬儀後に遺品整理の相談が発生", "生前相談から将来の葬儀受注に繋がる"]],
    ["FaBuilding", "② 老人ホーム紹介", ["入居前の生前整理が必要", "退去後の実家整理まで一貫対応"]],
    ["FaHome", "③ 不動産仲介・不動産売却", ["売却前には家財撤去が必須", "空き家相談から売却までサポート"]],
  ];
  for (let i = 0; i < IN.length; i++) {
    const y = 2.1 + i * 0.85;
    box(0.25, y, 2.5, 0.77, WHITE, "9DB0CF", { shadow: true });
    await img(IN[i][0], NAVY, 0.35, y + 0.17, 0.43);
    txt(IN[i][1], 0.88, y + 0.04, 1.85, 0.24, { size: 9.5, bold: true, color: NAVY });
    bullets(IN[i][2], 0.88, y + 0.28, 1.85, 0.47, 7);
    ln(2.75, y + 0.385, 0.12, 0, BLUE, { w: 2 });
  }
  ln(2.87, 2.485, 0, 1.7, BLUE, { w: 2 });
  ln(2.87, 3.3, 0.68, 0, BLUE, { w: 3, arrow: true });
  txt("相談・紹介", 2.85, 2.98, 0.75, 0.25, { size: 8.5, bold: true, color: BLUE, align: "center" });

  // ---------- 中心 ----------
  s.addShape(pres.shapes.OVAL, { x: 3.6, y: 2.05, w: 2.35, h: 2.35, fill: { color: "D6E1F0" }, line: { type: "none" }, objectName: "center_halo" });
  s.addShape(pres.shapes.OVAL, { x: 3.7, y: 2.15, w: 2.15, h: 2.15, fill: { color: NAVY }, line: { type: "none" }, objectName: "center" });
  txt("遺品整理", 3.7, 2.3, 2.15, 0.42, { size: 20, bold: true, color: WHITE, align: "center" });
  txt("（元請で受注）", 3.7, 2.7, 2.15, 0.25, { size: 9, bold: true, color: WHITE, align: "center" });
  await img("FaBoxOpen", "E8C690", 4.5, 2.95, 0.55);
  txt([{ text: "遺品の仕分け・整理", options: { breakLine: true } }, { text: "貴重品探索・供養品の選定", options: { breakLine: true } }, { text: "買取・処分・搬出" }], 3.75, 3.52, 2.05, 0.58, { size: 7.5, bold: true, color: WHITE, align: "center", valign: "top" });

  // ---------- 出口 ----------
  ln(5.95, 3.22, 0.85, 0, GREEN, { w: 3, arrow: true });
  txt("整理後の\nニーズ発生", 5.95, 2.72, 0.85, 0.42, { size: 8.5, bold: true, color: NAVY, align: "center" });
  box(7.05, 1.72, 2.7, 0.3, GREEN, null, { r: 0.04 });
  txt("出口（整理後のサービス）", 7.05, 1.72, 2.7, 0.3, { size: 10.5, bold: true, color: WHITE, align: "center" });
  const OUTS = [
    ["FaBroom", "⑦ 清掃業", ["ハウスクリーニング", "消臭・除菌・特殊清掃"]],
    ["FaPaintRoller", "⑧ リフォーム会社", ["住まいの再生・改修", "売却・賃貸前の価値向上"]],
    ["FaGem", "⑥ リユース・買取", ["遺品の買取で費用負担を軽減", "リユース・リサイクル販売"]],
    ["FaHardHat", "④ 解体業者", ["建物解体の前に残置物撤去", "解体工事まで一括で対応"]],
    ["FaTruck", "⑤ 一般廃棄物業者", ["一般廃棄物の収集運搬・処分", "適正処理でコンプライアンス確保"]],
  ];
  for (let i = 0; i < OUTS.length; i++) {
    const y = 2.1 + i * 0.55;
    box(7.05, y, 2.7, 0.5, WHITE, "8BBF8E", { shadow: true });
    await img(OUTS[i][0], GREEN, 7.15, y + 0.08, 0.34);
    txt(OUTS[i][1], 7.58, y + 0.01, 2.15, 0.2, { size: 8.5, bold: true, color: GREEN });
    bullets(OUTS[i][2], 7.58, y + 0.2, 2.15, 0.3, 6.5);
    ln(6.95, y + 0.25, 0.1, 0, GREEN, { w: 2 });
  }
  ln(6.95, 2.35, 0, 2.2, GREEN, { w: 2 });

  // ---------- 中心からの点線 ----------
  ln(3.05, 4.62, 0, 0.3, ORANGE, { dash: "dash", w: 1.75 });
  ln(3.05, 4.62, 1.05, 0, ORANGE, { dash: "dash", w: 1.75 });
  ln(4.1, 4.2, 0, 0.42, ORANGE, { dash: "dash", w: 1.75, arrow: true, flipV: true });
  txt("ケースにより\n直接依頼", 3.2, 4.22, 0.85, 0.36, { size: 7.5, bold: true, color: ORANGE, align: "right" });
  ln(5.0, 4.32, 0, 0.6, BLUE, { dash: "dash", w: 1.75, arrow: true });
  txt("必要に応じて\n連携・発注", 5.08, 4.42, 0.9, 0.36, { size: 7.5, bold: true, color: BLUE });

  // ---------- パターンA ----------
  box(0.25, 5.0, 4.15, 2.05, "FFF8F2", ORANGE, { lw: 1.25 });
  box(0.35, 4.88, 2.35, 0.24, NAVY, null, { r: 0.12 });
  txt("商流のパターン（主な2つの流れ）", 0.35, 4.88, 2.35, 0.24, { size: 7.5, bold: true, color: WHITE, align: "center" });
  txt("パターンA：解体会社からの残置物撤去依頼（下請けの流れ）", 0.3, 5.14, 4.05, 0.24, { size: 8, bold: true, color: ORANGE, align: "center" });
  const PA = [["解体業者", "FaHardHat", "解体工事の準備で\n残置物撤去が必要"], ["遺品整理会社", "FaBoxOpen", "残置物の撤去・処分"], ["解体業者", "FaHardHat", "解体工事を実施"]];
  const PAL = ["撤去を\n依頼", "完了\n報告"];
  for (let i = 0; i < 3; i++) {
    const x = 0.4 + i * 1.4;
    box(x, 5.42, 1.0, 1.1, WHITE, "F2C4A0");
    txt(PA[i][0], x, 5.44, 1.0, 0.24, { size: 8.5, bold: true, align: "center" });
    await img(PA[i][1], ORANGE, x + 0.31, 5.7, 0.38);
    txt(PA[i][2], x, 6.1, 1.0, 0.4, { size: 6.5, align: "center" });
    if (i < 2) {
      ln(x + 1.04, 6.0, 0.32, 0, ORANGE, { w: 2, arrow: true });
      txt(PAL[i], x + 1.0, 5.55, 0.4, 0.42, { size: 6, bold: true, color: ORANGE, align: "center" });
    }
  }
  box(0.4, 6.68, 3.85, 0.26, "FCE4D2", "F2C4A0");
  txt("解体業者の前工程として残置物撤去（遺品整理）を請け負うケース", 0.4, 6.68, 3.85, 0.26, { size: 7.5, bold: true, color: ORANGE, align: "center" });

  // ---------- パターンB ----------
  box(4.55, 5.0, 5.2, 2.05, "F5F8FC", BLUE, { lw: 1.25 });
  txt("パターンB：遺品整理を元請で受注し、解体まで一括対応する流れ（元請けの流れ）", 4.6, 5.12, 5.1, 0.24, { size: 7.5, bold: true, color: BLUE, align: "center" });
  const PB = [["お客様（ご遺族等）", "FaUserFriends", "遺品整理の相談・依頼"], ["遺品整理会社", "FaBoxOpen", "遺品整理を実施"], ["解体業者", "FaHardHat", "解体工事を実施"], ["不動産会社等", "FaHome", "売却・活用提案"]];
  const PBL = ["依頼\n（元請）", "必要に\n応じて\n発注", "連携"];
  for (let i = 0; i < 4; i++) {
    const x = 4.65 + i * 1.35;
    box(x, 5.42, 0.95, 1.1, WHITE, "A9BEDD");
    txt(PB[i][0], x, 5.44, 0.95, 0.24, { size: i === 0 ? 7 : 8.5, bold: true, align: "center" });
    await img(PB[i][1], i === 0 ? GREEN : BLUE, x + 0.29, 5.7, 0.38);
    txt(PB[i][2], x, 6.1, 0.95, 0.4, { size: 6.5, align: "center" });
    if (i < 3) {
      ln(x + 0.99, 6.0, 0.32, 0, BLUE, { w: 2, arrow: true });
      txt(PBL[i], x + 0.95, 5.5, 0.4, 0.48, { size: 5.5, bold: true, color: BLUE, align: "center" });
    }
  }
  box(4.85, 6.68, 4.6, 0.26, "E3ECF8", "A9BEDD");
  txt("遺品整理を起点に、解体・売却・活用までワンストップで提供できるケース", 4.85, 6.68, 4.6, 0.26, { size: 7.5, bold: true, color: NAVY, align: "center" });

  await pres.writeFile({ fileName: OUT });
  const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
  await applyTheme(OUT, { name: "Ihin Seiri Proposal", headFontFace: FONT, bodyFontFace: FONT, colors: { dk1: "000000", lt1: "FFFFFF", dk2: NAVY, lt2: "F3F4F6", accent1: NAVY, accent2: BLUE, accent3: "D7261E", accent4: ORANGE, accent5: GREEN, accent6: "5B3F8C", hlink: BLUE, folHlink: "5B3F8C" } });
  console.log("written:", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
