// REVIVAL LINK様 2回目商談 ご提案書 空パッケージ生成スクリプト
// いただいた骨子（P.1〜P.15）＋個社別ご提案パートのひな形を、先方に刺さる順に並べた枠だけのデッキ（中身は「（記入）」の仮置き、骨子の文言は発表者ノート）
// 使い方: node build_revival2_empty.js <出力パス.pptx>
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa");

const OUT = process.argv[2] || "../REVIVAL_LINK様_2回目商談_ご提案書_空パッケージ.pptx";
const PG = process.argv[3];
const ASSET = path.join(__dirname, "assets", "kikaku");

const NAVY = "1F2D50", NAVY2 = "2B3F6E", BLUE = "2E5FA8", LBLUE = "EAF0F8", RED = "D7261E", YEL = "FFF000";
const GREEN = "2E7D4F", LGREEN = "EAF4EE", ORANGE = "E07B2E", LORANGE = "FDF0E4", PURPLE = "5B3F8C";
const GRAY = "444444", MGRAY = "8A8F99", LGRAY = "F3F4F6", LINE = "C9D1DE", DARK = "222222", WHITE = "FFFFFF";
const FONT = "Meiryo";
const W = 10.833, H = 7.5;
const MODEL = "※船井総研作成のモデル数値です。成果を保証するものではありません。";

const ICON_CACHE = {};
async function prepIcons(list) {
  for (const [name, color] of list) {
    const key = name + "_" + color;
    if (ICON_CACHE[key]) continue;
    const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(FA[name], { color: "#" + color, size: 256 }));
    ICON_CACHE[key] = "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
  }
}

async function main() {
  const ICONS = { home: "FaHome", building: "FaBuilding", tools: "FaTools", bed: "FaBed", wifi: "FaWifi", paper: "FaNewspaper", box: "FaBoxOpen",
    search: "FaSearchLocation", yen: "FaYenSign", users: "FaUsers", law: "FaBalanceScale", shake: "FaHandshake", key: "FaKey", hammer: "FaHammer",
    chart: "FaChartLine", map: "FaMapMarkedAlt", store: "FaStore", ok: "FaCheckCircle", recycle: "FaRecycle", link: "FaLink", truck: "FaTruck" };
  const pairs = [];
  Object.values(ICONS).forEach((n) => [WHITE, NAVY, BLUE, MGRAY, RED, GREEN, ORANGE].forEach((c) => pairs.push([n, c])));
  await prepIcons(pairs);
  const IC = (k, c) => ICON_CACHE[ICONS[k] + "_" + c];

  const pres = new pptxgen();
  pres.defineLayout({ name: "KIKAKU", width: W, height: H });
  pres.layout = "KIKAKU";
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "リバイバルリンク株式会社様 ご提案書";
  pres.author = "株式会社船井総合研究所";
  pres.company = "株式会社船井総合研究所";
  pres.defineSlideMaster({ title: "CONTENT", background: { path: path.join(ASSET, "bg_content.jpg") }, objects: [],
    slideNumber: { x: 10.1, y: 7.16, w: 0.65, h: 0.28, fontFace: FONT, fontSize: 9, color: WHITE, align: "center" } });
  pres.defineSlideMaster({ title: "COVER", background: { path: path.join(ASSET, "bg_cover.jpg") }, objects: [] });
  pres.defineSlideMaster({ title: "DIVIDER", background: { color: NAVY }, objects: [] });
  pres.defineSlideMaster({ title: "IMAGE", background: { color: WHITE }, objects: [] });

  let oid = 0;
  const nm = (p) => `${p}_${++oid}`;
  function rect(s, x, y, w, h, fill, o = {}) {
    s.addShape(o.round ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE, { x, y, w, h, fill: fill ? { color: fill } : { type: "none" },
      line: o.line ? { color: o.line, width: o.lw || 1, dashType: o.dash || "solid" } : { type: "none" }, rectRadius: o.round ? (o.r || 0.08) : undefined, objectName: nm("box") });
  }
  function oval(s, x, y, d, fill, o = {}) { s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1 } : { type: "none" }, objectName: nm("oval") }); }
  function txt(s, t, x, y, w, h, o = {}) {
    if (Array.isArray(t)) {
      const out = [];
      t.forEach((r) => { const ps = r.text.split("\n"); ps.forEach((p, k) => { const op = Object.assign({ align: o.align || "left" }, r.options); if (k < ps.length - 1) op.breakLine = true; if (p !== "" || k < ps.length - 1) out.push({ text: p, options: op }); }); });
      t = out;
    }
    s.addText(t, { x, y, w, h, fontFace: o.font || FONT, fontSize: o.size || 12, bold: !!o.bold, color: o.color || DARK, align: o.align || "left", valign: o.valign || "top",
      margin: o.margin === undefined ? 0 : o.margin, lineSpacingMultiple: o.ls || 1.15, isTextBox: true, fit: "none", objectName: nm("text") });
  }
  const B = (t, o = {}) => ({ text: t, options: Object.assign({ bold: true }, o) });
  const N = (t, o = {}) => ({ text: t, options: Object.assign({ bold: false }, o) });
  function icon(s, k, x, y, d, color, bg) {
    if (bg) { oval(s, x, y, d, bg); const p = d * 0.24; s.addImage({ data: IC(k, color), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: nm("icon") }); }
    else s.addImage({ data: IC(k, color), x, y, w: d, h: d, objectName: nm("icon") });
  }
  function num(s, n, x, y, d, fill) { oval(s, x, y, d, fill || NAVY); txt(s, String(n), x, y, d, d, { size: Math.round(d * 28), bold: true, color: WHITE, align: "center", valign: "middle" }); }
  function arrowR(s, x, y, w, h, c) { s.addShape(pres.shapes.RIGHT_ARROW, { x, y, w, h, fill: { color: c || RED }, line: { type: "none" }, objectName: nm("arrow") }); }
  function chev(s, x, y, w, h, fill, t, first, size) {
    s.addShape(first ? pres.shapes.PENTAGON : pres.shapes.CHEVRON, { x, y, w, h, fill: { color: fill }, line: { type: "none" }, objectName: nm("chev") });
    txt(s, t, x + (first ? 0.05 : h * 0.4), y, w - (first ? h * 0.45 : h * 0.8), h, { size: size || 11, bold: true, color: WHITE, align: "center", valign: "middle" });
  }
  function line(s, x, y, w, h, c, wd, o = {}) { s.addShape(pres.shapes.LINE, { x, y, w, h, flipV: !!o.flipV, flipH: !!o.flipH, line: { color: c || LINE, width: wd || 1, dashType: o.dash || "solid" }, objectName: nm("line") }); }
  function page(header, head, sub, notes) {
    const s = pres.addSlide({ masterName: "CONTENT" });
    txt(s, header, 0.2, 0.17, 6.5, 0.55, { size: 20, color: WHITE, valign: "middle" });
    txt(s, head, 0.35, 0.98, 10.13, 0.45, { size: head.length > 30 ? 16 : head.length > 25 ? 18 : 21, bold: true, color: NAVY, align: "center", valign: "middle" });
    if (sub) txt(s, sub, 0.35, 1.45, 10.13, 0.34, { size: sub.length > 52 ? 11 : sub.length > 44 ? 12 : 13, bold: true, color: DARK, align: "center", valign: "middle" });
    if (notes) s.addNotes(notes);
    return s;
  }
  function band(s, t, y, h, size) {
    rect(s, 0.35, y, 10.13, h || 0.62, NAVY);
    txt(s, t, 0.35, y, 10.13, h || 0.62, { size: size || 15, bold: true, color: YEL, align: "center", valign: "middle" });
  }
  function note(s, t, y) { txt(s, t, 0.35, y, 10.13, 0.25, { size: 8, color: MGRAY, valign: "middle" }); }
  function imagePage(p) {
    const f = path.join(PG, `p-${String(p).padStart(2, "0")}.jpg`);
    if (!fs.existsSync(f)) throw new Error("missing page image: " + f);
    const s = pres.addSlide({ masterName: "IMAGE" });
    s.addImage({ path: f, x: 0, y: 0, w: W, h: H, objectName: `kikaku_p${p}` });
    s.addNotes(`企画書（PDF）${p}ページをそのまま使用。`);
  }
  const HEADER = "リバイバルリンク株式会社様へのご提案";
  const PH = "8A8F99";
  const ph = (s, t, x, y, w, h, o = {}) => txt(s, t, x, y, w, h, Object.assign({ size: 11, color: PH }, o));
  function fbox(s, x, y, w, h, title, c, o = {}) {
    rect(s, x, y, w, h, WHITE, { line: c, lw: 1.5 });
    rect(s, x, y, w, o.hh || 0.5, c);
    txt(s, title, x + 0.12, y, w - 0.24, o.hh || 0.5, { size: o.hs || 13, bold: true, color: WHITE, valign: "middle", align: o.align || "left" });
  }
  function bullets3(s, x, y, w, n, size) {
    ph(s, Array.from({ length: n }, (_, i) => N("・（記入）" + (i < n - 1 ? "\n" : ""))), x, y, w, 0.4 * n, { size: size || 11.5, ls: 1.6 });
  }
  function divider(n, t, sub) {
    const s = pres.addSlide({ masterName: "DIVIDER" });
    txt(s, `第${n}章`, 0.5, 2.2, 9.83, 0.7, { size: 22, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, t, 0.5, 2.95, 9.83, 1.1, { size: 34, bold: true, color: YEL, align: "center", valign: "middle" });
    txt(s, sub, 0.5, 4.15, 9.83, 0.5, { size: 15, color: WHITE, align: "center", valign: "middle" });
  }

  {
    const s = pres.addSlide({ masterName: "COVER" });
    txt(s, "リバイバルリンク株式会社様\nご提案書", 0.8, 2.3, 9.2, 1.9, { size: 34, bold: true, color: NAVY, align: "center", valign: "middle", ls: 1.25 });
    txt(s, "遺品整理を起点にした、売主直の仕入れモデルのご提案", 0.8, 4.2, 9.2, 0.5, { size: 16, bold: true, color: DARK, align: "center", valign: "middle" });
    txt(s, "株式会社船井総合研究所", 0.8, 5.6, 9.2, 0.4, { size: 14, color: DARK, align: "center", valign: "middle" });
    s.addNotes("表紙。日付・担当者名は必要に応じて追加。");
  }

  divider(1, "前回のふりかえり", "貴社の現状・目指す姿・課題");

  // 2 貴社の現状
  {
    const s = page(HEADER, "貴社の現状", "（記入：貴社の現状をひと言で）", "【記入】HP・会社概要・ヒアリングの内容をもとに記入。売上構成比が分からない場合は事業名だけでよい。");
    fbox(s, 0.45, 2.0, 5.6, 3.3, "事業構成", NAVY);
    for (let i = 0; i < 4; i++) {
      const y = 2.65 + i * 0.65;
      rect(s, 0.65, y, 3.9, 0.52, LGRAY, { round: true, r: 0.06 });
      ph(s, `事業${i + 1}：（記入）`, 0.8, y, 3.7, 0.52, { valign: "middle", size: 12 });
      rect(s, 4.65, y, 1.2, 0.52, LBLUE, { round: true, r: 0.06 });
      ph(s, "○○%", 4.65, y, 1.2, 0.52, { valign: "middle", align: "center", size: 13, bold: true });
    }
    fbox(s, 6.3, 2.0, 4.1, 3.3, "貴社の強み", BLUE);
    for (let i = 0; i < 3; i++) {
      const y = 2.65 + i * 0.85;
      num(s, i + 1, 6.48, y + 0.17, 0.4, BLUE);
      ph(s, "（記入）", 7.0, y, 3.25, 0.75, { valign: "middle", size: 12 });
    }
    [["商圏", "（記入：本社所在地・対応エリア）"], ["体制", "（記入：従業員数・車両・拠点）"], ["許認可", "（記入：宅建・古物商・産廃など）"]].forEach(([k, v], i) => {
      const y = 5.5 + i * 0.47;
      rect(s, 0.45, y, 1.5, 0.4, NAVY);
      txt(s, k, 0.45, y, 1.5, 0.4, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
      rect(s, 1.95, y, 8.45, 0.4, WHITE, { line: LINE });
      ph(s, v, 2.1, y, 8.2, 0.4, { valign: "middle", size: 11.5 });
    });
  }
  // 3 目的
  {
    const s = page(HEADER, "遺品整理事業に取り組む目的（貴社の目指す姿）", "（記入：ヒアリングで伺った目的をひと言で）", "【記入】ヒアリングで伺った言葉をなるべくそのまま使う。「実現できたときの状態」の見出しは内容に合わせて書き換える。");
    rect(s, 0.45, 2.0, 9.95, 1.7, LBLUE, { line: NAVY, lw: 1.5 });
    txt(s, "貴社が遺品整理事業に取り組む目的", 0.65, 2.1, 9.5, 0.4, { size: 13, bold: true, color: NAVY, valign: "middle" });
    ph(s, "（記入：目的を1文で）", 0.65, 2.55, 9.5, 1.0, { size: 20, bold: true, align: "center", valign: "middle" });
    s.addShape(pres.shapes.DOWN_ARROW, { x: 5.1, y: 3.8, w: 0.6, h: 0.4, fill: { color: RED }, line: { type: "none" } });
    txt(s, "目的を実現できたときの状態", 0.45, 4.25, 9.95, 0.35, { size: 13, bold: true, color: NAVY, align: "center", valign: "middle" });
    ["（見出し：例 売上）", "（見出し：例 本業）", "（見出し：例 組織）"].forEach((t, i) => {
      const x = 0.45 + i * 3.35;
      fbox(s, x, 4.7, 3.15, 2.0, t, [NAVY, BLUE, GREEN][i], { hh: 0.48, align: "center" });
      ph(s, "（記入）", x + 0.15, 5.25, 2.85, 1.35, { size: 12, valign: "middle", align: "center" });
    });
  }
  {
    const s = page(HEADER, "リバイバルリンク様の現状課題と構造的ボトルネック", "前回の面談で共有いただいた4大課題の構造的要因を整理。",
      "【いただいた骨子 P.3】\n1. 財務：融資1.6億円・在庫8,000万円により、手元現金が拘束され新規仕入れが停滞。\n2. エリア：都内拠点にもかかわらず、地盤のない北関東で他社が回らないエリアを泥臭く業者回り。\n3. 組織：仕入れが営業マン個人の腕・足に依存し、属人化と独立リスクを抱えている。\n4. 既存事業：堅実に黒字化している民泊事業（年商1億）をさらに伸ばす仕入れルートが不足。\n【トーク】「認識に相違はありませんか？」と確認しながら進める。");
    const hd = ["課題", "現状（前回伺った内容）", "構造的ボトルネック"];
    const cw = [1.6, 4.2, 4.33], cx = [0.35, 1.95, 6.15];
    hd.forEach((h, j) => { rect(s, cx[j], 2.0, cw[j] - 0.05, 0.45, NAVY); txt(s, h, cx[j], 2.0, cw[j] - 0.05, 0.45, { size: 12.5, bold: true, color: WHITE, align: "center", valign: "middle" }); });
    ["財務", "エリア", "組織", "既存事業（民泊）"].forEach((k, i) => {
      const y = 2.55 + i * 1.05;
      rect(s, cx[0], y, cw[0] - 0.05, 0.95, LBLUE);
      num(s, i + 1, cx[0] + 0.1, y + 0.27, 0.4, NAVY);
      txt(s, k, cx[0] + 0.55, y, cw[0] - 0.6, 0.95, { size: 12, bold: true, color: NAVY, valign: "middle" });
      rect(s, cx[1], y, cw[1] - 0.05, 0.95, WHITE, { line: LINE });
      ph(s, "（記入）", cx[1] + 0.15, y, cw[1] - 0.3, 0.95, { valign: "middle" });
      rect(s, cx[2], y, cw[2] - 0.05, 0.95, LGRAY);
      ph(s, "（記入）", cx[2] + 0.15, y, cw[2] - 0.3, 0.95, { valign: "middle" });
    });
    band(s, "（記入：4大課題をひと言で）", 6.8 - 0.15, 0.5, 14);
  }

  divider(2, "遺品整理に参入する「真の目的」", "売主直の仕入れと、貴社事業とのシナジー");

  {
    const s = page(HEADER, "本提案の趣旨：遺品整理事業に参入する「真の目的」", "遺品整理は単なる片付け業ではなく、「不動産・民泊物件を市場流通前の最上流（売主直）で自動仕入れする装置」である。",
      "【いただいた骨子 P.1】\n【左側：従来の商流】業者回り・出物待ち ➔ 競合多数・中途半端な仕入れ・価格競争。\n【右側：本モデルの商流】遺族からの「実家の片付け相談（Web）」 ➔ 競合ゼロで空き家・不動産・残置物を直接捕捉。\n【結論】「片付けで即日利益」＋「売主直の優良物件をノーリスク選別」。\n【トーク】前回の「不用品回収は売買が決まった後。遺品整理はその手前で取れる」、社長ご自身の実例（遺品整理業者の紹介で2,300万円→約1か月で3,500万円で売却）を起点にする。");
    [["従来の商流", MGRAY, LGRAY, 0.45], ["本モデルの商流", NAVY, LBLUE, 5.75]].forEach(([t, c, bg, x]) => {
      rect(s, x, 2.05, 4.65, 3.75, bg, { line: c, lw: 1.5 });
      rect(s, x, 2.05, 4.65, 0.52, c);
      txt(s, t, x, 2.05, 4.65, 0.52, { size: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
      [["入口", 2.75], ["物件と出会うタイミング", 3.75], ["結果", 4.75]].forEach(([k, y]) => {
        txt(s, k, x + 0.2, y, 4.25, 0.3, { size: 11, bold: true, color: c, valign: "middle" });
        rect(s, x + 0.2, y + 0.32, 4.25, 0.58, WHITE, { line: LINE, round: true, r: 0.06 });
        ph(s, "（記入）", x + 0.35, y + 0.32, 4.0, 0.58, { valign: "middle" });
      });
    });
    arrowR(s, 5.15, 3.6, 0.55, 0.6, RED);
    band(s, "（記入：結論）", 6.0, 0.7, 15);
  }

  {
    const s = page(HEADER, "リバイバルリンク様が目指すべき事業シナジー全体像", "フロント（遺品整理）で日銭を作り、バックエンド（買取再販・仲介・民泊）へ利益を流し込む三位一体モデル。",
      "【いただいた骨子 P.2】\nフロント：遺品整理・生前整理・ゴミ屋敷清掃（粗利率60%超の即時キャッシュ）。\nバックエンド①（不動産）：売主直の仲介（ノーリスク高粗利） または 厳選買取再販。\nバックエンド②（民泊）：格安空き家仕入れ ＋ 家具・古物リユースによる低コスト展開。");
    fbox(s, 0.45, 2.1, 3.9, 3.9, "フロント：遺品整理", NAVY, { hh: 0.6, hs: 15, align: "center" });
    icon(s, "box", 1.95, 2.9, 0.9, NAVY);
    bullets3(s, 0.7, 4.0, 3.45, 3, 12);
    [["バックエンド①：不動産", BLUE, 2.1, "building"], ["バックエンド②：民泊", GREEN, 4.15, "bed"]].forEach(([t, c, y, k]) => {
      fbox(s, 5.95, y, 4.45, 1.85, t, c, { hh: 0.5, hs: 14 });
      icon(s, k, 6.15, y + 0.75, 0.6, c);
      bullets3(s, 6.95, y + 0.65, 3.35, 2, 12);
      arrowR(s, 4.5, y + 0.65, 1.3, 0.5, RED);
      ph(s, "（流れるもの）", 4.45, y + 0.25, 1.45, 0.35, { size: 9.5, align: "center", valign: "middle" });
    });
    band(s, "（記入：三位一体モデルをひと言で）", 6.2, 0.62, 15);
  }

  // 5 つながり図
  {
    const s = page(HEADER, "遺品整理が、貴社の事業とつながる", "（記入：つながりをひと言で）", "【記入】貴社の事業を最大6つ。使わない枠は削除する。つながりが弱い事業は線を点線にする。");
    const CX = 5.42, CY = 4.45;
    const N6 = [[-3.6, -1.6], [3.6, -1.6], [-3.6, 1.55], [3.6, 1.55], [0, -2.0], [0, 2.05]];
    N6.forEach(([dx, dy], i) => {
      const bw = 2.9, bh = 0.85, x = CX + dx - bw / 2, y = CY + dy - bh / 2;
      const x1 = CX + (dx === 0 ? 0.001 : Math.sign(dx) * 0.9), y1 = CY + (dy > 0 ? 0.55 : -0.55) * (dx === 0 ? 1.72 : 1);
      const x2 = dx === 0 ? CX : (dx > 0 ? x : x + bw), y2 = dx === 0 ? (dy > 0 ? y : y + bh) : y + bh / 2;
      line(s, Math.min(x1, x2), Math.min(y1, y2), Math.abs(x2 - x1) || 0.001, Math.abs(y2 - y1) || 0.001, BLUE, 2, { flipV: (x2 - x1) * (y2 - y1) < 0 });
      rect(s, x, y, bw, bh, WHITE, { line: BLUE, lw: 1.5, round: true, r: 0.08 });
      ph(s, [B(`（事業${i + 1}）\n`, { fontSize: 11.5, color: DARK }), N("（つながり方）", { fontSize: 9.5 })], x + 0.15, y, bw - 0.3, bh, { valign: "middle", ls: 1.15 });
    });
    oval(s, CX - 0.95, CY - 0.95, 1.9, NAVY);
    txt(s, "遺品整理\n（新規事業）", CX - 0.95, CY - 0.5, 1.9, 1.0, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle", ls: 1.05 });
  }
  // 6 シナジー（何を共有するかで4つに分ける。向きはどちらでもよい）
  {
    const s = page(HEADER, "貴社ならではのシナジー", "（記入：シナジーをひと言で）", "【記入】4つの観点のうち、貴社に当てはまるものだけ残して使う。本業→遺品整理、遺品整理→本業のどちらの向きでも、同じ観点に書いてよい。業種ごとの当てはまり方は最後の「参考：業種別シナジー早見表」を参照。");
    const G = [["users", "① LTVの最大化", "同じお客様に、本業と遺品整理の両方を提供できる", "送客・紹介、既存顧客・会員の活用、クロスセル", NAVY],
      ["truck", "② 経営資源", "人・車両・ノウハウ・拠点を、両方の事業で使える", "空いている稼働の活用、設備投資を抑えた立ち上げ", BLUE],
      ["yen", "③ 収益・コスト", "お金の流れが良くなる", "外注の内製化、客単価・粗利の向上、集客費の共通化", GREEN],
      ["store", "④ 信頼・ブランド", "地域での立ち位置が強くなる", "専門店化、直接受注・脱下請け、ワンストップでの差別化", ORANGE]];
    G.forEach(([k, t, d, ex, c], i) => {
      const x = 0.45 + (i % 2) * 5.02, y = 1.95 + Math.floor(i / 2) * 2.5;
      rect(s, x, y, 4.88, 2.35, WHITE, { line: c, lw: 1.5 });
      rect(s, x, y, 4.88, 0.55, c);
      icon(s, k, x + 0.12, y + 0.08, 0.4, WHITE);
      txt(s, t, x + 0.62, y, 4.1, 0.55, { size: 14, bold: true, color: WHITE, valign: "middle" });
      txt(s, d, x + 0.18, y + 0.6, 4.55, 0.3, { size: 10.5, bold: true, color: c, valign: "middle" });
      txt(s, "例：" + ex, x + 0.18, y + 0.88, 4.55, 0.26, { size: 9, color: GRAY, valign: "middle" });
      line(s, x + 0.18, y + 1.18, 4.52, 0.001, LINE, 0.75, { dash: "dash" });
      ph(s, [B("（見出し）\n", { fontSize: 13, color: DARK }), N("（記入：貴社での具体的な内容）", { fontSize: 11 })], x + 0.18, y + 1.25, 4.52, 1.05, { ls: 1.4 });
    });
  }
  // 6-2 シナジー創出スキーム（大枠）
  {
    const s = page(HEADER, "シナジーを生むスキーム", "（記入：スキームをひと言で）", "【記入】横軸はお客様の流れ（接点→相談→遺品整理→整理後→その後）。各レーンの枠に、そのタイミングで誰が何をするかを書く。送客の矢印は向き・本数を貴社に合わせて動かす・消す。使わないレーン・枠は削除してよい。");
    const L0 = 1.75, CW = 1.72, X0 = 1.75;
    const COLS = ["① 接点", "② 相談", "③ 遺品整理", "④ 整理後", "⑤ その後"];
    COLS.forEach((c, i) => {
      s.addShape(i === 0 ? pres.shapes.PENTAGON : pres.shapes.CHEVRON, { x: X0 + i * CW, y: 1.95, w: CW + 0.08, h: 0.42, fill: { color: i === 2 ? BLUE : NAVY }, line: { type: "none" } });
      txt(s, c, X0 + i * CW + 0.15, 1.95, CW - 0.25, 0.42, { size: 11.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    });
    const LANES = [["貴社\n本業", NAVY, LBLUE], ["貴社\n遺品整理事業", BLUE, "EEF3FB"], ["提携先・\n外注先", ORANGE, LORANGE]];
    LANES.forEach(([t, c, bg], i) => {
      const y = 2.5 + i * 1.3;
      rect(s, 0.35, y, 1.32, 1.18, c);
      txt(s, t, 0.35, y, 1.32, 1.18, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle", ls: 1.15 });
      rect(s, 1.72, y, 8.76, 1.18, bg);
      for (let k = 0; k < 5; k++) {
        rect(s, X0 + k * CW + 0.07, y + 0.1, CW - 0.14, 0.98, WHITE, { line: c, round: true, r: 0.06 });
        ph(s, [B("（誰が・何を）\n", { fontSize: 10, color: DARK }), N("（記入）", { fontSize: 9.5 })], X0 + k * CW + 0.15, y + 0.15, CW - 0.3, 0.88, { ls: 1.3, align: "center", valign: "middle" });
      }
    });
    // 送客の矢印（例）
    [[X0 + 1 * CW + CW / 2, 3.6, 0.32, "送客"], [X0 + 3 * CW + CW / 2, 3.6, 0.32, "送客"]].forEach(([cx, y, h, t], i) => {
      s.addShape(i === 0 ? pres.shapes.DOWN_ARROW : pres.shapes.UP_ARROW, { x: cx - 0.17, y: y - 0.02, w: 0.34, h: 0.42, fill: { color: RED }, line: { type: "none" } });
      txt(s, t, cx + 0.2, y, 0.6, 0.35, { size: 9, bold: true, color: RED, valign: "middle" });
    });
    rect(s, 0.35, 6.45, 10.13, 0.5, LGRAY);
    txt(s, [B("スキームのポイント　", { color: NAVY }), N("（記入：このスキームで生まれる効果。例：本業のお客様を遺品整理へ送客し、整理後の案件を本業で受注する）", { color: PH })], 0.5, 6.45, 9.9, 0.5, { size: 11, valign: "middle" });
  }
  divider(3, "貴社の課題への解決策", "財務・エリア・組織・民泊");

  {
    const s = page(HEADER, "解決策①【財務・エリア】：在庫リスクゼロで即日キャッシュ化＆都内近郊の反響獲得", "現金を寝かせず日銭を作り、足を使った飛び込みから「Web反響仕入れ」へシフト。",
      "【いただいた骨子 P.4】\n先出しキャッシュゼロ：片付け完了時に代金回収（粗利約18万円/件）。在庫リスクなし。\n都内オフィスの利点を活用：Web・MEOにより「都内在住で実家（都内・近郊・北関東）の処分に困っている売主」を直接集客。\n【トーク】実家が北関東でも、相続する子どもは都内に住んでいることが多い。");
    [["yen", "財務：先出しキャッシュゼロ", NAVY, 0.45], ["map", "エリア：都内オフィスの利点を活用", BLUE, 5.5]].forEach(([k, t, c, x]) => {
      fbox(s, x, 2.0, 4.9, 4.3, "", c, { hh: 0.6 });
      icon(s, k, x + 0.15, 2.08, 0.44, WHITE);
      txt(s, t, x + 0.7, 2.0, 4.1, 0.6, { size: 14, bold: true, color: WHITE, valign: "middle" });
      [["Before", MGRAY, LGRAY], ["After", RED, "FDECEC"]].forEach(([b, bc, bg], i) => {
        const y = 2.8 + i * 0.85;
        rect(s, x + 0.2, y, 0.95, 0.65, bc);
        txt(s, b, x + 0.2, y, 0.95, 0.65, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle" });
        rect(s, x + 1.15, y, 3.55, 0.65, bg);
        ph(s, "（記入）", x + 1.3, y, 3.3, 0.65, { valign: "middle" });
      });
      line(s, x + 0.2, 4.6, 4.5, 0.001, LINE, 0.75, { dash: "dash" });
      bullets3(s, x + 0.25, 4.75, 4.4, 3, 11.5);
    });
    band(s, "（記入：ひと言で）", 6.45, 0.55, 14);
  }

  {
    const s = page(HEADER, "解決策②【組織体制】：属人化を排除する「組織営業（0次接客・即時見積）」", "営業マンの力量や独立に左右されず、誰でも成約率80%を叩き出す型を導入。",
      "【いただいた骨子 P.5】\n0次接客（電話）：5分以内の初回コールでヒアリング・写真送付誘導。\n即時見積Excelシート：現場調査時、平米・間取り・荷物量を入れるだけでその場で見積提示・即決クロージング。\n【トーク】前回の「営業1人で月1件、育っても独立した方が得になる」への回答。");
    const ST = ["問い合わせ", "0次接客（電話）", "現地調査・即時見積", "即決クロージング"];
    ST.forEach((t, i) => chev(s, 0.45 + i * 2.5, 2.05, 2.55, 0.6, i === 1 || i === 2 ? NAVY : BLUE, t, i === 0, 12));
    ST.forEach((t, i) => {
      const x = 0.45 + i * 2.5;
      rect(s, x + 0.05, 2.85, 2.35, 2.9, WHITE, { line: LINE });
      txt(s, "使う仕組み", x + 0.15, 2.92, 2.15, 0.3, { size: 10.5, bold: true, color: NAVY, valign: "middle" });
      ph(s, "（記入）", x + 0.15, 3.25, 2.15, 0.6, { size: 12, bold: true, valign: "middle" });
      line(s, x + 0.15, 3.9, 2.15, 0.001, LINE, 0.75, { dash: "dash" });
      bullets3(s, x + 0.15, 4.0, 2.15, 3, 10.5);
    });
    band(s, "（記入：誰でも成果が出る型をひと言で）", 6.05, 0.65, 15);
  }

  {
    const s = page(HEADER, "解決策③【民泊連携】：主力事業「民泊（年商1億）」との強力な相乗効果", "他社には絶対に真似できない「空き家×民泊転用×家具リユース」の独自モデル。",
      "【いただいた骨子 P.6】\n物件仕入れ：売却に困る空き家・古民家を安価に買取・サブリースし、インバウンド民泊へ転用。\nコスト削減：遺品整理で出る良質な家具・昭和レトロ調度品を民泊インテリアとして再活用。\n独自提案：「売りたくない」遺族に対し、「民泊運用代行」という第三の出口を提示して他社を排除。\n【トーク】前回の「堅実に拡大できるのは民泊。こちらを軸足にしたい」につなげる。");
    [["home", "物件仕入れ", NAVY], ["recycle", "コスト削減", BLUE], ["shake", "独自提案", GREEN]].forEach(([k, t, c], i) => {
      const x = 0.45 + i * 3.35;
      fbox(s, x, 2.0, 3.15, 2.45, "", c, { hh: 0.55 });
      icon(s, k, x + 0.15, 2.07, 0.4, WHITE);
      txt(s, t, x + 0.65, 2.0, 2.4, 0.55, { size: 14, bold: true, color: WHITE, valign: "middle" });
      ph(s, "（記入）", x + 0.18, 2.7, 2.8, 1.6, { size: 12, valign: "middle", align: "center" });
      s.addShape(pres.shapes.DOWN_ARROW, { x: x + 1.3, y: 4.55, w: 0.55, h: 0.4, fill: { color: RED }, line: { type: "none" } });
    });
    rect(s, 0.45, 5.05, 9.95, 0.85, NAVY, { round: true, r: 0.1 });
    icon(s, "bed", 0.7, 5.22, 0.5, WHITE);
    txt(s, "民泊（年商1億）", 1.35, 5.05, 3, 0.85, { size: 17, bold: true, color: WHITE, valign: "middle" });
    ph(s, "（記入：民泊にとっての効果）", 4.2, 5.05, 6.0, 0.85, { size: 12, color: "C9D1DE", valign: "middle" });
    band(s, "（記入：独自モデルをひと言で）", 6.1, 0.6, 15);
  }

  {
    const s = page(HEADER, "4大課題と、遺品整理で解決できること", "（記入：課題と、遺品整理で解決できることをひと言で）", "【ひな形4】前回伺った4大課題（P.3）と、解決策①〜③の対応をまとめる。");
    txt(s, "貴社の課題", 0.45, 1.9, 4.5, 0.35, { size: 14, bold: true, color: MGRAY, valign: "middle" });
    txt(s, "遺品整理で解決できること", 5.9, 1.9, 4.5, 0.35, { size: 14, bold: true, color: NAVY, valign: "middle" });
    ["財務", "エリア", "組織", "既存事業（民泊）"].forEach((k, i) => {
      const y = 2.32 + i * 1.12;
      rect(s, 0.45, y, 4.6, 0.98, LGRAY);
      icon(s, "search", 0.6, y + 0.12, 0.32, MGRAY);
      ph(s, [B(k + "\n", { fontSize: 13, color: DARK }), N("（記入：内容）", { fontSize: 10.5 })], 1.1, y + 0.08, 3.8, 0.85, { ls: 1.3 });
      arrowR(s, 5.2, y + 0.27, 0.55, 0.42, RED);
      rect(s, 5.9, y, 4.5, 0.98, LBLUE, { line: NAVY });
      num(s, i + 1, 6.05, y + 0.12, 0.36, NAVY);
      ph(s, [B("（解決策）\n", { fontSize: 13, color: NAVY }), N("（記入：内容）", { fontSize: 10.5 })], 6.55, y + 0.08, 3.75, 0.85, { ls: 1.3 });
    });
  }

  divider(4, "事業収支シミュレーション", "フロント収益・バックエンド収益・3年間の数値計画");

  {
    const s = page(HEADER, "【フロント収益】遺品整理事業 単体PLモデル", "月5〜10件の稼働で、年間1,000万〜2,000万円規模の安定粗利を確立。",
      "【いただいた骨子 P.7】\n平均客単価：30万円（家一棟丸ごと・高単価案件中心）\n売上原価：提携処分費・作業外注費（約40%）\n売上総利益（粗利）：約60%（18万円/件）\n月間5件稼働時：月商150万円／月粗利90万円\n月間10件稼働時：月商300万円／月粗利180万円\n【トーク】前回の「今は残置物撤去に1軒約40万円払っている」につなげる。");
    const cell = (t, i, j, hdr) => ({ text: t, options: { fontFace: FONT, fontSize: 12, valign: "middle", align: j === 0 ? "left" : "center", bold: hdr || j === 0, color: hdr ? WHITE : (j === 0 ? NAVY : PH), fill: { color: hdr ? NAVY : (j === 0 ? LBLUE : WHITE) }, border: { type: "solid", pt: 0.75, color: LINE } } });
    txt(s, "1件あたり", 0.45, 1.95, 4.4, 0.35, { size: 13, bold: true, color: NAVY, valign: "middle" });
    s.addTable([["項目", "金額・比率"], ["平均客単価", "○○万円"], ["売上原価", "○○%"], ["売上総利益（粗利）", "○○%（○○万円）"]].map((r, i) => r.map((c, j) => cell(c, i, j, i === 0))), { x: 0.45, y: 2.35, w: 4.4, colW: [2.2, 2.2], rowH: 0.62, objectName: "pl_unit" });
    txt(s, "稼働件数別", 5.25, 1.95, 5.15, 0.35, { size: 13, bold: true, color: NAVY, valign: "middle" });
    s.addTable([["", "月間5件稼働時", "月間10件稼働時"], ["月商", "○○万円", "○○万円"], ["月粗利", "○○万円", "○○万円"], ["年間粗利", "○○万円", "○○万円"]].map((r, i) => r.map((c, j) => cell(c, i, j, i === 0))), { x: 5.25, y: 2.35, w: 5.15, colW: [1.55, 1.8, 1.8], rowH: 0.62, objectName: "pl_volume" });
    rect(s, 0.45, 5.1, 9.95, 0.85, LGRAY);
    ph(s, "（記入：補足）", 0.65, 5.1, 9.6, 0.85, { valign: "middle" });
    band(s, "（記入：年間粗利の規模をひと言で）", 6.1, 0.6, 15);
    note(s, MODEL, 6.82);
  }

  {
    const s = page(HEADER, "【バックエンド収益】不動産仲介・買取・民泊への波及効果", "受託10件に1件が不動産・民泊案件化。ハズレ在庫を抱えず利益最大化。",
      "【いただいた骨子 P.8】\n年間受託想定：80件 ➔ 不動産派生：年間4〜8件（5〜10%）\n選択肢A（仲介）：手元資金ゼロで仲介手数料（1件あたり100万〜300万円）を獲得。\n選択肢B（厳選買取）：大幅指値が通る超優良物件のみを選定して自社買取（回転率重視）。\n選択肢C（民泊転用）：好立地・古民家は民泊物件として自社ポートフォリオへ組み入れ。\n【トーク】選択肢Bの例として社長ご自身の実例（2,300万円→3,500万円）を添える。");
    [["年間受託想定", "○○件", NAVY, 1.2], ["不動産派生", "年間○○件（○○%）", RED, 6.0]].forEach(([t, v, c, x]) => {
      rect(s, x, 2.0, 3.6, 1.1, WHITE, { line: c, lw: 1.5, round: true, r: 0.08 });
      txt(s, [B(t + "\n", { fontSize: 12, color: GRAY }), B(v, { fontSize: 20, color: PH })], x, 2.0, 3.6, 1.1, { align: "center", valign: "middle", ls: 1.2 });
    });
    arrowR(s, 4.95, 2.3, 0.9, 0.5, RED);
    [["A", "仲介", NAVY, "shake"], ["B", "厳選買取", BLUE, "key"], ["C", "民泊転用", GREEN, "bed"]].forEach(([a, t, c, k], i) => {
      const x = 0.45 + i * 3.35;
      fbox(s, x, 3.35, 3.15, 2.55, "", c, { hh: 0.55 });
      icon(s, k, x + 0.15, 3.42, 0.4, WHITE);
      txt(s, `選択肢${a}（${t}）`, x + 0.65, 3.35, 2.4, 0.55, { size: 14, bold: true, color: WHITE, valign: "middle" });
      ph(s, "（記入）", x + 0.18, 4.05, 2.8, 1.7, { size: 12, valign: "middle", align: "center" });
    });
    band(s, "（記入：ひと言で）", 6.1, 0.6, 15);
    note(s, MODEL, 6.82);
  }

  {
    const s = page(HEADER, "3年間の数値計画", "（記入：計画をひと言で）", "【ひな形12】数字は貴社の体制・商圏に合わせて記入。参考：企画書の3年モデル（売上3,600万→5,400万→1億円）。立上げスケジュールは「立ち上げ6ヶ月ロードマップ」に記載。");
    const T = [["", "1年目", "2年目", "3年目"], ["売上", "○○万円", "○○万円", "○○万円"], ["月間契約件数", "○○件", "○○件", "○○件"], ["営業利益率", "○○%", "○○%", "○○%"], ["体制", "○名", "○名", "○名"]];
    const data = T.map((r, i) => r.map((c, j) => ({ text: c, options: { fontFace: FONT, fontSize: 13, valign: "middle", align: "center", bold: i === 0 || j === 0, color: i === 0 ? WHITE : (j === 0 ? NAVY : PH), fill: { color: i === 0 ? NAVY : (j === 0 ? LBLUE : WHITE) }, border: { type: "solid", pt: 0.75, color: LINE } } })));
    s.addTable(data, { x: 0.45, y: 2.0, w: 6.4, colW: [1.9, 1.5, 1.5, 1.5], rowH: 0.7, objectName: "plan" });
    rect(s, 7.05, 2.0, 3.35, 3.5, WHITE, { line: GREEN, lw: 1.5 });
    rect(s, 7.05, 2.0, 3.35, 0.5, GREEN);
    txt(s, "本業への波及（上乗せ）", 7.05, 2.0, 3.35, 0.5, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    bullets3(s, 7.25, 2.7, 3.0, 5, 11.5);
    band(s, "（記入：結論のひと言）", 5.9, 0.62, 15);
    note(s, MODEL, 6.7);
  }

  divider(5, "立ち上げの不安を解消する", "過去のFCとの違い・リソース・外注スキーム");

  {
    const s = page(HEADER, "過去のFC参入経験（エコパートナーズ）との決定的な違い", "営業マンの査定センス頼みだった過去の失敗要因を、完全な「仕組み」で克服。",
      "【いただいた骨子 P.11】\n【過去のモデル】：安く買い叩く査定スキル依存／不用品回収中心／営業マン次第で失速。\n【今回のモデル】：定額パック・即時見積ツールによる明朗会計／家一棟・不動産狙い／組織営業化。\n【トーク】前回のお話（2023年頃にFC参入、本部の集客で1日2〜3件アポ、買取は営業マン次第、半年未満で撤退）を受けて違いを示す。");
    rect(s, 2.25, 2.0, 3.95, 0.55, MGRAY);
    txt(s, "過去のモデル", 2.25, 2.0, 3.95, 0.55, { size: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
    rect(s, 6.45, 2.0, 3.95, 0.55, NAVY);
    txt(s, "今回のモデル", 6.45, 2.0, 3.95, 0.55, { size: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
    ["価格の決め方", "主な商品", "営業", "（比較の軸）"].forEach((k, i) => {
      const y = 2.7 + i * 0.85;
      rect(s, 0.45, y, 1.75, 0.75, LBLUE);
      txt(s, k, 0.45, y, 1.75, 0.75, { size: 12, bold: true, color: NAVY, align: "center", valign: "middle" });
      rect(s, 2.25, y, 3.95, 0.75, LGRAY);
      ph(s, "（記入）", 2.4, y, 3.7, 0.75, { valign: "middle" });
      arrowR(s, 6.22, y + 0.2, 0.2, 0.35, RED);
      rect(s, 6.45, y, 3.95, 0.75, WHITE, { line: NAVY });
      ph(s, "（記入）", 6.6, y, 3.7, 0.75, { valign: "middle", bold: true });
    });
    band(s, "（記入：決定的な違いをひと言で）", 6.2, 0.62, 15);
  }

  {
    const s = page(HEADER, "立ち上げの障壁①：リソース不足（トラック・作業員・許認可）の解消", "自社でトラックや現場作業員を抱える必要は一切ありません。",
      "【いただいた骨子 P.9】\n三者間契約モデル：集客・現調・見積もりは自社、実作業・処分は地元の一般廃棄物・産廃許可業者へ委託。\n完全合法かつ固定費ゼロでスモールスタートが可能。\n【トーク】前回のご質問「トラックは必要か」への回答。前回のFCでは社員がけがをしたが、作業は委託できる。");
    // 三者間契約の図
    const P = [["お客様", 2.6, 2.05, MGRAY, "users"], ["貴社", 0.55, 4.45, NAVY, "building"], ["許可業者", 4.65, 4.45, ORANGE, "truck"]];
    line(s, 1.5, 2.95, 1.5, 1.5, LINE, 2, { flipH: true });
    line(s, 3.6, 2.95, 1.5, 1.5, LINE, 2);
    line(s, 2.4, 4.95, 2.25, 0.001, LINE, 2);
    P.forEach(([t, x, y, c, k]) => {
      rect(s, x, y, 1.85, 0.95, c, { round: true, r: 0.1 });
      icon(s, k, x + 0.12, y + 0.27, 0.4, WHITE);
      txt(s, t, x + 0.55, y, 1.25, 0.95, { size: 13, bold: true, color: WHITE, valign: "middle" });
    });
    ph(s, "（契約）", 1.0, 3.4, 1.2, 0.3, { size: 9.5, align: "center" });
    ph(s, "（契約）", 3.95, 3.4, 1.2, 0.3, { size: 9.5, align: "center" });
    ph(s, "（依頼）", 2.85, 5.0, 1.4, 0.3, { size: 9.5, align: "center" });
    txt(s, "三者間契約モデル", 0.45, 5.65, 6.1, 0.3, { size: 11, bold: true, color: NAVY, align: "center", valign: "middle" });
    fbox(s, 6.85, 2.05, 3.55, 1.8, "自社でやること", NAVY, { hh: 0.48, align: "center" });
    bullets3(s, 7.0, 2.65, 3.3, 3, 11.5);
    fbox(s, 6.85, 4.05, 3.55, 1.8, "委託すること", ORANGE, { hh: 0.48, align: "center" });
    bullets3(s, 7.0, 4.65, 3.3, 3, 11.5);
    band(s, "（記入：ひと言で）", 6.15, 0.62, 15);
  }

  // 6-3 外注スキーム（請負部分）
  {
    const s = page(HEADER, "遺品整理の外注スキーム", "（記入：自社が担う工程と、外注する工程をひと言で）", "【記入】工程（列）ごとに、自社と外注先のどちらが担うかを書く。担わないレーンの枠は消す。工程は貴社に合わせて増減してよい。上のモデル案件は、総量×単価＝売上、外注費、粗利（粗利率）を記入。本業・関連事業のレーンは、遺品整理の前後で本業につながる動きを書く（不要なら削除）。");
    rect(s, 0.35, 1.9, 1.35, 0.42, NAVY);
    txt(s, "モデル案件", 0.35, 1.9, 1.35, 0.42, { size: 11.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    rect(s, 1.7, 1.9, 6.0, 0.42, WHITE, { line: LINE });
    ph(s, [N("総量 ○○㎥ × 単価 ○○円/㎥ → ", { color: DARK }), B("売上 ○○万円", { color: RED }), N("　｜　外注費 ○○万円", { color: DARK })], 1.85, 1.9, 5.8, 0.42, { size: 11, valign: "middle" });
    rect(s, 7.8, 1.9, 2.68, 0.42, RED);
    txt(s, "粗利 ○○万円（粗利率○○%）", 7.8, 1.9, 2.68, 0.42, { size: 11.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    const STEPS = ["集客", "現地調査", "営業・見積", "仕分け指示", "現場作業", "処分", "完了確認"];
    const X0 = 1.75, CW = 1.247;
    STEPS.forEach((t, i) => {
      rect(s, X0 + i * CW + 0.03, 2.45, CW - 0.06, 0.38, "D9E2F0", { round: true, r: 0.08 });
      txt(s, `${["①", "②", "③", "④", "⑤", "⑥", "⑦"][i]} ${t}`, X0 + i * CW + 0.03, 2.45, CW - 0.06, 0.38, { size: 10, bold: true, color: NAVY, align: "center", valign: "middle" });
    });
    const LANES = [["自社\n遺品整理事業", NAVY, "EEF3FB"], ["外注\n（専門業者）", ORANGE, LORANGE], ["自社\n本業・関連事業", GREEN, LGREEN]];
    LANES.forEach(([t, c, bg], i) => {
      const y = 2.95 + i * 1.12;
      rect(s, 0.35, y, 1.35, 1.02, c);
      txt(s, t, 0.35, y, 1.35, 1.02, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle", ls: 1.15 });
      rect(s, 1.72, y, 8.76, 1.02, bg);
      for (let k = 0; k < 7; k++) {
        rect(s, X0 + k * CW + 0.05, y + 0.08, CW - 0.1, 0.86, WHITE, { line: c, round: true, r: 0.06 });
        ph(s, "（作業内容）\n（金額）", X0 + k * CW + 0.1, y + 0.12, CW - 0.2, 0.78, { size: 9, align: "center", valign: "middle", ls: 1.3 });
      }
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 6.33, w: 2.0, h: 0.62, fill: { color: NAVY }, line: { type: "none" } });
    txt(s, "スキームのポイント", 0.35, 6.33, 2.0, 0.62, { size: 11.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    for (let k = 0; k < 3; k++) {
      const x = 2.45 + k * 2.68;
      rect(s, x, 6.33, 2.6, 0.62, WHITE, { line: LINE });
      icon(s, "ok", x + 0.08, 6.49, 0.3, ORANGE);
      ph(s, "（記入）", x + 0.45, 6.33, 2.1, 0.62, { size: 10, valign: "middle" });
    }
  }
  divider(6, "商圏・競合と事業コンセプト", "勝てる市場か、どう勝つか");

  // 7 商圏
  {
    const s = page(HEADER, "貴社商圏の市場", "（記入：市場規模をひと言で）", "【記入】人口・世帯数は自治体の住民基本台帳、65歳以上の割合は自治体の統計、空き家数は総務省「住宅・土地統計調査」。市場規模＝人口×MS（国民1人あたりの遺品整理消費額）4,200円の船井総研モデル。");
    ["人口", "世帯数", "65歳以上の割合"].forEach((h, i) => {
      const x = 0.45 + i * 3.35;
      rect(s, x, 2.0, 3.15, 1.55, LGRAY, { round: true, r: 0.06 });
      rect(s, x, 2.0, 3.15, 0.42, NAVY, { round: true, r: 0.12 });
      txt(s, "○○市の" + h, x, 2.0, 3.15, 0.42, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
      ph(s, ["○○万人", "○○万世帯", "○○%"][i], x, 2.45, 3.15, 0.65, { size: 26, bold: true, align: "center", valign: "middle" });
      ph(s, "（○年○月○日現在）", x, 3.1, 3.15, 0.35, { size: 9.5, align: "center", valign: "middle" });
    });
    rect(s, 0.45, 3.8, 9.95, 2.1, WHITE, { line: NAVY, lw: 1.25 });
    rect(s, 0.45, 3.8, 9.95, 0.45, NAVY);
    txt(s, "遺品整理の市場規模（試算）", 0.45, 3.8, 9.95, 0.45, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    [["人口", "○○万人", PH], ["MS（1人あたり消費額）", "4,200円", NAVY], ["市場規模", "約○○億円", PH]].forEach(([t, v, c], i) => {
      const x = 0.85 + i * 3.25;
      txt(s, [B(t + "\n", { fontSize: 11, color: GRAY }), B(v, { fontSize: 22, color: c })], x, 4.4, 2.6, 1.0, { align: "center", valign: "middle", ls: 1.2 });
      if (i < 2) txt(s, i === 0 ? "×" : "＝", x + 2.6, 4.4, 0.65, 1.0, { size: 22, bold: true, color: GRAY, align: "center", valign: "middle" });
    });
    ph(s, "（記入：補足。例 空き家数○○戸、3年目の年商1億円はシェア約○%）", 0.6, 5.4, 9.65, 0.4, { size: 11.5, bold: true, align: "center", valign: "middle" });
    band(s, "（記入：結論のひと言）", 6.1, 0.6);
    note(s, "出典：（記入）。市場規模は人口×MS4,200円で算出した船井総研のモデル数値。", 6.82);
  }
  // 8 ポジショニング
  {
    const s = page(HEADER, "競合状況", "（記入：競合の状況をひと言で）", "【記入】商圏で「遺品整理 ○○市」「不用品回収 ○○市」をGoogle検索・Googleマップで調べた上位の会社を配置する。札は動かして使い、不要な札は削除する。");
    const X0 = 0.9, Y0 = 2.15, W = 9.0, Hh = 4.45;
    rect(s, X0, Y0, W, Hh, WHITE, { line: "C9D6E6" });
    line(s, X0 + W / 2, Y0, 0.001, Hh, "C9D6E6", 1);
    line(s, X0, Y0 + Hh / 2, W, 0.001, "C9D6E6", 1);
    txt(s, "遺品整理・高単価", X0, Y0 - 0.35, W, 0.3, { size: 11, bold: true, color: GRAY, align: "center", valign: "middle" });
    txt(s, "不用品回収・低単価", X0, Y0 + Hh + 0.05, W, 0.3, { size: 11, bold: true, color: GRAY, align: "center", valign: "middle" });
    txt(s, "地域\n密着", 0.25, Y0 + Hh / 2 - 0.3, 0.6, 0.6, { size: 11, bold: true, color: GRAY, align: "center", valign: "middle" });
    txt(s, "広域\n対応", 9.95, Y0 + Hh / 2 - 0.3, 0.6, 0.6, { size: 11, bold: true, color: GRAY, align: "center", valign: "middle" });
    const P = [[1.9, 2.55, RED, "推奨ポジション"], [4.6, 2.5, NAVY], [7.4, 2.75, NAVY], [3.1, 3.35, NAVY], [6.2, 3.55, NAVY], [8.3, 4.6, "2E8B8B"], [3.6, 4.95, "2E8B8B"], [5.6, 5.25, BLUE], [7.6, 5.75, BLUE], [1.5, 5.7, BLUE], [4.3, 6.0, BLUE]];
    P.forEach(([x, y, c, t], i) => {
      rect(s, x, y, 1.55, 0.36, c);
      txt(s, t || "（競合名）", x, y, 1.55, 0.36, { size: 10.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    });
  }
  // 9 競合一覧
  {
    const s = page(HEADER, "競合一覧", "（記入：競合の特徴をひと言で）", "【記入】周辺サービスには、遺品整理以外の対応（生前整理・買取・供養・特殊清掃・解体・不動産・相続など）を書く。");
    const hdr = ["会社・ブランド", "主な打ち出し", "公開価格", "周辺サービス", "弱点・示唆"];
    const rows = [hdr, ...Array.from({ length: 8 }, () => ["（記入）", "（記入）", "（記入）", "（記入）", "（記入）"])];
    const data = rows.map((r, i) => r.map((c) => ({ text: c, options: { fontFace: FONT, fontSize: i === 0 ? 11.5 : 10.5, bold: i === 0, valign: "middle", color: i === 0 ? WHITE : PH, fill: { color: i === 0 ? NAVY : (i % 2 ? WHITE : LGRAY) }, border: { type: "solid", pt: 0.75, color: LINE } } })));
    s.addTable(data, { x: 0.35, y: 2.0, w: 10.13, colW: [1.75, 2.45, 1.4, 2.3, 2.23], rowH: 0.5, objectName: "competitors" });
  }
  // 10 勝ち筋
  {
    const s = page(HEADER, "競合調査から見えた勝ち筋", "（記入：勝ち筋をひと言で）", "【記入】黒字＝競合の状況、赤字＝貴社の打ち手。4組まで。");
    for (let i = 0; i < 4; i++) {
      const y = 2.0 + i * 1.15;
      rect(s, 0.45, y, 9.95, 1.0, i % 2 ? WHITE : LGRAY);
      txt(s, "・（記入：競合の状況）", 0.65, y + 0.08, 9.6, 0.42, { size: 13, bold: true, color: DARK, valign: "middle" });
      txt(s, "→（記入：貴社の打ち手）", 0.95, y + 0.5, 9.3, 0.42, { size: 13, bold: true, color: RED, valign: "middle" });
    }
  }
  {
    const s = page(HEADER, "立ち上げの障壁②：Web集客・競合への不安解消", "大手ポータルや大手買取再販業者と真正面から戦わない局地戦戦略。",
      "【いただいた骨子 P.10】\nエリア・キーワードの絞り込み：「都内・近郊の家一棟片付け」「ゴミ屋敷」「実家じまい」など高単価ロングテールに特化。\nGoogleビジネスプロフィール（MEO）対策：地域一番店の口コミ・マップ上位表示を獲得。\n【トーク】前回のご質問「大手が入ってこない理由」につなげる。");
    [["search", "エリア・キーワードの絞り込み", NAVY, 0.45], ["map", "Googleビジネスプロフィール（MEO）対策", BLUE, 5.5]].forEach(([k, t, c, x]) => {
      fbox(s, x, 2.0, 4.9, 3.85, "", c, { hh: 0.6 });
      icon(s, k, x + 0.15, 2.08, 0.44, WHITE);
      txt(s, t, x + 0.7, 2.0, 4.1, 0.6, { size: 13.5, bold: true, color: WHITE, valign: "middle" });
      ph(s, "（記入）", x + 0.2, 2.75, 4.5, 0.9, { size: 13, bold: true, valign: "middle", align: "center" });
      line(s, x + 0.2, 3.75, 4.5, 0.001, LINE, 0.75, { dash: "dash" });
      bullets3(s, x + 0.25, 3.9, 4.4, 4, 11.5);
    });
    band(s, "（記入：局地戦戦略をひと言で）", 6.05, 0.65, 15);
  }

  // 11 事業コンセプト
  {
    const s = page(HEADER, "貴社版の事業コンセプト", "（記入：コンセプトをひと言で）", "【記入】屋号は本業と分けるかどうかも含めて記入。");
    [["users", "売り先（誰に）", NAVY], ["box", "売り物（何を）", BLUE], ["chart", "売り方（どう売る）", GREEN]].forEach(([k, t, c], i) => {
      const x = 0.45 + i * 3.35;
      rect(s, x, 2.0, 3.15, 3.15, WHITE, { line: c, lw: 1.5 });
      rect(s, x, 2.0, 3.15, 0.55, c);
      icon(s, k, x + 0.15, 2.07, 0.4, WHITE);
      txt(s, t, x + 0.6, 2.0, 2.5, 0.55, { size: 14, bold: true, color: WHITE, valign: "middle" });
      bullets3(s, x + 0.18, 2.75, 2.85, 3, 12);
    });
    rect(s, 0.45, 5.35, 9.95, 1.25, "FFFBE6", { line: "C9A227", dash: "dash" });
    txt(s, [B("屋号・コンセプト案　", { color: NAVY })], 0.65, 5.45, 3, 0.4, { size: 12.5, valign: "middle" });
    ph(s, "（記入：屋号案・キャッチコピー・本業との関係）", 0.65, 5.85, 9.6, 0.65, { size: 12, valign: "middle" });
  }
  divider(7, "立ち上げの進め方とご支援", "ロードマップ・支援内容・事例・まとめ");

  {
    const s = page(HEADER, "立ち上げ6ヶ月ロードマップ", "6ヶ月で「提携構築 ➔ HP完成 ➔ 集客・営業開始 ➔ 案件受注」までを計画的に完遂。",
      "【いただいた骨子 P.12】\n1〜2ヶ月目：商圏競合分析、商品・料金設計、提携処分業者（三者間契約）の開拓。\n3〜4ヶ月目：遺品整理専門HP制作ディレクション、営業ツール（Excel見積・トーク集）導入。\n5〜6ヶ月目：Web広告（リスティング・MEO）テスト運用開始、初回商談同行・改善。\n【トーク】前回のご質問「何から始めればよいか」への回答。");
    ["提携構築", "HP完成", "集客・営業開始", "案件受注"].forEach((t, i) => chev(s, 0.45 + i * 2.5, 1.95, 2.55, 0.5, [NAVY, NAVY2, BLUE, RED][i], t, i === 0, 12));
    const X0 = 2.35, MW = 1.34;
    rect(s, 0.45, 2.65, 9.95, 0.4, NAVY);
    for (let m = 0; m < 6; m++) txt(s, `${m + 1}ヶ月目`, X0 + m * MW, 2.65, MW, 0.4, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle" });
    ["（項目）", "（項目）", "（項目）", "（項目）", "（項目）", "（項目）", "（項目）"].forEach((t, i) => {
      const y = 3.1 + i * 0.42;
      rect(s, 0.45, y, 9.95, 0.4, i % 2 ? WHITE : LGRAY);
      ph(s, t, 0.55, y, 1.8, 0.4, { size: 10.5, bold: true, valign: "middle" });
      const st = Math.floor(i / 2.4) * 2;
      rect(s, X0 + st * MW + 0.05, y + 0.07, MW * 2 - 0.1, 0.26, "D5DAE3", { round: true, r: 0.05 });
      ph(s, "（内容）", X0 + st * MW + 0.05, y + 0.07, MW * 2 - 0.1, 0.26, { size: 9, align: "center", valign: "middle", color: GRAY });
    });
    band(s, "（記入：ひと言で）", 6.2, 0.6, 15);
  }

  {
    const s = page(HEADER, "船井総研の提供価値（支援パッケージ詳細）", "戦略立案からツール提供、現場の営業指導までワンストップで伴走。",
      "【いただいた骨子 P.13】\n1. 商品設計 ＋ HP制作ディレクション（約150万円）：勝てるコンセプト策定、Web制作会社ディレクション。\n2. マニュアル・オリジナルツール導入支援（100万円）：0次接客スクリプト、即時見積Excel、アライアンス資料。\n3. 業績アップ月次伴走コンサルティング（20万円/月）：月次定例MTG、進捗管理、広告・営業KPI改善。");
    ["商品設計 ＋ HP制作ディレクション", "マニュアル・オリジナルツール導入支援", "業績アップ月次伴走コンサルティング"].forEach((t, i) => {
      const y = 2.0 + i * 1.4;
      rect(s, 0.45, y, 9.95, 1.25, WHITE, { line: [NAVY, BLUE, GREEN][i], lw: 1.5 });
      rect(s, 0.45, y, 0.9, 1.25, [NAVY, BLUE, GREEN][i]);
      txt(s, String(i + 1), 0.45, y, 0.9, 1.25, { size: 28, bold: true, color: WHITE, align: "center", valign: "middle" });
      txt(s, t, 1.55, y + 0.1, 5.8, 0.45, { size: 15, bold: true, color: [NAVY, BLUE, GREEN][i], valign: "middle" });
      ph(s, "（記入：内容）", 1.55, y + 0.58, 5.8, 0.55, { size: 11.5, valign: "middle" });
      rect(s, 7.6, y + 0.3, 2.6, 0.65, LGRAY, { round: true, r: 0.1 });
      ph(s, "○○万円", 7.6, y + 0.3, 2.6, 0.65, { size: 17, bold: true, align: "center", valign: "middle" });
    });
    band(s, "（記入：ひと言で）", 6.3, 0.6, 15);
  }

  {
    const s = page(HEADER, "全国の成功事例（不動産業×遺品整理）", "同様の課題を抱えていた企業が、短期間で事業立ち上げと業績アップを実現。",
      "【いただいた骨子 P.14】\n事例A（地方不動産）：参入1年で片付け月8件受託、年間5件の媒介・買取案件を獲得し営業利益1,500万円増。\n事例B（買取再販）：即時見積ツールの導入により営業マンの成約率が35%から78%へ大幅向上。");
    ["事例A（○○）", "事例B（○○）"].forEach((t, i) => {
      const x = 0.45 + i * 5.05;
      fbox(s, x, 2.0, 4.9, 4.0, t, [NAVY, BLUE][i], { hh: 0.6, hs: 15, align: "center" });
      [["企業の概要", 2.75], ["取り組み", 3.55]].forEach(([k, y]) => {
        txt(s, k, x + 0.2, y, 4.5, 0.3, { size: 11, bold: true, color: [NAVY, BLUE][i], valign: "middle" });
        ph(s, "（記入）", x + 0.2, y + 0.3, 4.5, 0.45, { valign: "middle" });
      });
      rect(s, x + 0.2, 4.45, 4.5, 1.4, "FDECEC", { round: true, r: 0.08 });
      txt(s, "成果", x + 0.35, 4.5, 2, 0.3, { size: 11, bold: true, color: RED, valign: "middle" });
      ph(s, "（記入：大きな数字）", x + 0.35, 4.8, 4.2, 0.95, { size: 18, bold: true, align: "center", valign: "middle" });
    });
    band(s, "（記入：ひと言で）", 6.2, 0.6, 15);
  }

  {
    const s = page(HEADER, "まとめ・次回のアクション", "成功パターンの7割はすでに型化されています。残り3割を御社仕様に合わせて伴走します。",
      "【いただいた骨子 P.15】\n本日の提案骨子の総括。\n次回までの決定事項（社内体制の確定、支援開始月の決定）。\n「泥臭い業者回りと在庫リスクから脱却し、都内オフィスを活かした自社仕入れモデルへシフトしましょう」");
    fbox(s, 0.45, 2.0, 5.9, 3.7, "本日の提案骨子の総括", NAVY, { hh: 0.5 });
    for (let i = 0; i < 6; i++) { num(s, i + 1, 0.65, 2.68 + i * 0.5, 0.36, NAVY); ph(s, "（記入）", 1.15, 2.65 + i * 0.5, 5.0, 0.42, { valign: "middle" }); }
    fbox(s, 6.55, 2.0, 3.85, 3.7, "次回までの決定事項", RED, { hh: 0.5 });
    for (let i = 0; i < 3; i++) { rect(s, 6.8, 2.8 + i * 0.85, 0.32, 0.32, WHITE, { line: GRAY, lw: 1.25 }); ph(s, "（記入）", 7.25, 2.7 + i * 0.85, 3.0, 0.52, { valign: "middle", size: 12 }); }
    band(s, "（記入：締めの一文）", 5.95, 0.85, 16);
  }

  {
    const s = pres.addSlide({ masterName: "DIVIDER" });
    txt(s, "船井総研がご提供できること", 0.5, 2.6, 9.83, 1.2, { size: 34, bold: true, color: YEL, align: "center", valign: "middle" });
    txt(s, "（企画書の「船井総研がご提供できること」のページをここに差し込む）", 0.5, 3.9, 9.83, 0.6, { size: 14, color: WHITE, align: "center", valign: "middle" });
    s.addNotes("企画書の「船井総研がご提供できること」をそのまま最後に差し込む。");
  }

  await pres.writeFile({ fileName: OUT });
  console.log("written:", OUT);
}
main().catch((e) => { console.error(e); process.exit(1); });
