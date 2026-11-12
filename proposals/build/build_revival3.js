// REVIVAL LINK様 2回目商談 ご提案書（不動産・買取再販×遺品整理 版／肉付け）生成スクリプト。分からない箇所は「不明」（赤字）
// いただいた骨子（P.1〜P.15）＋個社別ご提案パートのひな形を、先方に刺さる順に並べた枠だけのデッキ（中身は「（記入）」の仮置き、骨子の文言は発表者ノート）
// 使い方: node build_revival2_empty.js <出力パス.pptx>
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa");

const OUT = process.argv[2] || "../REVIVAL_LINK様_2回目商談_ご提案書_不動産版.pptx";
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

  // 記入済みの文字（「不明」を含むものは赤で目立たせる）
  const isU = (t) => (typeof t === "string" ? t : JSON.stringify(t)).includes("不明");
  const fl = (s, t, x, y, w, h, o = {}) => txt(s, t, x, y, w, h, Object.assign({ size: 11.5 }, o, { color: isU(t) && typeof t === "string" ? RED : (o.color || DARK) }));
  const blist = (s, arr, x, y, w, size, ls) => txt(s, arr.map((t, i) => N("・" + t + (i < arr.length - 1 ? "\n" : ""), isU(t) ? { color: RED } : {})), x, y, w, 0.42 * arr.length, { size: size || 11.5, ls: ls || 1.5, color: DARK });
  const SRC = "前回のお話をもとに作成。";

  // 1 表紙
  {
    const s = pres.addSlide({ masterName: "COVER" });
    txt(s, "リバイバルリンク株式会社様\nご提案書", 0.8, 2.3, 9.2, 1.9, { size: 34, bold: true, color: NAVY, align: "center", valign: "middle", ls: 1.25 });
    txt(s, "遺品整理を起点にした、不動産・買取再販の仕入れモデルのご提案", 0.8, 4.2, 9.2, 0.5, { size: 16, bold: true, color: DARK, align: "center", valign: "middle" });
    txt(s, "株式会社船井総合研究所", 0.8, 5.6, 9.2, 0.4, { size: 14, color: DARK, align: "center", valign: "middle" });
    s.addNotes("表紙。日付・担当者名は必要に応じて追加。");
  }
  divider(1, "前回のふりかえり", "貴社の現状・目指す姿・課題");
  // 3 貴社の現状
  {
    const s = page(HEADER, "貴社の現状", "年商約2億円（不動産・民泊）。不動産は業者回りの仕入れで月1〜2件", SRC + "売上構成比は前回伺った概算。従業員数・車両・宅建業の免許は不明のため要確認。");
    fbox(s, 0.45, 2.0, 5.6, 3.3, "事業構成", NAVY);
    [["不動産（買取再販）：約1億円", "約50%"], ["民泊：約1億円（単体で黒字）", "約50%"], ["通信：ほぼ稼働なし", "不明"], ["ポスティング：ほぼ稼働なし", "不明"]].forEach(([a, b], i) => {
      const y = 2.65 + i * 0.65;
      rect(s, 0.65, y, 3.9, 0.52, LGRAY, { round: true, r: 0.06 });
      fl(s, a, 0.8, y, 3.7, 0.52, { valign: "middle", size: 12 });
      rect(s, 4.65, y, 1.2, 0.52, LBLUE, { round: true, r: 0.06 });
      fl(s, b, 4.65, y, 1.2, 0.52, { valign: "middle", align: "center", size: 13, bold: true });
    });
    fbox(s, 6.3, 2.0, 4.1, 3.3, "貴社の強み", BLUE);
    ["買取から再販までを自社で完結（仕入れ・販売とも月1〜2件）", "古物商許可を保有。残置物撤去・買取の事業経験がある", "民泊の運営ノウハウ（単体で黒字）"].forEach((t, i) => {
      const y = 2.65 + i * 0.85;
      num(s, i + 1, 6.48, y + 0.17, 0.4, BLUE);
      fl(s, t, 7.0, y, 3.3, 0.75, { valign: "middle", size: 11 });
    });
    [["商圏", "本社：東京都世田谷区／物件の仕入れは北関東が中心"], ["体制", "不動産の営業担当1名（従業員数・車両：不明）"], ["許認可", "古物商（宅建業の免許：不明）"]].forEach(([k, v], i) => {
      const y = 5.5 + i * 0.47;
      rect(s, 0.45, y, 1.5, 0.4, NAVY);
      txt(s, k, 0.45, y, 1.5, 0.4, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
      rect(s, 1.95, y, 8.45, 0.4, WHITE, { line: LINE });
      txt(s, v.includes("不明") ? [N(v.split("（")[0], { color: DARK }), N("（" + v.split("（")[1], { color: RED })] : v, 2.1, y, 8.2, 0.4, { valign: "middle", size: 11.5, color: DARK });
    });
  }
  // 4 目的
  {
    const s = page(HEADER, "遺品整理事業に取り組む目的（貴社の目指す姿）", "「現金をあまり使わずに伸ばしたい」「空き家単体では強みがない」というお考えから", SRC + "前回の社長のお言葉：「現金をあまり使わずに伸ばせる事業」「空き家単体で儲けるのは難しい。自社は何を強みにするのか」「稼げるのは空き家の周りの事業」。");
    rect(s, 0.45, 2.0, 9.95, 1.7, LBLUE, { line: NAVY, lw: 1.5 });
    txt(s, "貴社が遺品整理事業に取り組む目的", 0.65, 2.1, 9.5, 0.4, { size: 13, bold: true, color: NAVY, valign: "middle" });
    txt(s, "現金を寝かせずに、売主直で良い物件を仕入れられる事業構造をつくる", 0.65, 2.55, 9.5, 1.0, { size: 20, bold: true, color: NAVY, align: "center", valign: "middle" });
    s.addShape(pres.shapes.DOWN_ARROW, { x: 5.1, y: 3.8, w: 0.6, h: 0.4, fill: { color: RED }, line: { type: "none" } });
    txt(s, "目的を実現できたときの状態", 0.45, 4.25, 9.95, 0.35, { size: 13, bold: true, color: NAVY, align: "center", valign: "middle" });
    [["財務", "在庫を持たずに、片付けで先に現金が入る"], ["仕入れ", "売主直で、良い物件だけを選んで仕入れられる"], ["組織", "営業マン個人に頼らず、仕組みで回る"]].forEach(([t, d], i) => {
      const x = 0.45 + i * 3.35;
      fbox(s, x, 4.7, 3.15, 2.0, t, [NAVY, BLUE, GREEN][i], { hh: 0.48, align: "center" });
      fl(s, d, x + 0.15, 5.25, 2.85, 1.35, { size: 13, bold: true, valign: "middle", align: "center" });
    });
  }
  // 5 現状課題（P.3）
  {
    const s = page(HEADER, "リバイバルリンク様の現状課題と構造的ボトルネック", "前回の面談で共有いただいた4大課題の構造的要因を整理。",
      "【トーク】「認識に相違はありませんか？」と確認しながら進める。\n4行目は、いただいた骨子の「民泊」を「仕入れの質」に差し替え。");
    const hd = ["課題", "現状（前回伺った内容）", "構造的ボトルネック"];
    const cw = [1.6, 4.2, 4.33], cx = [0.35, 1.95, 6.15];
    hd.forEach((h, j) => { rect(s, cx[j], 2.0, cw[j] - 0.05, 0.45, NAVY); txt(s, h, cx[j], 2.0, cw[j] - 0.05, 0.45, { size: 12.5, bold: true, color: WHITE, align: "center", valign: "middle" }); });
    const R = [["財務", "融資約1.6億円・在庫約8,000万円・現金約8,000万円", "手元現金が拘束され新規仕入れが停滞"],
      ["エリア", "事務所は都内なのに、物件は北関東で、他社が回らない地域の業者を回って仕入れている", "遠方で業者のつながりも限られ、件数を増やせない（社長：このやり方では広げられない）"],
      ["組織", "営業1名で月1件程度。育っても独立した方が得になる", "仕入れが営業マン個人の腕・足に依存。属人化と独立リスク"],
      ["仕入れの質", "業者からの出物待ち。実需寄りの中途半端な物件を買い、売れ残りがある", "物件を選べる立場にない（出物を取り合う）"]];
    R.forEach(([k, a, b], i) => {
      const y = 2.55 + i * 1.05;
      rect(s, cx[0], y, cw[0] - 0.05, 0.95, LBLUE);
      num(s, i + 1, cx[0] + 0.1, y + 0.27, 0.4, NAVY);
      txt(s, k, cx[0] + 0.55, y, cw[0] - 0.6, 0.95, { size: 12, bold: true, color: NAVY, valign: "middle" });
      rect(s, cx[1], y, cw[1] - 0.05, 0.95, WHITE, { line: LINE });
      txt(s, a.includes("不明") ? [N(a.split("（")[0], {}), N("（" + a.split("（")[1], { color: RED })] : a, cx[1] + 0.15, y, cw[1] - 0.3, 0.95, { valign: "middle", size: 11.5, color: DARK });
      rect(s, cx[2], y, cw[2] - 0.05, 0.95, LGRAY);
      fl(s, b, cx[2] + 0.15, y, cw[2] - 0.3, 0.95, { valign: "middle", bold: true });
    });
    band(s, "現金・エリア・人に頼る仕入れから抜け出せていない", 6.65, 0.5, 14);
  }
  divider(2, "遺品整理に参入する「真の目的」", "売主直の仕入れと、不動産・買取再販とのシナジー");
  // 7 真の目的（P.1）
  {
    const s = page(HEADER, "本提案の趣旨：遺品整理事業に参入する「真の目的」", "遺品整理は単なる片付け業ではなく、「不動産を市場流通前の最上流（売主直）で自動仕入れする装置」である。",
      "【トーク】前回の「不用品回収は売買が決まった後。遺品整理はその手前で取れる」、社長ご自身の実例（遺品整理業者の紹介で2,300万円で仕入れ、約1か月で3,500万円で売却）を起点にする。");
    const V = [["業者回り・出物待ち", "売却が決まった後（市場流通後）", "競合多数・中途半端な仕入れ・価格競争"], ["遺族からの「実家の片付け相談（Web）」", "売却を決める前（市場流通前・最上流）", "競合ゼロで空き家・不動産・残置物を直接捕捉"]];
    [["従来の商流", MGRAY, LGRAY, 0.45], ["本モデルの商流", NAVY, LBLUE, 5.75]].forEach(([t, c, bg, x], m) => {
      rect(s, x, 2.05, 4.65, 3.75, bg, { line: c, lw: 1.5 });
      rect(s, x, 2.05, 4.65, 0.52, c);
      txt(s, t, x, 2.05, 4.65, 0.52, { size: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
      [["入口", 2.75], ["物件と出会うタイミング", 3.75], ["結果", 4.75]].forEach(([k, y], j) => {
        txt(s, k, x + 0.2, y, 4.25, 0.3, { size: 11, bold: true, color: c, valign: "middle" });
        rect(s, x + 0.2, y + 0.32, 4.25, 0.58, WHITE, { line: LINE, round: true, r: 0.06 });
        fl(s, V[m][j], x + 0.35, y + 0.32, 4.0, 0.58, { valign: "middle", bold: m === 1, color: m === 1 ? NAVY : DARK });
      });
    });
    arrowR(s, 5.15, 3.6, 0.55, 0.6, RED);
    band(s, "「片付けで即日利益」＋「売主直の優良物件をノーリスク選別」", 6.0, 0.7, 16);
  }
  {
    const s = page(HEADER, "不用品回収と遺品整理の違い：出会うタイミングが違う", "不用品回収は売却が決まった後。遺品整理は、売るかどうかを決める前に出会える。",
      "【トーク】前回、社長が一番腹落ちされたポイント（「不用品回収は売買が決まった後。遺品整理はその手前で取れる」「不用品回収を混ぜるやり方は儲からない」）。");
    const ST = ["相続の発生", "実家の片付け", "売却・活用の検討", "売却の決定", "引き渡し前の片付け"];
    ST.forEach((t, i) => chev(s, 0.45 + i * 2.0, 2.1, 2.05, 0.6, i === 1 ? RED : (i === 4 ? MGRAY : NAVY), t, i === 0, 11.5));
    [["遺品整理", RED, 1, "売却を決める前（最上流）"], ["不用品回収", MGRAY, 4, "売却が決まった後"]].forEach(([t, c, k, d]) => {
      const x = 0.45 + k * 2.0;
      s.addShape(pres.shapes.DOWN_ARROW, { x: x + 0.75, y: 2.8, w: 0.5, h: 0.45, fill: { color: c }, line: { type: "none" } });
      rect(s, x - 0.2, 3.35, 2.45, 0.55, c, { round: true, r: 0.1 });
      txt(s, t, x - 0.2, 3.35, 2.45, 0.55, { size: 14, bold: true, color: WHITE, align: "center", valign: "middle" });
      txt(s, d, x - 0.3, 3.95, 2.65, 0.35, { size: 10.5, bold: true, color: c, align: "center", valign: "middle" });
    });
    [["遺品整理で出会うと", NAVY, LBLUE, 0.45, ["売るかどうかを決める前に、売主と直接話せる", "片付けの現地調査で、物件の状態を先に見られる"]], ["不用品回収で出会うと", MGRAY, LGRAY, 5.5, ["売却先が決まった後なので、仕入れにはつながらない", "1件の単価が小さく、価格競争になりやすい"]]].forEach(([t, c, bg, x, L]) => {
      rect(s, x, 4.5, 4.9, 1.45, bg, { line: c });
      txt(s, t, x + 0.2, 4.55, 4.5, 0.35, { size: 12.5, bold: true, color: c, valign: "middle" });
      blist(s, L, x + 0.25, 4.92, 4.5, 11.5);
    });
    band(s, "売主と最初に出会えるのは、不用品回収ではなく遺品整理", 6.15, 0.62, 15);
  }

  {
    const s = page(HEADER, "リバイバルリンク様が目指すべき事業シナジー全体像", "フロント（遺品整理）で日銭を作り、バックエンド（仲介・買取再販）へ利益を流し込むモデル。", "民泊は「その他の出口」としてついでに触れる程度の扱い。");
    fbox(s, 0.45, 2.1, 3.9, 3.5, "フロント：遺品整理", NAVY, { hh: 0.6, hs: 15, align: "center" });
    icon(s, "box", 1.95, 2.85, 0.85, NAVY);
    blist(s, ["遺品整理・生前整理・ゴミ屋敷清掃", "粗利率60%超の即時キャッシュ", "片付け完了時に代金回収"], 0.7, 3.9, 3.5, 12);
    [["バックエンド①：仲介", BLUE, 2.1, "shake", ["売主直の仲介（ノーリスク高粗利）", "手元資金を使わない"], "売却の相談"], ["バックエンド②：買取再販", RED, 3.85, "key", ["厳選買取再販", "指値が通る優良物件だけを買う"], "優良物件"]].forEach(([t, c, y, k, L, f]) => {
      fbox(s, 5.95, y, 4.45, 1.6, t, c, { hh: 0.5, hs: 14 });
      icon(s, k, 6.15, y + 0.7, 0.55, c);
      blist(s, L, 6.95, y + 0.6, 3.4, 11.5);
      arrowR(s, 4.5, y + 0.55, 1.3, 0.5, RED);
      fl(s, f, 4.4, y + 0.15, 1.55, 0.35, { size: 9.5, bold: true, align: "center", valign: "middle", color: RED });
    });
    rect(s, 0.45, 5.72, 9.95, 0.38, LGRAY, { round: true, r: 0.08 });
    txt(s, [B("その他の出口　", { color: GRAY }), N("好立地・古民家などは、民泊・賃貸での活用も提案できる（貴社の民泊事業へつなぐことも可能）", { color: DARK })], 0.65, 5.72, 9.6, 0.38, { size: 10.5, valign: "middle" });
    band(s, "フロントで日銭、バックエンドの仲介・買取再販で大きな利益", 6.25, 0.6, 15);
  }

  // 9 つながり図
  {
    const s = page(HEADER, "遺品整理が、貴社の事業とつながる", "遺品整理を入口に、仲介・買取再販を中心とした不動産事業につながる", "士業（相続・登記）との連携は未検討のため不明。民泊・賃貸はついでの出口。");
    const CX = 5.42, CY = 4.45;
    const N6 = [[-3.6, -1.6], [3.6, -1.6], [-3.6, 1.55], [3.6, 1.55], [0, -2.0], [0, 2.05]];
    const D = [["不動産仲介", "片付け後の売却相談を仲介"], ["買取再販", "指値の通る優良物件を厳選して買取"], ["古物買取", "家具・骨董品を買取、片付け費用と相殺"], ["リフォーム", "買取後の再生（既存の取引先へ外注）"], ["相続・登記", "士業との連携：不明"], ["民泊・賃貸", "ついでの出口：活用の提案"]];
    N6.forEach(([dx, dy], i) => {
      const bw = 2.9, bh = 0.85, x = CX + dx - bw / 2, y = CY + dy - bh / 2;
      const x1 = CX + (dx === 0 ? 0.001 : Math.sign(dx) * 0.9), y1 = CY + (dy > 0 ? 0.55 : -0.55) * (dx === 0 ? 1.72 : 1);
      const x2 = dx === 0 ? CX : (dx > 0 ? x : x + bw), y2 = dx === 0 ? (dy > 0 ? y : y + bh) : y + bh / 2;
      line(s, Math.min(x1, x2), Math.min(y1, y2), Math.abs(x2 - x1) || 0.001, Math.abs(y2 - y1) || 0.001, BLUE, 2, { flipV: (x2 - x1) * (y2 - y1) < 0, dash: i === 5 ? "dash" : "solid" });
      rect(s, x, y, bw, bh, WHITE, { line: BLUE, lw: 1.5, round: true, r: 0.08 });
      txt(s, [B(D[i][0] + "\n", { fontSize: 12, color: NAVY }), N(D[i][1], { fontSize: 10, color: isU(D[i][1]) ? RED : DARK })], x + 0.15, y, bw - 0.3, bh, { valign: "middle", ls: 1.15 });
    });
    oval(s, CX - 0.95, CY - 0.95, 1.9, NAVY);
    txt(s, "遺品整理\n（フロント）", CX - 0.95, CY - 0.5, 1.9, 1.0, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle", ls: 1.05 });
  }
  // 10 シナジー4観点
  {
    const s = page(HEADER, "貴社ならではのシナジー", "遺品整理と不動産・買取再販は、同じお客様・同じ物件でつながる", "");
    const G = [["users", "① LTVの最大化", "同じお客様に、本業と遺品整理の両方を提供できる", NAVY, "1件の片付けから、売却・活用まで", "片付けを依頼したご遺族に、仲介・買取再販まで提案できる"],
      ["truck", "② 経営資源", "人・車両・ノウハウ・拠点を、両方の事業で使える", BLUE, "今ある許可・拠点・ノウハウを使う", "古物商許可、都内オフィス、買取再販のノウハウ、リフォーム業者とのつながりを活かせる"],
      ["yen", "③ 収益・コスト", "お金の流れが良くなる", GREEN, "先に現金が入り、撤去費が売上に変わる", "完了時に代金回収。今は1軒約40万円払っている撤去費が自社の売上に。買取物件の残置物も自社で"],
      ["store", "④ 信頼・ブランド", "地域での立ち位置が強くなる", ORANGE, "売主直・ワンストップで選ばれる", "業者回りではなく、売主から直接相談が来る。片付けから売却まで一社で対応"]];
    G.forEach(([k, t, d, c, h, b], i) => {
      const x = 0.45 + (i % 2) * 5.02, y = 1.95 + Math.floor(i / 2) * 2.5;
      rect(s, x, y, 4.88, 2.35, WHITE, { line: c, lw: 1.5 });
      rect(s, x, y, 4.88, 0.55, c);
      icon(s, k, x + 0.12, y + 0.08, 0.4, WHITE);
      txt(s, t, x + 0.62, y, 4.1, 0.55, { size: 14, bold: true, color: WHITE, valign: "middle" });
      txt(s, d, x + 0.18, y + 0.6, 4.55, 0.3, { size: 10, bold: true, color: c, valign: "middle" });
      line(s, x + 0.18, y + 0.95, 4.52, 0.001, LINE, 0.75, { dash: "dash" });
      txt(s, [B(h + "\n", { fontSize: 13, color: DARK }), N(b, { fontSize: 11, color: DARK })], x + 0.18, y + 1.02, 4.52, 1.28, { ls: 1.35, color: DARK });
    });
  }
  // 11 シナジーを生むスキーム
  {
    const s = page(HEADER, "シナジーを生むスキーム", "Webの片付け相談から、物件ごとに仲介・買取再販の出口へつなぐ", "提携先からの紹介ルート（葬儀社・士業など）は未検討のため不明。");
    const CW = 1.72, X0 = 1.75;
    ["① 接点", "② 相談", "③ 遺品整理", "④ 整理後", "⑤ その後"].forEach((c, i) => {
      s.addShape(i === 0 ? pres.shapes.PENTAGON : pres.shapes.CHEVRON, { x: X0 + i * CW, y: 1.95, w: CW + 0.08, h: 0.42, fill: { color: i === 2 ? BLUE : NAVY }, line: { type: "none" } });
      txt(s, c, X0 + i * CW + 0.15, 1.95, CW - 0.25, 0.42, { size: 11.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    });
    const LANES = [["貴社\n本業", NAVY, LBLUE], ["貴社\n遺品整理事業", BLUE, "EEF3FB"], ["提携先・\n外注先", ORANGE, LORANGE]];
    const C = [[["―", ""], ["―", ""], ["物件の確認", "現調に同行"], ["出口の判断", "仲介・厳選買取"], ["再販", "リフォーム後に売却"]],
      [["Web・MEO", "「実家の片付け」の問い合わせ"], ["0次接客", "5分以内の電話・写真送付誘導"], ["即時見積・即決", "作業管理・家具の買取"], ["今後のヒアリング", "売却・活用の意向"], ["口コミ・紹介", "次の案件へ"]],
      [["紹介ルート", "不明"], ["―", ""], ["実作業・処分", "三者間契約"], ["―", ""], ["―", ""]]];
    LANES.forEach(([t, c, bg], i) => {
      const y = 2.5 + i * 1.3;
      rect(s, 0.35, y, 1.32, 1.18, c);
      txt(s, t, 0.35, y, 1.32, 1.18, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle", ls: 1.15 });
      rect(s, 1.72, y, 8.76, 1.18, bg);
      for (let k = 0; k < 5; k++) {
        rect(s, X0 + k * CW + 0.07, y + 0.1, CW - 0.14, 0.98, WHITE, { line: c, round: true, r: 0.06 });
        const [h, d] = C[i][k];
        txt(s, [B(h + (d ? "\n" : ""), { fontSize: 10.5, color: h === "―" ? MGRAY : DARK }), N(d, { fontSize: 9, color: isU(d) ? RED : GRAY })], X0 + k * CW + 0.12, y + 0.15, CW - 0.24, 0.88, { ls: 1.25, align: "center", valign: "middle", color: DARK });
      }
    });
    const cx = X0 + 3 * CW + CW / 2;
    s.addShape(pres.shapes.UP_ARROW, { x: cx - 0.17, y: 3.58, w: 0.34, h: 0.42, fill: { color: RED }, line: { type: "none" } });
    txt(s, "送客", cx + 0.2, 3.6, 0.6, 0.35, { size: 9, bold: true, color: RED, valign: "middle" });
    rect(s, 0.35, 6.45, 10.13, 0.5, LGRAY);
    txt(s, [B("スキームのポイント　", { color: NAVY }), N("片付けで先に利益を得ながら、売主直で物件を見てから出口（仲介・買取再販）を選べる", { color: DARK })], 0.5, 6.45, 9.9, 0.5, { size: 11, valign: "middle" });
  }
  divider(3, "貴社の課題への解決策", "財務・エリア・組織・仕入れの質");
  // 13 解決策①（P.4）
  {
    const s = page(HEADER, "解決策①【財務・エリア】：在庫リスクゼロで即日キャッシュ化＆都内近郊の反響獲得", "現金を寝かせず日銭を作り、足を使った飛び込みから「Web反響仕入れ」へシフト。", "");
    const D = [["yen", "財務：先出しキャッシュゼロ", NAVY, 0.45, "物件を買って在庫を持つ（在庫約8,000万円）", "片付け完了時に代金回収（粗利約18万円/件）", ["在庫リスクなし", "融資・手元現金を使わずに始められる", "片付けの利益で毎月の現金が回る"]],
      ["map", "エリア：都内オフィスの利点を活用", BLUE, 5.5, "北関東の業者回り・出物待ち", "都内でWeb・MEOの反響", ["「都内在住で実家（都内・近郊・北関東）の処分に困っている売主」を直接集客", "事務所から回れる範囲で営業できる"]]];
    D.forEach(([k, t, c, x, bf, af, L]) => {
      fbox(s, x, 2.0, 4.9, 4.3, "", c, { hh: 0.6 });
      icon(s, k, x + 0.15, 2.08, 0.44, WHITE);
      txt(s, t, x + 0.7, 2.0, 4.1, 0.6, { size: 14, bold: true, color: WHITE, valign: "middle" });
      [["Before", MGRAY, LGRAY, bf], ["After", RED, "FDECEC", af]].forEach(([b, bc, bg, v], i) => {
        const y = 2.8 + i * 0.85;
        rect(s, x + 0.2, y, 0.95, 0.65, bc);
        txt(s, b, x + 0.2, y, 0.95, 0.65, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle" });
        rect(s, x + 1.15, y, 3.55, 0.65, bg);
        fl(s, v, x + 1.3, y, 3.3, 0.65, { valign: "middle", size: 11, bold: i === 1 });
      });
      line(s, x + 0.2, 4.6, 4.5, 0.001, LINE, 0.75, { dash: "dash" });
      blist(s, L, x + 0.25, 4.72, 4.4, 11);
    });
    band(s, "現金を寝かせず日銭を作り、「Web反響仕入れ」へシフト", 6.45, 0.55, 14);
  }
  // 14 解決策②（P.5）
  {
    const s = page(HEADER, "解決策②【組織体制】：属人化を排除する「組織営業（0次接客・即時見積）」", "営業マンの力量や独立に左右されず、誰でも成約率80%を叩き出す型を導入。", "【トーク】前回の「営業1人で月1件、育っても独立した方が得になる」への回答。問い合わせの受付方法（電話・フォーム・LINE）は未定のため不明。");
    const ST = ["問い合わせ", "0次接客（電話）", "現地調査・即時見積", "即決クロージング"];
    const D = [["専門HP・MEO", ["高単価ロングテールのキーワード", "地域一番店の口コミ", "受付方法：不明"]], ["5分以内の初回コール", ["ヒアリング", "写真送付誘導", "現地調査の日程を確定"]], ["即時見積Excelシート", ["平米・間取り・荷物量を入力", "その場で見積提示", "誰が調査しても同じ金額"]], ["その場で即決", ["定額パックで明朗会計", "金額が出るので迷わせない", "成約率80%"]]];
    ST.forEach((t, i) => chev(s, 0.45 + i * 2.5, 2.05, 2.55, 0.6, i === 1 || i === 2 ? NAVY : BLUE, t, i === 0, 12));
    ST.forEach((t, i) => {
      const x = 0.45 + i * 2.5;
      rect(s, x + 0.05, 2.85, 2.35, 2.9, WHITE, { line: LINE });
      txt(s, "使う仕組み", x + 0.15, 2.92, 2.15, 0.3, { size: 10.5, bold: true, color: NAVY, valign: "middle" });
      fl(s, D[i][0], x + 0.15, 3.25, 2.15, 0.6, { size: 12.5, bold: true, valign: "middle", color: NAVY });
      line(s, x + 0.15, 3.9, 2.15, 0.001, LINE, 0.75, { dash: "dash" });
      blist(s, D[i][1], x + 0.15, 4.0, 2.2, 10.5);
    });
    band(s, "営業マンの力量・独立に左右されず、誰でも成約率80%", 6.05, 0.65, 15);
  }
  {
    const s = page(HEADER, "解決策③【仕入れの質】：売主直だから、物件を見てから選べる", "業者からの出物待ちではなく、売主から直接相談が来るから、良い物件だけを選んで仕入れられる。",
      "【前回のお話】実需寄りの中途半端な物件を買ってしまい、売れ残っている。一番良い仕入れは遺品整理業者からの紹介だった。");
    const V = [["業者が持ってくる物件から選ぶ", "業者を通すぶん高く、指値が通りにくい", "他社と取り合い（他社が回らない北関東へ）"], ["片付けの現調で、物件を見てから選べる", "売主と直接話せるので、指値が通りやすい", "売りに出る前なので、競合が少ない"]];
    [["業者からの出物待ち", MGRAY, LGRAY, 0.45], ["売主直（遺品整理から）", NAVY, LBLUE, 5.5]].forEach(([t, c, bg, x], m) => {
      fbox(s, x, 2.0, 4.9, 2.6, t, c, { hh: 0.55, align: "center", hs: 14 });
      [["物件の選び方", 2.7], ["価格", 3.3], ["競合", 3.9]].forEach(([k, y], j) => {
        txt(s, k, x + 0.2, y, 1.4, 0.5, { size: 11, bold: true, color: c, valign: "middle" });
        rect(s, x + 1.6, y, 3.1, 0.5, WHITE, { line: LINE, round: true, r: 0.06 });
        fl(s, V[m][j], x + 1.7, y, 2.95, 0.5, { valign: "middle", size: 10.5, bold: m === 1, color: m === 1 ? NAVY : DARK });
      });
    });
    arrowR(s, 5.35, 3.0, 0.2, 0.5, RED);
    rect(s, 0.45, 4.8, 9.95, 1.15, "FFFBE6", { line: "C9A227", dash: "dash" });
    txt(s, "社長ご自身の実例（一番良かった仕入れ）", 0.65, 4.85, 6, 0.35, { size: 12, bold: true, color: NAVY, valign: "middle" });
    txt(s, [N("遺品整理業者からの紹介で仕入れた物件　→　", { color: DARK }), B("2,300万円で仕入れ、約1か月で3,500万円で売却", { color: RED })], 0.65, 5.2, 9.6, 0.65, { size: 13, valign: "middle" });
    band(s, "良い物件だけを選んで仕入れ、実需寄りの物件をつかまない", 6.15, 0.62, 15);
  }

  // 16 課題と解決の対応
  {
    const s = page(HEADER, "4大課題と、遺品整理で解決できること", "4大課題は、遺品整理を入口にすることで同時に解決できる", "");
    txt(s, "貴社の課題", 0.45, 1.9, 4.5, 0.35, { size: 14, bold: true, color: MGRAY, valign: "middle" });
    txt(s, "遺品整理で解決できること", 5.9, 1.9, 4.5, 0.35, { size: 14, bold: true, color: NAVY, valign: "middle" });
    [["財務", "現金が拘束され新規仕入れが停滞", "在庫リスクゼロで即日キャッシュ化", "解決策①"], ["エリア", "遠い北関東の業者回りでは件数を増やせない", "都内近郊のWeb反響獲得", "解決策①"], ["組織", "属人化と独立リスク", "組織営業（0次接客・即時見積）", "解決策②"], ["仕入れの質", "実需寄りの物件をつかみ売れ残る", "売主直で、物件を見てから選ぶ", "解決策③"]].forEach(([k, a, b, r], i) => {
      const y = 2.32 + i * 1.12;
      rect(s, 0.45, y, 4.6, 0.98, LGRAY);
      icon(s, "search", 0.6, y + 0.12, 0.32, MGRAY);
      txt(s, [B(k + "\n", { fontSize: 13, color: DARK }), N(a, { fontSize: 11, color: DARK })], 1.1, y + 0.08, 3.8, 0.85, { ls: 1.3, color: DARK });
      arrowR(s, 5.2, y + 0.27, 0.55, 0.42, RED);
      rect(s, 5.9, y, 4.5, 0.98, LBLUE, { line: NAVY });
      num(s, i + 1, 6.05, y + 0.12, 0.36, NAVY);
      txt(s, [B(b + "\n", { fontSize: 13, color: NAVY }), N(r, { fontSize: 10.5, color: GRAY })], 6.55, y + 0.08, 3.75, 0.85, { ls: 1.3, color: DARK });
    });
  }
  divider(4, "事業収支シミュレーション", "遺品整理の収益・不動産への波及・3年間の数値計画");
  // 18 単体PL（P.7）
  {
    const s = page(HEADER, "【フロント収益】遺品整理事業 単体PLモデル", "1件30万円・月10件で、年間売上3,600万円・粗利2,420万円（企画書の1年目モデル）",
      "企画書の数値計画（1年目）を、1件あたり・月間・年間に割り戻したもの。1件あたり・月間は年間÷120件・÷12で計算。\n前回の骨子の「原価約40%・粗利約60%（18万円）」は、企画書のモデル（処分費30%＋消耗品、粗利67%）に合わせて差し替え。\n【要確認】1年目の営業利益：表の経費を合計すると2,420万−1,834万＝586万円（16.3%）となり、企画書の106万円（2.94%）と合わない。\n【トーク】前回の「今は残置物撤去に1軒約40万円払っている」につなげる。");
    const C = (t, o = {}) => ({ text: t, options: Object.assign({ fontFace: FONT, fontSize: 10.5, valign: "middle", align: "center", color: DARK, fill: { color: WHITE }, border: { type: "solid", pt: 0.75, color: LINE } }, o) });
    const H = (t) => C(t, { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 11 });
    const L = (t, o = {}) => C(t, Object.assign({ align: "left", bold: true, color: NAVY }, o));
    const SUM = { bold: true, fill: { color: LBLUE }, color: NAVY };
    const CAT = (t, n) => C(t, { rowspan: n, bold: true, color: WHITE, fill: { color: NAVY2 }, fontSize: 11 });
    const rows = [
      [H("項目"), H(""), H("1件あたり"), H("月間（10件）"), H("年間（1年目・120件）"), H("備考")],
      [L("売上", { colspan: 2, fontSize: 12 }), C("30.0万円", { bold: true, fontSize: 12 }), C("300万円", { bold: true, fontSize: 12 }), C("3,600万円", { bold: true, fontSize: 12 }), C("客単価30万円", { fontSize: 9.5, color: GRAY })],
      [CAT("原価", 2), L("処分費"), C("9.0万円"), C("90万円"), C("1,080万円"), C("売上の30%", { fontSize: 9.5, color: GRAY })],
      [L("その他消耗品費"), C("0.8万円"), C("8.3万円"), C("100万円"), C("資材等", { fontSize: 9.5, color: GRAY })],
      [L("粗利", Object.assign({ colspan: 2 }, SUM)), C("20.2万円", SUM), C("201.7万円", SUM), C("2,420万円", SUM), C("粗利率67%", Object.assign({}, SUM, { fontSize: 9.5 }))],
      [CAT("経費", 5), L("アルバイト事務"), C("1.0万円"), C("10.0万円"), C("120万円"), C("事務1名", { fontSize: 9.5, color: GRAY })],
      [L("広告宣伝費"), C("4.0万円"), C("40.3万円"), C("484万円"), C("粗利の20%", { fontSize: 9.5, color: GRAY })],
      [L("その他事業経費"), C("1.3万円"), C("12.5万円"), C("150万円"), C("", {})],
      [L("人件費"), C("5.0万円"), C("50.0万円"), C("600万円"), C("専任1名", { fontSize: 9.5, color: GRAY })],
      [L("作業員アルバイト"), C("4.0万円"), C("40.0万円"), C("480万円"), C("2名", { fontSize: 9.5, color: GRAY })],
      [L("営業利益", { colspan: 2, fill: { color: "E8EFE9" }, color: GREEN, fontSize: 12 }), C("0.9万円", { bold: true, fill: { color: "E8EFE9" }, color: GREEN }), C("8.8万円", { bold: true, fill: { color: "E8EFE9" }, color: GREEN }), C("106万円", { bold: true, fill: { color: "E8EFE9" }, color: GREEN, fontSize: 12 }), C("営業利益率2.94%", { bold: true, fill: { color: "E8EFE9" }, color: GREEN, fontSize: 9.5 })],
    ];
    s.addTable(rows, { x: 0.45, y: 1.95, w: 9.95, colW: [0.75, 2.05, 1.65, 1.75, 2.05, 1.7], rowH: 0.37, objectName: "pl_unit" });
    rect(s, 0.45, 6.12, 9.95, 0.6, "FDECEC", { line: RED, dash: "dash" });
    txt(s, [B("要確認　", { color: RED }), N("1年目の営業利益は、表の経費を合計して計算すると586万円（16.3%）。企画書の106万円（2.94%）と合わないため、どちらが正しいか確認が必要", { color: DARK })], 0.6, 6.12, 9.7, 0.6, { size: 10.5, valign: "middle", ls: 1.2 });
    note(s, MODEL + "　出典：企画書の数値計画（3年間）の1年目。", 6.82);
  }
  // 20 3年間の数値計画
  {
    const s = page(HEADER, "3年間の数値計画", "3年目には売上1億円を達成し、営業利益率23.1%の高収益体制へ（企画書の3年モデル）",
      "企画書の数値計画（3年間）をそのまま掲載。古物商許可は取得済みのため、初期投資の古物商許可取得費用（5〜8万円）は不要。\n【要確認】1年目の営業利益：表の経費を合計すると586万円（16.3%）となり、企画書の106万円（2.94%）と合わない。2年目・3年目は計算と一致。\n体制（人件費の専任人数）は貴社の担当者が決まっていないため、企画書のモデルのまま。");
    const C = (t, o = {}) => ({ text: t, options: Object.assign({ fontFace: FONT, fontSize: 9.5, valign: "middle", align: "center", color: DARK, fill: { color: WHITE }, border: { type: "solid", pt: 0.75, color: LINE } }, o) });
    const G = (t) => C(t, { fontSize: 8, color: GRAY });
    const L = (t, o = {}) => C(t, Object.assign({ align: "left", bold: true, color: NAVY }, o));
    const SUM = { bold: true, fill: { color: LBLUE }, color: NAVY };
    const OP = { bold: true, fill: { color: "E8EFE9" }, color: GREEN };
    const CAT = (t, n) => C(t, { rowspan: n, bold: true, color: WHITE, fill: { color: NAVY2 }, fontSize: 10.5 });
    const Y = (t, c) => C(t, { colspan: 2, bold: true, color: WHITE, fill: { color: c }, fontSize: 11 });
    const rows = [
      [C("項目", { colspan: 2, bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 11 }), Y("1年目", NAVY), Y("2年目", "B8892B"), Y("3年目", BLUE)],
      [L("売上", { colspan: 2, fontSize: 11 }), C("3,600万円", { colspan: 2, bold: true, fontSize: 11 }), C("5,400万円", { colspan: 2, bold: true, fontSize: 11 }), C("1億円", { colspan: 2, bold: true, fontSize: 11 })],
      [CAT("原価", 3), L("処分費"), C("1,080万円"), G("30%"), C("1,620万円"), G("30%"), C("3,000万円"), G("30%")],
      [L("その他消耗品費"), C("100万円"), G("資材等"), C("150万円"), G("件数増加分"), C("200万円"), G("件数増加分")],
      [L("粗利", SUM), C("2,420万円", SUM), C("67%", SUM), C("3,630万円", SUM), C("67%", SUM), C("6,800万円", SUM), C("68%", SUM)],
      [CAT("経費", 5), L("アルバイト事務"), C("120万円"), G("事務1名"), C("120万円"), G("事務1名"), C("120万円"), G("事務1名")],
      [L("広告宣伝費"), C("484万円"), G("粗利の20%"), C("726万円"), G("粗利の20%"), C("1,360万円"), G("粗利の20%")],
      [L("その他事業経費"), C("150万円"), G(""), C("200万円"), G(""), C("250万円"), G("")],
      [L("人件費"), C("600万円"), G("専任1名"), C("1,200万円"), G("専任2名"), C("1,800万円"), G("専任3名")],
      [L("作業員アルバイト"), C("480万円"), G("2名"), C("480万円"), G("2名"), C("960万円"), G("4名")],
      [L("営業利益", Object.assign({ colspan: 2 }, OP)), C("106万円", Object.assign({ colspan: 2 }, OP)), C("904万円", Object.assign({ colspan: 2 }, OP)), C("2,310万円", Object.assign({ colspan: 2 }, OP))],
      [L("営業利益率", Object.assign({ colspan: 2 }, OP)), C("2.94%", Object.assign({ colspan: 2 }, OP)), C("16.74%", Object.assign({ colspan: 2 }, OP)), C("23.10%", Object.assign({ colspan: 2 }, OP))],
    ];
    s.addTable(rows, { x: 0.35, y: 1.92, w: 7.5, colW: [0.55, 1.45, 1.0, 0.85, 1.0, 0.85, 1.0, 0.8], rowH: 0.34, objectName: "plan" });
    // 右：初期投資・投資回収
    rect(s, 8.05, 1.92, 2.43, 2.15, WHITE, { line: NAVY, lw: 1.5 });
    rect(s, 8.05, 1.92, 2.43, 0.45, NAVY);
    txt(s, "初期投資", 8.05, 1.92, 2.43, 0.45, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, [N("HP制作費\n", { fontSize: 10, color: DARK }), B("150万円", { fontSize: 17, color: RED })], 8.2, 2.42, 2.15, 0.75, { valign: "middle", ls: 1.1 });
    line(s, 8.2, 3.2, 2.13, 0.001, LINE, 0.75);
    txt(s, [N("古物商許可取得費用\n", { fontSize: 10, color: DARK }), B("取得済み（0円）", { fontSize: 13, color: GREEN })], 8.2, 3.25, 2.15, 0.75, { valign: "middle", ls: 1.1 });
    s.addShape(pres.shapes.DOWN_ARROW, { x: 9.0, y: 4.12, w: 0.55, h: 0.35, fill: { color: NAVY }, line: { type: "none" } });
    rect(s, 8.05, 4.52, 2.43, 1.9, WHITE, { line: NAVY, lw: 1.5 });
    rect(s, 8.05, 4.52, 2.43, 0.45, NAVY);
    txt(s, "投資回収", 8.05, 4.52, 2.43, 0.45, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, [B("2年目", { fontSize: 22, color: RED }), N("で回収", { fontSize: 12, color: RED })], 8.15, 5.0, 2.25, 0.6, { align: "center", valign: "middle" });
    txt(s, "2年目の営業利益904万円により、初期投資を回収見込み", 8.2, 5.6, 2.15, 0.75, { size: 9.5, color: DARK, valign: "middle" });
    txt(s, [B("要確認　", { color: RED }), N("1年目の営業利益は、表の経費を合計すると586万円（16.3%）。企画書の106万円（2.94%）と合わない", { color: DARK })], 0.35, 6.12, 7.5, 0.4, { size: 9.5, valign: "middle" });
    txt(s, [B("【特徴】", { color: DARK }), B("初期投資が低く、社員1名からスタート可能。20%以上の利益率を目指せる", { color: RED })], 0.35, 6.48, 10.13, 0.35, { size: 12, valign: "middle" });
    note(s, MODEL + "　出典：企画書の数値計画（3年間）。", 6.85);
  }
  // 19 バックエンド（P.8）
  {
    const s = page(HEADER, "【バックエンド収益】不動産仲介・買取再販への波及効果", "受託10件に1件が不動産案件化。ハズレ在庫を抱えず利益最大化。", "");
    [["年間受託想定", "80件", NAVY, 1.2], ["不動産派生", "年間4〜8件（5〜10%）", RED, 6.0]].forEach(([t, v, c, x]) => {
      rect(s, x, 2.0, 3.6, 1.1, WHITE, { line: c, lw: 1.5, round: true, r: 0.08 });
      txt(s, [B(t + "\n", { fontSize: 12, color: GRAY }), B(v, { fontSize: v.length > 8 ? 16 : 20, color: c })], x, 2.0, 3.6, 1.1, { align: "center", valign: "middle", ls: 1.2 });
    });
    arrowR(s, 4.95, 2.3, 0.9, 0.5, RED);
    const D = ["手元資金ゼロで仲介手数料（1件あたり100万〜300万円）を獲得", "大幅指値が通る超優良物件のみを選定して自社買取（回転率重視）\n例：2,300万円で仕入れ、約1か月で3,500万円で売却（社長の実例）", "好立地・古民家は、民泊・賃貸での活用も提案できる（ついでの出口）"];
    [["A", "仲介", NAVY, "shake"], ["B", "厳選買取再販", RED, "key"], ["C", "その他", MGRAY, "bed"]].forEach(([a, t, c, k], i) => {
      const x = 0.45 + i * 3.35;
      fbox(s, x, 3.35, 3.15, 2.55, "", c, { hh: 0.55 });
      icon(s, k, x + 0.15, 3.42, 0.4, WHITE);
      txt(s, `選択肢${a}（${t}）`, x + 0.65, 3.35, 2.4, 0.55, { size: 14, bold: true, color: WHITE, valign: "middle" });
      fl(s, D[i], x + 0.18, 4.0, 2.8, 1.8, { size: 11.5, valign: "middle", ls: 1.35 });
    });
    band(s, "物件ごとに出口を選び、ハズレ在庫を抱えず利益最大化", 6.1, 0.6, 15);
    note(s, MODEL, 6.82);
  }
  // 21 買取再販・仲介のインパクト（船井総研 買取再販・空き家再生事業KPIを参考）
  {
    const s = page(HEADER, "買取再販・仲介のインパクト（1件あたり・年間）", "遺品整理の件数 × 買取再販のKPIで、年間の上乗せを試算",
      "【参考】船井総研「買取再販・空き家再生事業」のモデル（セミナー資料）\n・1棟の販売価格1,600万円、粗利300万円（税抜価格の18%）。仕入〜販売まで6か月、年2回転。\n・エンドユーザーからの仕入KPI：反響→査定面談率80%（買取チラシ・売却サイト）→買取率15%／媒介取得率40%。\n【この試算の前提】\n・不動産の相談率5〜10%は前回の骨子の数字（遺品整理の件数は企画書の3か年PL：120件／180件／約333件）。遺品整理の相談は売却サイト等の反響に近いとみて査定面談率80%を使用。\n・媒介1件の手数料：1,600万円の物件・片手・宅建業法の上限（1,600万円×3%＋6万円＝54万円、税抜）。媒介はすべて成約した場合。\n・買取再販の粗利は「販売価格−仕入価格」ではなく、リフォーム費・諸経費を引いた後の粗利（モデル300万円）。\n・社長の実例（2,300万円→3,500万円、差額約1,200万円）は諸経費を引く前の差額で、モデルより大きい好事例。\n・買取には仕入資金（1棟あたり約1,300万円が目安）が必要。在庫期間は6か月が目標。");
    // 上：KPIの流れ
    const ST = [["遺品整理", "（企画書の件数）", NAVY], ["不動産の相談", "5〜10%", BLUE], ["査定面談", "80%", BLUE], ["買取", "15%", RED], ["媒介", "40%", GREEN]];
    ST.forEach(([t, r, c], i) => {
      const x = 0.45 + i * 2.0;
      chev(s, x, 1.95, 2.05, 0.5, c, t, i === 0, 12);
      txt(s, i === 0 ? r : (i >= 3 ? "査定の " : "") + r, x + 0.1, 2.48, 1.9, 0.28, { size: 10, bold: true, color: c, align: "center", valign: "middle" });
    });
    // 左：1件あたり
    rect(s, 0.45, 2.9, 3.55, 3.1, WHITE, { line: NAVY, lw: 1.5 });
    rect(s, 0.45, 2.9, 3.55, 0.42, NAVY);
    txt(s, "1件あたりの粗利（モデル）", 0.45, 2.9, 3.55, 0.42, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, [B("買取再販 1棟\n", { fontSize: 11, color: RED }), B("粗利 300万円\n", { fontSize: 18, color: RED }), N("販売価格1,600万円の18%\n仕入〜販売 6か月（年2回転）", { fontSize: 9.5, color: DARK })], 0.6, 3.38, 3.3, 1.25, { ls: 1.15, valign: "middle" });
    line(s, 0.6, 4.68, 3.25, 0.001, LINE, 0.75, { dash: "dash" });
    txt(s, [B("媒介 1件\n", { fontSize: 11, color: GREEN }), B("仲介手数料 約54万円\n", { fontSize: 15, color: GREEN }), N("1,600万円の物件・片手（税抜）", { fontSize: 9.5, color: DARK })], 0.6, 4.72, 3.3, 0.85, { ls: 1.15, valign: "middle" });
    txt(s, "参考：社長の実例は差額約1,200万円（諸経費を引く前）", 0.6, 5.6, 3.3, 0.35, { size: 8.5, color: GRAY, valign: "middle" });
    // 右：年間
    const C = (t, o = {}) => ({ text: t, options: Object.assign({ fontFace: FONT, fontSize: 10.5, valign: "middle", align: "center", color: DARK, fill: { color: WHITE }, border: { type: "solid", pt: 0.75, color: LINE } }, o) });
    const L = (t, o = {}) => C(t, Object.assign({ align: "left", bold: true, color: NAVY, fill: { color: LBLUE } }, o));
    const H = (t) => C(t, { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 11 });
    const R = { bold: true, color: RED, fill: { color: "FDECEC" }, fontSize: 9.5 };
    const rows = [
      [H("年間（試算）"), H("1年目"), H("2年目"), H("3年目")],
      [L("遺品整理の件数"), C("120件"), C("180件"), C("約333件")],
      [L("不動産の相談（5〜10%）"), C("6〜12件"), C("9〜18件"), C("17〜33件")],
      [L("査定面談（80%）"), C("約5〜10件"), C("約7〜14件"), C("約13〜27件")],
      [L("買取（査定の15%）"), C("約0.7〜1.4棟"), C("約1.1〜2.2棟"), C("約2〜4棟")],
      [L("媒介（査定の40%）"), C("約2〜4件"), C("約3〜6件"), C("約5〜11件")],
      [L("不動産の粗利（上乗せ・約）", { color: RED, fill: { color: "FDECEC" } }), C("320〜640万円", R), C("480〜960万円", R), C("890〜1,770万円", R)],
      [L("参考：遺品整理単体の営業利益", { fontSize: 9.5 }), C("106万円", { color: GRAY }), C("904万円", { color: GRAY }), C("2,310万円", { color: GRAY })],
    ];
    s.addTable(rows, { x: 4.2, y: 2.9, w: 6.2, colW: [2.3, 1.3, 1.3, 1.3], rowH: 0.385, objectName: "impact" });
    band(s, "遺品整理の件数が増えるほど、不動産の粗利が上乗せされる", 6.15, 0.6, 15);
    note(s, "※船井総研の買取再販・空き家再生事業モデル（1棟1,600万円・粗利300万円、査定面談率80%・買取率15%・媒介取得率40%）を当てはめた試算。相談率5〜10%は仮定。成果を保証するものではありません。", 6.8);
  }
  divider(5, "立ち上げの不安を解消する", "過去のFCとの違い・リソース・外注スキーム・原価");
  // 22 過去のFCとの違い（P.11）
  {
    const s = page(HEADER, "過去のFC参入経験（エコパートナーズ）との決定的な違い", "営業マンの査定センス頼みだった過去の失敗要因を、完全な「仕組み」で克服。", "【トーク】前回のお話（2023年頃にFC参入、古物商取得・トラック・社員1名、本部の集客で1日2〜3件アポ、買取は営業マン次第、社員の骨折もあり半年未満で撤退）を受けて違いを示す。");
    rect(s, 2.25, 2.0, 3.95, 0.55, MGRAY);
    txt(s, "過去のモデル", 2.25, 2.0, 3.95, 0.55, { size: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
    rect(s, 6.45, 2.0, 3.95, 0.55, NAVY);
    txt(s, "今回のモデル", 6.45, 2.0, 3.95, 0.55, { size: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
    [["価格の決め方", "安く買い叩く査定スキル依存", "定額パック・即時見積ツールによる明朗会計"], ["主な商品", "不用品回収中心", "家一棟・不動産狙い"], ["営業", "営業マン次第で失速", "組織営業化"], ["現場作業", "社員が作業（けがが撤退の一因に）", "作業は許可業者へ委託（三者間契約）"]].forEach(([k, a, b], i) => {
      const y = 2.7 + i * 0.85;
      rect(s, 0.45, y, 1.75, 0.75, LBLUE);
      txt(s, k, 0.45, y, 1.75, 0.75, { size: 12, bold: true, color: NAVY, align: "center", valign: "middle" });
      rect(s, 2.25, y, 3.95, 0.75, LGRAY);
      fl(s, a, 2.4, y, 3.7, 0.75, { valign: "middle" });
      arrowR(s, 6.22, y + 0.2, 0.2, 0.35, RED);
      rect(s, 6.45, y, 3.95, 0.75, WHITE, { line: NAVY });
      fl(s, b, 6.6, y, 3.7, 0.75, { valign: "middle", bold: true, color: NAVY });
    });
    band(s, "査定センス頼みから、完全な「仕組み」へ", 6.2, 0.62, 15);
  }
  // 23 リソース不足（P.9）
  {
    const s = page(HEADER, "立ち上げの障壁①：リソース不足（トラック・作業員・許認可）の解消", "自社でトラックや現場作業員を抱える必要は一切ありません。", "【トーク】前回のご質問「トラックは必要か」への回答。地元の提携先（許可業者）は未開拓のため不明。");
    const P = [["お客様", 2.6, 2.05, MGRAY, "users"], ["貴社", 0.55, 4.45, NAVY, "building"], ["許可業者", 4.65, 4.45, ORANGE, "truck"]];
    line(s, 1.5, 2.95, 1.5, 1.5, LINE, 2, { flipH: true });
    line(s, 3.6, 2.95, 1.5, 1.5, LINE, 2);
    line(s, 2.4, 4.95, 2.25, 0.001, LINE, 2);
    P.forEach(([t, x, y, c, k]) => {
      rect(s, x, y, 1.85, 0.95, c, { round: true, r: 0.1 });
      icon(s, k, x + 0.12, y + 0.27, 0.4, WHITE);
      txt(s, t, x + 0.55, y, 1.25, 0.95, { size: 13, bold: true, color: WHITE, valign: "middle" });
    });
    fl(s, "片付けの契約", 0.75, 3.4, 1.5, 0.3, { size: 9.5, align: "center", bold: true });
    fl(s, "処分の契約", 3.9, 3.4, 1.5, 0.3, { size: 9.5, align: "center", bold: true });
    fl(s, "作業の依頼", 2.85, 5.0, 1.4, 0.3, { size: 9.5, align: "center", bold: true });
    txt(s, "三者間契約モデル", 0.45, 5.65, 6.1, 0.3, { size: 11, bold: true, color: NAVY, align: "center", valign: "middle" });
    fbox(s, 6.85, 2.05, 3.55, 1.8, "自社でやること", NAVY, { hh: 0.48, align: "center" });
    blist(s, ["集客", "現調", "見積もり"], 7.0, 2.65, 3.3, 12);
    fbox(s, 6.85, 4.05, 3.55, 1.8, "委託すること", ORANGE, { hh: 0.48, align: "center" });
    blist(s, ["実作業", "処分", "委託先：地元の一般廃棄物・産廃許可業者"], 7.0, 4.65, 3.35, 11.5);
    band(s, "完全合法かつ固定費ゼロでスモールスタートが可能", 6.15, 0.62, 15);
  }
  // 24 外注スキーム
  {
    const s = page(HEADER, "遺品整理の外注スキーム", "集客〜見積とお客様対応は自社、作業・処分は外注", "モデル案件の数字は単体PL（客単価30万円・原価約40%）に合わせた。外注の各工程の金額は提携先が未開拓のため不明。");
    rect(s, 0.35, 1.9, 1.35, 0.42, NAVY);
    txt(s, "モデル案件", 0.35, 1.9, 1.35, 0.42, { size: 11.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    rect(s, 1.7, 1.9, 6.0, 0.42, WHITE, { line: LINE });
    txt(s, [N("客単価 30万円 → ", { color: DARK }), B("売上 30万円", { color: RED }), N("　｜　外注費 約12万円（約40%）", { color: DARK })], 1.85, 1.9, 5.8, 0.42, { size: 11, valign: "middle" });
    rect(s, 7.8, 1.9, 2.68, 0.42, RED);
    txt(s, "粗利 18万円（粗利率約60%）", 7.8, 1.9, 2.68, 0.42, { size: 11.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    const STEPS = ["集客", "現地調査", "営業・見積", "仕分け指示", "現場作業", "処分", "完了確認"];
    const X0 = 1.75, CW = 1.247;
    STEPS.forEach((t, i) => {
      rect(s, X0 + i * CW + 0.03, 2.45, CW - 0.06, 0.38, "D9E2F0", { round: true, r: 0.08 });
      txt(s, `${["①", "②", "③", "④", "⑤", "⑥", "⑦"][i]} ${t}`, X0 + i * CW + 0.03, 2.45, CW - 0.06, 0.38, { size: 10, bold: true, color: NAVY, align: "center", valign: "middle" });
    });
    const C = [["Web・MEO\n広告費 月30〜40万円", "現調", "即時見積\n即決", "買取品の選別", "―", "―", "立会い\n代金回収"],
      ["―", "―", "―", "仕分け作業\n金額：不明", "搬出・清掃\n金額：不明", "許可業者\n金額：不明", "―"],
      ["―", "物件の確認", "売却・活用の\n意向確認", "家具・古物の\n買取", "―", "―", "仲介・買取再販\nの提案"]];
    const LANES = [["自社\n遺品整理事業", NAVY, "EEF3FB"], ["外注\n（専門業者）", ORANGE, LORANGE], ["自社\n本業・関連事業", GREEN, LGREEN]];
    LANES.forEach(([t, c, bg], i) => {
      const y = 2.95 + i * 1.12;
      rect(s, 0.35, y, 1.35, 1.02, c);
      txt(s, t, 0.35, y, 1.35, 1.02, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle", ls: 1.15 });
      rect(s, 1.72, y, 8.76, 1.02, bg);
      for (let k = 0; k < 7; k++) {
        rect(s, X0 + k * CW + 0.05, y + 0.08, CW - 0.1, 0.86, WHITE, { line: c, round: true, r: 0.06 });
        const v = C[i][k];
        txt(s, v.includes("不明") ? [N(v.split("\n")[0] + "\n", { color: DARK }), N(v.split("\n")[1], { color: RED })] : v, X0 + k * CW + 0.08, y + 0.12, CW - 0.16, 0.78, { size: 9, align: "center", valign: "middle", ls: 1.25, color: v === "―" ? MGRAY : DARK });
      }
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 6.33, w: 2.0, h: 0.62, fill: { color: NAVY }, line: { type: "none" } });
    txt(s, "スキームのポイント", 0.35, 6.33, 2.0, 0.62, { size: 11.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    ["自社は集客〜見積とお客様対応に集中", "トラック・作業員を持たずに始められる", "現調の段階で物件を確認できる"].forEach((t, k) => {
      const x = 2.45 + k * 2.68;
      rect(s, x, 6.33, 2.6, 0.62, WHITE, { line: LINE });
      icon(s, "ok", x + 0.08, 6.49, 0.3, ORANGE);
      fl(s, t, x + 0.45, 6.33, 2.1, 0.62, { size: 10, valign: "middle" });
    });
  }
  {
    const s = page(HEADER, "30万円で本当にできるのか：原価のつくり方", "今は残置物撤去に1軒約40万円。原価を下げる3つの方法で、片付けを利益の出る事業にする。", "【前回のお話】撤去は業者に依頼して1軒約40万円が多い。「本当に30万円でできるのか」。原価率約40%はP.7の数字。処分単価・買取単価は提携先が未開拓のため不明。");
    rect(s, 0.45, 2.0, 3.0, 1.5, LGRAY, { round: true, r: 0.08 });
    txt(s, [B("今の撤去費（1軒）\n", { fontSize: 12, color: GRAY }), B("約40万円", { fontSize: 24, color: DARK })], 0.45, 2.0, 3.0, 1.5, { align: "center", valign: "middle", ls: 1.3 });
    arrowR(s, 3.6, 2.5, 0.8, 0.5, RED);
    rect(s, 4.55, 2.0, 3.0, 1.5, LBLUE, { line: NAVY, round: true, r: 0.08 });
    txt(s, [B("遺品整理の客単価\n", { fontSize: 12, color: NAVY }), B("30万円", { fontSize: 24, color: NAVY })], 4.55, 2.0, 3.0, 1.5, { align: "center", valign: "middle", ls: 1.3 });
    rect(s, 7.7, 2.0, 2.7, 1.5, "FDECEC", { line: RED, round: true, r: 0.08 });
    txt(s, [B("原価率の目安\n", { fontSize: 12, color: RED }), B("約40%", { fontSize: 24, color: RED })], 7.7, 2.0, 2.7, 1.5, { align: "center", valign: "middle", ls: 1.3 });
    [["store", "処分業者の選定", NAVY, ["エリアの一般廃棄物許可業者を比較", "立米単価・車両1台あたりの単価を確認", "処分単価：不明"]], ["recycle", "分別で処分量を減らす", BLUE, ["鉄・木材などは買取業者へ", "処分に回る量を物理的に減らす", "買取単価：不明"]], ["yen", "買取で費用を相殺", GREEN, ["売れる物は買い取って費用を相殺", "古物商許可を活用", "最初に家に入る会社が買い取れる"]]].forEach(([k, t, c, L], i) => {
      const x = 0.45 + i * 3.35;
      fbox(s, x, 3.75, 3.15, 2.2, "", c, { hh: 0.55 });
      icon(s, k, x + 0.15, 3.82, 0.4, WHITE);
      txt(s, t, x + 0.65, 3.75, 2.45, 0.55, { size: 13, bold: true, color: WHITE, valign: "middle" });
      blist(s, L, x + 0.18, 4.45, 2.9, 10.5);
    });
    band(s, "原価を下げる3つの方法で、30万円でも利益が出る形にする", 6.15, 0.62, 15);
  }

  divider(6, "商圏・競合と事業コンセプト", "勝てる市場か、どう勝つか");
  // 26 商圏
  {
    const s = page(HEADER, "貴社商圏の市場", "世田谷区だけで、遺品整理の市場は約38.8億円", "商圏の範囲（世田谷区のみか、近隣区まで含むか）は不明。空き家数は未調査のため不明。");
    [["人口", "92.3万人"], ["世帯数", "50.3万世帯"], ["65歳以上の割合", "20.7%"]].forEach(([h, v], i) => {
      const x = 0.45 + i * 3.35;
      rect(s, x, 2.0, 3.15, 1.55, LGRAY, { round: true, r: 0.06 });
      rect(s, x, 2.0, 3.15, 0.42, NAVY, { round: true, r: 0.12 });
      txt(s, "世田谷区の" + h, x, 2.0, 3.15, 0.42, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
      txt(s, v, x, 2.45, 3.15, 0.65, { size: 26, bold: true, color: NAVY, align: "center", valign: "middle" });
      txt(s, "（2025年1月1日現在）", x, 3.1, 3.15, 0.35, { size: 9.5, color: GRAY, align: "center", valign: "middle" });
    });
    rect(s, 0.45, 3.8, 9.95, 2.1, WHITE, { line: NAVY, lw: 1.25 });
    rect(s, 0.45, 3.8, 9.95, 0.45, NAVY);
    txt(s, "遺品整理の市場規模（試算）", 0.45, 3.8, 9.95, 0.45, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    [["人口", "92.3万人", NAVY], ["MS（1人あたり消費額）", "4,200円", NAVY], ["市場規模", "約38.8億円", RED]].forEach(([t, v, c], i) => {
      const x = 0.85 + i * 3.25;
      txt(s, [B(t + "\n", { fontSize: 11, color: GRAY }), B(v, { fontSize: 22, color: c })], x, 4.4, 2.6, 1.0, { align: "center", valign: "middle", ls: 1.2 });
      if (i < 2) txt(s, i === 0 ? "×" : "＝", x + 2.6, 4.4, 0.65, 1.0, { size: 22, bold: true, color: GRAY, align: "center", valign: "middle" });
    });
    txt(s, [N("3年目の年商1億円は、市場の約2.6%　｜　", { color: DARK }), N("空き家数：不明", { color: RED })], 0.6, 5.4, 9.65, 0.4, { size: 11.5, bold: true, align: "center", valign: "middle" });
    band(s, "地元・世田谷区だけでも、年商1億円を十分に狙える市場", 6.1, 0.6);
    note(s, "出典：世田谷区 住民基本台帳（2025年1月1日現在）、世田谷区の統計。市場規模は人口×MS4,200円で算出した船井総研のモデル数値。", 6.82);
  }
  // 27 競合状況
  {
    const s = page(HEADER, "競合状況", "競合の配置：不明（商圏の競合調査はこれから実施）", "「遺品整理 世田谷区」「不用品回収 世田谷区」でGoogle検索・マップ上位を調べて配置する。プログレス様は前回「地域の一番店」としてお伝えした会社だが、位置づけは調査後に確定。");
    const X0 = 0.9, Y0 = 2.15, W2 = 9.0, Hh = 4.45;
    rect(s, X0, Y0, W2, Hh, WHITE, { line: "C9D6E6" });
    line(s, X0 + W2 / 2, Y0, 0.001, Hh, "C9D6E6", 1);
    line(s, X0, Y0 + Hh / 2, W2, 0.001, "C9D6E6", 1);
    txt(s, "遺品整理・高単価", X0, Y0 - 0.35, W2, 0.3, { size: 11, bold: true, color: GRAY, align: "center", valign: "middle" });
    txt(s, "不用品回収・低単価", X0, Y0 + Hh + 0.05, W2, 0.3, { size: 11, bold: true, color: GRAY, align: "center", valign: "middle" });
    txt(s, "地域\n密着", 0.25, Y0 + Hh / 2 - 0.3, 0.6, 0.6, { size: 11, bold: true, color: GRAY, align: "center", valign: "middle" });
    txt(s, "広域\n対応", 9.95, Y0 + Hh / 2 - 0.3, 0.6, 0.6, { size: 11, bold: true, color: GRAY, align: "center", valign: "middle" });
    const P = [[1.9, 2.55, RED, "推奨ポジション"], [7.0, 2.75, NAVY, "プログレス様（位置：不明）"], [4.6, 2.5, MGRAY, "不明"], [3.1, 3.35, MGRAY, "不明"], [6.2, 3.55, MGRAY, "不明"], [3.6, 4.95, MGRAY, "不明"], [5.6, 5.25, MGRAY, "不明"], [7.6, 5.75, MGRAY, "不明"], [1.5, 5.7, MGRAY, "不明"]];
    P.forEach(([x, y, c, t]) => {
      const w = t.length > 8 ? 2.6 : 1.55;
      rect(s, x, y, w, 0.36, c);
      txt(s, t, x, y, w, 0.36, { size: 10.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    });
  }
  // 28 競合一覧
  {
    const s = page(HEADER, "競合一覧", "競合の特徴：不明（商圏の競合調査はこれから実施）", "周辺サービスには、遺品整理以外の対応（生前整理・買取・供養・特殊清掃・解体・不動産・相続など）を書く。特に「不動産まで対応しているか」を確認する。");
    const hdr = ["会社・ブランド", "主な打ち出し", "公開価格", "周辺サービス", "弱点・示唆"];
    const rows = [hdr, ["プログレス様", "不明", "不明", "不明", "不明"], ...Array.from({ length: 7 }, () => ["不明", "不明", "不明", "不明", "不明"])];
    const data = rows.map((r, i) => r.map((c, j) => ({ text: c, options: { fontFace: FONT, fontSize: i === 0 ? 11.5 : 10.5, bold: i === 0 || (j === 0 && c !== "不明"), valign: "middle", color: i === 0 ? WHITE : (c === "不明" ? RED : DARK), fill: { color: i === 0 ? NAVY : (i % 2 ? WHITE : LGRAY) }, border: { type: "solid", pt: 0.75, color: LINE } } })));
    s.addTable(data, { x: 0.35, y: 2.0, w: 10.13, colW: [1.75, 2.45, 1.4, 2.3, 2.23], rowH: 0.5, objectName: "competitors" });
  }
  {
    const s = page(HEADER, "空き家買取の競合との違い", "Web広告・葬儀社のパイプ・規模で戦う会社と、同じ土俵に立たない。", "【前回のお話】社長が挙げられた空き家の会社：Web広告が強い会社（アルバリンク様）、葬儀社とのパイプが強い会社（アックスライフ様）、規模の大きい会社（社名は不明）。");
    const hd = ["", "主な仕入れの入口", "強み", "貴社との違い"];
    const cw = [2.0, 2.65, 2.65, 2.65];
    let x = 0.45;
    hd.forEach((h, j) => { rect(s, x, 2.0, cw[j] - 0.05, 0.45, j === 0 ? WHITE : NAVY); if (h) txt(s, h, x, 2.0, cw[j] - 0.05, 0.45, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle" }); x += cw[j]; });
    const R = [["Web広告型\n（アルバリンク様）", "Web広告", "Web広告の集客力", "広告費の勝負。売却を決めてから相談が来る"],
      ["葬儀社パイプ型\n（アックスライフ様）", "葬儀社からの紹介", "葬儀社とのパイプ", "紹介元のある会社とは同じ戦い方をしない"],
      ["規模型\n（社名：不明）", "不明", "資金力・規模", "規模で勝負しない"],
      ["貴社\n（遺品整理起点）", "遺品整理の相談（Web）", "片付けの段階で売主と出会える", "片付けで利益を出しながら仕入れる"]];
    R.forEach((r, i) => {
      const y = 2.55 + i * 0.85, me = i === 3;
      let xx = 0.45;
      cw.forEach((w, j) => {
        rect(s, xx, y, w - 0.05, 0.75, j === 0 ? (me ? RED : LBLUE) : (me ? "FDECEC" : WHITE), { line: me ? RED : LINE });
        if (j === 0) txt(s, r[0], xx, y, w - 0.05, 0.75, { size: 10.5, bold: true, color: me ? WHITE : (isU(r[0]) ? RED : NAVY), align: "center", valign: "middle", ls: 1.1 });
        else fl(s, r[j], xx + 0.12, y, w - 0.3, 0.75, { valign: "middle", bold: me, size: 10.5, color: me ? RED : DARK });
        xx += w;
      });
    });
    band(s, "遺品整理を入口にすることが、空き家買取での貴社の強みになる", 6.1, 0.62, 15);
  }

  // 29 勝ち筋
  {
    const s = page(HEADER, "競合調査から見えた勝ち筋", "大手と真正面から戦わない局地戦（競合の状況は調査後に更新）", "黒字＝競合の状況、赤字＝貴社の打ち手。競合の状況のうち「不明」は調査後に更新する。");
    [["大手ポータル・大手買取再販業者は、広域・ビッグキーワードで集客", "「都内・近郊の家一棟片付け」「ゴミ屋敷」「実家じまい」など高単価ロングテールに特化"],
     ["Googleマップ上位の状況：不明", "MEO対策で、地域一番店の口コミ・マップ上位表示を獲得"],
     ["不動産まで対応する競合の有無：不明", "片付けから仲介・買取再販までワンストップで対応"],
     ["競合の価格の打ち出し：不明", "定額パック・即時見積による明朗会計"]].forEach(([a, b], i) => {
      const y = 2.0 + i * 1.15;
      rect(s, 0.45, y, 9.95, 1.0, i % 2 ? WHITE : LGRAY);
      txt(s, "・" + a, 0.65, y + 0.08, 9.6, 0.42, { size: 13, bold: true, color: isU(a) ? MGRAY : DARK, valign: "middle" });
      txt(s, "→" + b, 0.95, y + 0.5, 9.3, 0.42, { size: 13, bold: true, color: RED, valign: "middle" });
    });
  }
  // 30 Web集客・競合（P.10）
  {
    const s = page(HEADER, "立ち上げの障壁②：Web集客・競合への不安解消", "大手ポータルや大手買取再販業者と真正面から戦わない局地戦戦略。", "【トーク】前回のご質問「大手が入ってこない理由」：地域ごとに廃棄物のルールが違い全国展開しにくい、1件の売上が小さい、と前回お伝えした内容につなげる。");
    [["search", "エリア・キーワードの絞り込み", NAVY, 0.45, "高単価ロングテールに特化", ["都内・近郊の家一棟片付け", "ゴミ屋敷", "実家じまい", "広告費の目安：月30〜40万円"]],
     ["map", "Googleビジネスプロフィール（MEO）対策", BLUE, 5.5, "地域一番店の口コミ・マップ上位表示", ["Googleビジネスプロフィールの整備", "口コミの獲得", "作業事例・写真の投稿", "対象エリア：不明"]]].forEach(([k, t, c, x, h, L]) => {
      fbox(s, x, 2.0, 4.9, 3.85, "", c, { hh: 0.6 });
      icon(s, k, x + 0.15, 2.08, 0.44, WHITE);
      txt(s, t, x + 0.7, 2.0, 4.1, 0.6, { size: 13.5, bold: true, color: WHITE, valign: "middle" });
      fl(s, h, x + 0.2, 2.75, 4.5, 0.9, { size: 15, bold: true, valign: "middle", align: "center", color: c });
      line(s, x + 0.2, 3.75, 4.5, 0.001, LINE, 0.75, { dash: "dash" });
      blist(s, L, x + 0.25, 3.9, 4.4, 12);
    });
    band(s, "大手と真正面から戦わず、局地戦で地域一番店を取る", 6.05, 0.65, 15);
  }
  // 31 事業コンセプト
  {
    const s = page(HEADER, "貴社版の事業コンセプト", "都内の「実家の片付け」から、売主直で物件と出会う", "屋号は本業と分けるかどうかも含めて未定のため不明。");
    [["users", "売り先（誰に）", NAVY, ["都内在住で、実家の処分に困っているご遺族", "実家は都内・近郊・北関東", "生前整理を考えるご家族"]], ["box", "売り物（何を）", BLUE, ["家一棟丸ごとの遺品整理・生前整理・ゴミ屋敷清掃", "定額パック・明朗会計", "片付け後の売却のご相談"]], ["chart", "売り方（どう売る）", GREEN, ["Web・MEOの反響", "0次接客・即時見積・即決", "仲介・買取再販へつなぐ"]]].forEach(([k, t, c, L], i) => {
      const x = 0.45 + i * 3.35;
      rect(s, x, 2.0, 3.15, 3.15, WHITE, { line: c, lw: 1.5 });
      rect(s, x, 2.0, 3.15, 0.55, c);
      icon(s, k, x + 0.15, 2.07, 0.4, WHITE);
      txt(s, t, x + 0.6, 2.0, 2.5, 0.55, { size: 14, bold: true, color: WHITE, valign: "middle" });
      blist(s, L, x + 0.18, 2.75, 2.85, 11.5, 1.6);
    });
    rect(s, 0.45, 5.35, 9.95, 1.25, "FFFBE6", { line: "C9A227", dash: "dash" });
    txt(s, [B("屋号・コンセプト案　", { color: NAVY })], 0.65, 5.45, 3, 0.4, { size: 12.5, valign: "middle" });
    txt(s, [N("屋号：不明（本業と分けるかを含めて、ご一緒に検討）\n", { color: RED }), N("コンセプト：片付けから、売却まで。ご実家のことを一社で", { color: DARK })], 0.65, 5.85, 9.6, 0.7, { size: 12, valign: "middle", ls: 1.35 });
  }
  divider(7, "立ち上げの進め方とご支援", "ロードマップ・支援内容・事例・まとめ");
  // 33 ロードマップ（P.12）
  {
    const s = page(HEADER, "立ち上げ6ヶ月ロードマップ", "6ヶ月で「提携構築 ➔ HP完成 ➔ 集客・営業開始 ➔ 案件受注」までを計画的に完遂。", "【トーク】前回のご質問「何から始めればよいか」への回答。開始月は不明。");
    ["提携構築", "HP完成", "集客・営業開始", "案件受注"].forEach((t, i) => chev(s, 0.45 + i * 2.5, 1.95, 2.55, 0.5, [NAVY, NAVY2, BLUE, RED][i], t, i === 0, 12));
    const X0 = 2.35, MW = 1.34;
    rect(s, 0.45, 2.65, 9.95, 0.4, NAVY);
    for (let m = 0; m < 6; m++) txt(s, `${m + 1}ヶ月目`, X0 + m * MW, 2.65, MW, 0.4, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle" });
    [["商圏競合分析", 0, "競合調査"], ["商品・料金設計", 0, "定額パック・料金表"], ["提携処分業者の開拓", 0, "三者間契約"], ["専門HP制作", 2, "HP制作ディレクション"], ["営業ツール導入", 2, "Excel見積・トーク集"], ["Web広告テスト運用", 4, "リスティング・MEO"], ["初回商談同行・改善", 4, "同行・振り返り"]].forEach(([t, st, d], i) => {
      const y = 3.1 + i * 0.42;
      rect(s, 0.45, y, 9.95, 0.4, i % 2 ? WHITE : LGRAY);
      txt(s, t, 0.55, y, 1.8, 0.4, { size: 10, bold: true, color: DARK, valign: "middle" });
      rect(s, X0 + st * MW + 0.05, y + 0.07, MW * 2 - 0.1, 0.26, [NAVY, BLUE, RED][st / 2], { round: true, r: 0.05 });
      txt(s, d, X0 + st * MW + 0.05, y + 0.07, MW * 2 - 0.1, 0.26, { size: 9, bold: true, color: WHITE, align: "center", valign: "middle" });
    });
    band(s, "6ヶ月で提携構築からHP完成、集客・営業開始、案件受注まで", 6.2, 0.6, 15);
  }
  // 34 支援パッケージ（P.13）
  {
    const s = page(HEADER, "船井総研の提供価値（支援パッケージ詳細）", "戦略立案からツール提供、現場の営業指導までワンストップで伴走。", "");
    [["商品設計 ＋ HP制作ディレクション", "勝てるコンセプト策定、Web制作会社ディレクション", "約150万円"], ["マニュアル・オリジナルツール導入支援", "0次接客スクリプト、即時見積Excel、アライアンス資料", "100万円"], ["業績アップ月次伴走コンサルティング", "月次定例MTG、進捗管理、広告・営業KPI改善", "20万円/月"]].forEach(([t, d, p], i) => {
      const y = 2.0 + i * 1.4, c = [NAVY, BLUE, GREEN][i];
      rect(s, 0.45, y, 9.95, 1.25, WHITE, { line: c, lw: 1.5 });
      rect(s, 0.45, y, 0.9, 1.25, c);
      txt(s, String(i + 1), 0.45, y, 0.9, 1.25, { size: 28, bold: true, color: WHITE, align: "center", valign: "middle" });
      txt(s, t, 1.55, y + 0.1, 5.8, 0.45, { size: 15, bold: true, color: c, valign: "middle" });
      fl(s, d, 1.55, y + 0.58, 5.8, 0.55, { size: 11.5, valign: "middle" });
      rect(s, 7.6, y + 0.3, 2.6, 0.65, LGRAY, { round: true, r: 0.1 });
      txt(s, p, 7.6, y + 0.3, 2.6, 0.65, { size: 18, bold: true, color: c, align: "center", valign: "middle" });
    });
    band(s, "戦略立案からツール提供、営業指導までワンストップで伴走", 6.3, 0.6, 15);
  }
  // 35 成功事例（P.14）
  {
    const s = page(HEADER, "全国の成功事例（不動産業×遺品整理）", "同様の課題を抱えていた企業が、短期間で事業立ち上げと業績アップを実現。", "事例A・Bの社名・所在地・出典、資料への掲載許可は不明。");
    [["事例A（地方不動産）", NAVY, "地方の不動産会社（社名・所在地：不明）", "遺品整理事業に参入", "営業利益1,500万円増", "参入1年で片付け月8件受託、年間5件の媒介・買取案件を獲得"],
     ["事例B（買取再販）", BLUE, "買取再販会社（社名・所在地：不明）", "即時見積ツールの導入", "成約率 35% → 78%", "営業マンの成約率が大幅に向上"]].forEach(([t, c, a, b, r, rd], i) => {
      const x = 0.45 + i * 5.05;
      fbox(s, x, 2.0, 4.9, 4.0, t, c, { hh: 0.6, hs: 15, align: "center" });
      [["企業の概要", 2.75, a], ["取り組み", 3.55, b]].forEach(([k, y, v]) => {
        txt(s, k, x + 0.2, y, 4.5, 0.3, { size: 11, bold: true, color: c, valign: "middle" });
        txt(s, v.includes("不明") ? [N(v.split("（")[0], { color: DARK }), N("（" + v.split("（")[1], { color: RED })] : v, x + 0.2, y + 0.3, 4.5, 0.45, { valign: "middle", size: 11.5, color: DARK });
      });
      rect(s, x + 0.2, 4.45, 4.5, 1.4, "FDECEC", { round: true, r: 0.08 });
      txt(s, "成果", x + 0.35, 4.5, 2, 0.3, { size: 11, bold: true, color: RED, valign: "middle" });
      txt(s, [B(r + "\n", { fontSize: 20, color: RED }), N(rd, { fontSize: 10.5, color: DARK })], x + 0.35, 4.8, 4.2, 0.95, { align: "center", valign: "middle", ls: 1.2 });
    });
    band(s, "同様の課題を抱えていた企業が、短期間で業績アップを実現", 6.2, 0.6, 15);
  }
  // 36 まとめ（P.15）
  {
    const s = page(HEADER, "まとめ・次回のアクション", "成功パターンの7割はすでに型化されています。残り3割を御社仕様に合わせて伴走します。", "次回の日程は前回決まっていないため不明。");
    fbox(s, 0.45, 2.0, 5.9, 3.7, "本日の提案骨子の総括", NAVY, { hh: 0.5 });
    ["目的：遺品整理は、売主直で不動産を仕入れる装置", "課題：財務・エリア・組織・仕入れの質を解決", "収支：遺品整理で年間粗利1,000万〜2,000万円規模", "不安：トラック・作業員は不要。過去のFCとは仕組みが違う", "競合：大手と戦わない局地戦で地域一番店へ", "進め方：6ヶ月で立ち上げ、船井総研が伴走"].forEach((t, i) => { num(s, i + 1, 0.65, 2.68 + i * 0.5, 0.36, NAVY); fl(s, t, 1.15, 2.65 + i * 0.5, 5.1, 0.42, { valign: "middle", size: 11 }); });
    fbox(s, 6.55, 2.0, 3.85, 3.7, "次回までの決定事項", RED, { hh: 0.5 });
    ["社内体制の確定", "支援開始月の決定", "次回の日程：不明"].forEach((t, i) => { rect(s, 6.8, 2.8 + i * 0.85, 0.32, 0.32, WHITE, { line: GRAY, lw: 1.25 }); fl(s, t, 7.25, 2.7 + i * 0.85, 3.0, 0.52, { valign: "middle", size: 13, bold: true }); });
    band(s, "泥臭い業者回りと在庫リスクから脱却し、\n都内オフィスを活かした自社仕入れモデルへシフトしましょう", 5.95, 0.85, 15);
  }
  // 37 船井総研がご提供できること
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
