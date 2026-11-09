// REVIVAL LINK株式会社様 ご提案書 生成スクリプト
// 使い方: PPTX_SKILL=<pptxスキルのパス> node build_revival.js <出力パス.pptx> <企画書ページ画像のフォルダ>
// 構成：前半＝最新の企画書（PDF）のページをそのまま画像で使用／後半＝貴社向けの新規スライド（編集可能）／最後＝企画書の「船井総研がご提供できること」をそのまま
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa");

const OUT = process.argv[2] || "REVIVAL_LINK様_ご提案書.pptx";
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
  pres.title = "REVIVAL LINK株式会社様 遺品整理事業のご提案";
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
    txt(s, head, 0.35, 0.98, 10.13, 0.45, { size: 21, bold: true, color: NAVY, align: "center", valign: "middle" });
    if (sub) txt(s, sub, 0.35, 1.45, 10.13, 0.34, { size: 13, bold: true, color: DARK, align: "center", valign: "middle" });
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
  const HEADER = "REVIVAL LINK様へのご提案";

  // =================================================================
  // 表紙
  // =================================================================
  {
    const s = pres.addSlide({ masterName: "COVER" });
    txt(s, "REVIVAL LINK株式会社様", 0.75, 2.5, 7.5, 0.5, { size: 20, bold: true, color: NAVY, font: "Yu Mincho" });
    txt(s, "遺品整理事業\nご提案書", 0.75, 3.0, 7.5, 1.7, { size: 40, color: DARK, font: "Yu Mincho", ls: 1.1 });
    txt(s, "株式会社 船井総合研究所\n専門工事支援部\n伊木　颯志朗", 0.75, 4.7, 5, 0.85, { size: 12, color: DARK, font: "Yu Mincho", ls: 1.2 });
  }

  // =================================================================
  // 前半：企画書の内容（PDFのページをそのまま）
  // =================================================================
  [2, 3, 4, 5].forEach(imagePage);                     // 会社紹介・コンサルタント
  for (let p = 7; p <= 19; p++) imagePage(p);           // 遺品整理とは・市場性・ビジネスモデル
  for (let p = 45; p <= 76; p++) imagePage(p);          // 許可・4つの柱・集客・アライアンス・営業・見積・現場・成功事例

  // =================================================================
  // 後半：REVIVAL LINK様へのご提案（新規・編集可能）
  // =================================================================
  {
    const s = pres.addSlide({ masterName: "DIVIDER" });
    txt(s, "REVIVAL LINK株式会社様への\nご提案", 0.5, 2.4, 9.83, 2.4, { size: 40, bold: true, color: YEL, align: "center", valign: "middle", ls: 1.2 });
  }
  // 貴社の事業と強み
  {
    const s = page(HEADER, "貴社の事業と強み", "空き家の「仕入れ」から「再生」「活用」まで、自社で一気通貫に手がけられることが最大の強みです",
      "【出典】REVIVAL LINK株式会社 公式サイト（https://revival-link.jp/）の事業内容・会社概要。対応エリア・実績件数・人員はサイトに記載がないため【要ヒアリング】。");
    const G = [["search", "仕入れ", ["空き家買取・再生活用", "不動産売買・収益物件サポート"], NAVY], ["tools", "再生", ["リノベーション", "内装リフォーム・家具家電の設置"], BLUE], ["bed", "活用", ["インバウンド向け民泊", "シェアハウス運営", "収益物件の出口戦略"], GREEN]];
    G.forEach(([k, t, items, c], i) => {
      const x = 0.45 + i * 3.45;
      rect(s, x, 2.0, 3.0, 2.55, WHITE, { line: c, lw: 1.5 });
      rect(s, x, 2.0, 3.0, 0.55, c);
      icon(s, k, x + 0.15, 2.07, 0.4, WHITE);
      txt(s, t, x + 0.6, 2.0, 2.3, 0.55, { size: 16, bold: true, color: WHITE, valign: "middle" });
      txt(s, items.map((it, j) => N("・" + it + (j < items.length - 1 ? "\n" : ""))), x + 0.2, 2.7, 2.7, 1.75, { size: 12, color: DARK, ls: 1.5 });
      if (i < 2) arrowR(s, x + 3.05, 3.05, 0.35, 0.45, LINE);
    });
    rect(s, 0.45, 4.7, 9.9, 0.55, LGRAY);
    txt(s, [B("支える事業　", { color: NAVY }), N("ポスティング・販促支援（地域密着の配布網）　／　通信サービス（入居時の通信導入・法人向け提案）", { color: DARK })], 0.6, 4.7, 9.7, 0.55, { size: 11.5, valign: "middle" });
    txt(s, "貴社の3つの強み", 0.45, 5.38, 5, 0.32, { size: 12.5, bold: true, color: NAVY, valign: "middle" });
    [["空き家の買取から再生・運営まで自社で完結"], ["相続・権利調整案件の対応力と士業連携"], ["自社のポスティング網による地域販促"]].forEach(([t], i) => {
      const x = 0.45 + i * 3.32;
      rect(s, x, 5.75, 3.18, 0.85, LBLUE, { round: true, r: 0.06 });
      num(s, i + 1, x + 0.12, 5.98, 0.4);
      txt(s, t, x + 0.62, 5.75, 2.5, 0.85, { size: 11.5, bold: true, color: NAVY, valign: "middle" });
    });
    note(s, "※貴社公式サイトの記載をもとに整理", 6.75);
  }
  // 不動産会社としてのシナジー（企画書の不動産パートをそのまま）
  [24, 25, 26].forEach(imagePage);
  // 遺品整理が貴社の事業とつながる
  {
    const s = page(HEADER, "遺品整理が、貴社の事業とつながる", "遺品整理は、空き家を「最初に」見つける入口。貴社の既存事業の多くと直接つながります",
      "【話し方】中心の遺品整理と、貴社の5事業のつながりを1枚で示す。通信サービスは関係が弱いので灰色。士業ネットワークはアライアンスとして既存の資産に含める。");
    const CX = 5.42, CY = 4.45;
    oval(s, CX - 0.95, CY - 0.95, 1.9, NAVY);
    icon(s, "box", CX - 0.25, CY - 0.75, 0.5, WHITE);
    txt(s, "遺品整理\n（新規事業）", CX - 0.95, CY - 0.18, 1.9, 0.75, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle", ls: 1.05 });
    const N5 = [["search", "空き家買取・再生活用", "相続空き家の仕入れの入口", -3.6, -1.6, BLUE], ["building", "不動産売買・収益物件", "片付け後の売却・買取相談", 3.6, -1.6, BLUE], ["bed", "民泊・シェアハウス運営", "再生した物件の活用先", 3.6, 1.55, GREEN], ["paper", "ポスティング・販促支援", "遺品整理チラシを自社で配布", -3.6, 1.55, ORANGE], ["law", "士業ネットワーク", "相続案件の相互紹介", 0, -2.0, PURPLE], ["wifi", "通信サービス", "（直接のつながりは小さい）", 0, 2.05, MGRAY]];
    N5.forEach(([k, t, d, dx, dy, c]) => {
      const bw = 2.9, bh = 0.85, x = CX + dx - bw / 2, y = CY + dy - bh / 2;
      const x1 = CX + (dx === 0 ? 0.001 : Math.sign(dx) * 0.9), y1 = CY + (dy > 0 ? 0.55 : -0.55) * (dx === 0 ? 1.72 : 1);
      const x2 = dx === 0 ? CX : (dx > 0 ? x : x + bw), y2 = dx === 0 ? (dy > 0 ? y : y + bh) : y + bh / 2;
      line(s, Math.min(x1, x2), Math.min(y1, y2), Math.abs(x2 - x1) || 0.001, Math.abs(y2 - y1) || 0.001, c, 2, { flipV: (x2 - x1) * (y2 - y1) < 0, dash: c === MGRAY ? "dash" : "solid" });
      rect(s, x, y, bw, bh, WHITE, { line: c, lw: 1.5, round: true, r: 0.08 });
      icon(s, k, x + 0.1, y + 0.2, 0.45, WHITE, c);
      txt(s, [B(t + "\n", { fontSize: 11.5, color: c === MGRAY ? MGRAY : DARK }), N(d, { fontSize: 9.5, color: c === MGRAY ? MGRAY : GRAY })], x + 0.65, y, bw - 0.7, bh, { valign: "middle", ls: 1.15 });
    });
  }
  // シナジー①②
  {
    const s = page(HEADER, "シナジー①②　仕入れの入口と、残置物整理の内製化", "空き家事業の「上流」と「現場」を、遺品整理で強くします",
      "【根拠】①は不動産会社向けDL・企画書の不動産パート（相続物件と“市場に出る前”に出会う、受付帳の様式変更）。②は企画書の不動産パートのモデル（1現場30〜50万円で外注）。貴社の年間買取件数・残置物整理の現状は【要ヒアリング】。");
    const C = [
      ["search", "① 空き家の仕入れの入口になる", BLUE, ["相続が起きたご家族が最初に困るのは「実家の片付け」", "片付けの段階でご家族と出会い、空き家の今後を相談される立場になれる", "受付帳の様式変更（2026年10月〜）の後も使える、新しい“風上”の仕入れルート"], "ポータル・一括査定に出る前に、相続空き家と出会える"],
      ["box", "② 残置物整理を内製化する", GREEN, ["買取・再生物件の残置物整理を、自社の遺品整理部門で行う", "1現場30〜50万円の外注費が、自社の売上・粗利に変わる", "日程を自社で調整でき、買取から再生・活用までのスピードが上がる"], "外注費の削減と、再生までのスピードアップを同時に"],
    ];
    C.forEach(([k, t, c, items, sum], i) => {
      const x = 0.45 + i * 5.02;
      rect(s, x, 2.0, 4.88, 4.85, WHITE, { line: c, lw: 1.5 });
      rect(s, x, 2.0, 4.88, 0.65, c);
      icon(s, k, x + 0.15, 2.1, 0.45, WHITE);
      txt(s, t, x + 0.7, 2.0, 4.1, 0.65, { size: 15, bold: true, color: WHITE, valign: "middle" });
      items.forEach((it, j) => {
        const y = 2.85 + j * 0.9;
        s.addImage({ data: IC("ok", c), x: x + 0.22, y: y + 0.08, w: 0.3, h: 0.3, objectName: nm("ok") });
        txt(s, it, x + 0.62, y, 4.1, 0.8, { size: 12, color: DARK, ls: 1.3 });
      });
      rect(s, x + 0.2, 5.65, 4.48, 0.95, i ? LGREEN : LBLUE, { round: true, r: 0.06 });
      txt(s, sum, x + 0.3, 5.65, 4.28, 0.95, { size: 12.5, bold: true, color: c, align: "center", valign: "middle", ls: 1.25 });
    });
  }
  // シナジー③④
  {
    const s = page(HEADER, "シナジー③④　自社の販促網と、士業ネットワーク", "貴社がすでに持っている資産が、そのまま遺品整理の集客力になります",
      "【根拠】③チラシは「遺品整理」「保存版」を大きく、配布数は世帯数×80%が目安（2027年予測レポート）。世田谷区の世帯数は502,617世帯（2025年1月1日、世田谷区）。④アライアンスの紹介料は粗利の3〜5%が目安（DL資料）。");
    const C = [
      ["paper", "③ 自社のポスティング網で集客する", ORANGE, ["戸建向けの遺品整理チラシ（保存版）を、自社のポスティング網で配布", "外注の配布費がかからず、集客コストを抑えられる", "配布数は世帯数×80%が目安。エリア別の反響を見て頻度を調整"], "WEBと紙の両方で、地域の相談を集められる"],
      ["law", "④ 士業ネットワークをアライアンスに", PURPLE, ["連携している税理士・司法書士・弁護士を、遺品整理の紹介元に", "相続の相談から、片付け・空き家の買取・活用まで一貫して受けられる", "紹介料は粗利の3〜5%が目安。成約時のみ費用が発生"], "相見積もりにならない、信頼ベースの紹介案件が増える"],
    ];
    C.forEach(([k, t, c, items, sum], i) => {
      const x = 0.45 + i * 5.02;
      rect(s, x, 2.0, 4.88, 4.85, WHITE, { line: c, lw: 1.5 });
      rect(s, x, 2.0, 4.88, 0.65, c);
      icon(s, k, x + 0.15, 2.1, 0.45, WHITE);
      txt(s, t, x + 0.7, 2.0, 4.1, 0.65, { size: 15, bold: true, color: WHITE, valign: "middle" });
      items.forEach((it, j) => {
        const y = 2.85 + j * 0.9;
        s.addImage({ data: IC("ok", c === ORANGE ? ORANGE : NAVY), x: x + 0.22, y: y + 0.08, w: 0.3, h: 0.3, objectName: nm("ok") });
        txt(s, it, x + 0.62, y, 4.1, 0.8, { size: 12, color: DARK, ls: 1.3 });
      });
      rect(s, x + 0.2, 5.65, 4.48, 0.95, i ? "F1EDF7" : LORANGE, { round: true, r: 0.06 });
      txt(s, sum, x + 0.3, 5.65, 4.28, 0.95, { size: 12.5, bold: true, color: c, align: "center", valign: "middle", ls: 1.25 });
    });
  }
  // 片付けた後の出口
  {
    const s = page(HEADER, "片付けた後の出口が、すべて貴社にある", "「出口を持つ遺品整理会社」は、地域にほとんどありません",
      "【話し方】一般的な遺品整理業者は片付けで終わる。貴社は片付けた後の空き家を、買取再販・民泊・シェアハウス・売買仲介・リノベーションに自社でつなげられる。解体・土地活用は提携先と連携。");
    chev(s, 0.45, 2.1, 2.6, 0.75, NAVY, "相続・空き家のご相談", true, 12);
    chev(s, 2.95, 2.1, 2.6, 0.75, BLUE, "遺品整理・家財整理", false, 12);
    chev(s, 5.45, 2.1, 2.6, 0.75, NAVY2, "物件の状況を把握", false, 12);
    chev(s, 7.95, 2.1, 2.45, 0.75, RED, "貴社が最適な出口を提案", false, 11);
    const E = [["key", "買取再販", "立地・状態の良い物件を買い取り、再生して再販", BLUE], ["bed", "民泊・シェアハウス", "再生した物件を自社で運営し、収益化", GREEN], ["building", "売買仲介", "相続人から直接、売却の相談を受ける", NAVY], ["tools", "リノベーション", "住み継ぐ・貸すための再生を提案", ORANGE], ["hammer", "解体・土地活用", "古家は提携先と連携して対応", MGRAY]];
    E.forEach(([k, t, d, c], i) => {
      const x = 0.45 + i * 2.0;
      rect(s, x, 3.15, 1.88, 2.6, LGRAY, { round: true, r: 0.06 });
      icon(s, k, x + 0.62, 3.3, 0.65, WHITE, c);
      txt(s, t, x + 0.08, 4.05, 1.72, 0.5, { size: 13, bold: true, color: c === MGRAY ? GRAY : c, align: "center", valign: "middle" });
      txt(s, d, x + 0.12, 4.55, 1.64, 1.1, { size: 10, color: GRAY, align: "center", ls: 1.3 });
    });
    txt(s, "※解体・土地活用は提携先と連携", 0.45, 5.8, 10, 0.25, { size: 8.5, color: MGRAY, align: "right", valign: "middle" });
    band(s, "片付けで信頼を得た会社に、「この家をどうするか」の相談が集まります。", 6.15, 0.62);
  }
  // 本業込みの利益
  {
    const s = page(HEADER, "本業込みの利益シミュレーション", "遺品整理の利益に、残置物の内製化と空き家事業への波及が上乗せされます",
      "【根拠】遺品整理10件の粗利180万円（30万円×粗利率60%×10件）、売買仲介1件（2,000万円×（3%＋6万円）×両手＝132万円）は企画書の「本業込みの利益」と同じモデル数値。残置物の内製化・買取再販・民泊活用は貴社の実績から試算するため【要ヒアリング】。");
    txt(s, "遺品整理10件あたりの粗利（モデル）", 0.45, 1.98, 6, 0.35, { size: 13, bold: true, color: NAVY, valign: "middle" });
    const R = [["遺品整理（10件）", "180万円", 180, NAVY, false], ["売買仲介（1件・両手）", "132万円", 132, BLUE, false], ["残置物整理の内製化", "貴社の外注件数から試算", 0, GREEN, true], ["買取再販・民泊活用", "貴社の実績から試算", 0, ORANGE, true]];
    R.forEach(([t, v, n, c, ph], i) => {
      const y = 2.5 + i * 0.75;
      txt(s, t, 0.45, y, 2.6, 0.55, { size: 12, bold: true, color: DARK, valign: "middle" });
      if (!ph) { rect(s, 3.1, y + 0.1, (n / 180) * 3.6, 0.38, c); txt(s, v, 3.2 + (n / 180) * 3.6, y, 1.6, 0.55, { size: 13, bold: true, color: c, valign: "middle" }); }
      else { rect(s, 3.1, y + 0.08, 4.4, 0.42, WHITE, { line: c, dash: "dash" }); txt(s, v + "【要ヒアリング】", 3.1, y + 0.08, 4.4, 0.42, { size: 10.5, bold: true, color: c, align: "center", valign: "middle" }); }
    });
    rect(s, 7.9, 2.45, 2.5, 2.95, LBLUE, { round: true, r: 0.06 });
    txt(s, [B("1件あたりの粗利\n", { fontSize: 12, color: NAVY }), B("18万円\n", { fontSize: 16, color: MGRAY }), B("↓\n", { fontSize: 14, color: MGRAY }), B("約31万円〜", { fontSize: 24, color: RED })], 7.9, 2.55, 2.5, 2.4, { align: "center", valign: "middle", ls: 1.2 });
    txt(s, "（遺品整理＋売買仲介のみで試算）", 7.9, 4.95, 2.5, 0.35, { size: 8.5, color: GRAY, align: "center", valign: "middle" });
    rect(s, 0.45, 5.6, 9.95, 0.6, LGRAY);
    txt(s, [B("本業込みの利益 ＝ ", { color: NAVY }), B("遺品整理の粗利", { color: BLUE }), B(" ＋ 残置物の内製化 ＋ 仲介手数料 ＋ 買取再販・民泊の利益", { color: NAVY })], 0.45, 5.6, 9.95, 0.6, { size: 13, align: "center", valign: "middle" });
    note(s, "※仲介手数料は宅建業法の上限（売買価格2,000万円×3%＋6万円、両手）で計算。" + MODEL, 6.35);
  }
  // 貴社商圏の市場
  {
    const s = page(HEADER, "貴社商圏の市場", "世田谷区だけで、遺品整理の市場は約39億円。3年目の年商1億円はシェア約2.6%です",
      "【出典】世田谷区「住民基本台帳」人口923,210人・世帯数502,617世帯（2025年1月1日現在）、老年人口割合20.7%（2026年1月1日現在）。市場規模＝人口×MS（国民1人あたりの遺品整理消費額）4,200円の船井総研モデル。空き家数は【要確認：総務省「住宅・土地統計調査」】。対応エリア（周辺区を含むか）は【要ヒアリング】。");
    const ST = [["人口", "92.3", "万人", "2025年1月1日現在"], ["世帯数", "50.3", "万世帯", "2025年1月1日現在"], ["65歳以上の割合", "20.7", "%", "2026年1月1日現在"]];
    ST.forEach(([h, v, u, d], i) => {
      const x = 0.45 + i * 3.35;
      rect(s, x, 2.0, 3.15, 1.55, LGRAY, { round: true, r: 0.06 });
      rect(s, x, 2.0, 3.15, 0.42, NAVY, { round: true, r: 0.12 });
      txt(s, "世田谷区の" + h, x, 2.0, 3.15, 0.42, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
      txt(s, [B(v, { fontSize: 28, color: BLUE }), B(" " + u, { fontSize: 13, color: BLUE })], x, 2.45, 3.15, 0.65, { align: "center", valign: "middle" });
      txt(s, d, x, 3.1, 3.15, 0.35, { size: 9.5, color: GRAY, align: "center", valign: "middle" });
    });
    rect(s, 0.45, 3.8, 9.95, 2.1, WHITE, { line: NAVY, lw: 1.25 });
    rect(s, 0.45, 3.8, 9.95, 0.45, NAVY);
    txt(s, "遺品整理の市場規模（試算）", 0.45, 3.8, 9.95, 0.45, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    const F = [["人口", "92.3万人"], ["MS（1人あたり消費額）", "4,200円"], ["市場規模", "約38.8億円"]];
    F.forEach(([t, v], i) => {
      const x = 0.85 + i * 3.25;
      txt(s, [B(t + "\n", { fontSize: 11, color: GRAY }), B(v, { fontSize: 22, color: i === 2 ? RED : NAVY })], x, 4.4, 2.6, 1.0, { align: "center", valign: "middle", ls: 1.2 });
      if (i < 2) txt(s, i === 0 ? "×" : "＝", x + 2.6, 4.4, 0.65, 1.0, { size: 22, bold: true, color: GRAY, align: "center", valign: "middle" });
    });
    txt(s, "商圏20万人のモデル（市場8.4億円）の約4.6倍の市場が、世田谷区だけにあります", 0.6, 5.4, 9.65, 0.4, { size: 11.5, bold: true, color: NAVY, align: "center", valign: "middle" });
    band(s, "3年目の年商1億円は、世田谷区の市場の約2.6%。十分に狙える規模です。", 6.1, 0.6);
    note(s, "出典：世田谷区「住民基本台帳」。市場規模は人口×MS4,200円で算出した船井総研のモデル数値。", 6.82);
  }
  // 貴社版の事業コンセプト
  {
    const s = page(HEADER, "貴社版の事業コンセプト", "売り先・売り物・売り方をそろえ、空き家事業とは別の「遺品整理専門店」として打ち出します",
      "【要ヒアリング】屋号案、対応エリア、ターゲットの優先順位。屋号は空き家買取の会社名のままだと「買い叩かれそう」と警戒されやすいため、遺品整理専門の屋号を別に立てる。");
    const C3 = [["users", "売り先（誰に）", ["世田谷区・周辺の相続世代（50〜70代）", "実家が空き家になったご家族", "賃貸オーナー（退去後の残置物）"], NAVY], ["box", "売り物（何を）", ["遺品整理の標準パック", "供養・買取・特殊清掃のオプション", "空き家の買取査定・活用相談（無料）"], BLUE], ["chart", "売り方（どう売る）", ["専門HP（SEO・MEO・WEB広告）", "自社ポスティングで戸建に配布", "士業アライアンスからの紹介"], GREEN]];
    C3.forEach(([k, t, items, c], i) => {
      const x = 0.45 + i * 3.35;
      rect(s, x, 2.0, 3.15, 3.15, WHITE, { line: c, lw: 1.5 });
      rect(s, x, 2.0, 3.15, 0.55, c);
      icon(s, k, x + 0.15, 2.07, 0.4, WHITE);
      txt(s, t, x + 0.6, 2.0, 2.5, 0.55, { size: 14, bold: true, color: WHITE, valign: "middle" });
      txt(s, items.map((it, j) => N("・" + it + (j < items.length - 1 ? "\n" : ""))), x + 0.18, 2.7, 2.85, 2.35, { size: 11.5, color: DARK, ls: 1.6 });
    });
    rect(s, 0.45, 5.35, 9.95, 1.25, "FFFBE6", { line: "C9A227", dash: "dash" });
    txt(s, [B("屋号・コンセプト案　", { color: NAVY }), N("例：「○○遺品整理センター（運営：REVIVAL LINK株式会社）」\n空き家事業とは別の屋号・専門HPで、片付けの専門家として信頼を得てから、空き家の買取・活用につなげる", { color: DARK })], 0.65, 5.35, 9.6, 1.25, { size: 11.5, valign: "middle", ls: 1.4 });
  }
  // 料金設計
  {
    const s = page(HEADER, "貴社版の料金設計", "お客様には分かりやすいパック料金、社内は立米単価で原価を管理します",
      "【出典】パック料金・オプションは2027年予測レポートのモデル。【要確認】世田谷区周辺の競合の料金を調査し、東京23区の相場に合わせて調整する。");
    const T = [["間取り", "基本物量の目安", "作業人数", "作業時間", "料金（税込）"], ["1K", "段ボール20箱程度", "1名", "1〜3時間", "30,000円〜"], ["1LDK", "段ボール50箱程度", "2〜3名", "3〜6時間", "80,000円〜"], ["2LDK", "段ボール100箱程度", "3〜4名", "5〜8時間", "150,000円〜"], ["3LDK", "段ボール150箱程度", "4〜6名", "6〜9時間", "200,000円〜"], ["4LDK以上", "段ボール200箱以上", "5名〜", "1日〜2日", "300,000円〜"]];
    const data = T.map((r, i) => r.map((c, j) => ({ text: c, options: { fontFace: FONT, fontSize: 11.5, valign: "middle", align: "center", bold: i === 0 || j === 0, color: i === 0 ? WHITE : (j === 0 ? NAVY : DARK), fill: { color: i === 0 ? NAVY : (j === 0 ? LBLUE : (i % 2 ? WHITE : LGRAY)) }, border: { type: "solid", pt: 0.75, color: LINE } } })));
    s.addTable(data, { x: 0.45, y: 2.0, w: 6.3, colW: [1.1, 1.75, 1.05, 1.1, 1.3], rowH: 0.42, objectName: "pack" });
    txt(s, "仕分け・搬出・処分・簡易清掃・運搬まで含んだパック料金（東京23区の相場で要調整）", 0.45, 4.6, 6.3, 0.3, { size: 9, color: MGRAY, valign: "middle" });
    txt(s, "オプション", 0.45, 4.95, 3, 0.32, { size: 12.5, bold: true, color: NAVY, valign: "middle" });
    [["供養（お焚き上げ）", "15,000円〜", 0], ["特殊清掃・脱臭", "30,000円〜", 0], ["空き家の買取査定", "無料", 1], ["活用・リノベーション相談", "無料", 1]].forEach(([t, v, hl], i) => {
      const x = 0.45 + i * 1.6;
      rect(s, x, 5.32, 1.52, 0.75, hl ? LBLUE : LGRAY, { round: true, r: 0.08, line: hl ? BLUE : undefined });
      txt(s, [B(t + "\n", { fontSize: 9.5, color: DARK }), B(v, { fontSize: 11, color: BLUE })], x, 5.32, 1.52, 0.75, { align: "center", valign: "middle", ls: 1.1 });
    });
    txt(s, "■ 貴社ならではの入口商品（無料相談で空き家事業へつなぐ）", 0.45, 6.15, 6.5, 0.28, { size: 9.5, bold: true, color: BLUE, valign: "middle" });
    rect(s, 7.05, 2.0, 3.35, 4.45, LGRAY, { round: true, r: 0.06 });
    txt(s, "社内の原価管理", 7.05, 2.1, 3.35, 0.4, { size: 13, bold: true, color: NAVY, align: "center", valign: "middle" });
    txt(s, [B("140サイズ段ボール10箱\n", { color: DARK }), B("＝ 1立米\n\n", { color: BLUE, fontSize: 16 }), B("1立米単価の目安\n", { color: DARK }), B("12,000〜15,000円\n", { color: RED, fontSize: 16 }), N("（紹介案件は7,000〜10,000円）\n\n", { color: GRAY, fontSize: 10 }), N("見積シートで物量を箱数に換算し、誰が調査しても同じ金額を出せる仕組みにする", { color: GRAY, fontSize: 10.5 })], 7.2, 2.55, 3.05, 3.8, { size: 12, align: "center", valign: "top", ls: 1.3 });
  }
  // 数値計画
  {
    const s = page(HEADER, "貴社版の数値計画", "企画書の3年モデルをベースに、貴社の体制と商圏に合わせて計画します",
      "【出典】売上・営業利益率・体制は企画書の3年モデル。貴社の人員配置・投資余力・残置物の内製化分は【要ヒアリング】で調整する。");
    const T = [["", "1年目", "2年目", "3年目"], ["売上", "3,600万円", "5,400万円", "1億円"], ["月間契約件数", "10件", "15件", "28件"], ["営業利益率", "2.9%", "16.7%", "23.1%"], ["専任スタッフ", "1名", "2名", "3名"], ["世田谷区でのシェア", "約0.9%", "約1.4%", "約2.6%"]];
    const data = T.map((r, i) => r.map((c, j) => ({ text: c, options: { fontFace: FONT, fontSize: 12.5, valign: "middle", align: "center", bold: i === 0 || j === 0 || i === 1, color: i === 0 ? WHITE : (i === 1 && j > 0 ? RED : (j === 0 ? NAVY : DARK)), fill: { color: i === 0 ? NAVY : (j === 0 ? LBLUE : (i % 2 ? WHITE : LGRAY)) }, border: { type: "solid", pt: 0.75, color: LINE } } })));
    s.addTable(data, { x: 0.45, y: 2.0, w: 6.4, colW: [2.2, 1.4, 1.4, 1.4], rowH: 0.52, objectName: "plan" });
    note(s, MODEL + "シェアは世田谷区の市場規模（約38.8億円）に対する割合。", 5.2);
    rect(s, 7.1, 2.0, 3.3, 3.12, WHITE, { line: GREEN, lw: 1.5 });
    rect(s, 7.1, 2.0, 3.3, 0.48, GREEN);
    txt(s, "貴社で上乗せできるもの", 7.1, 2.0, 3.3, 0.48, { size: 12.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, [N("・残置物整理の内製化による外注費の削減\n"), N("・遺品整理から生まれる空き家の買取・再販\n"), N("・民泊・シェアハウスとしての活用益\n"), N("・売買仲介の手数料")], 7.25, 2.6, 3.05, 2.0, { size: 11, color: DARK, ls: 1.5 });
    txt(s, "【要ヒアリング】で試算", 7.1, 4.65, 3.3, 0.35, { size: 10, bold: true, color: GREEN, align: "center", valign: "middle" });
    band(s, "遺品整理単体で3年1億円、さらに空き家事業への波及で全社の利益を伸ばします。", 5.6, 0.7, 14);
  }

  // =================================================================
  // 最後：船井総研がご提供できること（企画書をそのまま）
  // =================================================================
  for (let p = 77; p <= 96; p++) imagePage(p);

  await pres.writeFile({ fileName: OUT });
  console.log("written:", OUT);
}
main().catch((e) => { console.error(e); process.exit(1); });
