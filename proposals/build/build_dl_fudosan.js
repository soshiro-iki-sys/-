// 不動産会社向けダウンロードレポート 生成スクリプト
// 使い方: PPTX_SKILL=<pptxスキルのパス> node build_dl_fudosan.js <出力パス.pptx>
// トンマナ：「2027年 遺品整理業界の予測」レポートに準拠（16:9・白背景＋青いクリスタル装飾・游ゴシック・
//          大きな青い章番号＋見出し1行＋リード文2行＋紺の見出し帯／淡いグレーのカード）
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa");

const OUT = process.argv[2] || "不動産会社向け_遺品整理DLレポート.pptx";
const ASSET = path.join(__dirname, "assets", "report27");

// ---------- パレット（レポートから採色） ----------
const BLUE = "3333FF", NAVY = "1E2A5A", CARD = "F3F4F7", LAV = "EEEFFF", RED = "FF0000", TAGRED = "E60012";
const DARK = "1A1A1A", GRAY = "595959", MGRAY = "8C8C99", LINE = "C8CBD6", WHITE = "FFFFFF", YEL = "FFF100";
const FONT = "Yu Gothic";
const MODEL = "※船井総研作成のモデル数値です。成果を保証するものではありません。";

const THEME = {
  name: "Funai Report 2027", headFontFace: FONT, bodyFontFace: FONT,
  colors: { dk1: "000000", lt1: "FFFFFF", dk2: NAVY, lt2: CARD, accent1: BLUE, accent2: NAVY, accent3: RED, accent4: "7A7DFF", accent5: "5B6B9A", accent6: "A0A4C8", hlink: BLUE, folHlink: NAVY },
};

const ICON_CACHE = {};
async function prepIcons(list) {
  for (const [name, color] of list) {
    const key = name + "_" + color;
    if (ICON_CACHE[key]) continue;
    const Comp = FA[name];
    if (!Comp) throw new Error("icon not found: " + name);
    const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
    const buf = await sharp(Buffer.from(svg)).png().toBuffer();
    ICON_CACHE[key] = "image/png;base64," + buf.toString("base64");
  }
}

async function main() {
  const ICONS = {
    building: "FaBuilding", yen: "FaYenSign", clock: "FaHourglassHalf", home: "FaHome", users: "FaUsers", law: "FaBalanceScale",
    box: "FaBoxOpen", search: "FaSearchLocation", file: "FaFileAlt", shake: "FaHandshake", hammer: "FaHammer", key: "FaKey",
    tools: "FaTools", chart: "FaChartLine", store: "FaStore", laptop: "FaLaptop", paper: "FaNewspaper", phone: "FaPhoneAlt",
    talk: "FaComments", list: "FaClipboardList", gem: "FaGem", map: "FaMapMarkedAlt", truck: "FaTruck", mail: "FaEnvelope",
    db: "FaDatabase", lightbulb: "FaLightbulb", ok: "FaCheckCircle", ng: "FaTimesCircle", recycle: "FaRecycle", pray: "FaPrayingHands",
  };
  const COLORS = [WHITE, BLUE, NAVY, RED, MGRAY, GRAY];
  const pairs = [];
  Object.values(ICONS).forEach((n) => COLORS.forEach((c) => pairs.push([n, c])));
  await prepIcons(pairs);
  const IC = (k, c) => ICON_CACHE[ICONS[k] + "_" + c];

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "不動産会社のための「遺品整理×空き家」新・仕入れ戦略";
  pres.author = "株式会社船井総合研究所";
  pres.company = "株式会社船井総合研究所";

  pres.defineSlideMaster({ title: "CONTENT", background: { path: path.join(ASSET, "bg_content.jpg") }, objects: [] });
  pres.defineSlideMaster({ title: "CONTENTS", background: { path: path.join(ASSET, "bg_contents.jpg") }, objects: [] });
  pres.defineSlideMaster({ title: "COVER", background: { path: path.join(ASSET, "bg_cover.jpg") }, objects: [] });

  let oid = 0;
  const nm = (p) => `${p}_${++oid}`;
  function rect(s, x, y, w, h, fill, o = {}) {
    s.addShape(o.round ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE, {
      x, y, w, h, fill: fill ? { color: fill } : { type: "none" },
      line: o.line ? { color: o.line, width: o.lw || 1, dashType: o.dash || "solid" } : { type: "none" },
      rectRadius: o.round ? (o.r || 0.08) : undefined, objectName: nm("box"),
    });
  }
  function oval(s, x, y, d, fill) { s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: nm("oval") }); }
  function txt(s, t, x, y, w, h, o = {}) {
    if (Array.isArray(t)) {
      const out = [];
      t.forEach((r) => {
        const parts = r.text.split("\n");
        parts.forEach((p, k) => {
          const op = Object.assign({ align: o.align || "left" }, r.options);
          if (k < parts.length - 1) op.breakLine = true;
          if (p !== "" || k < parts.length - 1) out.push({ text: p, options: op });
        });
      });
      t = out;
    }
    s.addText(t, {
      x, y, w, h, fontFace: FONT, fontSize: o.size || 12, bold: !!o.bold, color: o.color || DARK,
      align: o.align || "left", valign: o.valign || "top", margin: o.margin === undefined ? 0 : o.margin,
      lineSpacingMultiple: o.ls || 1.15, isTextBox: true, fit: "none", objectName: nm("text"),
    });
  }
  const B = (t, o = {}) => ({ text: t, options: Object.assign({ bold: true }, o) });
  const N = (t, o = {}) => ({ text: t, options: Object.assign({ bold: false }, o) });
  function icon(s, k, x, y, d, color, bg) {
    if (bg) { oval(s, x, y, d, bg); const p = d * 0.24; s.addImage({ data: IC(k, color), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: nm("icon") }); }
    else s.addImage({ data: IC(k, color), x, y, w: d, h: d, objectName: nm("icon") });
  }
  function num(s, n, x, y, d, fill) {
    oval(s, x, y, d, fill || BLUE);
    txt(s, String(n), x, y, d, d, { size: Math.round(d * 30), bold: true, color: WHITE, align: "center", valign: "middle" });
  }
  function arrowR(s, x, y, w, h, c) { s.addShape(pres.shapes.RIGHT_ARROW, { x, y, w, h, fill: { color: c || LINE }, line: { type: "none" }, objectName: nm("arrow") }); }
  function arrowD(s, x, y, w, h, c) { s.addShape(pres.shapes.DOWN_ARROW, { x, y, w, h, fill: { color: c || LINE }, line: { type: "none" }, objectName: nm("arrow") }); }
  function chev(s, x, y, w, h, fill, first) { s.addShape(first ? pres.shapes.PENTAGON : pres.shapes.CHEVRON, { x, y, w, h, fill: { color: fill }, line: { type: "none" }, objectName: nm("chev") }); }
  function hline(s, x, y, w, c, wd, dash) { s.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color: c || LINE, width: wd || 0.75, dashType: dash || "solid" }, objectName: nm("line") }); }

  // ページ：大きな青い章番号＋章名、見出し1行、リード文2行
  function page(ch, chTitle, head, lead, notes) {
    const s = pres.addSlide({ masterName: "CONTENT" });
    txt(s, `${ch}.`, 1.22, 0.33, 0.75, 0.68, { size: 40, bold: true, color: BLUE, valign: "bottom" });
    txt(s, chTitle, 1.98, 0.42, 9, 0.55, { size: 22, bold: true, color: DARK, valign: "middle" });
    txt(s, head, 0.8, 1.42, 11.9, 0.5, { size: 19, bold: true, color: DARK, valign: "middle" });
    if (lead) txt(s, lead, 0.8, 1.98, 11.75, 0.62, { size: 11.5, color: GRAY, ls: 1.45 });
    if (notes) s.addNotes(notes);
    return s;
  }
  function band(s, t, y, h) {
    rect(s, 0.72, y, 11.66, h || 0.44, NAVY, { round: true, r: 0.1 });
    txt(s, t, 0.72, y, 11.66, h || 0.44, { size: 13.5, bold: true, color: YEL, align: "center", valign: "middle" });
  }
  function conclusion(s, t, y) { txt(s, [B("結論：", { color: BLUE }), B(t)], 0.75, y, 11.8, 0.4, { size: 13.5, valign: "middle" }); }
  function source(s, t, y) { txt(s, t, 0.75, y, 11.8, 0.28, { size: 8, color: MGRAY, valign: "middle" }); }
  function headCard(s, x, y, w, h, title, hh) {
    rect(s, x, y, w, h, CARD, { round: true, r: 0.06 });
    rect(s, x, y, w, hh || 0.44, NAVY, { round: true, r: 0.12 });
    txt(s, title, x, y, w, hh || 0.44, { size: 12.5, bold: true, color: WHITE, align: "center", valign: "middle" });
  }
  function bullets(s, items, x, y, w, h, o = {}) {
    txt(s, items.map((t, i) => N("・" + t + (i < items.length - 1 ? "\n" : ""), { color: o.color || GRAY })), x, y, w, h, { size: o.size || 10.5, ls: o.ls || 1.35, valign: o.valign || "top" });
  }
  function table(s, rows, x, y, colW, o = {}) {
    const fs = o.size || 10.5;
    const data = rows.map((r, i) => r.map((cell, j) => {
      const isObj = typeof cell === "object" && cell !== null;
      const base = { fontFace: FONT, fontSize: fs, valign: "middle", margin: o.margin || [2, 6, 2, 6], color: GRAY, border: { type: "solid", pt: 0.75, color: LINE }, align: j === 0 ? "left" : (o.align || "center"), fill: { color: WHITE } };
      if (i === 0) Object.assign(base, { bold: true, color: WHITE, fill: { color: NAVY }, align: "center" });
      else if (o.firstCol && j === 0) Object.assign(base, { bold: true, color: DARK, fill: { color: LAV }, align: "center" });
      if (isObj) Object.assign(base, cell.options || {});
      return { text: isObj ? cell.text : String(cell), options: base };
    }));
    s.addTable(data, { x, y, colW, rowH: o.rowH, autoPage: false, objectName: nm("table") });
  }

  // =================================================================
  // 1 表紙
  // =================================================================
  {
    const s = pres.addSlide({ masterName: "COVER" });
    rect(s, 2.05, 1.55, 2.9, 0.42, BLUE, { round: true, r: 0.2 });
    txt(s, "不動産会社様向けレポート", 2.05, 1.55, 2.9, 0.42, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, [B("不動産会社のための\n", { color: NAVY }), B("「遺品整理×空き家」", { color: BLUE }), B("新・仕入れ戦略", { color: NAVY })], 2.05, 2.15, 10.5, 1.9, { size: 38, ls: 1.2, valign: "middle" });
    hline(s, 2.1, 4.25, 9.6, LINE, 1);
    txt(s, "2027年、売却相談を“市場に出る前”に掴む方法", 2.05, 4.4, 10.5, 0.55, { size: 20, bold: true, color: DARK, valign: "middle" });
    txt(s, "ポータル・一括査定に頼らない、相続物件の仕入れルートのつくり方", 2.05, 4.95, 10.5, 0.4, { size: 13, color: GRAY, valign: "middle" });
    txt(s, "株式会社 船井総合研究所", 2.05, 6.3, 6, 0.4, { size: 12, bold: true, color: NAVY, valign: "middle" });
    s.addNotes("表紙。テーマは「相続物件の仕入れルート」。遺品整理を“本業の入口”として位置づける。");
  }

  // =================================================================
  // 2 目次
  // =================================================================
  {
    const s = pres.addSlide({ masterName: "CONTENTS" });
    const T = ["不動産会社を取り巻く「仕入れ」の変化", "なぜ不動産会社が遺品整理をやるべきか", "不動産会社版 遺品整理事業モデル", "成功するための取組み", "船井総研がお手伝いできること"];
    T.forEach((t, i) => {
      const y = 0.72 + i * 1.07;
      txt(s, `${i + 1}.`, 3.45, y, 0.85, 0.7, { size: 40, bold: true, color: "8C8C8C", valign: "bottom" });
      txt(s, t, 4.25, y + 0.12, 8, 0.5, { size: 17, bold: true, color: DARK, valign: "middle" });
    });
  }

  // =================================================================
  // 1章
  // =================================================================
  const C1 = "不動産会社を取り巻く「仕入れ」の変化";
  {
    const s = page(1, C1, [B("売却物件の"), B("「奪い合い」", { color: RED }), B("が、年々厳しくなっている")],
      "ポータルサイトや一括査定の普及で、売主との最初の接点は「複数社が比べられる場」に移りました。広告費をかけても、他社と同じ土俵で物件を奪い合う構造から抜け出しにくくなっています。",
      "【話し方】不動産会社の3つの悩み（仕入れ競争・集客コスト・案件の長期化）に共感してもらうページ。吹き出しは「よくあるお悩み」の例。");
    const C = [
      ["search", "① 仕入れ競争の激化", ["ポータル・一括査定経由の案件が増え、複数社で奪い合いに", "査定額の高さで選ばれ、媒介を取っても売れ残る"]],
      ["yen", "② 集客コストの高騰", ["反響サイト・Web広告・チラシの費用が上昇", "1件あたりの獲得コストが年々上がっている"]],
      ["clock", "③ 案件の長期化", ["売却前の片付けや相続手続きで案件が止まる", "残置物撤去は外注で、日程も費用も読めない"]],
    ];
    C.forEach(([k, t, b], i) => {
      const x = 0.72 + i * 3.98;
      headCard(s, x, 2.78, 3.7, 1.85, t);
      icon(s, k, x + 0.22, 3.48, 0.62, WHITE, BLUE);
      bullets(s, b, x + 1.0, 3.28, 2.55, 1.25, { size: 10.5, valign: "middle" });
    });
    txt(s, "よくあるお悩み", 0.75, 4.9, 3, 0.3, { size: 11, bold: true, color: NAVY, valign: "middle" });
    ["一括査定の反響が、以前より減ってきた", "媒介は取れても、なかなか専任が取れない", "残置物の業者待ちで、引き渡しが遅れる"].forEach((t, i) => {
      const x = 0.72 + i * 3.98;
      s.addShape(pres.shapes.ROUNDED_RECTANGULAR_CALLOUT, { x, y: 5.25, w: 3.7, h: 0.6, fill: { color: WHITE }, line: { color: LINE, width: 1 }, objectName: nm("bubble") });
      txt(s, `「${t}」`, x + 0.05, 5.25, 3.6, 0.52, { size: 10.5, bold: true, color: DARK, align: "center", valign: "middle" });
    });
    conclusion(s, "同じ土俵で戦う限り、仕入れコストは上がり続ける。「市場に出る前」の接点が必要です。", 6.2);
  }
  {
    const s = page(1, C1, [B("不動産情報を"), B("「風上」", { color: RED }), B("で掴む方法の一つが、使えなくなった")],
      "相続登記をきっかけに法務局で不動産登記受付帳を開示請求し、所有者へアプローチする手法がありました。2026年10月1日以降の登記分から受付帳の様式が変わり、この方法は事実上使えなくなっています。",
      "【出典】不動産登記規則の一部を改正する省令（令和7年法務省令第49号、2025年10月10日公布・2026年10月1日施行）。東京法務局の案内：2026年10月1日以降に出された登記は、開示請求時に「登記の目的」「不動産所在事項」が表示されない（9月30日までの受付分は従来様式）。");
    hline(s, 1.4, 3.15, 10.5, LINE, 4);
    [["これまで", 2.2, MGRAY], ["2026年10月1日〜", 6.4, RED], ["これから", 10.6, NAVY]].forEach(([t, cx, c]) => {
      oval(s, cx - 0.17, 2.98, 0.34, c);
      txt(s, t, cx - 1.5, 2.6, 3.0, 0.34, { size: 12.5, bold: true, color: c, align: "center", valign: "middle" });
    });
    const BX = [
      ["相続登記をきっかけに、不動産登記受付帳を開示請求", "所有者へDM・訪問でアプローチし、売却相談を“風上”で獲得できた", CARD, LINE, DARK],
      ["受付帳に「登記の目的」「不動産所在事項」が表示されない", "相続登記の情報から所有者を探す方法は、事実上使えなくなる", "FFF0F0", "F5B5B5", RED],
      ["ポータル・反響広告・既存の人脈での競争がさらに激化", "相続物件と“市場に出る前”に出会う、新しい入口が必要になる", LAV, "C9CBF5", NAVY],
    ];
    BX.forEach(([h, b, fill, ln, hc], i) => {
      const x = 0.72 + i * 3.98;
      rect(s, x, 3.55, 3.7, 1.95, fill, { line: ln, round: true, r: 0.06 });
      txt(s, h, x + 0.2, 3.68, 3.3, 0.75, { size: 12, bold: true, color: hc, valign: "middle", ls: 1.25 });
      hline(s, x + 0.2, 4.5, 3.3, ln, 0.75, "dash");
      txt(s, b, x + 0.2, 4.6, 3.3, 0.8, { size: 10.5, color: GRAY, ls: 1.35 });
    });
    arrowR(s, 4.48, 4.35, 0.2, 0.3, LINE); arrowR(s, 8.46, 4.35, 0.2, 0.3, LINE);
    source(s, "出典：不動産登記規則の一部を改正する省令（令和7年法務省令第49号）、東京法務局「不動産登記受付帳の様式変更について」", 5.6);
    band(s, "だからこそ今、新しい「風上チャネル」が必要です。その一つが「空き家 × 遺品整理」です。", 6.15, 0.5);
  }
  {
    const s = page(1, C1, [B("相続不動産は増え続け、"), B("最初の相談先は「片付け」の会社", { color: RED })],
      "多死社会・空き家・相続登記の義務化で、相続物件は今後も増え続けます。ただし、ご家族が最初に困るのは「売却」ではなく「実家の片付け」です。",
      "【出典】厚生労働省「令和7年人口動態統計（確定数）」、社人研「日本の将来推計人口」、総務省「令和5年住宅・土地統計調査」。相続登記は過去の相続分も対象。");
    const ST = [["2025年の年間死亡数", "158.9", "万人", "ピークは2040年（約167万人）の見込み"], ["空き家総数", "900", "万戸", "うち放置空き家は385万戸（空き家率13.8%）"], ["相続登記の申請義務", "3", "年以内", "過去の相続分も対象。怠ると10万円以下の過料"]];
    ST.forEach(([h, v, u, d], i) => {
      const x = 0.72 + i * 3.98;
      headCard(s, x, 2.72, 3.7, 1.62, h, 0.4);
      txt(s, [B(v, { fontSize: 30, color: BLUE }), B(" " + u, { fontSize: 13, color: BLUE })], x, 3.12, 3.7, 0.65, { align: "center", valign: "middle" });
      txt(s, d, x + 0.15, 3.8, 3.4, 0.42, { size: 10, color: GRAY, align: "center", valign: "middle" });
    });
    txt(s, "相続が起きてから、ご家族が動く順番", 0.75, 4.5, 6, 0.3, { size: 11.5, bold: true, color: NAVY, valign: "middle" });
    const F = [["相続発生", MGRAY], ["① 実家の片付け", BLUE], ["② 相続の手続き", "7A7DFF"], ["③ 売却・活用の検討", NAVY]];
    F.forEach(([t, c], i) => {
      const x = 0.72 + i * 2.95;
      chev(s, x, 4.88, 3.05, 0.62, c, i === 0);
      txt(s, t, x + (i === 0 ? 0.1 : 0.35), 4.88, 2.55, 0.62, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    });
    rect(s, 3.75, 5.62, 2.9, 0.42, WHITE, { line: BLUE, lw: 1.25, round: true, r: 0.2 });
    txt(s, "遺品整理の会社が出会うのはここ", 3.75, 5.62, 2.9, 0.42, { size: 10.5, bold: true, color: BLUE, align: "center", valign: "middle" });
    rect(s, 9.6, 5.62, 2.8, 0.42, WHITE, { line: NAVY, lw: 1.25, round: true, r: 0.2 });
    txt(s, "不動産会社に相談が来るのはここ", 9.6, 5.62, 2.8, 0.42, { size: 10.5, bold: true, color: NAVY, align: "center", valign: "middle" });
    conclusion(s, "売却の相談より先に、ご家族は「片付け」で業者と出会っている。ここが本当の川上です。", 6.3);
    source(s, "出典：厚生労働省「令和7年人口動態統計（確定数）」、社人研「日本の将来推計人口」、総務省「令和5年住宅・土地統計調査」", 6.75);
  }

  // =================================================================
  // 2章
  // =================================================================
  const C2 = "なぜ不動産会社が遺品整理をやるべきか";
  {
    const s = page(2, C2, [B("相続した不動産は、まず"), B("“片付け”", { color: RED }), B("から")],
      "遺品整理の現場では、ご家族の状況と物件の状態を、誰よりも先に、誰よりも詳しく知ることができます。片付けを入口にすれば、売却・活用のご相談を自然な流れで受けられます。",
      "【話し方】売り込むのではなく、片付けの延長で「この家をどうするか」を一緒に考える立場になれる、という点を強調。");
    const F = ["相続発生・ご相談", "遺品整理・家財整理", "物件の状況を把握", "貴社が最適プランを提案", "ご契約・実行"];
    F.forEach((t, i) => {
      const x = 0.72 + i * 2.36;
      chev(s, x, 2.8, 2.45, 0.7, i === 1 ? BLUE : (i === 3 ? NAVY : "A6A9D8"), i === 0);
      txt(s, [B(`0${i + 1}\n`, { fontSize: 9 }), B(t, { fontSize: 11.5 })], x + (i === 0 ? 0.1 : 0.35), 2.8, 1.85, 0.7, { color: WHITE, align: "center", valign: "middle", ls: 1.05 });
    });
    txt(s, "片付けた後の5つの出口", 0.75, 3.75, 5, 0.32, { size: 12, bold: true, color: NAVY, valign: "middle" });
    const E = [["key", "A", "売却仲介", "相続人から直接、媒介を受ける"], ["home", "B", "買取・再販", "条件の良い物件を仕入れる"], ["tools", "C", "リフォーム・\nリノベーション", "実家を「活かす」提案に"], ["building", "D", "賃貸仲介・管理", "貸す選択肢で管理収入をつくる"], ["hammer", "E", "解体・土地活用", "古家は解体して土地で活かす"]];
    E.forEach(([k, a, t, d], i) => {
      const x = 0.72 + i * 2.36;
      rect(s, x, 4.15, 2.2, 2.0, CARD, { round: true, r: 0.06 });
      icon(s, k, x + 0.78, 4.3, 0.62, WHITE, i === 0 || i === 1 ? BLUE : NAVY);
      txt(s, [B(a + "  ", { color: BLUE }), B(t, { color: DARK })], x + 0.1, 5.0, 2.0, 0.55, { size: 12, align: "center", valign: "middle", ls: 1.05 });
      txt(s, d, x + 0.12, 5.55, 1.96, 0.5, { size: 9.5, color: GRAY, align: "center", valign: "top", ls: 1.25 });
    });
    band(s, "片付けで信頼を得た会社に、「この家をどうするか」の相談が集まります。", 6.35, 0.5);
  }
  {
    const s = page(2, C2, [B("遺品整理で生まれる、"), B("2つの利益", { color: RED })],
      "一つは、今まで外注していた残置物撤去を自社で行うこと。もう一つは、遺品整理のお客様から本業の案件が生まれることです。どちらも不動産会社だからこそ得られる効果です。",
      "【根拠】①は96p企画書の不動産パートのモデル例（売買・仲介で月5件の残置物撤去を外注している会社）。②は同企画書の収益シミュレーション（年120件×不動産相談5〜10%）。いずれも船井総研のモデル数値。");
    // ①
    headCard(s, 0.72, 2.75, 5.6, 3.35, "① 残置物撤去の内製化", 0.46);
    txt(s, "売買・仲介で発生する残置物撤去を外注している場合", 0.95, 3.32, 5.1, 0.35, { size: 10.5, color: GRAY, align: "center", valign: "middle" });
    rect(s, 1.05, 3.72, 4.95, 0.6, WHITE, { round: true, r: 0.1 });
    txt(s, "月5件 × 30〜50万円 × 12ヶ月", 1.05, 3.72, 4.95, 0.6, { size: 15, bold: true, color: DARK, align: "center", valign: "middle" });
    arrowD(s, 3.32, 4.38, 0.4, 0.28, BLUE);
    txt(s, [B("年間 ", { fontSize: 14 }), B("1,800〜3,000", { fontSize: 28, color: BLUE }), B(" 万円", { fontSize: 14 })], 0.95, 4.68, 5.1, 0.62, { align: "center", valign: "middle" });
    txt(s, "の外注費を、自社の売上・粗利に変えられる。既存案件で仕事を確保できるため、立上げ時から現場経験を積めます。", 1.0, 5.3, 5.05, 0.7, { size: 10, color: GRAY, align: "center", ls: 1.35 });
    // ②
    headCard(s, 6.78, 2.75, 5.6, 3.35, "② 遺品整理から本業の案件が生まれる", 0.46);
    txt(s, [B("遺品整理 年120件 × 不動産相談 5〜10% ＝ "), B("年6〜12件", { color: BLUE })], 6.95, 3.32, 5.3, 0.35, { size: 11.5, align: "center", valign: "middle" });
    const RK = [["S・A", "立地良・築浅", "買取再販", "200〜900万円", BLUE], ["B", "標準", "売買仲介", "300〜600万円", "7A7DFF"], ["C", "立地悪・築古", "解体", "60〜300万円", MGRAY]];
    RK.forEach(([r, c, o, v, col], i) => {
      const y = 3.75 + i * 0.5;
      rect(s, 7.0, y, 0.8, 0.42, col, { round: true, r: 0.08 });
      txt(s, r, 7.0, y, 0.8, 0.42, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
      txt(s, [N(c + "  →  ", { color: GRAY }), B(o, { color: DARK })], 7.95, y, 2.4, 0.42, { size: 10.5, valign: "middle" });
      txt(s, v, 10.3, y, 1.9, 0.42, { size: 13, bold: true, color: col === MGRAY ? GRAY : col, align: "right", valign: "middle" });
    });
    rect(s, 7.0, 5.32, 5.15, 0.62, WHITE, { round: true, r: 0.1 });
    txt(s, [B("年間粗利（シナジー分） ", { fontSize: 11.5 }), B("約560〜1,800万円", { fontSize: 20, color: BLUE })], 7.0, 5.32, 5.15, 0.62, { align: "center", valign: "middle" });
    source(s, MODEL + "物件価格・取引形態・地域により変動します。", 6.15);
    band(s, "外注費の内製化 ＋ 新たな不動産案件の獲得で、本業の利益が大きく伸びます。", 6.5, 0.48);
  }
  {
    const s = page(2, C2, [B("2027年は、"), B("「地域一番店」が不在", { color: RED }), B("の成長期")],
      "遺品整理業者は不動産の出口を持たず、不動産会社は片付けを持っていません。両方を持つ会社は、まだ地域にほとんどありません。ここから1〜2年で勢力図が固まります。",
      "【出典】2027年遺品整理業界の予測レポート（ライフサイクル・シェア区分）。MS＝国民1人あたりの遺品整理消費額4,200円、市場規模は商圏人口20万人×MSで算出した船井総研のモデル数値。");
    const L = [["導入期", "一部の専門業者のみ"], ["成長期【現在〜2027年】", "地域一番店が不在。\n先行者がシェアを取れる"], ["成長後期", "大手の参入で\n価格競争が激化"], ["成熟期", "供給過多。後発からの\n一番店化は相当困難"]];
    L.forEach(([t, d], i) => {
      const x = 0.72 + i * 2.95, on = i === 1;
      chev(s, x, 2.75, 3.05, 1.0, on ? BLUE : CARD, i === 0);
      txt(s, [B(t + "\n", { fontSize: 12.5, color: on ? WHITE : DARK }), N(d, { fontSize: 9.5, color: on ? WHITE : GRAY })], x + (i === 0 ? 0.1 : 0.45), 2.75, 2.4, 1.0, { align: "center", valign: "middle", ls: 1.1 });
    });
    // 左：両方を持つ会社
    txt(s, "片付けと不動産、両方を持つ会社はほとんどいない", 0.75, 3.95, 5.5, 0.32, { size: 11.5, bold: true, color: NAVY, valign: "middle" });
    const O = { text: "○", options: { bold: true, color: BLUE, fontSize: 14 } }, X = { text: "×", options: { bold: true, color: MGRAY, fontSize: 14 } };
    table(s, [["", "遺品整理業者", "不動産会社", "両方を持つ会社"], ["片付け（川上の接点）", O, X, O], ["不動産の出口（売却・買取）", X, O, O], [{ text: "地域での数", options: { bold: true, color: DARK, fill: { color: LAV }, align: "center" } }, "増加中", "多数", { text: "ほぼ不在", options: { bold: true, color: RED } }]], 0.72, 4.35, [2.3, 1.1, 1.05, 1.15], { size: 10, rowH: 0.42, firstCol: true });
    // 右：シェア表
    txt(s, "商圏人口20万人（市場規模 約8.4億円）でのシェア別売上", 6.6, 3.95, 5.9, 0.32, { size: 11.5, bold: true, color: NAVY, valign: "middle" });
    const SH = [["トップシェア", "2億1,840万円", "26.00%"], ["影響シェア", "9,960万円", "11.86%"], ["存在シェア", "5,400万円", "6.43%"], ["拠点シェア", "3,600万円", "4.29%"]];
    table(s, [["シェア区分", "売上（年間）", "シェア"], ...SH.map((r, i) => i === 0 ? r.map((c) => ({ text: c, options: { bold: true, color: BLUE, fill: { color: LAV } } })) : r)], 6.6, 4.35, [2.2, 2.0, 1.6], { size: 10.5, rowH: 0.34 });
    source(s, "※MS（国民1人あたりの遺品整理消費額）4,200円×商圏人口20万人で算出。" + MODEL, 6.12);
    band(s, "地域一番店が不在の2027年が、シェアを取りにいく最大の好機です。", 6.5, 0.48);
  }

  // =================================================================
  // 3章
  // =================================================================
  const C3 = "不動産会社版 遺品整理事業モデル";
  {
    const s = page(3, C3, [B("遺品整理の利益に、"), B("本業の利益が上乗せされる", { color: RED })],
      "商圏人口20万人・平均客単価30万円で、3年目に年商1億円を目指すモデルです。不動産会社は既存の車両・人員・顧客を活かせるうえ、遺品整理の利益に本業の利益が加わります。",
      "【要確認】PLは2027年予測レポート・企画書と同じモデル数値。1年目の営業利益106万円は、表の経費合計から計算すると586万円になるため、出す前に内訳を確認する。\n【本業込みの利益】遺品整理10件の粗利180万円（30万円×粗利率60%×10件）＋不動産売却1件（2,000万円×（3%＋6万円）×両手＝132万円）＋解体1件（150万円×粗利30%＝45万円）＝357万円。");
    // 左：PL
    txt(s, "3年間の数値計画", 0.75, 2.72, 5, 0.32, { size: 12, bold: true, color: NAVY, valign: "middle" });
    s.addChart(pres.charts.BAR, [{ name: "売上（万円）", labels: ["1年目", "2年目", "3年目"], values: [3600, 5400, 10000] }], {
      x: 0.72, y: 3.05, w: 5.6, h: 2.6, barDir: "col", barGapWidthPct: 45, chartColors: [NAVY], showValue: true, dataLabelPosition: "outEnd",
      dataLabelFormatCode: "[>=10000]\"1億円\";#,##0\"万円\"", dataLabelFontSize: 12, dataLabelFontBold: true, dataLabelColor: NAVY, dataLabelFontFace: FONT,
      catAxisLabelFontSize: 11, catAxisLabelFontFace: FONT, catAxisLabelColor: DARK, valAxisHidden: true, valGridLine: { style: "none" },
      catAxisLineShow: true, showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 12000, objectName: "sales_chart",
    });
    txt(s, "営業利益率", 0.75, 5.72, 1.1, 0.36, { size: 10, color: GRAY, valign: "middle" });
    [["2.9%", 1.95], ["16.7%", 3.48], ["23.1%", 5.0]].forEach(([v, x], i) => {
      rect(s, x, 5.72, 1.15, 0.36, i === 2 ? BLUE : LAV, { round: true, r: 0.15 });
      txt(s, v, x, 5.72, 1.15, 0.36, { size: 12, bold: true, color: i === 2 ? WHITE : BLUE, align: "center", valign: "middle" });
    });
    // 右：本業込みの利益
    headCard(s, 6.78, 2.75, 5.6, 3.38, "遺品整理10件あたりの粗利（モデル）", 0.44);
    txt(s, [B("本業込みの利益 ＝ ", { color: DARK }), B("遺品整理の粗利", { color: BLUE }), B(" ＋ 仲介手数料 ＋ 工事の粗利", { color: DARK })], 6.9, 3.27, 5.35, 0.38, { size: 11, align: "center", valign: "middle" });
    const BR = [["遺品整理（10件）", "180万円", 180], ["不動産売却（1件・両手）", "132万円", 132], ["解体工事（1件）", "45万円", 45]];
    BR.forEach(([t, v, n], i) => {
      const y = 3.78 + i * 0.5;
      txt(s, t, 6.95, y, 2.15, 0.38, { size: 10, color: GRAY, valign: "middle" });
      rect(s, 9.1, y + 0.05, (n / 180) * 2.0, 0.28, i === 0 ? NAVY : BLUE);
      txt(s, v, 9.15 + (n / 180) * 2.0, y, 1.0, 0.38, { size: 11, bold: true, color: DARK, valign: "middle" });
    });
    hline(s, 6.95, 5.33, 5.25, LINE, 0.75, "dash");
    txt(s, [B("1件あたりの粗利  ", { fontSize: 11 }), B("18万円", { fontSize: 15, color: GRAY }), B("  →  ", { fontSize: 14, color: GRAY }), B("約36万円", { fontSize: 22, color: BLUE })], 6.85, 5.42, 5.45, 0.6, { align: "center", valign: "middle" });
    source(s, "※仲介手数料は宅建業法の上限（売買価格2,000万円×3%＋6万円、両手）で計算。" + MODEL, 6.2);
    band(s, "既存の車両・人員・顧客を活かせる不動産会社は、低リスクで早く立ち上げられます。", 6.52, 0.46);
  }
  {
    const s = page(3, C3, [B("「内訳」ではなく"), B("「パック料金」", { color: RED }), B("で売る")],
      "お客様が知りたいのは「どれくらいの量で、おおよそいくらになるのか」です。交通費・作業費・処理費などの内訳を並べると、かえって不安を招きます。間取り×物量で総額が分かるパック料金にします。",
      "【出典】2027年予測レポートのパック料金表・オプション。不動産会社ならではのオプションとして、売却査定・買取相談、空き家管理を前に出す。許可は三社間契約で始められる（一般廃棄物は許可業者がお客様と直接契約）。");
    txt(s, "間取り・物量別 パック料金表（基本料金の例）", 0.75, 2.72, 7, 0.32, { size: 12, bold: true, color: NAVY, valign: "middle" });
    table(s, [["間取り", "基本物量の目安", "作業人数", "作業時間の目安", "料金（税込）"], ["1K", "段ボール20箱程度", "1名", "1〜3時間", "30,000円〜"], ["1LDK", "段ボール50箱程度", "2〜3名", "3〜6時間", "80,000円〜"], ["2LDK", "段ボール100箱程度", "3〜4名", "5〜8時間", "150,000円〜"], ["3LDK", "段ボール150箱程度", "4〜6名", "6〜9時間", "200,000円〜"], ["4LDK以上", "段ボール200箱以上", "5名〜", "1日〜2日", "300,000円〜"]], 0.72, 3.08, [1.15, 2.0, 1.15, 1.6, 1.6], { size: 10.5, rowH: 0.34, firstCol: true });
    txt(s, "仕分け・搬出・処分・簡易清掃・運搬まで含んだパック料金", 0.75, 5.17, 7.5, 0.26, { size: 9, color: MGRAY, valign: "middle" });
    txt(s, "オプション（必要な場合のみ追加）", 0.75, 5.48, 7, 0.3, { size: 12, bold: true, color: NAVY, valign: "middle" });
    const OP = [["供養（お焚き上げ）", "15,000円〜", 0], ["特殊清掃・脱臭", "30,000円〜", 0], ["解体", "100,000円〜", 0], ["売却査定・買取", "無料でご相談", 1], ["空き家管理", "月5,000円〜", 1]];
    OP.forEach(([t, v, hl], i) => {
      const x = 0.72 + i * 1.52;
      rect(s, x, 5.82, 1.44, 0.62, hl ? LAV : CARD, { round: true, r: 0.08, line: hl ? BLUE : undefined });
      txt(s, [B(t + "\n", { fontSize: 9.5, color: DARK }), B(v, { fontSize: 10, color: BLUE })], x, 5.82, 1.44, 0.62, { align: "center", valign: "middle", ls: 1.1 });
    });
    txt(s, "■ 不動産会社ならではのオプション", 0.75, 6.5, 6, 0.26, { size: 9, bold: true, color: BLUE, valign: "middle" });
    // 右：三社間契約
    headCard(s, 8.45, 2.75, 3.93, 3.95, "許可がなくても始められる", 0.44);
    txt(s, "三社間契約で、法令に沿ってスモールスタート", 8.55, 3.25, 3.73, 0.34, { size: 10.5, bold: true, color: NAVY, align: "center", valign: "middle" });
    rect(s, 9.55, 3.68, 1.75, 0.5, WHITE, { line: NAVY, round: true, r: 0.1 });
    txt(s, "お客様", 9.55, 3.68, 1.75, 0.5, { size: 12, bold: true, color: NAVY, align: "center", valign: "middle" });
    rect(s, 8.62, 4.72, 1.6, 0.6, NAVY, { round: true, r: 0.1 });
    txt(s, "貴社\n（遺品整理）", 8.62, 4.72, 1.6, 0.6, { size: 10, bold: true, color: WHITE, align: "center", valign: "middle", ls: 1.05 });
    rect(s, 10.6, 4.72, 1.6, 0.6, BLUE, { round: true, r: 0.1 });
    txt(s, "一般廃棄物\n許可業者", 10.6, 4.72, 1.6, 0.6, { size: 10, bold: true, color: WHITE, align: "center", valign: "middle", ls: 1.05 });
    s.addShape(pres.shapes.LINE, { x: 9.42, y: 4.18, w: 0.55, h: 0.54, flipH: true, line: { color: NAVY, width: 1.5 }, objectName: nm("ln") });
    s.addShape(pres.shapes.LINE, { x: 10.88, y: 4.18, w: 0.52, h: 0.54, line: { color: BLUE, width: 1.5 }, objectName: nm("ln") });
    txt(s, "①サービス契約\n仕分け・搬出・買取", 8.5, 4.2, 1.3, 0.45, { size: 8, color: NAVY, valign: "middle", ls: 1.05 });
    txt(s, "②直接契約\n収集運搬・処分", 11.2, 4.2, 1.15, 0.45, { size: 8, color: BLUE, valign: "middle", ls: 1.05 });
    hline(s, 8.65, 5.5, 3.55, LINE, 0.75, "dash");
    txt(s, [B("古物商許可", { color: NAVY }), N("：買取を行う場合に必要。審査 約40日、申請手数料19,000円", { color: GRAY })], 8.65, 5.58, 3.55, 0.6, { size: 9.5, ls: 1.3 });
    txt(s, "※許可の要件は自治体により異なります", 8.65, 6.22, 3.55, 0.3, { size: 8, color: MGRAY, valign: "middle" });
  }

  // =================================================================
  // 4章
  // =================================================================
  const C4 = "成功するための取組み";
  {
    const s = page(4, C4, [B("屋号もHPも、本業とは分けて"), B("「専門店」", { color: RED }), B("に")],
      "不動産会社の名前のまま遺品整理を打ち出すと、「売却を勧められそう」と警戒されます。遺品整理専門の屋号とHPを別に立て、まずは片付けの専門家として信頼を得ることが出発点です。",
      "【話し方】売却相談は「片付けで信頼を得た後」に自然に生まれる。最初から不動産を前に出さないことがポイント。HPの4条件は2027年予測レポートより。");
    // 左：before / after
    rect(s, 0.72, 2.78, 5.6, 1.55, CARD, { round: true, r: 0.06 });
    rect(s, 0.9, 2.9, 2.1, 0.32, MGRAY, { round: true, r: 0.15 });
    txt(s, "不動産会社の一部門として", 0.9, 2.9, 2.1, 0.32, { size: 9.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    icon(s, "ng", 0.95, 3.38, 0.42, MGRAY);
    txt(s, "「遺品整理もやっています」", 1.5, 3.3, 4.6, 0.4, { size: 13.5, bold: true, color: GRAY, valign: "middle" });
    txt(s, "売り込みを警戒され、相談のハードルが上がる。片付け業者として比較もされにくい", 1.5, 3.72, 4.65, 0.55, { size: 10, color: GRAY, ls: 1.3 });
    arrowD(s, 3.32, 4.4, 0.4, 0.3, BLUE);
    rect(s, 0.72, 4.78, 5.6, 1.55, WHITE, { round: true, r: 0.06, line: BLUE, lw: 1.5 });
    rect(s, 0.9, 4.9, 2.1, 0.32, BLUE, { round: true, r: 0.15 });
    txt(s, "遺品整理の専門店として", 0.9, 4.9, 2.1, 0.32, { size: 9.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    icon(s, "ok", 0.95, 5.38, 0.42, BLUE);
    txt(s, "専門の屋号・専門HPで「指名」される", 1.5, 5.3, 4.7, 0.4, { size: 13.5, bold: true, color: NAVY, valign: "middle" });
    txt(s, "片付けの専門家として信頼され、相談が集まる。運営会社として貴社名を出し、安心感にもつなげる", 1.5, 5.72, 4.7, 0.55, { size: 10, color: GRAY, ls: 1.3 });
    // 右：HPの4条件
    txt(s, "反響が取れる専門HPの4つの条件", 6.8, 2.72, 5.5, 0.32, { size: 12, bold: true, color: NAVY, valign: "middle" });
    const H = [["迷わない導線設計", "トップで専門性・料金・対応エリアを伝え、相談まで一直線に"], ["選びやすい商品設計", "間取り別のパック料金で「何を・いくらで」が一目でわかる"], ["安心できる信頼設計", "施工事例・お客様の声・Google口コミ・スタッフ・許可を明示"], ["問い合わせしやすいUI", "スマホで電話・LINE・見積フォームにすぐ届く"]];
    H.forEach(([t, d], i) => {
      const y = 3.12 + i * 0.82;
      rect(s, 6.78, y, 5.6, 0.72, CARD, { round: true, r: 0.08 });
      num(s, i + 1, 6.92, y + 0.14, 0.44);
      txt(s, [B(t + "\n", { fontSize: 12, color: DARK }), N(d, { fontSize: 9.5, color: GRAY })], 7.52, y, 4.75, 0.72, { valign: "middle", ls: 1.2 });
    });
    band(s, "まずは片付けで信頼を得る。売却の相談は、その先に自然と生まれます。", 6.5, 0.48);
  }
  {
    const s = page(4, C4, [B("集客は"), B("3本柱", { color: RED }), B("、営業は"), B("契約率80%", { color: RED })],
      "専門店HPを土台に、WEB・チラシ・提携の3本柱で反響をとります。営業は案件の規模で進め方を分け、最初の電話から即決までを仕組みにします。広告宣伝費は粗利の15〜20%が目安です。",
      "【出典】2027年予測レポート（5つの集客施策・営業KPI）。営業KPIは2年目の年間目安：反響230組→現場調査225組（98%）→契約180組（80%）＝月15組。");
    txt(s, "集客の3本柱", 0.75, 2.72, 4, 0.32, { size: 12, bold: true, color: NAVY, valign: "middle" });
    const P = [["laptop", "WEB", ["SEO・MEO・AIO対策", "Google検索広告・Meta広告"]], ["paper", "チラシ", ["「遺品整理」「保存版」を大きく", "配布数は世帯数×80%が目安"]], ["shake", "提携", ["司法書士・税理士・葬儀社・ケアマネと相互紹介", "自社の売却・買取物件の残置物も初期の仕事に"]]];
    P.forEach(([k, t, b], i) => {
      const y = 3.1 + i * 1.05;
      rect(s, 0.72, y, 5.6, 0.95, CARD, { round: true, r: 0.08 });
      icon(s, k, 0.88, y + 0.17, 0.6, WHITE, i === 2 ? BLUE : NAVY);
      txt(s, t, 1.6, y, 1.0, 0.95, { size: 14, bold: true, color: i === 2 ? BLUE : NAVY, valign: "middle" });
      bullets(s, b, 2.55, y + 0.05, 3.7, 0.85, { size: 10, ls: 1.3, valign: "middle" });
    });
    txt(s, "案件の規模で、営業方法を分ける", 6.8, 2.72, 5.5, 0.32, { size: 12, bold: true, color: NAVY, valign: "middle" });
    const SL = [["30万円未満", "オンライン中心で効率化", BLUE, "電話ヒアリング → アンケート → 写真 → オンライン見積 → 電話で契約"], ["30万円以上", "現地訪問で信頼と提案力を強化", NAVY, "電話ヒアリング → 事前準備 → 現地確認 → その場で見積・契約"]];
    SL.forEach(([t, sub, c, f], i) => {
      const x = 6.78 + i * 2.85;
      rect(s, x, 3.1, 2.75, 0.62, c, { round: true, r: 0.1 });
      txt(s, [B(t + "\n", { fontSize: 13 }), N(sub, { fontSize: 9 })], x, 3.1, 2.75, 0.62, { color: WHITE, align: "center", valign: "middle", ls: 1.05 });
      rect(s, x, 3.78, 2.75, 1.05, CARD, { round: true, r: 0.06 });
      txt(s, f, x + 0.15, 3.83, 2.45, 0.95, { size: 10, color: GRAY, valign: "middle", ls: 1.4 });
    });
    txt(s, "営業KPI（2年目・年間の目安）", 6.8, 5.05, 5.5, 0.3, { size: 11, bold: true, color: NAVY, valign: "middle" });
    [["反響", "230組", "A6A9D8"], ["現場調査", "225組", "7A7DFF"], ["契約", "180組", BLUE]].forEach(([t, v, c], i) => {
      const x = 6.78 + i * 1.9;
      chev(s, x, 5.42, 2.0, 0.6, c, i === 0);
      txt(s, [B(t + "  ", { fontSize: 10 }), B(v, { fontSize: 14 })], x + (i === 0 ? 0.05 : 0.3), 5.42, 1.5, 0.6, { color: WHITE, align: "center", valign: "middle" });
    });
    txt(s, "＝ 月15組 × 客単価30万円 ＝ 月商450万円", 6.8, 6.07, 5.6, 0.3, { size: 10.5, bold: true, color: DARK, align: "right", valign: "middle" });
    band(s, "最初の電話の印象で、契約率が決まります。", 6.52, 0.46);
  }
  {
    const s = page(4, C4, [B("顧客リストは、不動産案件の"), B("「宝の山」", { color: RED })],
      "遺品整理のお客様の多くは、片付けの後に「この家をどうするか」を考え始めます。現場では売り込まずにご家族の予定をお聞きし、契約後も関係を続けることで、売却・解体のご相談につなげます。",
      "【話し方】トーク例は「売り込まない聞き方」の例。売却の提案は、ご家族から相談が出てから。追客は顧客管理シートで仕組みにする（DL資料の営業STEP04）。");
    txt(s, "片付けの後に続く、ご家族との関係", 0.75, 2.72, 6, 0.32, { size: 12, bold: true, color: NAVY, valign: "middle" });
    const T = [["box", "作業完了のご報告", "貴重品・思い出の品の確認、作業報告書のお渡し"], ["file", "相続手続きのご相談", "司法書士・税理士など提携先のご紹介"], ["key", "空き家の管理・活用", "空き家管理、賃貸・リフォームのご提案"], ["home", "売却・解体のご相談", "査定・買取・解体まで貴社でワンストップ"]];
    T.forEach(([k, t, d], i) => {
      const y = 3.1 + i * 0.82;
      rect(s, 0.72, y, 6.0, 0.72, i === 3 ? LAV : CARD, { round: true, r: 0.08, line: i === 3 ? BLUE : undefined });
      icon(s, k, 0.88, y + 0.12, 0.48, WHITE, i === 3 ? BLUE : NAVY);
      txt(s, [B(t + "\n", { fontSize: 12, color: i === 3 ? BLUE : DARK }), N(d, { fontSize: 9.5, color: GRAY })], 1.5, y, 5.1, 0.72, { valign: "middle", ls: 1.2 });
      if (i < 3) arrowD(s, 1.03, y + 0.72, 0.18, 0.1, LINE);
    });
    txt(s, "現場での、売り込まない聞き方（例）", 7.2, 2.72, 5.2, 0.32, { size: 12, bold: true, color: NAVY, valign: "middle" });
    ["ご実家は、今後どうされるご予定ですか？", "相続のお手続きで、お困りのことはありませんか？", "しばらく空き家になる場合は、管理のご相談も承っています。"].forEach((t, i) => {
      const y = 3.12 + i * 0.86;
      s.addShape(pres.shapes.ROUNDED_RECTANGULAR_CALLOUT, { x: 7.2, y, w: 5.18, h: 0.66, fill: { color: WHITE }, line: { color: LINE, width: 1 }, objectName: nm("bubble") });
      txt(s, `「${t}」`, 7.3, y, 5.0, 0.56, { size: 11, bold: true, color: DARK, valign: "middle" });
    });
    rect(s, 7.2, 5.72, 5.18, 0.6, CARD, { round: true, r: 0.08 });
    icon(s, "db", 7.35, 5.82, 0.4, NAVY);
    txt(s, "顧客管理シートで、案件・ご家族の予定・連絡履歴を記録し、追客を仕組みにする", 7.85, 5.72, 4.45, 0.6, { size: 10, color: GRAY, valign: "middle", ls: 1.3 });
    band(s, "片付けのお客様の名簿が、そのまま貴社の「仕入れリスト」になります。", 6.52, 0.46);
  }

  // =================================================================
  // 5章
  // =================================================================
  const C5 = "船井総研がお手伝いできること";
  {
    const s = page(5, C5, "事業計画から商品設計・集客・営業まで、一気通貫で立上げをサポートします",
      "「遺品整理 新規事業立上げコンサルティングパッケージ」として、4つの領域でご支援します。協力業者探しや単価調整など、立上げ時につまずきやすい実務までお手伝いします。",
      "【出典】2027年予測レポートの4つのソリューションと立上げスケジュール（※仮のスケジュール）。");
    const SV = [["chart", "① 事業戦略策定", "3ヵ年経営計画・販促計画・財務分析"], ["gem", "② 商品設計提案", "遺品整理・不用品回収・買取の商品設計、協力業者探し"], ["laptop", "③ 集客戦略策定", "専門HP制作ディレクション・チラシ制作・WEB分析"], ["talk", "④ 営業戦略策定", "0次接客トーク・見積ツール・顧客管理の仕組み"]];
    SV.forEach(([k, t, d], i) => {
      const y = 2.78 + i * 0.92;
      rect(s, 0.72, y, 5.0, 0.8, CARD, { round: true, r: 0.08 });
      icon(s, k, 0.86, y + 0.12, 0.56, WHITE, i % 2 ? BLUE : NAVY);
      txt(s, [B(t + "\n", { fontSize: 12.5, color: DARK }), N(d, { fontSize: 9.5, color: GRAY })], 1.55, y, 4.1, 0.8, { valign: "middle", ls: 1.2 });
    });
    // 右：スケジュール
    const X0 = 7.55, MW = 0.8;
    txt(s, "立上げスケジュール（6か月）", 6.1, 2.72, 5, 0.3, { size: 12, bold: true, color: NAVY, valign: "middle" });
    rect(s, 6.1, 3.08, 6.28, 0.36, NAVY);
    for (let m = 0; m < 6; m++) txt(s, `${m + 1}か月目`, X0 + m * MW, 3.08, MW, 0.36, { size: 9.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    const G = [["業者探し", 0, 2, "協力業者・原価確認", "5B9BD5"], ["商品設計", 0, 2, "パック・オプション", "70AD47"], ["ブランド設計", 0, 2, "屋号・ロゴ・コンセプト", "4472C4"], ["HP制作", 0.4, 2.5, "TOP公開→全ページ公開", "ED7D31"], ["営業体制", 1.5, 6, "営業ツール・営業研修", "E6A700"], ["集客施策", 1.5, 6, "チラシ・WEB広告・SEO・MEO", "44546A"], ["アライアンス", 1.0, 3.0, "リストアップ・資料", "E05A5A"], ["", 3.8, 6, "提携に向けた営業活動", "E05A5A"]];
    let row = 0;
    G.forEach(([t, a, b, l, c]) => {
      if (t) row++;
      const y = 3.48 + (row - 1) * 0.42;
      if (t) { rect(s, 6.1, y, 6.28, 0.4, row % 2 ? CARD : WHITE); txt(s, t, 6.15, y, 1.4, 0.4, { size: 9.5, bold: true, color: DARK, valign: "middle" }); }
      rect(s, X0 + a * MW + 0.02, y + 0.06, (b - a) * MW - 0.04, 0.28, c, { round: true, r: 0.05 });
      txt(s, l, X0 + a * MW, y + 0.06, (b - a) * MW, 0.28, { size: 8, bold: true, color: WHITE, align: "center", valign: "middle" });
    });
    source(s, "※仮のスケジュールです。状況により変動いたします。", 6.4);
    txt(s, "マニュアル4種（事業概要・営業・現場・集客）と、見積ツール・トークスクリプト・業者間提携のご提案書などのツールも、貴社仕様でご提供します。", 0.75, 6.55, 11.8, 0.4, { size: 10.5, bold: true, color: NAVY, valign: "middle" });
  }
  {
    const s = page(5, C5, [B("まずは、"), B("貴社の商圏で試算", { color: RED }), B("してみませんか？")],
      "遺品整理の市場規模や、残置物撤去の内製化による効果は、商圏と貴社の案件数によって大きく変わります。個別の無料経営相談で、貴社の場合の数字を一緒に整理します。",
      "【案内】電話：0120-958-270（平日9:45〜17:30、土日祝・年末年始を除く）。WEB：https://www.funaisoken.co.jp/form/consulting");
    txt(s, "無料経営相談で分かること", 0.75, 2.72, 6, 0.32, { size: 12, bold: true, color: NAVY, valign: "middle" });
    const Q = [["map", "貴社商圏の遺品整理市場の規模", "商圏人口から市場規模と目標シェアを試算"], ["yen", "残置物撤去の内製化で残る利益", "今の外注件数・単価から自社化の効果を試算"], ["building", "不動産案件への波及効果", "売却・買取・解体につながる件数と粗利を整理"]];
    Q.forEach(([k, t, d], i) => {
      const x = 0.72 + i * 3.98;
      rect(s, x, 3.1, 3.7, 1.75, CARD, { round: true, r: 0.06 });
      num(s, i + 1, x + 0.18, 3.25, 0.44);
      icon(s, k, x + 3.0, 3.25, 0.5, BLUE);
      txt(s, t, x + 0.2, 3.8, 3.3, 0.45, { size: 12.5, bold: true, color: DARK, valign: "middle" });
      txt(s, d, x + 0.2, 4.25, 3.3, 0.55, { size: 10, color: GRAY, ls: 1.3 });
    });
    // 連絡先
    rect(s, 0.72, 5.08, 11.66, 1.55, WHITE, { line: LINE, round: true, r: 0.06 });
    txt(s, "お電話でのお問い合わせ・ご相談", 1.0, 5.18, 5.2, 0.32, { size: 11.5, bold: true, color: DARK, align: "center", valign: "middle" });
    icon(s, "phone", 1.55, 5.62, 0.42, BLUE);
    txt(s, "0120-958-270", 2.05, 5.52, 4.0, 0.62, { size: 28, bold: true, color: BLUE, valign: "middle" });
    txt(s, "【受付時間】平日 9時45分〜17時30分（土日祝、年末年始を除く）", 1.0, 6.15, 5.2, 0.3, { size: 8.5, color: GRAY, align: "center", valign: "middle" });
    s.addShape(pres.shapes.LINE, { x: 6.55, y: 5.25, w: 0, h: 1.2, line: { color: LINE, width: 0.75 }, objectName: nm("ln") });
    txt(s, "WEBでのお問い合わせ・ご相談", 6.85, 5.18, 5.3, 0.32, { size: 11.5, bold: true, color: DARK, align: "center", valign: "middle" });
    rect(s, 7.95, 5.58, 3.1, 0.55, NAVY, { round: true, r: 0.08 });
    txt(s, "無料 経営相談に申し込む", 7.95, 5.58, 3.1, 0.55, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, "https://www.funaisoken.co.jp/form/consulting", 6.85, 6.15, 5.3, 0.3, { size: 9, color: BLUE, align: "center", valign: "middle" });
    txt(s, "※本資料の数値は船井総研作成のモデル数値であり、成果を保証するものではありません。本資料には生成AIにより生成した情報が含まれている場合があります。", 0.75, 6.72, 11.8, 0.28, { size: 7.5, color: MGRAY, valign: "middle" });
  }

  await pres.writeFile({ fileName: OUT });
  if (process.env.PPTX_SKILL) {
    const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
    await applyTheme(OUT, THEME);
  }
  console.log("written:", OUT);
}
main().catch((e) => { console.error(e); process.exit(1); });
