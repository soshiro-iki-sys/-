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
const DARK = "1A1A1A", GRAY = "262626", MGRAY = "4D4D4D", LINE = "C8CBD6", WHITE = "FFFFFF", YEL = "FFF100";
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
  pres.title = "不動産会社のための遺品整理参入ガイド";
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
    if (head) txt(s, head, 0.8, 1.42, 11.9, 0.5, { size: 19, bold: true, color: DARK, valign: "middle" });
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
    rect(s, 2.05, 1.7, 2.9, 0.42, BLUE, { round: true, r: 0.2 });
    txt(s, "不動産会社様向けレポート", 2.05, 1.7, 2.9, 0.42, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, [B("不動産会社のための\n", { color: NAVY }), B("遺品整理", { color: BLUE }), B("参入ガイド", { color: NAVY })], 2.05, 2.3, 10.5, 1.9, { size: 42, ls: 1.2, valign: "middle" });
    hline(s, 2.1, 4.4, 9.6, LINE, 1);
    txt(s, "相続物件と“市場に出る前”に出会い、残置物撤去の内製化と仕入れ強化を同時に実現する", 2.05, 4.58, 10.8, 0.55, { size: 19, bold: true, color: DARK, valign: "middle" });
    txt(s, "株式会社 船井総合研究所", 2.05, 6.3, 6, 0.4, { size: 12, bold: true, color: NAVY, valign: "middle" });
    s.addNotes("表紙。テーマは「相続物件の仕入れルート」。遺品整理を“本業の入口”として位置づけ、残置物撤去の内製化と仕入れ強化のシナジーを伝える。");
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
  // 4章（2027年予測レポートの p6・p11・p12 と同じ内容）
  // =================================================================
  const C4 = "成功するための取組み";
  {
    const s = page(4, C4, null, null, "【出典】2027年予測レポート p6「2027年 実施いただきたいこと」と同じ内容。");
    txt(s, "これからやるべきこと　それは、", 0.8, 1.25, 11, 0.45, { size: 19, bold: true, color: DARK, valign: "middle" });
    txt(s, "「遺品整理専門店」を立ち上げ、地域一番店を取りにいくこと", 0.8, 1.78, 11.5, 0.5, { size: 21, bold: true, color: RED, valign: "middle" });
    txt(s, "便利屋の延長ではなく、遺品整理に絞り込んだ「専門店」として打ち出すことが出発点です。顧客が求める「分かりやすさ」「納得感」「グリーフケア（心のケア）」に応えるため、売り先・売り物・売り方の3要素を一致させます。", 0.8, 2.3, 11.75, 0.62, { size: 11.5, color: GRAY, ls: 1.45 });
    const T3 = [["① 売り先", "（ターゲット・商圏）", ["実家を相続する50〜60代の一般客", "賃貸オーナーなどの法人（貸室の残置物）", "商圏人口20万人エリア"]], ["② 売り物", "（商品・価格）", ["遺品整理標準パック（仕分け・搬出・処分・供養・簡易清掃）", "価格帯：5万円（1K）〜150万円（4LDK）"]], ["③ 売り方", "（販促・営業体制）", ["チラシ・WEB・アライアンス先からの紹介で集客", "現地確認＆即時クロージング型の営業"]]];
    T3.forEach(([t, sub, b], i) => {
      const x = 0.72 + i * 3.98;
      rect(s, x, 3.03, 3.7, 1.5, CARD, { round: true, r: 0.06 });
      rect(s, x, 3.03, 3.7, 0.42, NAVY, { round: true, r: 0.12 });
      txt(s, [B(t + " ", { fontSize: 13 }), N(sub, { fontSize: 9.5 })], x, 3.03, 3.7, 0.42, { color: WHITE, align: "center", valign: "middle" });
      bullets(s, b, x + 0.2, 3.55, 3.35, 0.95, { size: 10.5, ls: 1.3 });
    });
    txt(s, "2027年に実施いただきたい5つのこと", 0.75, 4.65, 8, 0.32, { size: 12.5, bold: true, color: NAVY, valign: "middle" });
    const F5 = [["専門店ロゴ・屋号をつくる", "草引き・剪定・害虫駆除などは打ち出さず、遺品整理の客層に絞る", 1], ["専門店HPを立ち上げる", "単価の低い便利屋商品を外し、遺品整理・不用品回収・買取の3つの受け皿をつくる", 1], ["価格を明瞭にする", "「基本パック＋オプション」で、間取りと物量から総額が分かる料金表にする", 0], ["5つの施策で集客する", "SEO・MEO・WEB広告・チラシ・アライアンス。広告宣伝費は粗利の15〜20%が目安", 0], ["即決できる営業の仕組みをつくる", "0次接客・営業ツール・顧客管理シートで、契約率80%を目指す", 0]];
    F5.forEach(([t, d, must], i) => {
      const x = 0.72 + i * 2.38;
      rect(s, x, 5.02, 2.2, 1.95, WHITE, { line: "C9CBF5", round: true, r: 0.06 });
      num(s, i + 1, x + 0.15, 5.14, 0.42);
      if (must) { rect(s, x + 1.45, 5.18, 0.62, 0.28, TAGRED, { round: true, r: 0.06 }); txt(s, "必須", x + 1.45, 5.18, 0.62, 0.28, { size: 9, bold: true, color: WHITE, align: "center", valign: "middle" }); }
      txt(s, t, x + 0.15, 5.62, 1.95, 0.5, { size: 11, bold: true, color: NAVY, valign: "top", ls: 1.15 });
      txt(s, d, x + 0.15, 6.12, 1.95, 0.8, { size: 9.5, color: GRAY, ls: 1.25 });
    });
  }
  {
    const s = page(4, C4, "集客：「反響が取れる専門店HP」を土台に、5つの施策で安定的に案件を獲得する",
      "集客の一丁目一番地は、専門店の屋号とHPです。そのうえで5つの施策を組み合わせ、反響を最大化します。広告宣伝費は粗利の15〜20%が目安です（2年目モデルで年726万円）。",
      "【出典】2027年予測レポート p11と同じ内容。");
    txt(s, "反響が取れるHPの4つの条件", 0.75, 2.72, 6, 0.32, { size: 12.5, bold: true, color: NAVY, valign: "middle" });
    const H4 = [["1. 迷わない導線設計", "トップで専門性・料金・対応エリアを伝え、相談まで一直線に"], ["2. 選びやすい商品設計", "間取り別のパック料金で「何を・いくらで」が一目でわかる"], ["3. 安心できる信頼設計", "施工事例・お客様の声・Google口コミ・スタッフ・許可を明示"], ["4. 問い合わせしやすいUI", "スマホで電話・LINE・見積フォームにすぐ届く"]];
    H4.forEach(([t, d], i) => {
      const x = 0.72 + i * 2.95;
      rect(s, x, 3.08, 2.82, 0.85, LAV, { round: true, r: 0.06 });
      txt(s, [B(t + "\n", { fontSize: 11.5, color: NAVY }), N(d, { fontSize: 9.5, color: GRAY })], x + 0.15, 3.08, 2.6, 0.85, { valign: "middle", ls: 1.2 });
    });
    txt(s, "5つの集客施策", 0.75, 4.08, 6, 0.32, { size: 12.5, bold: true, color: NAVY, valign: "middle" });
    const P5 = [["01 SEO対策", "検索で見つけてもらう", ["お役立ち記事の作成", "地域ページ・事例・FAQの充実", "長期的に反響が増える「資産型」の集客"]], ["02 MEO対策", "地域の地図検索で選ばれる", ["Googleビジネスプロフィールの最適化", "写真・投稿の定期更新", "口コミの獲得", "SEOより上位表示しやすい"]], ["03 WEB広告・SNS", "今すぐ客を集める", ["Google検索広告で顕在客に", "Meta広告で検討前の層に", "Threadsでビフォーアフターと費用を発信（他社例：3か月で反響月20〜30件）"]], ["04 チラシ・販促", "地域に認知してもらう", ["「遺品整理」「保存版」を大きく", "料金表と追加料金の考え方を掲載", "配布数は世帯数×80%、エリア別反響率で頻度を調整"]], ["05 アライアンス提携", "紹介で安定的に獲得", ["不動産・葬儀・司法書士・ケアマネ・解体などと相互紹介", "提携先の信用を守ることが最大の営業フック"]]];
    P5.forEach(([t, sub, b], i) => {
      const x = 0.72 + i * 2.38;
      rect(s, x, 4.45, 2.22, 2.5, WHITE, { line: "C9CBF5", round: true, r: 0.06 });
      rect(s, x, 4.45, 2.22, 0.62, NAVY, { round: true, r: 0.1 });
      txt(s, [B(t + "\n", { fontSize: 12 }), N(sub, { fontSize: 9 })], x, 4.45, 2.22, 0.62, { color: WHITE, align: "center", valign: "middle", ls: 1.05 });
      bullets(s, b, x + 0.12, 5.15, 2.0, 1.75, { size: 9.5, ls: 1.25 });
    });
  }
  {
    const s = page(4, C4, "営業：案件規模で営業方法を分け、契約率80%・月15組の契約を実現する",
      "契約率を高めるカギは、0次接客（電話ヒアリング）から即決までの流れを仕組みにすること。さらに遺品整理を入口に関連サービスを段階的に提案し、1件あたりの粗利を最大化します。",
      "【出典】2027年予測レポート p12と同じ内容。");
    const SL = [["30万円未満の案件", "オンライン中心で効率化", BLUE, ["0次接客（電話ヒアリング）", "アンケート送付（SMS・LINE）", "写真の提出（お客様）", "オンライン見積", "電話でクロージング・契約", "作業実施"], "訪問を省いたコスト削減を価格に反映。即日見積で契約率アップ"], ["30万円以上の案件", "現地訪問で信頼と提案力を強化", NAVY, ["0次接客（電話ヒアリング）", "事前の情報収集・日程調整", "現地確認（訪問）", "見積提示・追加提案", "その場でのクロージング・契約", "作業実施"], "コミコミ価格と作業スケジュールを明示し、買取・供養・解体などを追加提案"]];
    SL.forEach(([t, sub, c, st, pt], i) => {
      const x = 0.72 + i * 3.88;
      rect(s, x, 2.78, 3.7, 4.12, CARD, { round: true, r: 0.06 });
      rect(s, x, 2.78, 3.7, 0.62, c, { round: true, r: 0.1 });
      txt(s, [B(t + "\n", { fontSize: 13 }), N(sub, { fontSize: 9.5 })], x, 2.78, 3.7, 0.62, { color: WHITE, align: "center", valign: "middle", ls: 1.05 });
      st.forEach((v, k) => {
        txt(s, [B(`STEP${k + 1}`, { color: BLUE, fontSize: 10 }), N("　" + v, { color: DARK, fontSize: 10.5 })], x + 0.2, 3.52 + k * 0.4, 3.4, 0.36, { valign: "middle" });
      });
      txt(s, [B("ポイント：", { color: NAVY }), N(pt, { color: GRAY })], x + 0.2, 6.0, 3.35, 0.8, { size: 9.5, ls: 1.3 });
    });
    txt(s, "遺品整理から広がる7つの収益機会", 8.5, 2.72, 4, 0.32, { size: 12.5, bold: true, color: NAVY, valign: "middle" });
    const R7 = ["リフォーム・リノベーション", "不動産売却・買取サポート", "解体工事", "特殊清掃・消臭・除菌", "供養・お焚き上げ", "買取・リユース", "遺品整理（基本サービス）"];
    R7.forEach((t, i) => {
      const n = 7 - i, w = 2.35 + n * 0.25, y = 3.1 + i * 0.5;
      rect(s, 8.5, y, w, 0.42, n === 1 ? NAVY : (n >= 5 ? BLUE : "6F73F5"), { round: true, r: 0.05 });
      txt(s, `${n}  ${t}`, 8.62, y, w - 0.15, 0.42, { size: 10.5, bold: true, color: WHITE, valign: "middle" });
    });
    txt(s, "顧客の課題をワンストップで解決し、粗利を最大化する", 8.5, 6.6, 4.3, 0.3, { size: 10.5, bold: true, color: NAVY, valign: "middle" });
  }

  // =================================================================
  // 5章（タイトルのみ・中身は別途差し込み）
  // =================================================================
  const C5 = "船井総研がお手伝いできること";
  page(5, C5, null, null, "【白紙】ご支援内容のページを差し込む。");
  page(5, C5, null, null, "【白紙】ご支援内容・無料経営相談のご案内を差し込む。");

  await pres.writeFile({ fileName: OUT });
  if (process.env.PPTX_SKILL) {
    const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
    await applyTheme(OUT, THEME);
  }
  console.log("written:", OUT);
}
main().catch((e) => { console.error(e); process.exit(1); });
