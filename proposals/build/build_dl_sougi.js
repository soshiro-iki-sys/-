// 葬儀会社向けダウンロードレポート 生成スクリプト
// 使い方: PPTX_SKILL=<pptxスキルのパス> node build_dl_sougi.js <出力パス.pptx>
// トンマナ：「2027年 遺品整理業界の予測」レポートに準拠（16:9・白背景＋青いクリスタル装飾・游ゴシック・
//          大きな青い章番号＋見出し1行＋リード文2行＋紺の見出し帯／淡いグレーのカード）
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa");

const OUT = process.argv[2] || "葬儀会社向け_遺品整理DLレポート.pptx";
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
    db: "FaDatabase", lightbulb: "FaLightbulb", ok: "FaCheckCircle", ng: "FaTimesCircle", recycle: "FaRecycle", pray: "FaPrayingHands", heart: "FaHandHoldingHeart", idcard: "FaAddressCard",
  };
  const COLORS = [WHITE, BLUE, NAVY, RED, MGRAY, GRAY];
  const pairs = [];
  Object.values(ICONS).forEach((n) => COLORS.forEach((c) => pairs.push([n, c])));
  await prepIcons(pairs);
  const IC = (k, c) => ICON_CACHE[ICONS[k] + "_" + c];

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "葬儀会社のための遺品整理参入ガイド";
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
    txt(s, "葬儀会社様向けレポート", 2.05, 1.7, 2.9, 0.42, { size: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, [B("葬儀会社のための\n", { color: NAVY }), B("遺品整理", { color: BLUE }), B("参入ガイド", { color: NAVY })], 2.05, 2.3, 10.5, 1.9, { size: 42, ls: 1.2, valign: "middle" });
    hline(s, 2.1, 4.4, 9.6, LINE, 1);
    txt(s, "葬儀後サポート・会員組織を活かし、本業との相乗効果でアフター収益をつくる", 2.05, 4.58, 10.8, 0.55, { size: 19, bold: true, color: DARK, valign: "middle" });
    txt(s, "株式会社 船井総合研究所", 2.05, 6.3, 6, 0.4, { size: 12, bold: true, color: NAVY, valign: "middle" });
    s.addNotes("表紙。テーマは「葬儀後のアフター収益」。遺品整理を葬儀後サポートとして位置づけ、本業とのシナジーを伝える。");
  }

  // =================================================================
  // 2 目次
  // =================================================================
  {
    const s = pres.addSlide({ masterName: "CONTENTS" });
    const T = ["葬儀会社を取り巻く変化", "なぜ葬儀会社が遺品整理をやるべきか", "葬儀会社版 遺品整理事業モデル", "成功するための取組み", "船井総研がお手伝いできること"];
    T.forEach((t, i) => {
      const y = 0.72 + i * 1.07;
      txt(s, `${i + 1}.`, 3.45, y, 0.85, 0.7, { size: 40, bold: true, color: "8C8C8C", valign: "bottom" });
      txt(s, t, 4.25, y + 0.12, 8, 0.5, { size: 17, bold: true, color: DARK, valign: "middle" });
    });
  }

  // =================================================================
  // 1章
  // =================================================================
  const C1 = "葬儀会社を取り巻く変化";
  {
    const s = page(1, C1, [B("1事業所あたりの売上は、"), B("2015年から約26%減っている", { color: RED })],
      "経済産業省の統計を、葬儀を営む事業所1か所あたりに直すと、2024年の売上高は約2.0億円で、2015年の約2.8億円から約26%減りました。年間の取扱件数も約190件から約167件へ減り、1件あたりの単価も下がっています。",
      "【出典】経済産業省「特定サービス産業動態統計調査」（葬儀業）の長期データ。互助会保証株式会社「葬祭業の推移」（2025年6月）に掲載の年次値（売上高・取扱件数・調査企業の当該業務を営む事業所数）から、1事業所あたりの売上高（売上高÷事業所数）・取扱件数（取扱件数÷事業所数）を算出。\n【注意】調査対象は売上の大きい企業の事業所で、葬儀業全体ではない。事業所数は調査対象の追加でも増えるため、1事業所あたりの減少には調査範囲の変化の影響も含まれる。");
    txt(s, "葬儀を営む1事業所あたりの年間売上高", 0.75, 2.72, 7, 0.32, { size: 12, bold: true, color: NAVY, valign: "middle" });
    rect(s, 5.15, 2.7, 2.15, 0.36, WHITE, { line: RED, lw: 1.25, round: true, r: 0.2 });
    txt(s, [B("2015年比 ", { color: DARK }), B("約26%減", { color: RED })], 5.15, 2.7, 2.15, 0.36, { size: 12, align: "center", valign: "middle" });
    s.addChart(pres.charts.BAR, [{ name: "1事業所あたり売上高（億円）", labels: ["2015", "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024"], values: [2.75, 2.62, 2.59, 2.50, 2.40, 2.02, 1.92, 2.01, 2.04, 2.03] }], {
      x: 0.72, y: 3.15, w: 6.6, h: 2.95, barDir: "col", barGapWidthPct: 45, chartColors: [NAVY], showValue: true, dataLabelPosition: "outEnd",
      dataLabelFormatCode: "0.00", dataLabelFontSize: 10, dataLabelFontBold: true, dataLabelColor: NAVY, dataLabelFontFace: FONT,
      catAxisLabelFontSize: 10, catAxisLabelFontFace: FONT, catAxisLabelColor: DARK, valAxisHidden: true, valGridLine: { style: "none" },
      showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 3.2, objectName: "site_chart",
    });
    txt(s, "（単位：億円）", 5.6, 6.02, 1.7, 0.25, { size: 8.5, color: MGRAY, align: "right", valign: "middle" });
    txt(s, "1事業所あたりの推移", 7.65, 2.72, 5, 0.32, { size: 12, bold: true, color: NAVY, valign: "middle" });
    const HL = (t) => ({ text: t, options: { bold: true, color: BLUE, fill: { color: LAV } } });
    table(s, [["年", "年間売上高", "年間取扱件数", "1件あたり"], ["2015年", "2.75億円", "190件", "144.5万円"], ["2019年", "2.40億円", "179件", "134.3万円"], ["2020年", "2.02億円", "172件", "117.4万円"], ["2023年", "2.04億円", "172件", "118.5万円"], [HL("2024年"), HL("2.03億円"), HL("167件"), HL("121.5万円")]], 7.65, 3.08, [1.0, 1.3, 1.3, 1.12], { size: 10.5, rowH: 0.42, firstCol: true });
    conclusion(s, "1か所あたりの件数も単価も減っている。葬儀だけでは、売上を維持しにくい構造です。", 6.25);
    source(s, "出典：経済産業省「特定サービス産業動態統計調査」（葬儀業）をもとに作成。売上高・取扱件数を、調査対象の当該業務を営む事業所数で割って算出。調査対象は売上の大きい企業の事業所。", 6.7);
  }
  {
    const s = page(1, C1, [B("集客は紹介サイト頼みで、"), B("施行後の接点がない", { color: RED })],
      "葬儀の小規模化に加え、紹介サイト経由の受注が増え、価格で比べられる場面が多くなりました。葬儀が終わるとご遺族との接点が途切れ、会員やOB顧客も活かしきれていません。",
      "【話し方】葬儀会社の3つの悩み（単価・集客コスト・施行後の接点）に共感してもらうページ。吹き出しは「よくあるお悩み」の例。");
    const C = [
      ["yen", "① 葬儀単価の下落", ["家族葬化・価格競争で収益が減少", "会葬者が減り、返礼品・飲食も縮小"]],
      ["search", "② 集客コストの高騰", ["紹介サイトの手数料や広告費が上昇", "価格で比べられ、単価を上げにくい"]],
      ["users", "③ 施行後の接点不足", ["葬儀後、ご遺族との接点が途切れる", "会員・OB顧客を活かせていない"]],
    ];
    C.forEach(([k, t, b], i) => {
      const x = 0.72 + i * 3.98;
      headCard(s, x, 2.78, 3.7, 1.85, t);
      icon(s, k, x + 0.22, 3.48, 0.62, WHITE, BLUE);
      bullets(s, b, x + 1.0, 3.28, 2.55, 1.25, { size: 10.5, valign: "middle" });
    });
    txt(s, "よくあるお悩み", 0.75, 4.9, 3, 0.3, { size: 11, bold: true, color: NAVY, valign: "middle" });
    ["紹介サイト経由の葬儀は、手数料で利益が残らない", "施主から片付けを相談されても、紹介して終わり", "会員さんに、葬儀以外で提供できるものがない"].forEach((t, i) => {
      const x = 0.72 + i * 3.98;
      s.addShape(pres.shapes.ROUNDED_RECTANGULAR_CALLOUT, { x, y: 5.25, w: 3.7, h: 0.6, fill: { color: WHITE }, line: { color: LINE, width: 1 }, objectName: nm("bubble") });
      txt(s, `「${t}」`, x + 0.05, 5.25, 3.6, 0.52, { size: 10, bold: true, color: DARK, align: "center", valign: "middle" });
    });
    conclusion(s, "葬儀だけで売上を伸ばすのは難しい。葬儀の「後」に、ご遺族との接点をつくる必要があります。", 6.2);
  }
  {
    const s = page(1, C1, [B("ご遺族の相談は、"), B("葬儀の後にこそ続いている", { color: RED })],
      "多死社会・空き家・相続登記の義務化で、葬儀の後にご遺族が向き合う課題は増えています。実家の片付け、遺品の供養、相続、空き家。葬儀を任された会社は、その相談を最初に受けられる立場にあります。",
      "【出典】厚生労働省「令和7年人口動態統計（確定数）」、社人研「日本の将来推計人口」、総務省「令和5年住宅・土地統計調査」。相続登記は過去の相続分も対象。");
    const ST = [["2025年の年間死亡数", "158.9", "万人", "ピークは2040年（約167万人）の見込み"], ["空き家総数", "900", "万戸", "うち放置空き家は385万戸（空き家率13.8%）"], ["相続登記の申請義務", "3", "年以内", "過去の相続分も対象。怠ると10万円以下の過料"]];
    ST.forEach(([h, v, u, d], i) => {
      const x = 0.72 + i * 3.98;
      headCard(s, x, 2.72, 3.7, 1.62, h, 0.4);
      txt(s, [B(v, { fontSize: 30, color: BLUE }), B(" " + u, { fontSize: 13, color: BLUE })], x, 3.12, 3.7, 0.65, { align: "center", valign: "middle" });
      txt(s, d, x + 0.15, 3.8, 3.4, 0.42, { size: 10, color: GRAY, align: "center", valign: "middle" });
    });
    txt(s, "葬儀の後に、ご遺族が向き合うこと", 0.75, 4.5, 6, 0.3, { size: 11.5, bold: true, color: NAVY, valign: "middle" });
    const F = [["葬儀", NAVY], ["四十九日・法要", "5B6B9A"], ["遺品整理・実家の片付け", BLUE], ["相続・空き家", "7A7DFF"]];
    F.forEach(([t, c], i) => {
      const x = 0.72 + i * 2.95;
      chev(s, x, 4.88, 3.05, 0.62, c, i === 0);
      txt(s, t, x + (i === 0 ? 0.1 : 0.35), 4.88, 2.55, 0.62, { size: 12.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    });
    rect(s, 1.3, 5.62, 4.6, 0.42, WHITE, { line: NAVY, lw: 1.25, round: true, r: 0.2 });
    txt(s, "これまで、葬儀会社が関わってきたところ", 1.3, 5.62, 4.6, 0.42, { size: 10.5, bold: true, color: NAVY, align: "center", valign: "middle" });
    rect(s, 7.2, 5.62, 4.6, 0.42, WHITE, { line: BLUE, lw: 1.25, round: true, r: 0.2 });
    txt(s, "ご遺族の困りごとは、ここから続く", 7.2, 5.62, 4.6, 0.42, { size: 10.5, bold: true, color: BLUE, align: "center", valign: "middle" });
    conclusion(s, "葬儀を任された会社は、「その後」の相談を最初に受けられる立場にあります。", 6.3);
    source(s, "出典：厚生労働省「令和7年人口動態統計（確定数）」、社人研「日本の将来推計人口」、総務省「令和5年住宅・土地統計調査」", 6.75);
  }

  // =================================================================
  // 2章
  // =================================================================
  const C2 = "なぜ葬儀会社が遺品整理をやるべきか";
  {
    const s = page(2, C2, [B("ご遺族の"), B("「その後」", { color: RED }), B("を、自社で支えられる")],
      "葬儀を任されたご遺族は、その後の困りごとも「まずは葬儀会社に」と相談してくださいます。遺品整理を自社で持てば、葬儀から法要、片付け、供養まで、ご遺族を一貫して支えられます。",
      "【話し方】売り込むのではなく、ご遺族に寄り添うサポートを自社で完結できる、という点を強調する。相続・不動産は提携先（士業・不動産会社）へつなぐ。");
    const F = ["葬儀施行", "法要のご案内", "遺品整理・形見分け", "供養・お焚き上げ", "相続・不動産のご相談"];
    F.forEach((t, i) => {
      const x = 0.72 + i * 2.36;
      chev(s, x, 2.8, 2.45, 0.7, i === 2 ? BLUE : (i === 3 ? NAVY : "A6A9D8"), i === 0);
      txt(s, [B(`0${i + 1}\n`, { fontSize: 9 }), B(t, { fontSize: 11.5 })], x + (i === 0 ? 0.1 : 0.35), 2.8, 1.85, 0.7, { color: WHITE, align: "center", valign: "middle", ls: 1.05 });
    });
    txt(s, "※相続・不動産は、提携先（司法書士・税理士・不動産会社など）と連携して対応", 7.55, 3.55, 5, 0.26, { size: 8.5, color: MGRAY, align: "right", valign: "middle" });
    txt(s, "遺品整理に活きる、葬儀会社の4つの強み", 0.75, 3.85, 6, 0.32, { size: 12, bold: true, color: NAVY, valign: "middle" });
    const E = [["shake", "ご遺族との信頼関係", "葬儀を任された安心感があり、\n相談のハードルが低い"], ["pray", "供養の知識とつながり", "お焚き上げ・寺院とのつながりで、\n遺品を丁寧に供養できる"], ["heart", "ご遺族対応に慣れたスタッフ", "悲しみに寄り添う姿勢を、\n片付けの現場でも活かせる"], ["idcard", "会員組織・OB顧客", "葬儀以外のご案内で、\n会員様との接点を増やせる"]];
    E.forEach(([k, t, d], i) => {
      const x = 0.72 + i * 2.95;
      rect(s, x, 4.25, 2.8, 1.9, CARD, { round: true, r: 0.06 });
      icon(s, k, x + 1.08, 4.38, 0.62, WHITE, i % 2 ? NAVY : BLUE);
      txt(s, t, x + 0.1, 5.05, 2.6, 0.4, { size: 12, bold: true, color: DARK, align: "center", valign: "middle" });
      txt(s, d, x + 0.1, 5.45, 2.6, 0.62, { size: 9.5, color: GRAY, align: "center", ls: 1.3 });
    });
    band(s, "葬儀で信頼を得た会社だからこそ、「その後」も安心して任せていただけます。", 6.35, 0.5);
  }
  {
    const s = page(2, C2, [B("遺品整理で生まれる、"), B("2つの効果", { color: RED })],
      "一つは、葬儀の後に新しい売上が生まれること。もう一つは、これまで外部に紹介していた片付けを自社で行い、ご遺族からの信頼を守れることです。",
      "【話し方】①は考え方のみ（数字は入れない）。平均客単価30万円は船井総研のモデル数値。②は経営相談マニュアルの葬儀業の悩み「外部紹介で終わる／紹介先の雑な対応で長年の信頼を損なう懸念」に答える。");
    headCard(s, 0.72, 2.75, 5.6, 3.4, "① アフター売上の創出", 0.46);
    const FX = [["年間の施行件数", NAVY], ["遺品整理のご相談", BLUE], ["客単価", NAVY]];
    FX.forEach(([t, c], i) => {
      const x = 0.95 + i * 1.82;
      rect(s, x, 3.45, 1.5, 0.75, WHITE, { line: c, lw: 1.5, round: true, r: 0.1 });
      txt(s, t, x, 3.45, 1.5, 0.75, { size: 11.5, bold: true, color: c, align: "center", valign: "middle" });
      if (i < 2) txt(s, "×", x + 1.5, 3.45, 0.32, 0.75, { size: 16, bold: true, color: GRAY, align: "center", valign: "middle" });
    });
    txt(s, "※平均客単価30万円（船井総研モデル）", 0.95, 4.25, 5.1, 0.28, { size: 8.5, color: MGRAY, align: "right", valign: "middle" });
    arrowD(s, 3.32, 4.55, 0.4, 0.28, BLUE);
    rect(s, 1.2, 4.9, 4.6, 0.55, BLUE, { round: true, r: 0.15 });
    txt(s, "葬儀の後に生まれる、新しい売上", 1.2, 4.9, 4.6, 0.55, { size: 14, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, "供養・仏壇・生前整理などのオプションで、1件あたりの単価も上がります", 0.95, 5.55, 5.15, 0.5, { size: 10, color: GRAY, align: "center", ls: 1.3 });
    headCard(s, 6.78, 2.75, 5.6, 3.4, "② 紹介の品質を守る", 0.46);
    rect(s, 6.98, 3.38, 5.2, 1.15, "FFF0F0", { line: "F5B5B5", round: true, r: 0.06 });
    txt(s, "これまで：外部の業者を紹介", 7.15, 3.43, 4.9, 0.35, { size: 11, bold: true, color: RED, valign: "middle" });
    txt(s, "紹介先の雑な対応や追加請求があると、紹介した葬儀会社の信頼まで損なうおそれがある", 7.15, 3.78, 4.9, 0.7, { size: 10, color: GRAY, ls: 1.3 });
    arrowD(s, 9.38, 4.6, 0.4, 0.28, BLUE);
    rect(s, 6.98, 4.95, 5.2, 1.1, LAV, { line: "C9CBF5", round: true, r: 0.06 });
    txt(s, "これから：自社で遺品整理を提供", 7.15, 5.0, 4.9, 0.35, { size: 11, bold: true, color: BLUE, valign: "middle" });
    txt(s, "品質を自社で管理でき、会員特典として満足度・継続率の向上にもつながる", 7.15, 5.35, 4.9, 0.65, { size: 10, color: GRAY, ls: 1.3 });
    band(s, "アフター売上をつくりながら、ご遺族からの信頼も守れます。", 6.45, 0.48);
  }
  {
    const s = page(2, C2, [B("2027年は、"), B("「地域一番店」が不在", { color: RED }), B("の成長期")],
      "遺品整理業者は葬儀でのご遺族との関係を持たず、葬儀会社は片付けを持っていません。両方を持つ会社は、まだ地域にほとんどありません。ここから1〜2年で勢力図が固まります。",
      "【出典】2027年遺品整理業界の予測レポート（ライフサイクル・シェア区分）。MS＝国民1人あたりの遺品整理消費額4,200円、市場規模は商圏人口20万人×MSで算出した船井総研のモデル数値。");
    const L = [["導入期", "一部の専門業者のみ"], ["成長期【現在〜2027年】", "地域一番店が不在。\n先行者がシェアを取れる"], ["成長後期", "大手の参入で\n価格競争が激化"], ["成熟期", "供給過多。後発からの\n一番店化は相当困難"]];
    L.forEach(([t, d], i) => {
      const x = 0.72 + i * 2.95, on = i === 1;
      chev(s, x, 2.75, 3.05, 1.0, on ? BLUE : CARD, i === 0);
      txt(s, [B(t + "\n", { fontSize: 12.5, color: on ? WHITE : DARK }), N(d, { fontSize: 9.5, color: on ? WHITE : GRAY })], x + (i === 0 ? 0.1 : 0.45), 2.75, 2.4, 1.0, { align: "center", valign: "middle", ls: 1.1 });
    });
    txt(s, "葬儀と遺品整理、両方を持つ会社はほとんどいない", 0.75, 3.95, 5.6, 0.32, { size: 11.5, bold: true, color: NAVY, valign: "middle" });
    const O = { text: "○", options: { bold: true, color: BLUE, fontSize: 14 } }, X = { text: "×", options: { bold: true, color: MGRAY, fontSize: 14 } };
    table(s, [["", "遺品整理業者", "葬儀会社", "両方を持つ会社"], ["ご遺族との信頼関係（葬儀）", X, O, O], ["遺品整理・片付け", O, X, O], [{ text: "地域での数", options: { bold: true, color: DARK, fill: { color: LAV }, align: "center" } }, "増加中", "多数", { text: "ほぼ不在", options: { bold: true, color: RED } }]], 0.72, 4.35, [2.3, 1.1, 1.05, 1.15], { size: 10, rowH: 0.42, firstCol: true });
    txt(s, "商圏人口20万人（市場規模 約8.4億円）でのシェア別売上", 6.6, 3.95, 5.9, 0.32, { size: 11.5, bold: true, color: NAVY, valign: "middle" });
    const SH = [["トップシェア", "2億1,840万円", "26.00%"], ["影響シェア", "9,960万円", "11.86%"], ["存在シェア", "5,400万円", "6.43%"], ["拠点シェア", "3,600万円", "4.29%"]];
    table(s, [["シェア区分", "売上（年間）", "シェア"], ...SH.map((r, i) => i === 0 ? r.map((c) => ({ text: c, options: { bold: true, color: BLUE, fill: { color: LAV } } })) : r)], 6.6, 4.35, [2.2, 2.0, 1.6], { size: 10.5, rowH: 0.34 });
    source(s, "※MS（国民1人あたりの遺品整理消費額）4,200円×商圏人口20万人で算出。" + MODEL, 6.12);
    band(s, "地域一番店が不在の2027年が、シェアを取りにいく最大の好機です。", 6.5, 0.48);
  }

  // =================================================================
  // 3章
  // =================================================================
  const C3 = "葬儀会社版 遺品整理事業モデル";
  {
    const s = page(3, C3, [B("葬儀単価に左右されない、"), B("もう一つの収益の柱", { color: RED })],
      "商圏人口20万人・平均客単価30万円で、3年目に年商1億円を目指すモデルです。葬儀会社は、会員名簿・スタッフ・車両・会館といった既存の資産を活かして立ち上げられます。",
      "【出典】PLは2027年予測レポート・企画書と同じモデル数値。");
    txt(s, "3年間の数値計画", 0.75, 2.72, 5, 0.32, { size: 12, bold: true, color: NAVY, valign: "middle" });
    s.addChart(pres.charts.BAR, [{ name: "売上（万円）", labels: ["1年目", "2年目", "3年目"], values: [3600, 5400, 10000] }], {
      x: 0.72, y: 3.05, w: 5.6, h: 2.6, barDir: "col", barGapWidthPct: 45, chartColors: [NAVY], showValue: true, dataLabelPosition: "outEnd",
      dataLabelFormatCode: "[>=10000]\"1億円\";#,##0\"万円\"", dataLabelFontSize: 12, dataLabelFontBold: true, dataLabelColor: NAVY, dataLabelFontFace: FONT,
      catAxisLabelFontSize: 11, catAxisLabelFontFace: FONT, catAxisLabelColor: DARK, valAxisHidden: true, valGridLine: { style: "none" },
      showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 12000, objectName: "sales_chart",
    });
    txt(s, "営業利益率", 0.75, 5.72, 1.1, 0.36, { size: 10, color: GRAY, valign: "middle" });
    [["2.9%", 1.95], ["16.7%", 3.48], ["23.1%", 5.0]].forEach(([v, x], i) => {
      rect(s, x, 5.72, 1.15, 0.36, i === 2 ? BLUE : LAV, { round: true, r: 0.15 });
      txt(s, v, x, 5.72, 1.15, 0.36, { size: 12, bold: true, color: i === 2 ? WHITE : BLUE, align: "center", valign: "middle" });
    });
    txt(s, "葬儀会社が活かせる、4つの資産", 6.8, 2.72, 5.5, 0.32, { size: 12, bold: true, color: NAVY, valign: "middle" });
    const AS = [["idcard", "会員名簿・OB顧客", "会員様へのご案内・会員特典で、立上げ初期の受注に"], ["heart", "スタッフ", "ご遺族対応・グリーフケアの経験を、片付けの現場でも"], ["truck", "車両", "搬出・運搬に活用し、初期投資を抑えられる"], ["building", "会館", "相談会・終活セミナーの会場として集客に"]];
    AS.forEach(([k, t, d], i) => {
      const y = 3.1 + i * 0.75;
      rect(s, 6.78, y, 5.6, 0.66, CARD, { round: true, r: 0.08 });
      icon(s, k, 6.92, y + 0.1, 0.46, WHITE, i % 2 ? NAVY : BLUE);
      txt(s, [B(t + "\n", { fontSize: 11.5, color: DARK }), N(d, { fontSize: 9.5, color: GRAY })], 7.5, y, 4.8, 0.66, { valign: "middle", ls: 1.15 });
    });
    source(s, MODEL, 6.18);
    band(s, "既存の資産を活かせる葬儀会社は、低リスクで早く立ち上げられます。", 6.5, 0.48);
  }
  {
    const s = page(3, C3, [B("「内訳」ではなく"), B("「パック料金」", { color: RED }), B("で売る")],
      "お客様が知りたいのは「どれくらいの量で、おおよそいくらになるのか」です。交通費・作業費・処理費などの内訳を並べると、かえって不安を招きます。間取り×物量で総額が分かるパック料金にします。",
      "【出典】2027年予測レポートのパック料金表・オプション。葬儀会社ならではのオプション（仏壇・位牌の供養、形見分け、生前整理）は料金を「ご相談」とした。会員特典の割引率などは各社で設定。");
    txt(s, "間取り・物量別 パック料金表（基本料金の例）", 0.75, 2.72, 7, 0.32, { size: 12, bold: true, color: NAVY, valign: "middle" });
    table(s, [["間取り", "基本物量の目安", "作業人数", "作業時間の目安", "料金（税込）"], ["1K", "段ボール20箱程度", "1名", "1〜3時間", "30,000円〜"], ["1LDK", "段ボール50箱程度", "2〜3名", "3〜6時間", "80,000円〜"], ["2LDK", "段ボール100箱程度", "3〜4名", "5〜8時間", "150,000円〜"], ["3LDK", "段ボール150箱程度", "4〜6名", "6〜9時間", "200,000円〜"], ["4LDK以上", "段ボール200箱以上", "5名〜", "1日〜2日", "300,000円〜"]], 0.72, 3.08, [1.15, 2.0, 1.15, 1.6, 1.6], { size: 10.5, rowH: 0.34, firstCol: true });
    txt(s, "仕分け・搬出・処分・簡易清掃・運搬まで含んだパック料金", 0.75, 5.17, 7.5, 0.26, { size: 9, color: MGRAY, valign: "middle" });
    txt(s, "オプション（必要な場合のみ追加）", 0.75, 5.48, 7, 0.3, { size: 12, bold: true, color: NAVY, valign: "middle" });
    const OP = [["供養（お焚き上げ）", "15,000円〜", 1], ["仏壇・位牌の供養", "ご相談", 1], ["形見分けのお手伝い", "ご相談", 1], ["生前整理", "ご相談", 1], ["特殊清掃・脱臭", "30,000円〜", 0]];
    OP.forEach(([t, v, hl], i) => {
      const x = 0.72 + i * 1.52;
      rect(s, x, 5.82, 1.44, 0.62, hl ? LAV : CARD, { round: true, r: 0.08, line: hl ? BLUE : undefined });
      txt(s, [B(t + "\n", { fontSize: 9.5, color: DARK }), B(v, { fontSize: 10, color: BLUE })], x, 5.82, 1.44, 0.62, { align: "center", valign: "middle", ls: 1.1 });
    });
    txt(s, "■ 葬儀会社ならではのオプション。会員様向けの特典（割引など）も設定できます", 0.75, 6.5, 7.5, 0.26, { size: 9, bold: true, color: BLUE, valign: "middle" });
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
