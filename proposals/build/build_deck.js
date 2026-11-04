// 遺品整理事業 参入のご提案（企画書）生成スクリプト
// 使い方: PPTX_SKILL=<pptxスキルのパス> node build_deck.js <出力パス.pptx>
// 構成：承認済みの55枚版の流れ ＋ 元の企画書（76p・96p）から不足情報を補足
// デザイン：船井総研の企画書の型（4:3・紺の見出し帯・紺×黄の中扉）＋アイコン・図解・ネイティブグラフ
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa");

const OUT = process.argv[2] || "遺品整理事業_企画書.pptx";

// ---------- パレット ----------
const NAVY = "1F2D50", NAVY2 = "2B3F6E", BLUE = "2E5FA8", LBLUE = "EAF0F8", SKY = "D6E3F3";
const RED = "D7261E", YEL = "FFF000", MARK = "FFF2A8", GOLD = "C9A227", ORANGE = "E07B2E", LORANGE = "FDF0E4";
const GREEN = "2E7D4F", LGREEN = "EAF4EE", PURPLE = "5B3F8C", BROWN = "7A5230", LBROWN = "F7F1EA";
const GRAY = "666666", MGRAY = "8A8F99", LGRAY = "F3F4F6", LINE = "C9D1DE", DARK = "222222", WHITE = "FFFFFF";
const FONT = "Meiryo", MINCHO = "Yu Mincho";
const MODEL_NOTE = "※船井総研作成のモデル数値です。成果を保証するものではありません。";

const THEME = {
  name: "Ihin Seiri Proposal", headFontFace: FONT, bodyFontFace: FONT,
  colors: { dk1: "000000", lt1: "FFFFFF", dk2: NAVY, lt2: "F3F4F6", accent1: NAVY, accent2: BLUE, accent3: RED, accent4: ORANGE, accent5: GREEN, accent6: PURPLE, hlink: BLUE, folHlink: PURPLE },
};

// ---------- アイコン（react-icons → PNG） ----------
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
    box: "FaBoxOpen", dolly: "FaDolly", truck: "FaTruck", recycle: "FaRecycle", heart: "FaHandHoldingHeart", clock: "FaClock", search: "FaSearch",
    users: "FaUsers", home: "FaHome", law: "FaBalanceScale", chart: "FaChartLine", warn: "FaExclamationTriangle", shake: "FaHandshake",
    building: "FaBuilding", hammer: "FaHammer", tools: "FaTools", pray: "FaPrayingHands", broom: "FaBroom", gem: "FaGem", yen: "FaYenSign", phone: "FaPhoneAlt",
    check: "FaClipboardCheck", list: "FaClipboardList", shield: "FaShieldAlt", contract: "FaFileContract", map: "FaMapMarkedAlt", store: "FaStore",
    paper: "FaNewspaper", user: "FaUserTie", bulb: "FaLightbulb", ok: "FaCheckCircle", laptop: "FaLaptop", horn: "FaBullhorn", friends: "FaUserFriends",
    talk: "FaComments", cal: "FaCalendarAlt", pdf: "FaFilePdf", xls: "FaFileExcel", ppt: "FaFilePowerpoint", doc: "FaFileWord", video: "FaFileVideo",
    img: "FaFileImage", line: "FaLine", mail: "FaEnvelope", coins: "FaCoins", key: "FaKey", hospital: "FaClinicMedical", city: "FaCity",
    link: "FaLink", award: "FaAward", target: "FaBullseye", wrench: "FaWrench", question: "FaQuestionCircle", rocket: "FaRocket", people: "FaPeopleCarry", photo: "FaImages", bag: "FaShoppingBag", ring: "FaRing", couch: "FaCouch", tv: "FaTv", palette: "FaPalette", guitar: "FaGuitar", camera: "FaCamera", robot: "FaRobot", thumbs: "FaThumbsUp", file: "FaFileAlt", book: "FaBookOpen",
  };
  const COLORS = [WHITE, NAVY, BLUE, RED, GREEN, ORANGE, PURPLE, BROWN, GOLD, MGRAY, YEL];
  const pairs = [];
  Object.values(ICONS).forEach((n) => COLORS.forEach((c) => pairs.push([n, c])));
  await prepIcons(pairs);
  const IC = (key, color) => ICON_CACHE[ICONS[key] + "_" + color];

  const pres = new pptxgen();
  pres.defineLayout({ name: "STD_43", width: 10, height: 7.5 });
  pres.layout = "STD_43";
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "遺品整理事業 参入のご提案";
  pres.author = "株式会社船井総合研究所";
  pres.company = "株式会社船井総合研究所";

  pres.defineSlideMaster({
    title: "CONTENT", background: { color: WHITE },
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
    title: "DIVIDER", background: { color: NAVY },
    objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.6, y: 2.5, w: 8.8, h: 2.5, fontFace: FONT, fontSize: 40, bold: true, color: YEL, align: "center", valign: "middle" }, text: "" } }],
  });
  pres.defineSlideMaster({ title: "BLANK", background: { color: WHITE }, objects: [] });

  // ---------- 基本部品 ----------
  let objId = 0;
  const nm = (p) => `${p}_${++objId}`;
  function content(section, title, notes) {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
    s.addText(title, { placeholder: "title" });
    if (notes) s.addNotes(notes);
    s._pageTitle = title;
    return s;
  }
  function divider(section, title, notes) {
    const s = pres.addSlide({ masterName: "DIVIDER", sectionTitle: section });
    s.addText(title, { placeholder: "title" });
    if (notes) s.addNotes(notes);
    return s;
  }
  function rect(s, x, y, w, h, fill, opt = {}) {
    s.addShape(opt.round ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE, {
      x, y, w, h, fill: fill ? { color: fill } : { type: "none" }, line: opt.line ? { color: opt.line, width: opt.lw || 1, dashType: opt.dash || "solid" } : { type: "none" },
      rectRadius: opt.round ? (opt.r || 0.08) : undefined, objectName: nm("box"),
      shadow: opt.shadow ? { type: "outer", color: "000000", opacity: 0.18, blur: 5, offset: 2, angle: 90 } : undefined,
    });
  }
  function oval(s, x, y, w, h, fill, opt = {}) {
    s.addShape(pres.shapes.OVAL, { x, y, w, h, fill: { color: fill }, line: opt.line ? { color: opt.line, width: opt.lw || 1 } : { type: "none" }, objectName: nm("oval") });
  }
  function txt(s, t, x, y, w, h, opt = {}) {
    if (Array.isArray(t)) {
      const out = [];
      t.forEach((r) => {
        const parts = r.text.split("\n");
        parts.forEach((p, k) => {
          const o = Object.assign({ align: opt.align || "left" }, r.options);
          if (k < parts.length - 1) o.breakLine = true;
          if (p !== "" || k < parts.length - 1) out.push({ text: p, options: o });
        });
      });
      t = out;
    }
    s.addText(t, {
      x, y, w, h, fontFace: opt.font || FONT, fontSize: opt.size || 12, bold: !!opt.bold, color: opt.color || DARK,
      align: opt.align || "left", valign: opt.valign || "top", margin: opt.margin === undefined ? 0.04 : opt.margin,
      lineSpacingMultiple: opt.ls || 1.12, paraSpaceAfter: opt.psa || 0, isTextBox: true, objectName: nm("text"), fit: "none",
    });
  }
  const B = (t, o = {}) => ({ text: t, options: Object.assign({ bold: true }, o) });
  const HL = (t, o = {}) => ({ text: t, options: Object.assign({ bold: true, highlight: MARK }, o) });
  function icon(s, k, x, y, d, color, bg) {
    if (bg) {
      oval(s, x, y, d, d, bg);
      const p = d * 0.22;
      s.addImage({ data: IC(k, color), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: nm("icon") });
    } else {
      s.addImage({ data: IC(k, color), x, y, w: d, h: d, objectName: nm("icon") });
    }
  }
  function num(s, n, x, y, d, fill, color) {
    oval(s, x, y, d, d, fill);
    txt(s, String(n), x, y, d, d, { size: Math.round(d * 28), bold: true, color: color || WHITE, align: "center", valign: "middle", margin: 0 });
  }
  function arrowR(s, x, y, w, h, color) { s.addShape(pres.shapes.RIGHT_ARROW, { x, y, w, h, fill: { color: color || NAVY }, line: { type: "none" }, objectName: nm("arrow") }); }
  function arrowD(s, x, y, w, h, color) { s.addShape(pres.shapes.DOWN_ARROW, { x, y, w, h, fill: { color: color || NAVY }, line: { type: "none" }, objectName: nm("arrow") }); }
  function chevron(s, x, y, w, h, fill, t, opt = {}) {
    s.addShape(opt.first ? pres.shapes.PENTAGON : pres.shapes.CHEVRON, { x, y, w, h, fill: { color: fill }, line: { type: "none" }, objectName: nm("chev") });
    txt(s, t, x + (opt.first ? 0.05 : h * 0.35), y, w - (opt.first ? h * 0.4 : h * 0.75), h, { size: opt.size || 11, bold: true, color: opt.color || WHITE, align: "center", valign: "middle", margin: 0 });
  }
  function line(s, x, y, w, h, color, width, opt = {}) {
    s.addShape(pres.shapes.LINE, { x, y, w, h, flipV: !!opt.flipV, line: { color, width: width || 1.5, dashType: opt.dash || "solid", endArrowType: opt.arrow ? "triangle" : undefined }, objectName: nm("line") });
  }
  // ページ見出し（中央揃え）：ページ名＋その下にメッセージ
  function msg(s, t) {
    txt(s, s._pageTitle, 0.4, 0.9, 9.2, 0.38, { size: 18, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    txt(s, t, 0.4, 1.27, 9.2, 0.28, { size: 11.5, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
  }
  function key(s, t, opt = {}) {
    txt(s, t, 0.3, opt.y || 6.45, 9.4, opt.h || 0.55, { size: opt.size || 18, bold: true, color: RED, align: "center", valign: "middle", margin: 0 });
  }
  function foot(s, t) { txt(s, t, 2.3, 7.12, 6.85, 0.3, { size: 8, color: GRAY, align: "right", valign: "middle", margin: 0 }); }
  function photo(s, x, y, w, h, label) {
    rect(s, x, y, w, h, "E8EBF0", { line: "C2C8D2" });
    icon(s, "photo", x + w / 2 - 0.22, y + h / 2 - 0.38, 0.44, MGRAY);
    txt(s, label, x, y + h / 2 + 0.08, w, 0.4, { size: 9.5, color: MGRAY, align: "center", margin: 0 });
  }
  function table(s, rows, x, y, w, colW, opt = {}) {
    const fs = opt.size || 11;
    const data = rows.map((r, i) => r.map((cell, j) => {
      const isObj = typeof cell === "object" && cell !== null;
      const base = { fontFace: FONT, fontSize: fs, valign: "middle", margin: opt.margin || [3, 6, 3, 6], border: { type: "solid", pt: 0.75, color: LINE }, align: j === 0 ? "left" : (opt.align || "center") };
      if (i === 0) Object.assign(base, { bold: true, color: WHITE, fill: { color: opt.head || NAVY }, align: "center" });
      else Object.assign(base, { color: DARK, fill: { color: (i % 2 === 0) ? LGRAY : WHITE } });
      if (opt.firstColFill && j === 0 && i > 0) Object.assign(base, { bold: true, fill: { color: LBLUE }, color: NAVY });
      if (isObj) Object.assign(base, cell.options || {});
      return { text: isObj ? cell.text : String(cell), options: base };
    }));
    s.addTable(data, { x, y, w, colW, rowH: opt.rowH, autoPage: false, objectName: nm("table") });
  }
  function frame(s, x, y, w, h, head, color, opt = {}) {
    rect(s, x, y, w, h, opt.bg || WHITE, { line: color, lw: 1.25 });
    rect(s, x, y, w, opt.hh || 0.42, color);
    txt(s, head, x, y, w, opt.hh || 0.42, { size: opt.hsize || 13, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  }
  function bubble(s, t, x, y, w, h, opt = {}) {
    s.addShape(pres.shapes.ROUNDED_RECTANGULAR_CALLOUT, { x, y, w, h, fill: { color: opt.fill || WHITE }, line: { color: opt.line || LINE, width: 1 }, objectName: nm("bubble") });
    txt(s, t, x + 0.05, y, w - 0.1, h * 0.85, { size: opt.size || 10.5, bold: true, color: opt.color || DARK, align: "center", valign: "middle", margin: 0 });
  }
  function dashBox(s, x, y, w, h, iconKey, label) {
    rect(s, x, y, w, h, "FAFAFA", { line: "9AA5B8", dash: "dash" });
    icon(s, iconKey, x + w / 2 - 0.35, y + h / 2 - 0.6, 0.7, MGRAY);
    txt(s, label, x + 0.1, y + h / 2 + 0.2, w - 0.2, 0.9, { size: 10.5, color: GRAY, align: "center", ls: 1.3 });
  }

  // =====================================================================
  // 表紙・はじめに
  // =====================================================================
  const S0 = "はじめに";
  pres.addSection({ title: S0 });
  {
    const s = pres.addSlide({ masterName: "BLANK", sectionTitle: S0 });
    s.addShape(pres.shapes.OVAL, { x: -2.5, y: 5.7, w: 16, h: 5, fill: { color: NAVY }, line: { type: "none" }, objectName: "cover_wave1" });
    s.addShape(pres.shapes.OVAL, { x: 1.5, y: 6.25, w: 14, h: 4, fill: { color: BLUE }, line: { type: "none" }, objectName: "cover_wave2" });
    s.addShape(pres.shapes.OVAL, { x: 6.2, y: -2.6, w: 7, h: 3.4, fill: { color: NAVY }, line: { type: "none" }, objectName: "cover_wave3" });
    s.addShape(pres.shapes.OVAL, { x: 7.4, y: -2.9, w: 6, h: 3.2, fill: { color: BLUE }, line: { type: "none" }, objectName: "cover_wave4" });
    s.addShape(pres.shapes.OVAL, { x: -0.5, y: -0.4, w: 1.6, h: 1.6, fill: { color: BLUE }, line: { type: "none" }, objectName: "cover_circle" });
    for (let i = 0; i < 12; i++) {
      const d = 0.13 - i * 0.008;
      s.addShape(pres.shapes.OVAL, { x: 0.75 + i * 0.2, y: 1.85 - d / 2, w: d, h: d, fill: { color: NAVY, transparency: i * 6 }, line: { type: "none" }, objectName: nm("dot") });
    }
    txt(s, "株式会社○○ 様", 0.7, 2.1, 8, 0.5, { size: 20, font: MINCHO });
    txt(s, "遺品整理事業\n参入のご提案", 0.7, 2.65, 8.6, 1.7, { size: 38, font: MINCHO, ls: 1.05 });
    txt(s, "株式会社 船井総合研究所\n専門工事支援部\n○○ ○○", 0.7, 4.6, 6, 0.95, { size: 12, ls: 1.2 });
    s.addNotes("【要カスタマイズ】会社名・担当者名を差し替える。\n事前経営相談：「遺品整理とは」〜「各業態におけるシナジー」まで。事後経営相談：「ご質問への回答」以降。");
  }
  {
    const s = content(S0, "会社紹介");
    txt(s, "株式会社船井総合研究所", 0.4, 1.0, 5.6, 0.5, { size: 18, bold: true, valign: "middle", margin: 0 });
    txt(s, "経営コンサルティング業界で初めて株式公開を果たした、日本最大級の経営コンサルティングファーム。現場の即時業績アップ支援に強みを持ち、社会的価値の高いグレートカンパニーの創造をミッションとしています。", 0.4, 1.5, 5.6, 1.1, { size: 11, ls: 1.3 });
    const ST = [["1970年", "創立"], ["約1,500名", "連結従業員数"], ["約1,000名", "コンサルタント数"], ["10,000社以上", "全国の支援先"]];
    ST.forEach((st, i) => {
      const x = 0.4 + (i % 2) * 2.85, y = 2.75 + Math.floor(i / 2) * 1.35;
      rect(s, x, y, 2.7, 1.2, LBLUE);
      txt(s, st[0], x, y + 0.12, 2.7, 0.6, { size: 24, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
      txt(s, st[1], x, y + 0.75, 2.7, 0.35, { size: 11, color: GRAY, align: "center", margin: 0 });
    });
    txt(s, "創設者　船井 幸雄　／　代表取締役社長　真貝 大介　／　資本金　3,000百万円", 0.4, 5.55, 5.6, 0.35, { size: 10, margin: 0 });
    photo(s, 6.3, 1.0, 3.3, 2.5, "大阪本社");
    txt(s, "大阪本社　イノゲート大阪 21階・22階\n大阪府大阪市北区梅田三丁目2番123号", 6.3, 3.55, 3.3, 0.55, { size: 9, color: GRAY, margin: 0 });
    photo(s, 6.3, 4.2, 3.3, 2.15, "東京本社");
    txt(s, "東京本社　八重洲セントラルタワー 35階\n東京都中央区八重洲二丁目2番1号", 6.3, 6.4, 3.3, 0.55, { size: 9, color: GRAY, margin: 0 });
  }
  {
    const s = content(S0, "会社紹介", "【要カスタマイズ】担当コンサルタントの写真・経歴・ミッションを差し替える。");
    s._pageTitle = "担当コンサルタント";
    msg(s, "遺品整理事業の専門チームが、貴社の立ち上げをご支援します");
    photo(s, 0.4, 1.7, 2.6, 2.6, "担当コンサルタント写真");
    txt(s, [B("専門工事支援部\n", { fontSize: 11, color: GRAY }), B("○○ ○○\n", { fontSize: 22, color: NAVY }), B("遺品整理士 など保有資格", { fontSize: 11, color: GRAY })], 3.2, 1.7, 3.5, 1.05, { ls: 1.2 });
    txt(s, "【経歴】\n新卒で船井総合研究所に入社。遺品整理・リフォーム業界を専門に、新規事業の立上げから既存事業の活性化まで、業績アップのご支援に従事。", 3.2, 2.8, 6.4, 0.7, { size: 11, ls: 1.3 });
    rect(s, 3.2, 3.55, 6.4, 0.75, LBLUE);
    txt(s, [B("ミッション　", { color: NAVY }), B("遺品整理を『片付け』から、安心できる住まいと暮らしの『再編集』へ", { color: DARK })], 3.35, 3.55, 6.2, 0.75, { size: 12, valign: "middle" });
    txt(s, "遺品整理事業の専門チーム（専門工事支援部）", 0.4, 4.55, 9, 0.4, { size: 14, bold: true, color: NAVY, margin: 0 });
    const M = [["阪口 和輝", "チームリーダー\n古民家鑑定士"], ["堂野 宗一郎", "コンサルタント\n遺品整理士"], ["廣瀬 聡太朗", "コンサルタント"], ["伊木 颯志朗", "コンサルタント"]];
    M.forEach((m, i) => {
      const x = 0.4 + i * 2.33;
      oval(s, x, 5.1, 0.95, 0.95, "E8EBF0", { line: "C2C8D2" });
      icon(s, "user", x + 0.25, 5.33, 0.45, MGRAY);
      txt(s, m[0], x + 1.05, 5.1, 1.25, 0.38, { size: 12, bold: true, valign: "middle", margin: 0 });
      txt(s, m[1], x + 1.05, 5.48, 1.25, 0.65, { size: 9, color: GRAY, margin: 0 });
    });
  }
  {
    const s = content(S0, "貴社のお話を伺います", "【使い方】事前経営相談ではこのページで伺い、手元でメモする。事後経営相談では伺った内容を書き込んだ版に差し替え、最初に確認する。他社名・外注先名は書かない。");
    msg(s, "まずは、貴社の状況とお考えをお聞かせください");
    const H = [["search", "セミナー・資料DLのきっかけ", "お申し込み・ダウンロードに至った理由"], ["building", "貴社の現状", "事業規模・売上構成、人員、許認可、車両・倉庫"], ["warn", "課題", "本業の課題、片付け・残置物の相談状況"], ["question", "知りたいこと・疑問点", "セミナーで知りたいこと／資料で気になった点"], ["target", "遺品整理事業の位置づけ", "会社の中でどう考えているか"]];
    H.forEach((h, i) => {
      const y = 1.7 + i * 1.05;
      icon(s, h[0], 0.45, y + 0.12, 0.62, WHITE, NAVY);
      txt(s, h[1], 1.25, y + 0.05, 2.7, 0.42, { size: 13, bold: true, color: NAVY, valign: "middle", margin: 0 });
      txt(s, h[2], 1.25, y + 0.47, 2.7, 0.5, { size: 9.5, color: GRAY, margin: 0 });
      for (let k = 0; k < 2; k++) line(s, 4.1, y + 0.45 + k * 0.45, 5.5, 0, "B8C0CC", 0.75, { dash: "dash" });
    });
  }

  // =====================================================================
  // 遺品整理とは
  // =====================================================================
  const S1 = "遺品整理とは";
  pres.addSection({ title: S1 });
  divider(S1, "遺品整理とは");
  {
    const s = content(S1, "遺品整理とは", "不用品回収との違い（想いの整理・貴重品の捜索・供養）を強調する。");
    msg(s, "遺品整理の流れ【3ステップ】");
    const C = [["box", "整理（仕分け）", "一つひとつ確認しながら「残す物」「貴重品」「処分する物」に仕分け", ORANGE], ["dolly", "搬出", "仕分けが終わった荷物を、建物を傷つけないよう丁寧に運び出し", GREEN], ["truck", "トラックへの積込", "安全に積み込み、許可業者による適正な処理・リサイクルへ", NAVY]];
    C.forEach((c, i) => {
      const x = 0.4 + i * 3.13;
      rect(s, x, 1.6, 2.85, 2.05, WHITE, { line: c[3], lw: 1.5, round: true, r: 0.06 });
      icon(s, c[0], x + 0.2, 1.78, 0.7, WHITE, c[3]);
      txt(s, `STEP ${i + 1}`, x + 1.05, 1.78, 1.7, 0.3, { size: 10, bold: true, color: c[3], margin: 0 });
      txt(s, c[1], x + 1.05, 2.06, 1.75, 0.42, { size: 15, bold: true, color: c[3], valign: "middle", margin: 0 });
      txt(s, c[2], x + 0.2, 2.65, 2.5, 0.95, { size: 10.5, ls: 1.25 });
      if (i < 2) arrowR(s, x + 2.9, 2.45, 0.2, 0.32, MGRAY);
    });
    txt(s, "不用品回収との違い", 0.4, 3.9, 9.2, 0.4, { size: 15, bold: true, color: NAVY, margin: 0 });
    rect(s, 0.4, 4.35, 4.45, 2.05, LGRAY);
    rect(s, 5.15, 4.35, 4.45, 2.05, LBLUE, { line: NAVY, lw: 1.5 });
    txt(s, "不用品回収", 0.4, 4.4, 4.45, 0.4, { size: 13, bold: true, color: MGRAY, align: "center", margin: 0 });
    txt(s, "遺品整理（専門店）", 5.15, 4.4, 4.45, 0.4, { size: 13, bold: true, color: NAVY, align: "center", margin: 0 });
    txt(s, "・不要な物を早く安く回収する\n・貴重品・思い出の品は対応しないことが多い\n・供養や心のケアは行わない\n・回収して終わり", 0.6, 4.85, 3.9, 1.5, { size: 11, color: GRAY, ls: 1.35 });
    txt(s, "・故人とご遺族の想いを整理する\n・1点ずつ手で確認し、貴重品を見つけてお渡し\n・供養・お焚き上げ、ご遺族への寄り添い\n・相続・売却・解体・リフォームまで相談できる", 5.5, 4.85, 4.0, 1.5, { size: 11, bold: true, color: NAVY, ls: 1.35 });
    s.addShape(pres.shapes.OVAL, { x: 4.6, y: 5.0, w: 0.8, h: 0.8, fill: { color: RED }, line: { color: WHITE, width: 2 }, objectName: "vs" });
    txt(s, "VS", 4.6, 5.0, 0.8, 0.8, { size: 15, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    key(s, "遺品整理は「片付け」ではなく、心に寄り添う専門サービスです。", { y: 6.5, size: 16 });
  }
  {
    const s = content(S1, "遺品整理ビジネスのミッション");
    rect(s, 0, 0.85, 10, 6.25, "F7F4EE");
    txt(s, [B("遺品整理", { color: DARK }), B("は、\n", { color: DARK, fontSize: 20 }), B("地域", { color: GREEN }), B("の", { color: DARK }), B("福祉", { color: GREEN }), B("と", { color: DARK }), B("空き家問題", { color: GREEN }), B("を\n", { color: DARK }), B("解決する事業です", { color: DARK })], 0.5, 1.05, 5.6, 1.8, { size: 26, font: MINCHO, ls: 1.15 });
    txt(s, "私たちは、故人とご遺族の想いに寄り添い、\n地域の未来を支える遺品整理を行っています。", 0.5, 2.9, 5.6, 0.7, { size: 12, ls: 1.3 });
    photo(s, 6.2, 1.05, 3.4, 2.45, "スタッフとご遺族の写真");
    line(s, 0.5, 3.95, 9.0, 0, "B9A98F", 1);
    rect(s, 2.2, 3.75, 5.6, 0.4, "F7F4EE");
    txt(s, "遺品整理をしたいお客様は、こんなお悩みを抱えています", 2.2, 3.75, 5.6, 0.4, { size: 13, bold: true, align: "center", valign: "middle", margin: 0 });
    const P = [["heart", "遺品を丁寧に扱ってくれる、\n安心できる業者を探している"], ["clock", "実家にわざわざ帰って\n片付ける時間がない"], ["search", "ネットで出てきた会社が\n安心できるか分からない"]];
    P.forEach((p, i) => {
      const x = 0.5 + i * 3.05;
      icon(s, p[0], x, 4.35, 0.75, GREEN, "E3EEDF");
      txt(s, p[1], x + 0.85, 4.33, 2.1, 0.8, { size: 11, bold: true, valign: "middle", ls: 1.2 });
    });
    oval(s, 0.5, 5.4, 1.25, 1.25, GREEN);
    txt(s, "私たちの\nミッション", 0.5, 5.4, 1.25, 1.25, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, [B("そんな業界の問題を解決するべく、\n", { color: DARK }), B("地域密着", { color: GREEN }), B("で", { color: DARK }), B("安心できる遺品整理業者", { color: GREEN }), B("を増やし、\n地域の福祉と未来を支えていきます。", { color: DARK })], 1.95, 5.38, 7.5, 1.3, { size: 15, font: MINCHO, ls: 1.25, valign: "middle" });
  }

  // =====================================================================
  // 市場性
  // =====================================================================
  const S2 = "市場性";
  pres.addSection({ title: S2 });
  divider(S2, "市場性");
  {
    const s = content(S2, "遺品整理市場の成長性");
    msg(s, "3つの追い風が、遺品整理市場を拡大しています");
    const W = [["users", "市場拡大の波", "多死社会の到来", "年間死亡者数は今後も増加し、遺品整理・実家整理の需要が拡大", ORANGE], ["home", "社会課題の波", "空き家 約900万戸", "放置空き家も増え、活用・売却・解体の需要が拡大", BLUE], ["law", "法改正の波", "相続登記の義務化", "相続後3年以内の登記が必要になり、放置が難しくなる", GREEN]];
    W.forEach((w, i) => {
      const x = 0.4 + i * 3.13;
      rect(s, x, 1.6, 2.94, 2.85, WHITE, { line: LINE, shadow: true });
      icon(s, w[0], x + 1.07, 1.75, 0.8, WHITE, w[4]);
      txt(s, w[1], x, 2.62, 2.94, 0.32, { size: 11, bold: true, color: w[4], align: "center", margin: 0 });
      txt(s, w[2], x, 2.95, 2.94, 0.45, { size: 16, bold: true, align: "center", valign: "middle", margin: 0 });
      txt(s, w[3], x + 0.2, 3.45, 2.54, 0.9, { size: 10.5, color: GRAY, align: "center", ls: 1.25 });
    });
    s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: 4.55, y: 4.58, w: 0.9, h: 0.32, rotate: 180, fill: { color: NAVY }, line: { type: "none" }, objectName: "tri" });
    rect(s, 0.4, 5.0, 9.2, 0.55, NAVY);
    txt(s, [B("3つの波が重なり、", { color: WHITE }), B("「放置」", { color: YEL }), B("から", { color: WHITE }), B("「活用・処分」", { color: YEL }), B("へ動く人が急増", { color: WHITE })], 0.4, 5.0, 9.2, 0.55, { size: 16, align: "center", valign: "middle", margin: 0 });
    const O = [["box", "遺品整理"], ["home", "実家整理"], ["wrench", "リフォーム"], ["hammer", "解体"], ["yen", "売却"], ["key", "賃貸活用"]];
    O.forEach((o, i) => {
      const x = 0.55 + i * 1.53;
      icon(s, o[0], x + 0.35, 5.7, 0.5, NAVY);
      txt(s, o[1], x, 6.25, 1.2, 0.35, { size: 11, bold: true, color: NAVY, align: "center", margin: 0 });
    });
    foot(s, "※空き家数：総務省「住宅・土地統計調査」。最新の公表値に合わせて更新してください。");
  }
  {
    const s = content(S2, "遺品整理業界の現状");
    msg(s, "遺品整理業者が抱える課題と、見えてきた市場のチャンス");
    const I = [["truck", "不用品回収を中心とした事業者が多い", "遺品整理という名前でも、実際は不用品回収の片付けが中心。「想いの整理」まで対応できる事業者は限られています。"], ["pray", "グリーフケアや供養への対応に差がある", "ご遺族への寄り添い方や供養の提案など、心のケアまで提供できる事業者はまだ多くありません。"], ["warn", "サービス品質にばらつきがある", "対応・料金・作業品質が業者によって大きく違い、お客様の満足度に影響しています。"]];
    I.forEach((it, i) => {
      const y = 1.65 + i * 1.4;
      rect(s, 0.4, y, 9.2, 1.25, i % 2 ? WHITE : LBROWN, { line: "D9C7B2" });
      icon(s, it[0], 0.6, y + 0.24, 0.78, WHITE, BROWN);
      txt(s, it[1], 1.6, y + 0.1, 7.8, 0.45, { size: 15, bold: true, color: BROWN, valign: "middle", margin: 0 });
      txt(s, it[2], 1.6, y + 0.55, 7.8, 0.65, { size: 11, ls: 1.25, margin: 0 });
    });
    key(s, "「片付け・回収」ではなく、心に寄り添う専門店が求められています。", { y: 6.0, size: 17 });
  }
  {
    const s = content(S2, "今が参入のタイミング", "ライフサイクルの考え方：成長期は一番店が決まっていない。安定期に入ると価格競争になり、一番店化は難しい。");
    msg(s, "商品のライフサイクルで見ると、遺品整理は「成長期」にあります");
    const P = [["導入期", 0.75, 1.3], ["成長期", 2.05, 1.6], ["成熟期", 3.65, 1.55], ["安定期", 5.2, 1.55], ["衰退期", 6.75, 1.45]];
    P.forEach((p, i) => {
      rect(s, p[1], 1.65, p[2], 4.3, i === 1 ? "FCE8E6" : (i % 2 ? LGRAY : WHITE));
      txt(s, p[0], p[1], 5.62, p[2], 0.32, { size: 11, bold: true, color: i === 1 ? RED : GRAY, align: "center", margin: 0 });
    });
    s.addChart(pres.charts.LINE, [{ name: "市場規模", labels: ["", "", "", "", "", "", "", "", "", "", "", "", "", ""], values: [2, 3, 5, 9, 15, 23, 30, 34, 36, 36, 35, 33, 30, 27] }], {
      x: 0.6, y: 1.65, w: 7.75, h: 3.95, lineSmooth: true, lineSize: 4, chartColors: [NAVY], lineDataSymbol: "none",
      catAxisHidden: true, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 40, objectName: "lifecycle",
    });
    rect(s, 2.3, 2.35, 1.35, 0.55, RED, { round: true });
    txt(s, "遺品整理\n（いまここ）", 2.3, 2.35, 1.35, 0.55, { size: 10, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, "リフォーム・新築など", 4.0, 1.7, 2.4, 0.3, { size: 10, color: GRAY, align: "center", margin: 0 });
    rect(s, 8.35, 1.65, 1.25, 2.05, RED);
    txt(s, [B("成長期\n", { fontSize: 13 }), B("地域一番店が不在。今なら一番店をとれる", { fontSize: 10 })], 8.38, 1.7, 1.19, 1.95, { color: WHITE, align: "center", valign: "middle" });
    rect(s, 8.35, 3.9, 1.25, 2.05, "8FA3C4");
    txt(s, [B("安定期\n", { fontSize: 13 }), B("大手参入で価格競争。一番店化は困難", { fontSize: 10 })], 8.38, 3.95, 1.19, 1.95, { color: WHITE, align: "center", valign: "middle" });
    key(s, "地域一番店をとれるのは、専門店がまだ少ない「今」です。", { y: 6.35, size: 17 });
    foot(s, "※商品のライフサイクル理論（船井流）に基づく考え方です。");
  }

  // =====================================================================
  // ビジネスモデル概要
  // =====================================================================
  const S3 = "ビジネスモデル概要";
  pres.addSection({ title: S3 });
  divider(S3, "ビジネスモデル概要");
  {
    const s = content(S3, "事業モデルの全体像");
    const T = [["map", "商圏人口", "20万人"], ["yen", "平均客単価", "30万円"], ["chart", "目指す売上", "1億円"]];
    T.forEach((t, i) => {
      const x = 0.55 + i * 3.15;
      rect(s, x, 1.05, 2.6, 1.45, i === 2 ? RED : LBLUE, { round: true });
      icon(s, t[0], x + 0.15, 1.2, 0.45, i === 2 ? WHITE : NAVY);
      txt(s, t[1], x + 0.7, 1.15, 1.8, 0.45, { size: 12, bold: true, color: i === 2 ? WHITE : NAVY, valign: "middle", margin: 0 });
      txt(s, t[2], x, 1.65, 2.6, 0.75, { size: 32, bold: true, color: i === 2 ? WHITE : RED, align: "center", valign: "middle", margin: 0 });
      if (i < 2) txt(s, i === 0 ? "×" : "＝", x + 2.6, 1.3, 0.55, 0.9, { size: 30, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    });
    txt(s, "※商圏人口や会社の認知度・実績により年商は変動します。数字はモデル数値であり、確約するものではありません。", 0.4, 2.6, 9.2, 0.3, { size: 9, color: GRAY, align: "center", margin: 0 });
    const P = [["users", "ターゲット", "親の家を相続する40〜70代。\n遠方に住むご家族も多い"], ["home", "主力商品", "単価の大きい\n「戸建の遺品整理」を中心に集客"], ["link", "バックエンド", "解体・リフォーム・不動産・供養へ\nつなげて1顧客あたりの利益を高める"]];
    P.forEach((p, i) => {
      const x = 0.4 + i * 3.13;
      rect(s, x, 3.1, 2.94, 2.75, WHITE, { line: LINE, shadow: true });
      icon(s, p[0], x + 1.1, 3.25, 0.75, WHITE, NAVY);
      txt(s, p[1], x, 4.08, 2.94, 0.4, { size: 15, bold: true, color: NAVY, align: "center", margin: 0 });
      txt(s, p[2], x + 0.15, 4.55, 2.64, 1.1, { size: 11, align: "center", ls: 1.3 });
    });
    txt(s, "不用品回収・出張買取にも対応しつつ、戸建の遺品整理を軸にします", 0.4, 5.95, 9.2, 0.35, { size: 11, color: GRAY, align: "center", margin: 0 });
    key(s, "初期投資が低く、社員1名からスタートできるビジネスモデルです。", { y: 6.4, size: 17 });
  }
  {
    const s = content(S3, "3年間の数値計画", "経費の内訳は「収支計画（3年間）」のページで説明する。");
    msg(s, "3年目に売上1億円・営業利益率20%超を目指します");
    s.addChart(pres.charts.BAR, [{ name: "売上（万円）", labels: ["1年目（立上げ期）", "2年目（拡大期）", "3年目（組織化期）"], values: [3600, 5400, 10000] }], {
      x: 0.4, y: 1.55, w: 5.6, h: 4.35, barDir: "col", barGapWidthPct: 60, chartColors: [NAVY], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 13, dataLabelFontBold: true, dataLabelColor: NAVY, dataLabelFormatCode: "#,##0\"万円\"",
      showTitle: true, title: "売上の推移", titleFontSize: 13, titleColor: NAVY, titleFontFace: "+mn-lt", catAxisLabelFontSize: 10, catAxisLabelColor: DARK, catAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, valAxisMaxVal: 12000, valAxisMinVal: 0, objectName: "sales_chart",
    });
    txt(s, "営業利益率", 0.4, 6.0, 1.0, 0.38, { size: 10, color: GRAY, valign: "middle", margin: 0 });
    [["2.9%", 1.35], ["16.7%", 3.0], ["23.1%", 4.65]].forEach((r, i) => {
      rect(s, r[1], 6.0, 1.05, 0.38, i === 2 ? RED : LBLUE, { round: true });
      txt(s, r[0], r[1], 6.0, 1.05, 0.38, { size: 13, bold: true, color: i === 2 ? WHITE : NAVY, align: "center", valign: "middle", margin: 0 });
    });
    frame(s, 6.3, 1.55, 3.3, 2.15, "初期投資", NAVY);
    txt(s, [B("HP制作費　", { fontSize: 11 }), B("約150万円\n", { fontSize: 18, color: RED }), B("古物商許可　", { fontSize: 11 }), B("5〜8万円", { fontSize: 18, color: RED })], 6.4, 2.05, 3.1, 1.55, { valign: "middle", ls: 1.4 });
    arrowD(s, 7.7, 3.78, 0.5, 0.35, NAVY);
    frame(s, 6.3, 4.2, 3.3, 1.75, "投資回収", RED);
    txt(s, [B("2年目", { fontSize: 26, color: RED }), B("で回収\n", { fontSize: 14, color: RED }), B("2年目の営業利益 904万円で回収見込み", { fontSize: 10 })], 6.4, 4.65, 3.1, 1.25, { align: "center", valign: "middle" });
    txt(s, "前提：商圏人口20万人／平均客単価30万円／契約率80%", 6.3, 6.0, 3.3, 0.38, { size: 9, color: GRAY, valign: "middle", margin: 0 });
    foot(s, MODEL_NOTE);
  }
  {
    const s = content(S3, "仕事の入口と出口（商流）");
    msg(s, "さまざまな入口から遺品整理につながり、整理後の出口サービスまでワンストップで提供", { size: 13 });
    txt(s, "入口（案件の発生源）", 0.4, 1.55, 2.7, 0.35, { size: 12, bold: true, color: NAVY, margin: 0 });
    const IN = [["pray", "葬儀会社", "葬儀後に片付けの相談"], ["hospital", "介護・老人ホーム", "入居前の生前整理"], ["building", "不動産会社", "売却前の家財撤去"]];
    IN.forEach((it, i) => {
      const y = 1.95 + i * 0.85;
      rect(s, 0.4, y, 2.7, 0.72, LBLUE, { round: true });
      icon(s, it[0], 0.5, y + 0.13, 0.46, NAVY);
      txt(s, it[1], 1.05, y + 0.04, 2.0, 0.35, { size: 11.5, bold: true, color: NAVY, margin: 0 });
      txt(s, it[2], 1.05, y + 0.37, 2.0, 0.3, { size: 9.5, color: GRAY, margin: 0 });
    });
    arrowR(s, 3.2, 2.85, 0.45, 0.4, NAVY);
    oval(s, 3.75, 1.85, 2.5, 2.5, NAVY);
    txt(s, [B("遺品整理\n", { fontSize: 18 }), B("（元請けで受注）\n", { fontSize: 11 }), B("仕分け・貴重品捜索\n買取・搬出", { fontSize: 9.5, color: SKY })], 3.75, 1.85, 2.5, 2.5, { color: WHITE, align: "center", valign: "middle" });
    arrowR(s, 6.35, 2.85, 0.45, 0.4, GREEN);
    txt(s, "出口（整理後のサービス）", 6.9, 1.55, 2.7, 0.35, { size: 12, bold: true, color: GREEN, margin: 0 });
    const OUT_ = [["gem", "買取・リユース"], ["broom", "清掃・特殊清掃"], ["hammer", "解体工事"], ["wrench", "リフォーム"], ["yen", "不動産売却"], ["recycle", "許可業者（処分）"]];
    OUT_.forEach((o, i) => {
      const x = 6.9 + (i % 2) * 1.37, y = 1.95 + Math.floor(i / 2) * 0.85;
      rect(s, x, y, 1.3, 0.72, LGREEN, { round: true });
      icon(s, o[0], x + 0.08, y + 0.2, 0.32, GREEN);
      txt(s, o[1], x + 0.42, y, 0.86, 0.72, { size: 9.5, bold: true, color: GREEN, valign: "middle", margin: 0 });
    });
    rect(s, 0.4, 4.65, 4.5, 1.7, LGRAY, { round: true });
    txt(s, "下請けの流れ（よくある形）", 0.55, 4.72, 4.2, 0.35, { size: 12, bold: true, color: MGRAY, margin: 0 });
    chevron(s, 0.55, 5.15, 1.45, 0.55, MGRAY, "解体業者", { first: true, size: 10 });
    chevron(s, 1.9, 5.15, 1.55, 0.55, MGRAY, "残置物撤去", { size: 10 });
    chevron(s, 3.35, 5.15, 1.45, 0.55, MGRAY, "解体工事", { size: 10 });
    txt(s, "単価が低く、後工程の利益は他社へ", 0.55, 5.8, 4.2, 0.4, { size: 10.5, color: GRAY, margin: 0 });
    rect(s, 5.1, 4.65, 4.5, 1.7, "FCE8E6", { round: true, line: RED });
    txt(s, "元請けの流れ（目指す形）", 5.25, 4.72, 4.2, 0.35, { size: 12, bold: true, color: RED, margin: 0 });
    chevron(s, 5.25, 5.15, 1.45, 0.55, RED, "お客様", { first: true, size: 10 });
    chevron(s, 6.6, 5.15, 1.55, 0.55, RED, "遺品整理", { size: 10 });
    chevron(s, 8.05, 5.15, 1.45, 0.55, RED, "解体・売却", { size: 10 });
    txt(s, "遺品整理を起点に、解体・売却・活用までご提案", 5.25, 5.8, 4.2, 0.4, { size: 10.5, color: RED, bold: true, margin: 0 });
    key(s, "元請けで受けると、後工程の大きな利益が自社に残ります。", { y: 6.45, size: 16 });
  }
  {
    const s = content(S3, "遺品整理から広がる7つの収益機会");
    msg(s, "関連サービスを段階的にご提案し、お客様の課題をワンストップで解決します", { size: 14 });
    const st = [["box", "遺品整理（基本サービス）", ORANGE], ["gem", "買取・リユース", GOLD], ["pray", "供養・お焚き上げ", "7FA650"], ["broom", "特殊清掃・消臭・除菌", GREEN], ["hammer", "解体工事", BLUE], ["building", "不動産売却・買取サポート", NAVY], ["wrench", "リフォーム・リノベーション", PURPLE]];
    st.forEach((t, i) => {
      const x = 0.4 + i * 0.33, y = 5.9 - i * 0.64, w = 3.7;
      rect(s, x, y, w, 0.56, t[2]);
      txt(s, String(i + 1), x, y, 0.42, 0.56, { size: 16, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      icon(s, t[0], x + 0.45, y + 0.12, 0.32, WHITE);
      txt(s, t[1], x + 0.85, y, w - 0.9, 0.56, { size: 11.5, bold: true, color: WHITE, valign: "middle", margin: 0 });
    });
    oval(s, 0.55, 1.75, 1.5, 1.5, NAVY);
    txt(s, "遺品整理から\n広がる\n7つの収益機会", 0.55, 1.75, 1.5, 1.5, { size: 10.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    rect(s, 6.2, 1.65, 3.4, 1.3, LGRAY, { round: true });
    txt(s, "遺品整理のみ", 6.2, 1.7, 3.4, 0.35, { size: 11, bold: true, color: MGRAY, align: "center", margin: 0 });
    txt(s, [B("売上 ", { fontSize: 11 }), B("30万円", { fontSize: 22 }), B("　粗利 10万円", { fontSize: 11 })], 6.2, 2.05, 3.4, 0.8, { color: MGRAY, align: "center", valign: "middle" });
    arrowD(s, 7.65, 3.0, 0.5, 0.4, RED);
    rect(s, 6.2, 3.45, 3.4, 2.75, "FCE8E6", { round: true, line: RED });
    txt(s, "関連サービスを提案した場合", 6.2, 3.5, 3.4, 0.35, { size: 11, bold: true, color: RED, align: "center", margin: 0 });
    txt(s, [B("売上 ", { fontSize: 11 }), B("490万円\n", { fontSize: 28, color: RED }), B("粗利 ", { fontSize: 11 }), B("180万円", { fontSize: 20, color: RED })], 6.2, 3.85, 3.4, 1.3, { align: "center", valign: "middle" });
    txt(s, "買取▲5万／特殊清掃20万／供養5万／解体150万／不動産売却サポート90万／リフォーム200万", 6.35, 5.2, 3.1, 0.9, { size: 9, color: GRAY, align: "center" });
    key(s, "遺品整理は、関連サービスにつながる「入口商品」です。", { y: 6.45, size: 16 });
    foot(s, MODEL_NOTE);
  }
  {
    const s = content(S3, "事業を動かす4つの柱");
    msg(s, "集客・営業・現場・提携の4つを仕組みにすることで、未経験からでも立ち上げられます", { size: 13.5 });
    const P = [["horn", "集客", "専門店の屋号とHPをつくり、WEB・チラシで反響をとる", "専門HP（10ページ構成）／SEO・MEO・Web広告／戸建向けチラシ", BLUE], ["phone", "営業", "最初の電話対応と、現地での即決で契約率80%を目指す", "5分以内の電話／30万円未満はオンライン完結／30万円以上は現地で即決", ORANGE], ["check", "現場", "誰がやっても同じ品質になる、作業のルールをつくる", "捨てる前に必ず確認／貴重品の徹底捜索／作業完了報告", GREEN], ["shake", "提携", "不動産・士業・葬儀社などから紹介が集まる仕組みをつくる", "地域の案件ハブになる／提携先の信用を守る／紹介手数料の設計", PURPLE]];
    P.forEach((p, i) => {
      const x = i % 2 ? 6.15 : 0.4, y = i < 2 ? 1.6 : 4.1, w = 3.45;
      rect(s, x, y, w, 2.2, WHITE, { line: p[4], lw: 1.5, round: true, r: 0.05 });
      icon(s, p[0], x + 0.15, y + 0.15, 0.6, WHITE, p[4]);
      txt(s, p[1], x + 0.85, y + 0.15, 2.5, 0.6, { size: 18, bold: true, color: p[4], valign: "middle", margin: 0 });
      txt(s, p[2], x + 0.15, y + 0.85, w - 0.3, 0.6, { size: 11, bold: true, ls: 1.2 });
      txt(s, p[3], x + 0.15, y + 1.45, w - 0.3, 0.7, { size: 9.5, color: GRAY, ls: 1.2 });
    });
    line(s, 3.85, 2.7, 0.45, 0.55, LINE, 2); line(s, 5.7, 2.7, 0.45, 0.55, LINE, 2, { flipV: true });
    line(s, 3.85, 4.75, 0.45, 0.45, LINE, 2, { flipV: true }); line(s, 5.7, 4.75, 0.45, 0.45, LINE, 2);
    oval(s, 4.1, 3.05, 1.8, 1.8, NAVY);
    txt(s, "遺品整理\n事業", 4.1, 3.05, 1.8, 1.8, { size: 15, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    key(s, "4つの柱を仕組みにすれば、未経験でも立ち上げられます。", { y: 6.45, size: 16 });
  }

  // =====================================================================
  // 許可・資格について
  // =====================================================================
  const S4 = "許可・資格について";
  pres.addSection({ title: S4 });
  divider(S4, "許可・資格について");
  {
    const s = content(S4, "資格・許可について", "【話し方】①左の図：契約は2本。お客様と当社が「サービス契約」、お客様と許可業者が「直接契約」→②右の流れ：1〜4は当社、5の回収・処分だけ許可業者（現場で引き渡す）→③下：従来モデルとの違いとメリット。前提条件（当社は廃棄物を運搬しない／現場で許可業者へ引き渡す）と自治体ごとのルールは事前に確認する。");
    const GOLDT = "B8892B";
    txt(s, [B("三社間契約", { color: NAVY }), B("なら、", { color: DARK, fontSize: 16 }), B("一廃許可がなくても", { color: GOLDT }), B("始められます", { color: DARK, fontSize: 16 })], 0.4, 0.98, 9.2, 0.55, { size: 21, align: "center", valign: "middle", margin: 0 });
    // ---- 左：三者の関係（誰と契約するか） ----
    txt(s, "誰と契約するか", 0.4, 1.58, 5.0, 0.3, { size: 11, bold: true, color: GRAY, valign: "middle", margin: 0 });
    line(s, 1.45, 2.67, 0.85, 1.13, NAVY, 2.75, { flipV: true });
    line(s, 3.5, 2.67, 0.9, 1.13, GREEN, 2.75);
    rect(s, 1.85, 1.92, 2.1, 0.75, LBLUE, { line: NAVY, lw: 1.25, round: true, r: 0.08 });
    icon(s, "users", 2.0, 2.06, 0.48, NAVY);
    txt(s, "お客様", 2.55, 1.92, 1.35, 0.75, { size: 15, bold: true, color: NAVY, valign: "middle", margin: 0 });
    txt(s, [B("① サービス契約\n", { fontSize: 11, color: NAVY }), B("仕分け・搬出・買取", { fontSize: 9, color: NAVY })], 0.2, 2.85, 1.45, 0.55, { align: "right", valign: "middle", margin: 0, ls: 1.05 });
    txt(s, [B("② 直接契約\n", { fontSize: 11, color: GREEN }), B("収集運搬・処分", { fontSize: 9, color: GREEN })], 4.1, 2.85, 1.45, 0.55, { valign: "middle", margin: 0, ls: 1.05 });
    rect(s, 0.4, 3.8, 2.1, 0.85, NAVY, { round: true, r: 0.08 });
    icon(s, "user", 0.52, 3.98, 0.5, WHITE);
    txt(s, "当社\n（遺品整理）", 1.08, 3.8, 1.4, 0.85, { size: 12, bold: true, color: WHITE, valign: "middle", margin: 0, ls: 1.05 });
    rect(s, 3.35, 3.8, 2.1, 0.85, GREEN, { round: true, r: 0.08 });
    icon(s, "recycle", 3.47, 3.98, 0.5, WHITE);
    txt(s, "一般廃棄物\n許可業者", 4.03, 3.8, 1.4, 0.85, { size: 12, bold: true, color: WHITE, valign: "middle", margin: 0, ls: 1.05 });
    line(s, 2.55, 4.22, 0.75, 0, MGRAY, 1.5, { dash: "dash", arrow: true });
    txt(s, "現場で\n引き渡し", 2.5, 4.27, 0.85, 0.38, { size: 8, color: GRAY, align: "center", valign: "middle", margin: 0, ls: 1.0 });
    // ---- 右：作業の流れと担当 ----
    txt(s, "誰が何をするか", 5.85, 1.58, 3.75, 0.3, { size: 11, bold: true, color: GRAY, valign: "middle", margin: 0 });
    const ST = [["list", "ご相談・お見積り", "当社"], ["box", "仕分け・整理", "当社"], ["dolly", "搬出・運搬", "当社"], ["shake", "許可業者へ引渡し", "当社"], ["recycle", "回収・処分", "許可業者"]];
    line(s, 6.13, 2.1, 0, 2.3, LINE, 2);
    ST.forEach((st, i) => {
      const y = 1.92 + i * 0.55, col = i === 4 ? GREEN : NAVY;
      rect(s, 5.85, y, 3.75, 0.47, i === 4 ? LGREEN : LBLUE, { round: true, r: 0.08 });
      num(s, i + 1, 5.95, y + 0.06, 0.35, col);
      icon(s, st[0], 6.42, y + 0.08, 0.3, col);
      txt(s, st[1], 6.82, y, 1.9, 0.47, { size: 11, bold: true, color: col, valign: "middle", margin: 0 });
      rect(s, 8.7, y + 0.1, 0.8, 0.27, col, { round: true, r: 0.12 });
      txt(s, st[2], 8.7, y + 0.1, 0.8, 0.27, { size: 8.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    });
    // ---- 下：従来モデルとの違い・メリット ----
    line(s, 0.4, 4.85, 9.2, 0, LINE, 0.75);
    rect(s, 0.4, 5.0, 2.75, 1.3, LGRAY, { round: true, r: 0.06 });
    txt(s, "✕ 従来モデル", 0.55, 5.06, 2.5, 0.34, { size: 12.5, bold: true, color: MGRAY, valign: "middle", margin: 0 });
    txt(s, "当社が回収・処分まで行う\n→ 許可がないと違法リスク", 0.55, 5.42, 2.5, 0.8, { size: 10, color: GRAY, valign: "middle", margin: 0, ls: 1.25 });
    arrowR(s, 3.22, 5.47, 0.3, 0.36, RED);
    rect(s, 3.6, 5.0, 2.75, 1.3, "FCE8E6", { line: RED, lw: 1.5, round: true, r: 0.06 });
    txt(s, "○ 三社間契約", 3.75, 5.06, 2.5, 0.34, { size: 12.5, bold: true, color: RED, valign: "middle", margin: 0 });
    txt(s, [B("処分は許可業者が直接契約\n", { color: DARK }), B("→ 法令に沿って運営できる", { color: RED })], 3.75, 5.42, 2.5, 0.8, { size: 10, valign: "middle", margin: 0, ls: 1.25 });
    txt(s, "導入のメリット", 6.6, 4.96, 3.0, 0.3, { size: 11, bold: true, color: NAVY, valign: "middle", margin: 0 });
    [["shield", "法令遵守で安心"], ["rocket", "早期に開始"], ["award", "信頼性UP"], ["chart", "営業に集中"]].forEach((m, i) => {
      const x = 6.6 + (i % 2) * 1.52, y = 5.3 + Math.floor(i / 2) * 0.52;
      rect(s, x, y, 1.45, 0.45, LBLUE, { round: true, r: 0.1 });
      icon(s, m[0], x + 0.08, y + 0.09, 0.27, NAVY);
      txt(s, m[1], x + 0.4, y, 1.03, 0.45, { size: 9, bold: true, color: NAVY, valign: "middle", margin: 0 });
    });
    key(s, "許可がなくても、法令に沿ってスモールスタートできます。", { y: 6.45, size: 17 });
  }
  {
    const s = content(S4, "資格・許可について", "【話し方】①遺品整理に伴う回収・運搬・処分には原則として一般廃棄物収集運搬許可が必要。ただし三社間契約なら許可がなくても始められる（左）→②許可を取ると回収・運搬まで自社で完結でき、コスト・日程・対応力が変わる（右）→③メリット3つ→④まずは許可なしで始め、拡大に合わせて検討。許可の要件・新規受付の有無は自治体ごとに異なるため事前に確認する。");
    txt(s, [B("一般廃棄物収集運搬許可", { color: NAVY }), B("を持つと、収益性が", { color: NAVY }), B("さらに向上", { color: ORANGE }), B("します", { color: NAVY })], 0.3, 0.9, 9.4, 0.38, { size: 18, align: "center", valign: "middle", margin: 0 });
    txt(s, "許可がなくても三社間契約で事業はできます。許可は「利益を伸ばすための次の一手」です。", 0.3, 1.27, 9.4, 0.28, { size: 11, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    // 左右の比較
    const M = [
      [0.4, "許可なし（提携モデル）", MGRAY, LGRAY, GREEN, "当社　→　許可業者　→　許可業者", [["契約", "お客様と直接契約＋許可業者と提携"], ["回収・運搬", "許可業者に外注"], ["コスト", "処分費は外注（一般相場）"], ["日程", "外注先の都合に合わせる"], ["利益率", "標準的な水準"]]],
      [5.15, "許可あり（自社完結モデル）", GREEN, LGREEN, NAVY, "当社　→　当社　→　当社", [["契約", "回収・運搬まで自社で対応"], ["回収・運搬", "自社で回収・運搬"], ["コスト", "自社運営で処分費を抑えやすい"], ["日程", "自社で柔軟に調整できる"], ["利益率", "外注コスト削減で向上"]]],
    ];
    M.forEach(([x, head, col, bg, fcol, who, rows], k) => {
      rect(s, x, 1.68, 4.45, 3.3, bg, { line: k ? GREEN : null, lw: 1.5, round: true, r: 0.05 });
      txt(s, head, x + 0.15, 1.75, 4.15, 0.36, { size: 14, bold: true, color: col, valign: "middle", margin: 0 });
      chevron(s, x + 0.15, 2.15, 1.45, 0.45, NAVY, "仕分け・搬出", { first: true, size: 9.5 });
      chevron(s, x + 1.5, 2.15, 1.45, 0.45, fcol, "回収・運搬", { size: 9.5 });
      chevron(s, x + 2.85, 2.15, 1.45, 0.45, fcol, "処分", { size: 9.5 });
      txt(s, who, x + 0.15, 2.62, 4.15, 0.26, { size: 9, bold: true, color: k ? NAVY : GREEN, align: "center", valign: "middle", margin: 0 });
      rows.forEach((r, i) => {
        const y = 2.98 + i * 0.39;
        rect(s, x + 0.15, y, 1.05, 0.33, WHITE, { round: true, r: 0.1 });
        txt(s, r[0], x + 0.15, y, 1.05, 0.33, { size: 9.5, bold: true, color: col, align: "center", valign: "middle", margin: 0 });
        txt(s, r[1], x + 1.3, y, 3.05, 0.33, { size: 10.5, bold: k === 1, color: k ? DARK : GRAY, valign: "middle", margin: 0 });
      });
    });
    s.addShape(pres.shapes.RIGHT_ARROW, { x: 4.88, y: 3.1, w: 0.25, h: 0.5, fill: { color: RED }, line: { type: "none" }, objectName: "arr_k" });
    // メリット3つ
    txt(s, "許可を取得するメリット", 0.4, 5.08, 9.2, 0.3, { size: 12, bold: true, color: GREEN, valign: "middle", margin: 0 });
    [["coins", "利益率の向上", "回収・運搬が自社で完結し、外注コストを抑えられる"], ["cal", "スケジュールの柔軟性", "外注先との調整が不要になり、日程を自社で決められる"], ["rocket", "サービスの幅が広がる", "スポット対応・即日対応など、多様なニーズに応えられる"]].forEach((m, i) => {
      const x = 0.4 + i * 3.1;
      icon(s, m[0], x, 5.45, 0.56, WHITE, GREEN);
      txt(s, m[1], x + 0.66, 5.4, 2.35, 0.3, { size: 11.5, bold: true, color: GREEN, valign: "middle", margin: 0 });
      txt(s, m[2], x + 0.66, 5.7, 2.35, 0.45, { size: 9, color: GRAY, valign: "top", margin: 0, ls: 1.15 });
    });
    key(s, "まずは三社間契約で始め、事業の拡大に合わせて許可取得を検討します。", { y: 6.42, size: 15 });
    foot(s, "※許可の要件・新規受付の有無は自治体により異なります。");
  }
  {
    const s = content(S4, "資格・許可について", "【話し方】①古物商許可とは、中古品を買い取り・販売するための公安委員会の許可→②片付け・搬出だけなら不要。「買取」をするなら必要（上段）→③対象はブランド品・貴金属・家具家電・骨董品など→④取ると買取利益が出て、お客様の費用も下がる→⑤取得は約40日・費用は2万円前後。取得後は本人確認・古物台帳・プレート掲示を必ず守る。");
    const GOLDT = "B8892B";
    txt(s, [B("「買取」", { color: GOLDT }), B("をするなら、", { color: NAVY }), B("古物商許可", { color: NAVY }), B("を取得しましょう", { color: NAVY })], 0.3, 0.9, 9.4, 0.38, { size: 18, align: "center", valign: "middle", margin: 0 });
    txt(s, "古物商許可とは：中古品（古物）を買い取り・販売するために必要な、公安委員会の許可です", 0.3, 1.27, 9.4, 0.28, { size: 10.5, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    // 上段：不要／必要
    rect(s, 0.3, 1.62, 4.4, 1.38, LGRAY, { round: true, r: 0.05 });
    txt(s, "許可は不要（遺品整理の作業だけなら）", 0.42, 1.66, 4.2, 0.3, { size: 10.5, bold: true, color: MGRAY, valign: "middle", margin: 0 });
    [["box", "遺品の\n仕分け"], ["dolly", "搬出・\n片付け"], ["broom", "清掃"], ["recycle", "廃棄物の\n分別"], ["truck", "許可業者へ\n引渡し"]].forEach((t, i) => {
      const x = 0.42 + i * 0.84;
      icon(s, t[0], x + 0.22, 2.0, 0.36, MGRAY);
      txt(s, t[1], x, 2.4, 0.8, 0.5, { size: 8, bold: true, color: GRAY, align: "center", valign: "top", margin: 0, ls: 1.05 });
    });
    arrowR(s, 4.76, 2.08, 0.26, 0.42, RED);
    rect(s, 5.08, 1.62, 4.62, 1.38, LBLUE, { line: NAVY, lw: 1.5, round: true, r: 0.05 });
    txt(s, "古物商許可が必要（「買取」を行う場合）", 5.2, 1.66, 4.4, 0.3, { size: 10.5, bold: true, color: NAVY, valign: "middle", margin: 0 });
    ["お客様から遺品を買い取る", "買い取った品をリユース販売する", "古物市場へ出品する", "リサイクルショップで販売する"].forEach((t, i) => {
      const x = 5.2 + (i % 2) * 2.25, y = 2.05 + Math.floor(i / 2) * 0.45;
      icon(s, "ok", x, y + 0.08, 0.24, NAVY);
      txt(s, t, x + 0.3, y, 1.95, 0.4, { size: 9, bold: true, valign: "middle", margin: 0 });
    });
    // 対象になるもの
    txt(s, "遺品整理で対象になるもの", 0.3, 3.08, 2.4, 0.3, { size: 10.5, bold: true, color: NAVY, valign: "middle", margin: 0 });
    [["bag", "ブランド品"], ["clock", "腕時計"], ["ring", "貴金属"], ["couch", "家具"], ["tv", "家電"], ["palette", "骨董品"], ["guitar", "楽器"], ["camera", "カメラ"], ["tools", "工具"], ["robot", "おもちゃ"]].forEach((t, i) => {
      const x = 0.3 + i * 0.945;
      rect(s, x, 3.4, 0.88, 0.62, WHITE, { line: LINE, round: true, r: 0.08 });
      icon(s, t[0], x + 0.3, 3.45, 0.28, NAVY);
      txt(s, t[1], x, 3.75, 0.88, 0.24, { size: 8, bold: true, align: "center", valign: "middle", margin: 0 });
    });
    // メリット
    txt(s, "取得するメリット（売上・利益アップにつながる理由）", 0.3, 4.1, 9.4, 0.3, { size: 10.5, bold: true, color: NAVY, valign: "middle", margin: 0 });
    [["yen", "買取利益が\n新たな収益に"], ["coins", "お客様の\n処分費用を軽減"], ["link", "「処分＋買取」の\nワンストップ提案"], ["award", "他社との\n差別化"]].forEach((m, i) => {
      const x = 0.3 + i * 2.37;
      rect(s, x, 4.42, 2.25, 0.62, WHITE, { line: NAVY, lw: 1, round: true, r: 0.08 });
      icon(s, m[0], x + 0.1, 4.48, 0.48, WHITE, NAVY);
      txt(s, m[1], x + 0.66, 4.42, 1.55, 0.62, { size: 9.5, bold: true, color: NAVY, valign: "middle", margin: 0, ls: 1.05 });
    });
    // 取得方法
    txt(s, "取得までの流れ（審査 約40日）", 0.3, 5.12, 6.0, 0.28, { size: 10.5, bold: true, color: NAVY, valign: "middle", margin: 0 });
    ["書類準備", "警察署へ申請", "審査 約40日", "許可証交付", "買取開始"].forEach((t, i) => chevron(s, 0.3 + i * 1.18, 5.42, 1.3, 0.38, i === 4 ? RED : NAVY, t, { first: i === 0, size: 8.5 }));
    rect(s, 0.3, 5.88, 2.3, 0.5, LBLUE, { round: true, r: 0.06 });
    txt(s, [B("費用 ", { fontSize: 9, color: NAVY }), B("申請手数料 約19,000円\n", { fontSize: 8 }), B("各種証明書 約1,000〜3,000円", { fontSize: 8 })], 0.38, 5.88, 2.2, 0.5, { valign: "middle", margin: 0, ls: 1.05 });
    rect(s, 2.7, 5.88, 3.6, 0.5, LBLUE, { round: true, r: 0.06 });
    txt(s, [B("主な必要書類 ", { fontSize: 9, color: NAVY }), B("申請書・住民票・身分証明書・誓約書・略歴書・法人登記事項証明書（法人）", { fontSize: 7.5 })], 2.78, 5.88, 3.48, 0.5, { valign: "middle", margin: 0, ls: 1.05 });
    rect(s, 6.45, 5.12, 3.25, 1.26, "FFF8E8", { line: GOLD, round: true, r: 0.05 });
    txt(s, "取得後に必ず守ること", 6.55, 5.15, 3.05, 0.26, { size: 10, bold: true, color: GOLDT, valign: "middle", margin: 0 });
    [["本人確認", "取引相手の身分証を確認・記録"], ["古物台帳への記録", "取引内容を記録し3年保管"], ["プレートの掲示", "営業所の見やすい場所に掲示"]].forEach((t, i) => {
      const y = 5.43 + i * 0.31;
      icon(s, "ok", 6.55, y + 0.04, 0.2, GOLD);
      txt(s, [B(t[0] + "　", { fontSize: 8.5 }), B(t[1], { fontSize: 7, color: GRAY })], 6.8, y, 2.85, 0.28, { valign: "middle", margin: 0 });
    });
    key(s, "必須ではありませんが、「買取」を加えると売上・利益がさらに伸びます。", { y: 6.45, size: 15 });
    foot(s, "※行政書士へ依頼する場合は別途報酬が必要です。");
  }

  // =====================================================================
  // 各業態におけるシナジー
  // =====================================================================
  const S5 = "各業態におけるシナジー";
  pres.addSection({ title: S5 });
  divider(S5, "各業態におけるシナジー");
  {
    const s = content(S5, "遺品整理がもたらすシナジー");
    msg(s, "どの業種にとっても、遺品整理には4つの共通する効果があります");
    const E = [["target", "商流の川上がとれる", "売却・解体・リフォームを決める前のお客様と、最初に出会える。相見積もりになりにくい", 0.4, 1.6, NAVY], ["store", "既存のお客様・資産が活きる", "既存顧客、車両、職人、店舗、会員組織などをそのまま活用できる", 6.2, 1.6, BLUE], ["coins", "外注していた作業が利益に", "残置物撤去や家財処分を外注している場合、その費用が自社の粗利になる", 0.4, 4.35, GREEN], ["friends", "地域の信頼と紹介が増える", "「困ったらあの会社」という存在になり、提携先・お客様からの紹介が増える", 6.2, 4.35, ORANGE]];
    line(s, 3.8, 2.75, 0.55, 0.45, LINE, 2); line(s, 5.65, 2.75, 0.55, 0.45, LINE, 2, { flipV: true });
    line(s, 3.8, 4.8, 0.55, 0.45, LINE, 2, { flipV: true }); line(s, 5.65, 4.8, 0.55, 0.45, LINE, 2);
    E.forEach((e) => {
      rect(s, e[3], e[4], 3.4, 1.85, WHITE, { line: e[5], lw: 1.5, round: true, r: 0.05 });
      icon(s, e[0], e[3] + 0.15, e[4] + 0.15, 0.55, WHITE, e[5]);
      txt(s, e[1], e[3] + 0.8, e[4] + 0.12, 2.5, 0.6, { size: 13, bold: true, color: e[5], valign: "middle", margin: 0 });
      txt(s, e[2], e[3] + 0.15, e[4] + 0.8, 3.1, 1.0, { size: 10.5, ls: 1.25 });
    });
    oval(s, 3.85, 2.85, 2.3, 2.3, NAVY);
    txt(s, [B("遺品整理\n", { fontSize: 18 }), B("＝本業の「入口」", { fontSize: 12, color: YEL })], 3.85, 2.85, 2.3, 2.3, { color: WHITE, align: "center", valign: "middle" });
    key(s, "遺品整理は、どの業種にとっても「本業の入口」になります。", { y: 6.45, size: 16 });
  }
  const INDS = [
    { key: "building", name: "不動産", issue: "仕入れ競争・広告費の高騰、案件の長期化", gain: "売却を考える前の所有者と出会える／残置物撤去の外注費が利益に", biz: "売買仲介・買取再販・賃貸管理" },
    { key: "hammer", name: "解体", issue: "価格競争、紹介依存、職人の稼働の波", gain: "解体前の案件を元請けでとれる／家財撤去の利益も自社に", biz: "解体工事・土地活用" },
    { key: "wrench", name: "住宅・リフォーム", issue: "新築縮小、大型案件不足、OB客との接点不足", gain: "相続した実家のリノベ相談の入口になる", biz: "リフォーム・リノベーション" },
    { key: "pray", name: "葬儀", issue: "家族葬化で単価下落、葬儀後の収益がない", gain: "葬儀後の困りごとを自社で受け、会員満足度が上がる", biz: "法要・供養・墓じまい・相続相談" },
    { key: "hospital", name: "介護（ケアマネ・包括）", issue: "片付けの相談先不足、業務負担の増加", gain: "施設入居前の生前整理に対応でき、家族の信頼が上がる", biz: "入居相談・地域連携" },
    { key: "tools", name: "便利業", issue: "小口案件ばかりで利益率が低い、ブランド力が弱い", gain: "高単価商品を持てる／専門店として信頼される", biz: "不用品回収・ハウスクリーニング" },
    { key: "recycle", name: "リサイクル", issue: "良質な仕入れが不安定、一般客との接点が少ない", gain: "良質なリユース品を直接仕入れられる", biz: "資源回収・リユース販売" },
    { key: "gem", name: "買取", issue: "高価値品の仕入れが不安定、買取できない物が多い", gain: "家一軒分から仕入れられ「全部お願いできる」が強みに", biz: "出張買取・店舗販売" },
  ];
  {
    const s = content(S5, "業種別シナジー一覧", "貴社の業種の行を指しながら説明し、次ページ以降の業種別ページへ。");
    msg(s, "8つの業種で、遺品整理が本業にもたらす効果");
    const X = [0.4, 2.45, 5.0, 7.75], W = [2.0, 2.5, 2.7, 1.85];
    ["業種", "本業の課題", "遺品整理で得られること", "つながる本業"].forEach((h, j) => {
      rect(s, X[j], 1.55, W[j], 0.42, NAVY);
      txt(s, h, X[j], 1.55, W[j], 0.42, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    });
    INDS.forEach((d, i) => {
      const y = 2.02 + i * 0.6, bg = i % 2 ? WHITE : LGRAY;
      W.forEach((w, j) => rect(s, X[j], y, w, 0.56, j === 0 ? LBLUE : bg));
      icon(s, d.key, X[0] + 0.08, y + 0.1, 0.36, NAVY);
      txt(s, d.name, X[0] + 0.5, y, 1.5, 0.56, { size: 10.5, bold: true, color: NAVY, valign: "middle", margin: 0 });
      txt(s, d.issue, X[1] + 0.05, y, W[1] - 0.1, 0.56, { size: 9, valign: "middle", margin: 0 });
      txt(s, d.gain, X[2] + 0.05, y, W[2] - 0.1, 0.56, { size: 9, bold: true, color: NAVY, valign: "middle", margin: 0 });
      txt(s, d.biz, X[3] + 0.05, y, W[3] - 0.1, 0.56, { size: 9, valign: "middle", margin: 0 });
    });
  }
  const INDUSTRIES = [
    { name: "不動産会社", icon: "building",
      issues: [["仕入れ競争の激化", "ポータル・一括査定経由で奪い合い"], ["集客コストの高騰", "Web広告・反響サイトの費用が上昇"], ["案件の長期化", "売却前の片付けや相続手続きで止まる"], ["残置物撤去の外注", "1現場30〜50万円で外注している"]],
      syn: [["相続案件の入口獲得", "売却を考える前の所有者と接点を持てる"], ["物件仕入れの強化", "相続物件の情報を競合より早く取得"], ["残置物撤去の内製化", "外注費を自社の粗利に転換"], ["売却スピードの向上", "片付けから売却まで一貫して進められる"]],
      flow: ["遺品整理", "不動産相談", "売却・買取再販", "リフォーム・解体"], sum: ["相続案件の入口獲得", "物件仕入れの強化", "売却スピード向上"] },
    { name: "解体業者", icon: "hammer",
      issues: [["利益率の低下", "解体単体の価格競争"], ["集客コストの増加", "ネット広告や紹介への依存"], ["新たな収益源の確保", "解体工事以外の売上の柱がほしい"], ["職人稼働の平準化", "工事の波で稼働が不安定"]],
      syn: [["解体前案件の獲得", "遺品整理を起点に解体案件を自社で獲得"], ["客単価の向上", "解体のみより受注単価が大きく向上"], ["粗利の向上", "外注していた家財撤去を自社で行う"], ["紹介依存からの脱却", "エンドユーザーから直接受注できる"]],
      flow: ["遺品整理", "家財撤去・処分", "解体工事", "不動産・売却"], sum: ["解体前案件の獲得", "客単価の向上", "粗利の向上", "安定受注"] },
    { name: "住宅・リフォーム業", icon: "wrench",
      issues: [["新築市場の縮小", "人口・世帯数の減少で需要が減少"], ["大型案件の不足", "小規模工事が増え利益が伸びない"], ["OB顧客との接点不足", "築年数が経ち関係が希薄化"], ["空き家問題の増加", "活用提案が求められている"]],
      syn: [["実家リノベ案件の獲得", "整理をきっかけに実家の活用相談へ"], ["相続空き家市場の攻略", "片付けから活用までワンストップ"], ["リフォーム受注の増加", "売るだけでなく「活かす」提案ができる"], ["他事業との連携強化", "不動産・解体・士業と連携し差別化"]],
      flow: ["遺品整理", "不動産・相続相談", "リノベ・リフォーム", "アフターサポート"], sum: ["実家リノベ案件の獲得", "相続空き家市場の攻略", "リフォーム受注の増加"] },
    { name: "葬儀業", icon: "pray",
      issues: [["葬儀単価の下落", "家族葬化・価格競争で収益が減少"], ["集客コストの高騰", "紹介サイトや広告費が上昇"], ["アフター収益の不足", "葬儀後の接点が少ない"], ["会員組織の活用不足", "会員・OB顧客を活かしきれていない"]],
      syn: [["葬儀後サポートの強化", "ご遺族の「その後の困りごと」を解決"], ["アフター売上の創出", "遺品整理・生前整理が新たな収益源に"], ["会員組織の活用", "会員特典で満足度と継続率が向上"], ["関連事業への展開", "墓じまい・不動産・相続相談へ"]],
      flow: ["葬儀施行", "遺品整理", "供養・生前整理", "不動産・相続連携"], sum: ["葬儀後サポート強化", "アフター売上創出", "会員組織の活用"] },
    { name: "ケアマネジャー・包括支援センター", short: "ケアマネ・包括支援", icon: "hospital",
      issues: [["相談先の不足", "片付けの専門業者の紹介先がない"], ["高齢者の住環境問題", "空き家・独居高齢者の家の片付け"], ["家族からの相談増加", "施設入居・相続などの相談が増加"], ["業務負担の増加", "片付け相談の対応に時間を取られる"]],
      syn: [["地域課題の解決", "空き家・放置物件の問題を解決"], ["生前整理への対応", "施設入居前の整理を支援"], ["家族支援の強化", "遠方の家族の負担を軽減"], ["地域包括ケアとの連携", "多職種と連携し支援の幅が広がる"]],
      flow: ["相談対応", "専門業者の紹介", "整理・片付け", "関連サービス連携"], sum: ["地域課題の解決", "生前整理への対応", "家族支援の強化"] },
    { name: "便利業", icon: "tools",
      issues: [["利益率が低い", "単価が低く利益が残りにくい"], ["案件数への依存", "小口案件が多く件数頼み"], ["生産性が低い", "移動や見積の時間が多い"], ["ブランド力が弱い", "専門性をアピールしづらい"]],
      syn: [["高単価商品化", "1件あたりの粗利が大幅にアップ"], ["顧客単価の向上", "生前整理・空き家片付けへ展開"], ["既存顧客の活用", "過去の小口顧客からリピート"], ["ブランド力の向上", "「遺品整理の専門店」として選ばれる"]],
      flow: ["生前整理・片付け", "遺品整理", "空き家片付け", "各種オプション"], sum: ["高単価商品化", "顧客単価の向上", "ブランド力の向上"] },
    { name: "リサイクル業", icon: "recycle",
      issues: [["良質な仕入れが不安定", "市場価格に左右されやすい"], ["利益率が低い", "回収・選別だけでは伸びにくい"], ["一般客との接点が少ない", "法人案件が中心"], ["回収できる物が限られる", "資源物以外は処分になる"]],
      syn: [["良質なリユース品を確保", "遺品整理の現場から直接回収"], ["お客様との直接取引", "一般家庭から直接受注できる"], ["リユース収益の最大化", "家具・家電・古物を適正に再販"], ["周辺サービスへ展開", "不用品回収・生前整理・解体へ"]],
      flow: ["遺品整理", "選別・仕分け", "リユース・買取", "資源回収・処分"], sum: ["良質な仕入れの確保", "利益率向上", "一般顧客の獲得"] },
    { name: "買取業", icon: "gem",
      issues: [["高価値品の仕入れが不安定", "持込み・出張買取に依存"], ["買取できない品物が多い", "対象外の品は機会損失に"], ["一度の売上が小さい", "買取だけでは利益拡大に限界"], ["相続世代との接点不足", "高単価案件の入口が少ない"]],
      syn: [["高価値品の仕入れ強化", "現場からブランド品・貴金属を直接買取"], ["買取できない物も対応", "処分・搬出まで「全部お願いできる」"], ["客単価・利益率アップ", "買取利益に整理・搬出の売上を追加"], ["次のビジネスへ展開", "生前整理・空き家整理・不動産紹介"]],
      flow: ["遺品整理相談", "査定・買取", "販売", "不用品回収・不動産紹介"], sum: ["高利益商品の安定仕入れ", "客単価向上", "継続的な収益機会の創出"] },
  ];
  function industrySlide(d) {
    const s = content(S5, "業種別シナジー｜" + (d.short || d.name));
    const nw = d.name.length * 18 / 72;
    icon(s, d.icon, 5 - nw / 2 - 0.55, 0.98, 0.45, WHITE, NAVY);
    txt(s, d.name, 0.4, 0.95, 9.2, 0.5, { size: 18, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    rect(s, 0.4, 1.65, 3.6, 3.45, LGRAY);
    txt(s, "経営課題", 0.55, 1.7, 3.3, 0.4, { size: 13, bold: true, color: MGRAY, margin: 0 });
    d.issues.forEach((it, i) => {
      const y = 2.15 + i * 0.73;
      icon(s, "warn", 0.55, y + 0.06, 0.3, MGRAY);
      txt(s, it[0], 0.95, y, 2.95, 0.36, { size: 11.5, bold: true, valign: "middle", margin: 0 });
      txt(s, it[1], 0.95, y + 0.35, 2.95, 0.33, { size: 9, color: GRAY, margin: 0 });
    });
    arrowR(s, 4.07, 3.1, 0.45, 0.55, RED);
    rect(s, 4.6, 1.65, 5.0, 3.45, LBLUE, { line: NAVY });
    txt(s, "遺品整理事業参入によるシナジー", 4.75, 1.7, 4.7, 0.4, { size: 13, bold: true, color: NAVY, margin: 0 });
    d.syn.forEach((sy, i) => {
      const x = 4.75 + (i % 2) * 2.4, y = 2.2 + Math.floor(i / 2) * 1.43;
      rect(s, x, y, 2.3, 1.3, WHITE, { round: true });
      num(s, i + 1, x + 0.08, y + 0.1, 0.36, NAVY);
      txt(s, sy[0], x + 0.5, y + 0.06, 1.75, 0.46, { size: 11, bold: true, color: NAVY, valign: "middle", margin: 0 });
      txt(s, sy[1], x + 0.1, y + 0.56, 2.1, 0.7, { size: 9.5, ls: 1.2, margin: 0 });
    });
    txt(s, "ビジネスモデルのイメージ", 0.4, 5.25, 4, 0.32, { size: 11, bold: true, color: PURPLE, margin: 0 });
    d.flow.forEach((f, i) => chevron(s, 0.4 + i * 2.28, 5.6, 2.42, 0.55, i === 0 ? PURPLE : "7E63A8", f, { first: i === 0, size: 10.5 }));
    rect(s, 0.4, 6.3, 9.2, 0.65, NAVY);
    const runs = [B("まとめ：", { color: WHITE })];
    d.sum.forEach((k) => runs.push(B("「" + k + "」", { color: YEL })));
    runs.push(B("を実現", { color: WHITE }));
    txt(s, runs, 0.55, 6.3, 8.9, 0.65, { size: 12, valign: "middle" });
  }
  function realEstateChange() {
    const s = content(S5, "補足｜不動産会社にとっての環境変化", "【要確認】法改正の出典（法務省の通達等）を確認し、出典欄に記載する。");
    msg(s, "不動産情報を「風上」で掴む方法の一つが、使えなくなります");
    line(s, 0.7, 2.25, 8.6, 0, LINE, 4);
    [["これまで", 1.1, MGRAY], ["2026年10月〜", 4.4, RED], ["これから", 7.7, NAVY]].forEach(([t, x, c]) => {
      oval(s, x + 0.5, 2.07, 0.36, 0.36, c);
      txt(s, t, x - 0.3, 1.65, 2.0, 0.35, { size: 12, bold: true, color: c, align: "center", margin: 0 });
    });
    const C = [["相続登記をきっかけに、法務局の不動産登記受付帳を開示請求し、所有者へDM・営業でアプローチできた", LGRAY, MGRAY], ["不動産登記受付帳の様式変更で「登記の目的」「不動産の所在」が表記されなくなり、営業目的での情報取得が事実上できなくなる", "FCE8E6", RED], ["ポータル・反響広告・既存の人脈での競争がさらに激化。新しい「風上の入口」が必要になる", LBLUE, NAVY]];
    C.forEach((c, i) => {
      const x = 0.4 + i * 3.13;
      rect(s, x, 2.7, 2.94, 1.95, c[1], { line: c[2] });
      txt(s, c[0], x + 0.15, 2.75, 2.64, 1.85, { size: 11, bold: true, valign: "middle", ls: 1.3 });
    });
    rect(s, 0.4, 4.9, 9.2, 1.4, NAVY);
    txt(s, [B("だからこそ今、新しい「風上チャネル」を持つ必要があります。\n", { color: WHITE, fontSize: 13 }), B("その新しい入口の一つが「空き家 × 遺品整理」です。", { color: YEL, fontSize: 20 })], 0.6, 4.9, 8.8, 1.4, { align: "center", valign: "middle", ls: 1.3 });
    foot(s, "出典：（要記入）");
  }
  function realEstateSim() {
    const s = content(S5, "補足｜不動産会社の収益シミュレーション");
    msg(s, "遺品整理を「不動産案件の入口」にすると、本業の利益がこれだけ増えます", { size: 14 });
    frame(s, 0.4, 1.6, 4.3, 4.65, "① 残置物撤去の内製化", NAVY, { hh: 0.5, hsize: 14 });
    txt(s, "売買・仲介で発生する残置物撤去を外注している場合", 0.55, 2.2, 4.0, 0.5, { size: 10.5, ls: 1.2 });
    rect(s, 0.6, 2.75, 3.9, 1.25, LBLUE);
    txt(s, [B("月5件 × 30〜50万円 × 12ヶ月\n", { fontSize: 12, color: NAVY }), B("＝ 年間 1,800〜3,000万円", { fontSize: 17, color: RED })], 0.6, 2.75, 3.9, 1.25, { align: "center", valign: "middle", ls: 1.3 });
    txt(s, "この外注費を、自社の粗利に転換できます。既存案件で仕事を確保できるため、立上げ時から現場経験を積めます。", 0.55, 4.15, 4.0, 1.2, { size: 10.5, ls: 1.3 });
    txt(s, "※A社様の例：月間5件の残置物撤去案件", 0.55, 5.7, 4.0, 0.35, { size: 9, color: GRAY, margin: 0 });
    frame(s, 5.0, 1.6, 4.6, 4.65, "② 遺品整理からつながる不動産案件", RED, { hh: 0.5, hsize: 14 });
    txt(s, "遺品整理 年間120件 × 不動産相談 5〜10% ＝ 年間6〜12件", 5.15, 2.2, 4.3, 0.5, { size: 10.5, bold: true, ls: 1.2 });
    const RK = [["S・A", "立地良・築浅", "買取再販", "200〜900万円", RED], ["B", "標準", "売買仲介", "300〜600万円", ORANGE], ["C", "立地悪・築古", "解体", "60〜300万円", MGRAY]];
    RK.forEach((r, i) => {
      const y = 2.8 + i * 0.75;
      rect(s, 5.15, y, 0.8, 0.62, r[4]);
      txt(s, r[0], 5.15, y, 0.8, 0.62, { size: 14, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      txt(s, r[1] + "\n→ " + r[2], 6.02, y, 1.7, 0.62, { size: 9.5, bold: true, valign: "middle", margin: 0 });
      txt(s, r[3], 7.7, y, 1.8, 0.62, { size: 12.5, bold: true, color: r[4], align: "right", valign: "middle", margin: 0 });
    });
    rect(s, 5.15, 5.1, 4.3, 0.9, "FCE8E6");
    txt(s, [B("年間粗利（シナジー分）\n", { fontSize: 10.5 }), B("約560万〜1,800万円", { fontSize: 19, color: RED })], 5.15, 5.1, 4.3, 0.9, { align: "center", valign: "middle" });
    key(s, "外注費の内製化 ＋ 新たな不動産案件の獲得で、売上アップ・利益拡大を実現します。", { y: 6.4, size: 14 });
    foot(s, MODEL_NOTE);
  }
  INDUSTRIES.forEach((d, i) => { industrySlide(d); if (i === 0) { realEstateChange(); realEstateSim(); } });

  // =====================================================================
  // ご質問への回答
  // =====================================================================
  const S6 = "ご質問への回答";
  pres.addSection({ title: S6 });
  divider(S6, "ご質問への回答", "事後経営相談はここから。最初に「貴社のお話を伺います」（書き込み版）で前回の内容を確認する。");
  {
    const s = content(S6, "いただいたご質問", "【要カスタマイズ】セミナーのアンケートや前回の質問を記入し、該当ページで回答する。");
    msg(s, "セミナー・前回の面談でいただいたご質問に、順にお答えします");
    [["いただいたご質問", 0.4, 3.7], ["回答の要点", 4.2, 4.3], ["ページ", 8.6, 1.0]].forEach(([h, X, W]) => {
      rect(s, X, 1.6, W, 0.45, NAVY);
      txt(s, h, X, 1.6, W, 0.45, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    });
    for (let i = 0; i < 6; i++) {
      const y = 2.15 + i * 0.75;
      icon(s, "question", 0.45, y + 0.17, 0.36, NAVY);
      txt(s, "Q" + (i + 1), 0.85, y, 0.5, 0.7, { size: 13, bold: true, color: NAVY, valign: "middle", margin: 0 });
      line(s, 0.4, y + 0.72, 9.2, 0, LINE, 0.75);
    }
  }
  {
    const s = content(S6, "収支計画（3年間）", "【要確認】1年目の営業利益：元資料の記載（106万円）と、経費の内訳の合計（586万円）が合わない。差額480万円の項目を確定させてから使う。");
    // 見出し
    txt(s, [B("数値計画", { fontSize: 26, color: NAVY }), B("（3年間）", { fontSize: 15, color: NAVY })], 0.4, 0.98, 3.0, 0.62, { valign: "middle", margin: 0 });
    line(s, 3.05, 1.06, 0, 0.48, MGRAY, 1);
    txt(s, [B("3年目には売上1億円を達成し、営業利益率", { color: DARK }), B("23.1%", { color: GOLD }), B("の高収益体制へ\n安定した成長と収益性の向上を実現します。", { color: DARK })], 3.15, 0.98, 4.2, 0.62, { size: 9.5, valign: "middle", margin: 0 });
    // 右上バッジ
    s.addShape(pres.shapes.PARALLELOGRAM, { x: 7.35, y: 0.9, w: 2.65, h: 0.75, fill: { color: NAVY }, line: { type: "none" }, objectName: "badge" });
    icon(s, "chart", 7.65, 1.03, 0.48, GOLD);
    txt(s, [B("3年目 売上 ", { color: WHITE, fontSize: 10 }), B("1億円\n", { color: GOLD, fontSize: 15 }), B("営業利益率 ", { color: WHITE, fontSize: 10 }), B("23.1%", { color: GOLD, fontSize: 15 })], 8.2, 0.9, 1.75, 0.75, { valign: "middle", margin: 0, ls: 1.0 });
    // 表
    const Y1 = NAVY, Y2 = "B8892B", Y3 = "2F4A80", GR = "E9EFF7", OP = "E6EDE4";
    const c = (t, o = {}) => ({ text: t, options: o });
    const bd = { type: "solid", pt: 0.75, color: "D3D9E2" };
    const base = (o) => Object.assign({ fontFace: FONT, fontSize: 9, color: DARK, valign: "middle", align: "left", margin: [2, 4, 2, 4], border: bd, fill: { color: WHITE } }, o);
    const H = (t, fill, cs) => c(t, base({ bold: true, color: WHITE, fontSize: 11, align: "center", fill: { color: fill }, colspan: cs }));
    const row = (label, vals, o = {}) => [c(label, base(Object.assign({ bold: !!o.bold }, o.lab || {}))), ...vals.flatMap((v) => [c(v[0], base(Object.assign({ bold: !!o.bold }, o.cell || {}))), c(v[1] || "", base(Object.assign({ fontSize: 7.5, margin: [2, 2, 2, 3] }, o.cell || {})))])];
    const grp = (t, rs) => c(t, base({ bold: true, fontSize: 11, color: WHITE, align: "center", fill: { color: NAVY }, rowspan: rs }));
    const rows = [
      [H("項目", NAVY, 2), H("1年目", Y1, 2), H("2年目", Y2, 2), H("3年目", Y3, 2)],
      [c("売上", base({ bold: true, fontSize: 12, align: "center", fill: { color: "F4F6FA" }, colspan: 2 })), ...["¥36,000,000", "¥54,000,000", "¥100,000,000"].map((v) => c(v, base({ bold: true, fontSize: 13, align: "center", fill: { color: "F4F6FA" }, colspan: 2 })))],
      [grp("原価", 3), ...row("処分費", [["¥10,800,000", "30%"], ["¥16,200,000", "30%"], ["¥30,000,000", "30%"]])],
      row("その他消耗品費", [["¥1,000,000", "資材等"], ["¥1,500,000", "件数増加分"], ["¥2,000,000", "件数増加分"]]),
      row("粗利", [["¥24,200,000", "67%"], ["¥36,300,000", "67%"], ["¥68,000,000", "68%"]], { bold: true, lab: { fill: { color: GR }, color: NAVY }, cell: { fill: { color: GR } } }),
      [grp("経費", 5), ...row("アルバイト事務", [["¥1,200,000", "事務1名"], ["¥1,200,000", "事務1名"], ["¥1,200,000", "事務1名"]])],
      row("広告宣伝費", [["¥4,840,000", "粗利の20%"], ["¥7,260,000", "粗利の20%"], ["¥13,600,000", "粗利の20%"]]),
      row("その他事業経費", [["¥1,500,000"], ["¥2,000,000"], ["¥2,500,000"]]),
      row("人件費", [["¥6,000,000", "専任1名"], ["¥12,000,000", "専任2名"], ["¥18,000,000", "専任3名"]]),
      row("作業員アルバイト", [["¥4,800,000", "2名"], ["¥4,800,000", "2名"], ["¥9,600,000", "4名"]]),
      [c("営業利益", base({ bold: true, fontSize: 11, align: "center", fill: { color: OP }, colspan: 2 })), ...["¥1,060,000", "¥9,040,000", "¥23,100,000"].flatMap((v) => [c(v, base({ bold: true, fontSize: 10, fill: { color: OP } })), c("", base({ fill: { color: OP } }))])],
      [c("営業利益率", base({ bold: true, fontSize: 11, align: "center", fill: { color: OP }, colspan: 2 })), ...["2.94%", "16.74%", "23.10%"].flatMap((v) => [c(v, base({ bold: true, fontSize: 10, fill: { color: OP } })), c("", base({ fill: { color: OP } }))])],
    ];
    const TX = 0.35, TY = 1.75, RH = 0.36;
    const colW = [0.45, 1.2, 1.0, 0.8, 1.0, 0.8, 1.0, 0.8];
    s.addTable(rows, { x: TX, y: TY, w: colW.reduce((a, b) => a + b, 0), colW, rowH: [0.38, 0.44, RH, RH, RH, RH, RH, RH, RH, RH, 0.38, 0.38], autoPage: false, objectName: "pl_table" });
    // 強調の赤枠（人件費 1年目／営業利益率）
    const rowY = (i) => TY + 0.38 + 0.44 + (i - 2) * RH;
    const redFrame = (x, y, w, h) => s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: WHITE, transparency: 100 }, line: { color: RED, width: 2.25 }, objectName: nm("redframe") });
    redFrame(TX + 0.45, rowY(8), 1.2 + 1.0 + 0.8, RH);
    redFrame(TX, rowY(11) + (0.38 - RH), 7.05, 0.38);
    // 右：初期投資・投資回収
    const PX = 7.55, PW = 2.1;
    rect(s, PX, 1.75, PW, 2.25, WHITE, { line: NAVY, lw: 1.25, shadow: true });
    rect(s, PX, 1.75, PW, 0.45, NAVY);
    txt(s, "初期投資", PX, 1.75, PW, 0.45, { size: 14, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    icon(s, "laptop", PX + 0.12, 2.38, 0.5, NAVY);
    txt(s, [B("HP制作費\n", { fontSize: 9.5 }), B("150", { fontSize: 16, color: RED }), B("万円", { fontSize: 10, color: RED })], PX + 0.68, 2.3, PW - 0.72, 0.65, { valign: "middle", margin: 0 });
    line(s, PX + 0.12, 3.03, PW - 0.24, 0, LINE, 0.75);
    icon(s, "list", PX + 0.15, 3.18, 0.45, NAVY);
    txt(s, [B("古物商許可取得費用\n", { fontSize: 9.5 }), B("5〜8", { fontSize: 18, color: RED }), B("万円", { fontSize: 10, color: RED })], PX + 0.68, 3.08, PW - 0.72, 0.85, { valign: "middle", margin: 0 });
    s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: PX + PW / 2 - 0.3, y: 4.1, w: 0.6, h: 0.25, rotate: 180, fill: { color: NAVY }, line: { type: "none" }, objectName: "tri_pl" });
    rect(s, PX, 4.45, PW, 1.95, WHITE, { line: NAVY, lw: 1.25, shadow: true });
    rect(s, PX, 4.45, PW, 0.45, NAVY);
    txt(s, "投資回収", PX, 4.45, PW, 0.45, { size: 14, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, [B("2年目", { fontSize: 22, color: RED }), B("で回収", { fontSize: 12, color: RED })], PX + 0.1, 4.98, 1.5, 0.6, { valign: "middle", margin: 0 });
    icon(s, "target", PX + 1.62, 5.05, 0.4, NAVY);
    txt(s, [B("2年目の営業利益 ", { fontSize: 9 }), B("904万円", { fontSize: 10, color: RED }), B("により、初期投資を回収見込み。", { fontSize: 9 })], PX + 0.1, 5.65, PW - 0.2, 0.7, { ls: 1.2, margin: 0 });
    // 特徴
    txt(s, [B("【特徴】", { color: DARK }), B("初期投資費用が低く、社員1名からスタート可能", { color: RED }), B("であり、", { color: DARK }), B("20%以上の利益率", { color: RED }), B("を目指せるビジネスモデルである。", { color: DARK })], 0.35, 6.4, 9.3, 0.6, { size: 11.5, valign: "middle", margin: 0, ls: 1.1 });
    foot(s, MODEL_NOTE);
  }
  {
    const s = content(S6, "人員体制と見積・処分の仕組み");
    msg(s, "トラック・倉庫・専門の目利きがなくても始められる3つの仕組み");
    frame(s, 0.4, 1.6, 3.0, 4.75, "① 人員体制", NAVY);
    [["1年目", 1, "営業＋現場作業"], ["2年目", 2, "営業＋現場管理"], ["3年目", 3, "案件管理＋現場監督"]].forEach((r, i) => {
      const y = 2.2 + i * 1.0;
      txt(s, r[0], 0.55, y, 0.75, 0.8, { size: 12, bold: true, color: NAVY, valign: "middle", margin: 0 });
      for (let k = 0; k < r[1]; k++) icon(s, "user", 1.3 + k * 0.36, y + 0.08, 0.32, NAVY);
      txt(s, "正社員" + r[1] + "名", 1.3, y + 0.45, 1.2, 0.3, { size: 9.5, bold: true, margin: 0 });
      txt(s, r[2], 2.4, y + 0.05, 0.95, 0.75, { size: 9.5, color: GRAY, valign: "middle", margin: 0 });
    });
    rect(s, 0.55, 5.3, 2.7, 0.9, LBLUE);
    txt(s, "作業は派遣・アルバイト中心。正社員は営業と管理に集中", 0.6, 5.3, 2.6, 0.9, { size: 10, bold: true, color: NAVY, valign: "middle" });
    frame(s, 3.5, 1.6, 3.0, 4.75, "② 見積の仕組み", ORANGE);
    [["box", "140サイズ10箱", "＝ 1立米"], ["truck", "2tトラック", "＝ 10〜12立米"], ["people", "作業量の目安", "1人1日 1.2立米"]].forEach((r, i) => {
      icon(s, r[0], 3.7, 2.2 + i * 0.8, 0.55, ORANGE);
      txt(s, [B(r[1] + "\n", { fontSize: 11 }), B(r[2], { fontSize: 15, color: ORANGE })], 4.35, 2.15 + i * 0.8, 2.1, 0.7, { valign: "middle" });
    });
    rect(s, 3.65, 4.65, 2.7, 1.55, LORANGE);
    txt(s, [B("見積 ＝ 物量 × 立米単価\n＋ オプション − 買取額\n", { fontSize: 10.5 }), B("立米単価 1.2〜1.5万円", { fontSize: 11, color: ORANGE })], 3.7, 4.65, 2.6, 1.55, { align: "center", valign: "middle", ls: 1.3 });
    frame(s, 6.6, 1.6, 3.0, 4.75, "③ 処分費を下げる仕組み", GREEN);
    ["国内で買取・販売", "海外へ輸出リユース", "まとめ売り（福袋）", "自社で再利用", "最後に廃棄"].forEach((r, i) => {
      rect(s, 6.8, 2.2 + i * 0.8, 2.6, 0.52, i === 4 ? MGRAY : GREEN, { round: true });
      txt(s, r, 6.8, 2.2 + i * 0.8, 2.6, 0.52, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      if (i < 4) arrowD(s, 7.97, 2.74 + i * 0.8, 0.26, 0.24, "8FBF9F");
    });
    key(s, "社員1名＋外部スタッフで、未経験から始められます。", { y: 6.45, size: 16 });
  }
  {
    const s = content(S6, "法令遵守と運営体制", "「100%合法」とは言わず、前提条件を守れば法令に沿って運営できると説明する。自治体ごとのルールは事前に確認する。");
    msg(s, "三社間契約の前提条件と、本業のブランドを守る運営体制");
    const L = [["law", "三社間契約で守ること", NAVY, ["お客様と許可業者が、収集運搬・処分の契約を直接結ぶ", "当社は仕分け・搬出までを担い、廃棄物を運搬しない", "廃棄物は現場で許可業者のパッカー車へ引き渡す", "自治体のルールを事前に確認する"]], ["shield", "本業のブランドを守る仕組み", GREEN, ["請負契約書・買取同意書を必ず交わす", "古物台帳で買取の記録を3年間保管", "損害賠償保険に加入する", "作業前の写真撮影・貴重品返却シートでトラブルを防ぐ"]]];
    L.forEach((l, i) => {
      const x = 0.4 + i * 4.7;
      rect(s, x, 1.6, 4.5, 3.75, WHITE, { line: l[2], lw: 1.5 });
      icon(s, l[0], x + 0.15, 1.72, 0.55, WHITE, l[2]);
      txt(s, l[1], x + 0.8, 1.72, 3.6, 0.55, { size: 14, bold: true, color: l[2], valign: "middle", margin: 0 });
      l[3].forEach((t, k) => {
        icon(s, "ok", x + 0.2, 2.52 + k * 0.7, 0.3, l[2]);
        txt(s, t, x + 0.6, 2.42 + k * 0.7, 3.8, 0.5, { size: 11, valign: "middle", margin: 0 });
      });
    });
    rect(s, 0.4, 5.55, 9.2, 0.75, LBLUE, { round: true });
    icon(s, "contract", 0.6, 5.67, 0.5, NAVY);
    txt(s, "運用マニュアル・契約書のひな形（請負契約書／買取同意書／業務提携契約書）をご用意しています", 1.25, 5.55, 8.2, 0.75, { size: 11.5, bold: true, color: NAVY, valign: "middle" });
    key(s, "法令に沿った運営体制を、最初から整えて始めます。", { y: 6.45, size: 16 });
  }
  {
    const s = content(S6, "競合環境と参入時期", "【要カスタマイズ】商圏の競合状況（Googleマップ・HPで調査）を記入する。");
    msg(s, "専門店がいない今、専門店として打ち出せば後発でも勝てます");
    dashBox(s, 0.4, 1.6, 4.5, 3.6, "map", "貴社商圏の競合状況\n商圏内の遺品整理業者 ○社（うち専門店 ○社）\n不用品回収・便利屋 ○社");
    txt(s, "勝つための3つの打ち手", 5.1, 1.6, 4.5, 0.4, { size: 14, bold: true, color: RED, margin: 0 });
    [["store", "「遺品整理専門店」の屋号とHPで打ち出す"], ["link", "供養・買取・不動産相談までワンストップで対応"], ["shake", "不動産・士業・葬儀社との提携で、相見積もりのない紹介を得る"]].forEach((t, i) => {
      const y = 2.1 + i * 1.03;
      rect(s, 5.1, y, 4.5, 0.9, i % 2 ? WHITE : "FCE8E6", { line: "F2B8B4" });
      icon(s, t[0], 5.25, y + 0.2, 0.5, WHITE, RED);
      txt(s, t[1], 5.9, y, 3.6, 0.9, { size: 11.5, bold: true, valign: "middle", margin: 0 });
    });
    rect(s, 0.4, 5.4, 9.2, 0.95, NAVY);
    icon(s, "clock", 0.6, 5.6, 0.55, YEL);
    txt(s, [B("準備（提携先の開拓・HP制作・営業研修）に約6ヶ月", { color: YEL }), B("かかるため、\n広告単価が上がる前に、早めに着手することが重要です。", { color: WHITE })], 1.3, 5.4, 8.2, 0.95, { size: 12, valign: "middle" });
    key(s, "専門店がいない今が、参入の時期です。", { y: 6.45, size: 16 });
  }

  // =====================================================================
  // 貴社で進めるには
  // =====================================================================
  const S7 = "貴社で進めるには";
  pres.addSection({ title: S7 });
  divider(S7, "貴社で進めるには");
  {
    const s = content(S7, "貴社商圏の市場", "【要カスタマイズ】地域データは自治体の統計・住宅土地統計調査などから取得する。");
    msg(s, "貴社の商圏で、遺品整理の市場規模を試算します");
    const K = [["users", "商圏人口", "○○万人"], ["pray", "年間死亡者数", "○,○○○人"], ["home", "65歳以上の単身世帯", "○,○○○世帯"], ["key", "空き家数", "○,○○○戸"]];
    K.forEach((k, i) => {
      const x = 0.4 + i * 2.33;
      rect(s, x, 1.6, 2.15, 1.5, LBLUE);
      icon(s, k[0], x + 0.12, 1.72, 0.4, NAVY);
      txt(s, k[1], x + 0.55, 1.68, 1.55, 0.48, { size: 10.5, bold: true, color: NAVY, valign: "middle", margin: 0 });
      txt(s, k[2], x, 2.25, 2.15, 0.7, { size: 22, bold: true, color: RED, align: "center", valign: "middle", margin: 0 });
    });
    dashBox(s, 0.4, 3.3, 5.0, 3.0, "map", "商圏地図（対象エリアと競合の位置）");
    frame(s, 5.6, 3.3, 4.0, 3.0, "市場規模の試算", NAVY);
    txt(s, [B("参考：商圏人口20万人で 約8.4億円\n\n", { fontSize: 10.5, color: GRAY }), B("貴社商圏\n", { fontSize: 12 }), B("約○億円\n", { fontSize: 26, color: RED }), B("3年目の目標シェア 約○%", { fontSize: 11 })], 5.7, 3.8, 3.8, 2.4, { align: "center", valign: "middle" });
  }
  {
    const s = content(S7, "貴社版の事業コンセプト", "【要カスタマイズ】ヒアリング内容から、売り先・売り物・売り方、強み、屋号案を記入する。");
    msg(s, "売り先・売り物・売り方を絞り込み、貴社の強みを活かした専門店にします", { size: 14 });
    const V = [["売り先", "誰に売るか", "例）親の家を相続する50〜60代、賃貸オーナー", 1.15, 1.65, NAVY], ["売り物", "何を売るか", "例）遺品整理の標準パック＋供養・買取", 3.15, 1.65, BLUE], ["売り方", "どう売るか", "例）現地確認・即決型、提携先からの紹介", 2.15, 3.3, GREEN]];
    V.forEach((v) => s.addShape(pres.shapes.OVAL, { x: v[3], y: v[4], w: 2.6, h: 2.6, fill: { color: v[5], transparency: 25 }, line: { color: WHITE, width: 2 }, objectName: nm("venn") }));
    txt(s, [B("売り先\n", { fontSize: 15 }), B("誰に売るか", { fontSize: 10 })], 1.3, 2.1, 1.5, 0.8, { color: WHITE, align: "center" });
    txt(s, [B("売り物\n", { fontSize: 15 }), B("何を売るか", { fontSize: 10 })], 4.1, 2.1, 1.5, 0.8, { color: WHITE, align: "center" });
    txt(s, [B("売り方\n", { fontSize: 15 }), B("どう売るか", { fontSize: 10 })], 2.7, 4.95, 1.5, 0.8, { color: WHITE, align: "center" });
    txt(s, "遺品整理\n専門店", 2.8, 3.2, 1.3, 0.75, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    V.forEach((v, i) => {
      const y = 1.6 + i * 1.18;
      rect(s, 6.15, y, 3.45, 1.05, WHITE, { line: v[5], lw: 1.25 });
      txt(s, v[0] + "｜貴社の場合", 6.25, y + 0.03, 3.3, 0.3, { size: 10.5, bold: true, color: v[5], margin: 0 });
      txt(s, v[2], 6.25, y + 0.35, 3.3, 0.65, { size: 9.5, color: GRAY, ls: 1.2, margin: 0 });
    });
    rect(s, 6.15, 5.15, 3.45, 1.15, "FFFBEA", { line: GOLD });
    txt(s, [B("屋号・コンセプト案\n", { fontSize: 10.5, color: GOLD }), B("例）「○○の遺品整理」\n想いをつなぐ、地域の片付け相談室", { fontSize: 9.5, color: GRAY })], 6.25, 5.15, 3.3, 1.15, { valign: "middle", ls: 1.2 });
    rect(s, 0.4, 6.15, 5.55, 0.8, LGRAY);
    txt(s, [B("貴社の強み：", { color: NAVY }), B("例）地域での施工実績・既存顧客数・資格・車両など", { color: GRAY })], 0.55, 6.15, 5.35, 0.8, { size: 10.5, valign: "middle" });
  }
  {
    const s = content(S7, "貴社版の料金設計", "【要カスタマイズ】料金はパック料金で見せる（内訳をそのまま見せると不安・不信感につながる）。箱数・立米を併記し、見積の計算と一致させる。平均単価30万円になるのは、戸建2LDK以上が中心になるため。");
    msg(s, "お客様には分かりやすいパック料金、社内は立米単価で原価を管理します", { size: 14 });
    const PR = [["1K", "30,000", "段ボール20箱\n（約2立米）", "1〜2名／1〜3時間"], ["1LDK", "80,000", "段ボール50箱\n（約5立米）", "2〜3名／3〜6時間"], ["2LDK", "150,000", "段ボール100箱\n（約10立米）", "3〜4名／5〜8時間"], ["3LDK", "200,000", "段ボール150箱\n（約15立米）", "4〜6名／6〜9時間"], ["4LDK以上・戸建", "300,000", "段ボール200箱以上\n（約20立米〜）", "5名〜／1〜2日"]];
    PR.forEach((p, i) => {
      const x = 0.4 + i * 1.86, hot = i === 2;
      rect(s, x, 1.65, 1.76, 2.95, WHITE, { line: hot ? RED : LINE, lw: hot ? 2 : 1, shadow: true });
      rect(s, x, 1.65, 1.76, 0.6, hot ? RED : NAVY);
      txt(s, p[0], x, 1.65, 1.76, 0.6, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      txt(s, [B(p[1], { fontSize: 18, color: hot ? RED : NAVY }), B("円〜", { fontSize: 10, color: hot ? RED : NAVY })], x, 2.35, 1.76, 0.6, { align: "center", valign: "middle", margin: 0 });
      txt(s, p[2], x + 0.05, 3.0, 1.66, 0.75, { size: 9.5, align: "center", ls: 1.2, margin: 0 });
      txt(s, p[3], x + 0.05, 3.8, 1.66, 0.6, { size: 9, color: GRAY, align: "center", ls: 1.2, margin: 0 });
      if (hot) { rect(s, x + 0.28, 1.56, 1.2, 0.22, YEL, { round: true }); txt(s, "よくあるご依頼", x + 0.28, 1.56, 1.2, 0.22, { size: 8.5, bold: true, align: "center", valign: "middle", margin: 0 }); }
    });
    txt(s, [HL("仕分け・搬出・処分・簡易清掃・運搬まで含んだパック料金（追加料金なし）", { color: DARK })], 0.4, 4.7, 9.2, 0.35, { size: 11.5, align: "center", margin: 0 });
    txt(s, "オプション（必要な場合のみ）", 0.4, 5.15, 3, 0.3, { size: 11, bold: true, color: NAVY, margin: 0 });
    const O = [["pray", "供養", "15,000円〜"], ["broom", "特殊清掃", "30,000円〜"], ["home", "原状回復", "40,000円〜"], ["hammer", "解体", "100,000円〜"], ["building", "不動産売却", "ご相談"], ["key", "空き家管理", "5,000円〜/月"], ["gem", "買取", "お値引き"]];
    O.forEach((o, i) => {
      const x = 0.4 + i * 1.33;
      rect(s, x, 5.5, 1.25, 1.35, LBLUE, { round: true });
      icon(s, o[0], x + 0.43, 5.58, 0.38, NAVY);
      txt(s, o[1], x, 5.98, 1.25, 0.35, { size: 10, bold: true, color: NAVY, align: "center", margin: 0 });
      txt(s, o[2], x, 6.33, 1.25, 0.4, { size: 9.5, align: "center", margin: 0 });
    });
  }
  {
    const s = content(S7, "貴社版の数値計画", "【要カスタマイズ】1年目は準備期間（約6ヶ月）を含めた月ごとの立ち上がりで示す。兼任スタート案と専任採用案の2つを並べる。");
    msg(s, "貴社の商圏と体制に合わせた3年間の計画");
    s.addChart(pres.charts.BAR, [{ name: "月間契約件数（目安）", labels: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"], values: [0, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10] }], {
      x: 0.4, y: 1.55, w: 4.8, h: 3.6, barDir: "col", chartColors: [NAVY], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 9, dataLabelColor: DARK,
      showTitle: true, title: "1年目の月間契約件数（例）", titleFontSize: 12, titleColor: NAVY, titleFontFace: "+mn-lt",
      catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, catAxisLabelColor: GRAY, valAxisLabelColor: GRAY, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
      valGridLine: { color: "E5E7EB", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false, objectName: "ramp_chart",
    });
    rect(s, 0.6, 5.2, 2.0, 0.35, LGRAY); txt(s, "準備期（1〜6ヶ月目）", 0.6, 5.2, 2.0, 0.35, { size: 9, bold: true, color: GRAY, align: "center", valign: "middle", margin: 0 });
    rect(s, 2.75, 5.2, 2.3, 0.35, LBLUE); txt(s, "営業開始（7ヶ月目〜）", 2.75, 5.2, 2.3, 0.35, { size: 9, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    table(s, [["", "兼任スタート案", "専任採用案"], ["1年目の体制", "既存営業1名が兼任", "専任営業1名を採用"], ["人件費", "本業と按分", "年600万円"], ["向いている会社", "片付け相談がすでにある", "早く立ち上げたい"], ["1年目売上", "○○万円", "○○万円"], ["3年目売上", "○○万円", "○○万円"]], 5.4, 1.6, 4.2, [1.3, 1.45, 1.45], { size: 10, rowH: 0.52, firstColFill: true });
    rect(s, 5.4, 4.9, 4.2, 1.1, "FAFAFA", { line: "9AA5B8", dash: "dash" });
    txt(s, "前提条件：商圏人口・平均単価・契約率・広告予算など", 5.5, 4.9, 4.0, 1.1, { size: 10, color: GRAY, valign: "middle" });
    key(s, "準備期間を含めても、無理のない立ち上がりを計画します。", { y: 6.35, size: 16 });
    foot(s, MODEL_NOTE);
  }
  {
    const s = content(S7, "本業込みの利益", "【要カスタマイズ】仲介手数料は宅建業法の上限（売買価格400万円超：3%＋6万円、税別）で計算。発生率は貴社の実績や業種に合わせて調整する。");
    msg(s, "遺品整理の利益に、本業への波及効果を加えて考えます");
    rect(s, 0.4, 1.55, 9.2, 0.7, LBLUE, { round: true });
    txt(s, [B("本業込みの利益 ＝ ", { color: NAVY }), B("遺品整理の粗利", { color: ORANGE }), B(" ＋ 不動産の発生率×", { color: NAVY }), B("仲介手数料", { color: BLUE }), B(" ＋ 工事の発生率×", { color: NAVY }), B("工事粗利", { color: GREEN })], 0.5, 1.55, 9.0, 0.7, { size: 13, align: "center", valign: "middle" });
    s.addChart(pres.charts.BAR, [
      { name: "遺品整理", labels: ["遺品整理のみ", "本業込み"], values: [180, 180] },
      { name: "不動産売買", labels: ["遺品整理のみ", "本業込み"], values: [0, 132] },
      { name: "解体工事", labels: ["遺品整理のみ", "本業込み"], values: [0, 45] },
    ], {
      x: 0.4, y: 2.45, w: 4.6, h: 3.8, barDir: "col", barGrouping: "stacked", chartColors: [ORANGE, BLUE, GREEN], showValue: true, dataLabelPosition: "ctr", dataLabelFontSize: 10, dataLabelColor: WHITE, dataLabelFontFace: "+mn-lt",
      showTitle: true, title: "10件受託した場合の粗利（万円）", titleFontSize: 11, titleColor: NAVY, titleFontFace: "+mn-lt", catAxisLabelFontSize: 11, catAxisLabelColor: DARK, catAxisLabelFontFace: "+mn-lt",
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: true, legendPos: "b", legendFontSize: 9, legendFontFace: "+mn-lt", objectName: "ltv_chart",
    });
    table(s, [["項目（10件）", "計算", "粗利"], ["遺品整理", "30万円×10件×粗利率60%", "180万円"], ["不動産売買（1件）", "2,000万円の仲介手数料（両手）", "132万円"], ["解体工事（1件）", "工事費150万円×粗利率30%", "45万円"], [{ text: "合計", options: { bold: true } }, "", { text: "357万円", options: { bold: true, color: RED } }]], 5.2, 2.5, 4.4, [1.3, 2.2, 0.9], { size: 9.5, rowH: 0.45 });
    rect(s, 5.2, 4.95, 4.4, 1.25, "FCE8E6", { round: true });
    txt(s, [B("1件あたりの粗利\n", { fontSize: 11 }), B("18万円 → ", { fontSize: 15, color: MGRAY }), B("約36万円", { fontSize: 24, color: RED })], 5.2, 4.95, 4.4, 1.25, { align: "center", valign: "middle" });
    key(s, "本業まで含めると、遺品整理単体の約2倍の利益が残ります。", { y: 6.4, size: 16 });
    foot(s, "※仲介手数料は宅建業法の上限（3%＋6万円、税別）で計算。モデル数値であり成果を保証するものではありません。");
  }
  {
    const s = content(S7, "集客の進め方", "【要カスタマイズ】広告予算。立上げ直後はテスト運用で反響の出るキーワード・エリアを見極め、効果の高い媒体に絞って拡大する。");
    msg(s, "4つのチャネルで、安定して反響をとります");
    const CH = [["laptop", "WEB", "専門HP・SEO・MEO・広告", "今すぐ依頼したい人を集める", "約3万円／件", BLUE], ["paper", "チラシ", "戸建・B4・保存版", "検索していない人に認知してもらう", "約5万円／件", ORANGE], ["shake", "提携先の紹介", "不動産・士業・葬儀社など", "相見積もりのない紹介を得る", "成約時のみ手数料", PURPLE], ["users", "既存のお客様", "本業の顧客・会員", "本業のお客様に案内する", "ほぼ0円", GREEN]];
    CH.forEach((c, i) => {
      const x = 0.4 + i * 2.33;
      rect(s, x, 1.6, 2.2, 2.9, WHITE, { line: c[5], lw: 1.5, round: true, r: 0.05 });
      icon(s, c[0], x + 0.75, 1.75, 0.7, WHITE, c[5]);
      txt(s, c[1], x, 2.52, 2.2, 0.4, { size: 14, bold: true, color: c[5], align: "center", margin: 0 });
      txt(s, c[2], x + 0.1, 2.9, 2.0, 0.35, { size: 9, color: GRAY, align: "center", margin: 0 });
      txt(s, c[3], x + 0.1, 3.25, 2.0, 0.6, { size: 10, bold: true, align: "center", ls: 1.2 });
      rect(s, x + 0.15, 3.95, 1.9, 0.4, LGRAY, { round: true });
      txt(s, c[4], x + 0.15, 3.95, 1.9, 0.4, { size: 10, bold: true, color: c[5], align: "center", valign: "middle", margin: 0 });
    });
    txt(s, "広告予算の考え方", 0.4, 4.7, 9.2, 0.35, { size: 13, bold: true, color: NAVY, margin: 0 });
    chevron(s, 0.4, 5.1, 3.25, 0.95, BLUE, "テスト期（〜3ヶ月）\n月20万円で反響を検証", { first: true, size: 10.5 });
    chevron(s, 3.45, 5.1, 3.3, 0.95, NAVY, "拡大期（4ヶ月目〜）\n粗利の15〜20%を目安に", { size: 10.5 });
    chevron(s, 6.55, 5.1, 3.05, 0.95, GREEN, "安定期\n紹介・SEOで広告費を削減", { size: 10.5 });
    key(s, "広告ですぐ反響をとり、紹介とSEOで広告費を下げていきます。", { y: 6.35, size: 16 });
  }
  {
    const s = content(S7, "WEB集客の全体像", "【話し方】①まずは反響が取れる専門HP（受け皿）をつくる。HPでお客様の6つの不安を解消する→②次にSEO・MEO・PPC広告で流入を増やす。PPCはすぐ、MEOは1〜3ヶ月、SEOは3〜6ヶ月で効果が出始める→③この順番を守ると、広告費のムダが出ない。");
    txt(s, [B("遺品整理のWEB集客の全体像", { color: NAVY }), B("【成功の2ステップ】", { color: NAVY })], 0.3, 0.9, 9.4, 0.38, { size: 18, align: "center", valign: "middle", margin: 0 });
    txt(s, [B("まずは", { color: NAVY }), B("反響が取れる専門HP", { color: RED }), B("を構築し、次に", { color: NAVY }), B("SEO・MEO・PPC広告", { color: RED }), B("で流入数を増やします", { color: NAVY })], 0.3, 1.27, 9.4, 0.28, { size: 11, align: "center", valign: "middle", margin: 0 });
    const PY = 1.65, PH = 3.85;
    // ---- STEP1 ----
    const AX = 0.3, AW = 4.55;
    rect(s, AX, PY, AW, PH, WHITE, { line: NAVY, lw: 1.25, round: true, r: 0.04 });
    rect(s, AX, PY, AW, 0.5, NAVY, { round: true, r: 0.04 });
    num(s, 1, AX + 0.12, PY + 0.07, 0.36, WHITE, NAVY);
    txt(s, "反響が取れる専門HPを構築する", AX + 0.55, PY, AW - 0.6, 0.5, { size: 13.5, bold: true, color: WHITE, valign: "middle", margin: 0 });
    txt(s, "＝ 信頼され、問い合わせが増える「受け皿」をつくる", AX, PY + 0.52, AW, 0.27, { size: 9, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    // HPイメージ
    const HX = AX + 0.12, HY = PY + 0.85, HW = 2.25, HH = 1.8;
    rect(s, HX, HY, HW, HH, WHITE, { line: "B8C0CC", lw: 1 });
    rect(s, HX, HY, HW, 0.22, "EEF1F6");
    txt(s, "○○遺品整理センター", HX + 0.06, HY, 1.3, 0.22, { size: 6.5, bold: true, color: NAVY, valign: "middle", margin: 0 });
    rect(s, HX + 1.45, HY + 0.03, 0.74, 0.16, GREEN, { round: true, r: 0.05 });
    txt(s, "☎ 0120-xxx-xxx", HX + 1.45, HY + 0.03, 0.74, 0.16, { size: 5.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    rect(s, HX, HY + 0.22, HW, 0.82, "E8EBF0");
    txt(s, "遺品整理・生前整理は\nプロにお任せください", HX + 0.08, HY + 0.28, 1.5, 0.45, { size: 8, bold: true, color: NAVY, valign: "middle", margin: 0, ls: 1.05 });
    oval(s, HX + 1.6, HY + 0.32, 0.58, 0.58, YEL);
    txt(s, "実績\n1,000件", HX + 1.6, HY + 0.32, 0.58, 0.58, { size: 6.5, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0, ls: 1.0 });
    ["見積無料", "追加料金なし", "即日対応"].forEach((t, i) => { rect(s, HX + 0.08 + i * 0.5, HY + 0.78, 0.46, 0.17, BLUE, { round: true, r: 0.05 }); txt(s, t, HX + 0.08 + i * 0.5, HY + 0.78, 0.46, 0.17, { size: 4.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 }); });
    txt(s, "選ばれる6つの理由", HX, HY + 1.07, HW, 0.2, { size: 6.5, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    ["user", "list", "coins", "pray", "shield", "award"].forEach((k, i) => icon(s, k, HX + 0.14 + i * 0.35, HY + 1.33, 0.26, NAVY));
    txt(s, "スマホでも見やすく", HX, HY + 1.6, HW, 0.18, { size: 6, color: GRAY, align: "center", valign: "middle", margin: 0 });
    // ポイント
    txt(s, "反響が取れるHPのポイント", AX + 2.5, HY - 0.02, 1.95, 0.26, { size: 9.5, bold: true, color: NAVY, valign: "middle", margin: 0 });
    ["サービス・料金が明確", "事例・お客様の声で信頼", "資格・実績をしっかり掲載", "電話・LINE・フォームの導線", "スマホで見やすい構成"].forEach((t, i) => {
      icon(s, "ok", AX + 2.5, HY + 0.32 + i * 0.3, 0.2, NAVY);
      txt(s, t, AX + 2.75, HY + 0.27 + i * 0.3, 1.75, 0.3, { size: 8.5, bold: true, valign: "middle", margin: 0 });
    });
    // お客様の不安
    rect(s, AX + 0.12, PY + 2.75, AW - 0.24, 0.24, NAVY, { round: true, r: 0.1 });
    txt(s, "HPで解消するお客様の6つの不安", AX + 0.12, PY + 2.75, AW - 0.24, 0.24, { size: 8.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    ["費用はいくら？", "どこに頼めばいい？", "信頼できる会社？", "作業の流れは？", "追加料金は？", "作業員は慣れている？"].forEach((t, i) => {
      const x = AX + 0.12 + (i % 3) * 1.45, y = PY + 3.05 + Math.floor(i / 3) * 0.38;
      rect(s, x, y, 1.39, 0.32, LBLUE, { round: true, r: 0.08 });
      txt(s, t, x, y, 1.39, 0.32, { size: 8.5, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    });
    // 矢印
    s.addShape(pres.shapes.RIGHT_ARROW, { x: 4.88, y: PY + 1.6, w: 0.24, h: 0.5, fill: { color: NAVY }, line: { type: "none" }, objectName: "arr_web" });
    // ---- STEP2 ----
    const BX = 5.15, BW = 4.55;
    rect(s, BX, PY, BW, PH, WHITE, { line: GREEN, lw: 1.25, round: true, r: 0.04 });
    rect(s, BX, PY, BW, 0.5, GREEN, { round: true, r: 0.04 });
    num(s, 2, BX + 0.12, PY + 0.07, 0.36, WHITE, GREEN);
    txt(s, [B("流入数を増やす", { fontSize: 13.5 }), B("（SEO・MEO・PPC広告）", { fontSize: 10 })], BX + 0.55, PY, BW - 0.6, 0.5, { color: WHITE, valign: "middle", margin: 0 });
    txt(s, "＝ さまざまな集客チャネルからHPへのアクセスを増やす", BX, PY + 0.52, BW, 0.27, { size: 9, bold: true, color: GREEN, align: "center", valign: "middle", margin: 0 });
    const CH = [
      ["SEO対策", "検索エンジンからの\n自然流入を増やす", "3〜6ヶ月", ["お役立ち記事で\n検索上位を狙う", "資産として長く\n集客できる"], PURPLE, "seo"],
      ["MEO対策", "Googleマップでの\n上位表示で地域集客", "1〜3ヶ月", ["Googleビジネス\nプロフィールを最適化", "口コミを増やし\n評価を上げる"], GREEN, "meo"],
      ["PPC広告", "今すぐ依頼したい人に\nすぐアプローチ", "すぐ", ["地域・キーワードを\n絞って配信", "短期間で反響を\n獲得できる"], ORANGE, "ppc"],
    ];
    CH.forEach(([n, role, time, pts, col, kind], i) => {
      const x = BX + 0.1 + i * 1.47, y = PY + 0.83, w = 1.41;
      rect(s, x, y, w, 2.97, WHITE, { line: col, lw: 1.25, round: true, r: 0.05 });
      txt(s, n, x, y + 0.03, w, 0.3, { size: 12, bold: true, color: col, align: "center", valign: "middle", margin: 0 });
      txt(s, role, x, y + 0.32, w, 0.36, { size: 7, bold: true, align: "center", valign: "middle", margin: 0, ls: 1.0 });
      // 画面イメージ
      const mx = x + 0.08, my = y + 0.72, mw = w - 0.16, mh = 1.0;
      rect(s, mx, my, mw, mh, "FAFBFD", { line: "C9D1DE" });
      if (kind === "seo" || kind === "ppc") {
        rect(s, mx + 0.06, my + 0.06, mw - 0.12, 0.17, WHITE, { line: "C9D1DE", round: true, r: 0.08 });
        icon(s, "search", mx + 0.1, my + 0.09, 0.11, MGRAY);
        txt(s, kind === "seo" ? "遺品整理 費用相場" : "遺品整理 ○○市", mx + 0.24, my + 0.06, mw - 0.3, 0.17, { size: 5.5, color: DARK, valign: "middle", margin: 0 });
      }
      if (kind === "seo") {
        ["遺品整理の費用相場とは？", "実家の片付けは何から？", "遺品整理と不用品回収の違い"].forEach((t, k) => {
          txt(s, t, mx + 0.06, my + 0.28 + k * 0.23, mw - 0.1, 0.13, { size: 5.5, bold: true, color: BLUE, valign: "middle", margin: 0 });
          rect(s, mx + 0.06, my + 0.42 + k * 0.23, mw * 0.7, 0.03, "D5DAE2");
        });
      } else if (kind === "meo") {
        rect(s, mx, my, mw, 0.5, "E3ECD9");
        line(s, mx, my + 0.18, mw, 0.15, "FFFFFF", 2); line(s, mx + 0.4, my, 0.25, 0.5, "FFFFFF", 2);
        icon(s, "map", mx + 0.42, my + 0.1, 0.22, RED); icon(s, "map", mx + 0.75, my + 0.06, 0.22, RED);
        txt(s, "○○遺品整理センター", mx + 0.05, my + 0.53, mw - 0.1, 0.15, { size: 5.5, bold: true, color: NAVY, valign: "middle", margin: 0 });
        txt(s, [B("4.9 ", { color: DARK }), B("★★★★★", { color: "F2B705" }), B(" (128)", { color: GRAY })], mx + 0.05, my + 0.68, mw - 0.1, 0.13, { size: 5.5, valign: "middle", margin: 0 });
        txt(s, "営業時間 8:00〜20:00", mx + 0.05, my + 0.82, mw - 0.1, 0.13, { size: 5, color: GRAY, valign: "middle", margin: 0 });
      } else {
        txt(s, "広告", mx + 0.06, my + 0.28, 0.3, 0.13, { size: 5, bold: true, color: DARK, valign: "middle", margin: 0 });
        txt(s, "遺品整理｜即日対応・見積無料", mx + 0.06, my + 0.42, mw - 0.1, 0.16, { size: 5.5, bold: true, color: BLUE, valign: "middle", margin: 0 });
        txt(s, "遺品整理の専門業者。仕分け・供養・買取までワンストップで安心対応。", mx + 0.06, my + 0.58, mw - 0.1, 0.24, { size: 4.8, color: DARK, valign: "top", margin: 0, ls: 1.0 });
        ["即日対応", "見積無料", "追加料金なし"].forEach((t, k) => { rect(s, mx + 0.06 + k * 0.4, my + 0.83, 0.37, 0.12, LBLUE, { round: true, r: 0.04 }); txt(s, t, mx + 0.06 + k * 0.4, my + 0.83, 0.37, 0.12, { size: 4, bold: true, color: BLUE, align: "center", valign: "middle", margin: 0 }); });
      }
      // 効果が出るまで
      rect(s, x + 0.1, y + 1.78, w - 0.2, 0.26, LGRAY, { round: true, r: 0.1 });
      txt(s, [B("効果まで ", { fontSize: 6.5, color: GRAY }), B(time, { fontSize: 10, color: col })], x + 0.1, y + 1.78, w - 0.2, 0.26, { align: "center", valign: "middle", margin: 0 });
      pts.forEach((p, j) => txt(s, "・" + p, x + 0.08, y + 2.1 + j * 0.42, w - 0.12, 0.42, { size: 7.5, valign: "top", margin: 0, ls: 1.05 }));
    });
    // ---- 下：集客の流れとゴール ----
    rect(s, 0.3, 5.6, 1.15, 0.72, NAVY, { round: true, r: 0.06 });
    txt(s, "集客の流れ\n（2ステップ）", 0.3, 5.6, 1.15, 0.72, { size: 9.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, ls: 1.05 });
    rect(s, 1.55, 5.6, 2.6, 0.72, WHITE, { line: NAVY, round: true, r: 0.06 });
    icon(s, "laptop", 1.65, 5.76, 0.4, NAVY);
    txt(s, [B("① 専門HPを構築\n", { fontSize: 10, color: NAVY }), B("問い合わせの「受け皿」を整える", { fontSize: 7.5, color: DARK })], 2.12, 5.6, 2.0, 0.72, { valign: "middle", margin: 0, ls: 1.05 });
    arrowR(s, 4.2, 5.82, 0.3, 0.28, NAVY);
    rect(s, 4.55, 5.6, 2.6, 0.72, WHITE, { line: GREEN, round: true, r: 0.06 });
    icon(s, "search", 4.65, 5.76, 0.4, GREEN);
    txt(s, [B("② 流入数を増やす\n", { fontSize: 10, color: GREEN }), B("SEO・MEO・PPCで反響を獲得", { fontSize: 7.5, color: DARK })], 5.12, 5.6, 2.0, 0.72, { valign: "middle", margin: 0, ls: 1.05 });
    arrowR(s, 7.2, 5.82, 0.3, 0.28, NAVY);
    rect(s, 7.55, 5.6, 2.15, 0.72, NAVY, { round: true, r: 0.06 });
    icon(s, "chart", 7.62, 5.78, 0.36, YEL);
    txt(s, [B("最終的なゴール\n", { fontSize: 8, color: YEL }), B("問い合わせ増加→\n成約・売上の拡大", { fontSize: 8, color: WHITE })], 8.03, 5.6, 1.65, 0.72, { valign: "middle", margin: 0, ls: 1.0 });
    key(s, "この2ステップを正しく実行し、費用対効果の高い集客と安定した売上アップを実現します。", { y: 6.42, size: 13.5 });
  }
  {
    const s = content(S7, "専門HPとチラシのつくり方");
    msg(s, "一丁目一番地は「専門店の屋号とHP」。チラシは戸建に絞って打ち出します", { size: 14 });
    frame(s, 0.4, 1.6, 4.5, 4.75, "反響がとれる専門HP", BLUE, { hh: 0.5, hsize: 14 });
    rect(s, 1.85, 2.25, 1.6, 0.45, BLUE);
    txt(s, "トップページ", 1.85, 2.25, 1.6, 0.45, { size: 10, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    line(s, 2.65, 2.7, 0, 0.2, BLUE, 1.25);
    line(s, 0.94, 2.9, 3.33, 0, BLUE, 1.25);
    ["料金", "作業の流れ", "作業事例", "お客様の声", "選ばれる理由"].forEach((t, i) => {
      const x = 0.55 + i * 0.83;
      line(s, x + 0.39, 2.9, 0, 0.15, BLUE, 1.25);
      rect(s, x, 3.05, 0.78, 0.45, LBLUE);
      txt(s, t, x, 3.05, 0.78, 0.45, { size: 8, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    });
    txt(s, "＋ ブログ／会社概要／よくある質問／無料相談 の全10ページ", 0.55, 3.55, 4.2, 0.3, { size: 9, color: GRAY, margin: 0 });
    ["8秒以内に「何屋か」「信頼できるか」を伝える", "料金はパック料金でわかりやすく", "事例・お客様の声・資格で信頼を強化", "電話・LINE・フォームの導線を全ページに"].forEach((t, i) => {
      icon(s, "ok", 0.6, 4.02 + i * 0.55, 0.28, BLUE);
      txt(s, t, 0.98, 3.95 + i * 0.55, 3.8, 0.42, { size: 10.5, valign: "middle", margin: 0 });
    });
    frame(s, 5.1, 1.6, 4.5, 4.75, "戸建向けのチラシ", ORANGE, { hh: 0.5, hsize: 14 });
    [["B4サイズ", "埋もれない大きさ"], ["厚手の紙", "保存してもらえる"], ["戸建のみ", "単価の高い層に"]].forEach((t, i) => {
      const x = 5.25 + i * 1.45;
      rect(s, x, 2.25, 1.35, 0.95, LORANGE, { round: true });
      txt(s, [B(t[0] + "\n", { fontSize: 12, color: ORANGE }), B(t[1], { fontSize: 9 })], x, 2.25, 1.35, 0.95, { align: "center", valign: "middle" });
    });
    txt(s, "紙面に入れる5つのポイント", 5.25, 3.35, 4.2, 0.35, { size: 11, bold: true, color: ORANGE, margin: 0 });
    ["「遺品整理」を大きく打ち出す", "誰のための・何を解決するサービスかを明確に", "「保存版」を強調する", "料金プランと追加料金の考え方を明記", "実績・口コミ・資格・スタッフの顔を掲載"].forEach((t, i) => {
      num(s, i + 1, 5.3, 3.8 + i * 0.48, 0.3, ORANGE);
      txt(s, t, 5.7, 3.75 + i * 0.48, 3.8, 0.4, { size: 10.5, valign: "middle", margin: 0 });
    });
    key(s, "ターゲットは明確。その対象に対して、専門店として打ち出します。", { y: 6.45, size: 16 });
  }
  // ---- アライアンス提携 ①：なぜアライアンスか ----
  {
    const s = content(S7, "アライアンス提携", "【話し方】①遺品整理のご相談は、最初に不動産会社・司法書士・ケアマネなど「身近な専門家」に入ることが多い→②その専門家とつながり、貴社が地域の相続・空き家案件の受け皿（案件ハブ）になる→③WEB・チラシと比べ、獲得コストは成約時のみ・相見積もりのない特命受注→④目安は紹介元15社で月5〜8件。単価30万円・契約率80%なら年間約1,400〜2,300万円の売上機会。\n【数値根拠】一番店マニュアル集客編：WEB獲得単価3万円／チラシ5万円／アライアンス紹介元15社につき月5〜8件。");
    msg(s, "相続・空き家の相談を最初に受ける専門家とつながり、広告に頼らない受注の柱をつくります");
    // 左：ネットワーク図
    txt(s, "地域の相続・空き家案件が、貴社に集まる", 0.3, 1.7, 4.55, 0.3, { size: 11, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.OVAL, { x: 0.3, y: 2.1, w: 4.55, h: 3.7, fill: { color: "FFF8E8" }, line: { type: "none" }, objectName: "net_bg" });
    const CX = 2.575, CY = 3.95;
    const P = [["building", "不動産会社", 0, -1.45], ["user", "司法書士\n行政書士", -1.5, -0.92], ["pray", "葬儀社", 1.5, -0.92], ["list", "税理士", -1.6, 0.05], ["hammer", "解体会社", 1.6, 0.05], ["hospital", "ケアマネ\n包括支援", -1.5, 1.0], ["wrench", "リフォーム", 1.5, 1.0], ["city", "役所", 0, 1.48]];
    P.forEach(([k, t, dx, dy]) => {
      const len = Math.hypot(dx, dy), ux = dx / len, uy = dy / len;
      const x1 = CX + ux * 0.62, y1 = CY + uy * 0.62, x2 = CX + dx - ux * (Math.abs(dy) > 1.2 ? 0.25 : 0.6), y2 = CY + dy - uy * 0.27;
      s.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1) || 0.001, h: Math.abs(y2 - y1) || 0.001, flipV: (x2 - x1) * (y2 - y1) < 0, line: { color: ORANGE, width: 1.5, endArrowType: "triangle" }, flipH: false, objectName: nm("net") });
      const bw = 1.12, bh = 0.46;
      rect(s, CX + dx - bw / 2, CY + dy - bh / 2, bw, bh, WHITE, { line: LINE, shadow: true, round: true, r: 0.08 });
      icon(s, k, CX + dx - bw / 2 + 0.06, CY + dy - 0.13, 0.26, NAVY);
      txt(s, t, CX + dx - bw / 2 + 0.34, CY + dy - bh / 2, bw - 0.36, bh, { size: 8, bold: true, valign: "middle", margin: 0, ls: 1.0 });
    });
    oval(s, CX - 0.6, CY - 0.6, 1.2, 1.2, NAVY);
    icon(s, "home", CX - 0.17, CY - 0.47, 0.34, WHITE);
    txt(s, "貴社", CX - 0.6, CY - 0.08, 1.2, 0.45, { size: 14, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, "※提携先は貴社の本業・地域に合わせて選定", 0.3, 5.82, 4.55, 0.25, { size: 8, color: GRAY, align: "center", margin: 0 });
    // 右：集客手法の比較
    txt(s, "他の集客手法との違い", 5.1, 1.7, 4.6, 0.3, { size: 11, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    const A = (t) => ({ text: t, options: { bold: true, color: RED, fill: { color: LORANGE } } });
    table(s, [
      ["", "WEB", "チラシ", { text: "アライアンス", options: { fill: { color: RED } } }],
      ["獲得コスト", "約3万円/件", "約5万円/件", A("成約時のみ")],
      ["案件の入り方", "毎月出稿で\n安定", "需要月に\n集中", A("毎月安定")],
      ["案件の特徴", "相見積もり\nが多い", "戸建・空き家\nが多い", A("競合のいない\n特命受注")],
    ], 5.1, 2.05, 4.6, [1.0, 1.15, 1.15, 1.3], { size: 9, rowH: [0.34, 0.42, 0.55, 0.55] });
    // 数値の目安
    rect(s, 5.1, 4.1, 4.6, 1.95, WHITE, { line: NAVY, lw: 1.25, round: true, r: 0.05 });
    txt(s, "紹介の目安", 5.1, 4.15, 4.6, 0.28, { size: 10, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    txt(s, [B("紹介元 ", { fontSize: 11 }), B("15", { fontSize: 26, color: NAVY }), B("社", { fontSize: 12 })], 5.15, 4.45, 1.9, 0.6, { align: "center", valign: "middle", margin: 0 });
    arrowR(s, 7.05, 4.6, 0.35, 0.3, ORANGE);
    txt(s, [B("毎月 ", { fontSize: 11 }), B("5〜8", { fontSize: 26, color: RED }), B("件", { fontSize: 12 })], 7.4, 4.45, 2.25, 0.6, { align: "center", valign: "middle", margin: 0 });
    line(s, 5.35, 5.15, 4.1, 0, LINE, 0.75, { dash: "dash" });
    txt(s, [B("年間 約1,400〜2,300万円", { fontSize: 15, color: RED }), B(" の売上機会", { fontSize: 11 })], 5.1, 5.2, 4.6, 0.45, { align: "center", valign: "middle", margin: 0 });
    txt(s, "※平均単価30万円・契約率80%で試算", 5.1, 5.65, 4.6, 0.3, { size: 8, color: GRAY, align: "center", valign: "middle", margin: 0 });
    key(s, "広告費をかけずに、高単価の案件が毎月入ってくる状態をつくります。", { y: 6.25, size: 15 });
  }
  // ---- アライアンス提携 ②：貴社での進め方 ----
  {
    const s = content(S7, "アライアンス構築の進め方", "【話し方】①まず提携先を選び、ご挨拶に伺う→②最初の紹介案件で信頼をつくってから提携契約→③作業報告・定期訪問で紹介を定着させる→④目安は紹介元15社で毎月5〜8件→⑤提携のカギは「紹介元の信用を守る」こと。紹介元が一番恐れるのは、変な業者を紹介して自社の信用を失うこと。\n【数値根拠】一番店マニュアル集客編：紹介元15社につき月5〜8件の紹介。\n【補足】マニュアルでは初回訪問から約6ヶ月後に提携化（1週間後お礼→1ヶ月後再訪→…→5ヶ月後試験紹介→6ヶ月後以降提携契約）。");
    msg(s, "提携先を選び、信頼を積み重ねることで、紹介が毎月入ってくる状態をつくります");
    const PH = [
      ["STEP 1", "提携先の選定", "map", "地域の不動産会社・士業・\n葬儀社・ケアマネなどを\n洗い出し、優先順位をつける"],
      ["STEP 2", "ご挨拶・初回訪問", "shake", "提携資料を持って\nご挨拶に伺い、\n困りごとをお聞きする"],
      ["STEP 3", "提携契約", "contract", "最初の紹介案件で\n信頼をつくり、\n条件をすり合わせる"],
      ["STEP 4", "紹介の定着", "recycle", "作業報告・定期訪問で\n関係を続け、\n紹介を増やしていく"],
    ];
    const X0 = 0.3, CW = 2.35;
    PH.forEach(([m, t, k, a], i) => {
      const x = X0 + i * CW, col = i === 3 ? RED : [NAVY, NAVY2, BLUE][i];
      chevron(s, x, 1.72, CW + 0.05, 0.5, col, "", { first: i === 0 });
      txt(s, [B(m + "\n", { fontSize: 8, color: WHITE }), B(t, { fontSize: 11, color: WHITE })], x + (i === 0 ? 0.05 : 0.18), 1.72, CW - 0.3, 0.5, { align: "center", valign: "middle", margin: 0, ls: 1.0 });
      rect(s, x + 0.05, 2.35, CW - 0.12, 1.75, WHITE, { line: LINE });
      icon(s, k, x + CW / 2 - 0.3, 2.47, 0.6, WHITE, col);
      txt(s, a, x + 0.12, 3.15, CW - 0.26, 0.9, { size: 9.5, align: "center", valign: "middle", margin: 0, ls: 1.2 });
    });
    // 目安
    txt(s, [B("目安　", { fontSize: 11, color: NAVY }), B("紹介元 ", { fontSize: 11 }), B("15社", { fontSize: 18, color: RED }), B(" で、紹介は ", { fontSize: 11 }), B("毎月5〜8件", { fontSize: 18, color: RED })], 0.3, 4.3, 9.4, 0.5, { align: "center", valign: "middle", margin: 0 });
    // カギ
    rect(s, 0.3, 5.0, 9.4, 1.75, LGRAY, { round: true, r: 0.05 });
    icon(s, "shield", 0.55, 5.45, 0.8, WHITE, NAVY);
    txt(s, [B("提携のカギは「紹介元の信用を守る」こと\n", { fontSize: 12.5, color: NAVY }), B("紹介元が一番恐れているのは、紹介した業者のせいで\n自社の信用を失うこと。「この会社なら安心して任せられる」\nと思っていただくことが、紹介の継続につながります。", { fontSize: 9, bold: false })], 1.5, 5.1, 4.5, 1.55, { valign: "middle", margin: 0, ls: 1.25 });
    [["clock", "迅速なレスポンス"], ["yen", "明朗な見積り"], ["check", "丁寧な作業と事後報告"]].forEach(([k, t], i) => {
      const y = 5.2 + i * 0.5;
      rect(s, 6.15, y, 3.4, 0.4, WHITE, { line: LINE, round: true, r: 0.1 });
      icon(s, k, 6.25, y + 0.07, 0.26, GREEN);
      txt(s, t, 6.6, y, 2.9, 0.4, { size: 10, bold: true, valign: "middle", margin: 0 });
    });
  }
  // ---- 営業 ①：営業KPI ----
  {
    const s = content(S7, "営業KPI", "【話し方】①売上は「集客件数×契約率×単価」で決まる→②契約率80%を基準に、1年目は営業1名・月10件契約からスタート→③3年目は3名体制・月28件で売上1億円→④一人当たりの契約件数を落とさないことが利益のカギ。\n【根拠】3年1億円モデル（単価30万円・契約率80%）。");
    msg(s, "契約率80%を基準に、少人数の営業体制で3年1億円をつくります");
    const R = (t) => ({ text: t, options: { bold: true, color: RED } });
    table(s, [
      ["", "1年目", "2年目", "3年目"],
      ["営業人員", "1名", "2名", "3名"],
      ["月間集客件数", "12.5件", "18.75件", "35件"],
      ["月間契約件数", "10件", "15件", "28件"],
      ["一人当たり契約件数（月）", "10件", "7.5件", "約9.3件"],
      ["単価", "30万円", "30万円", "30万円"],
      ["売上", R("3,600万円"), R("5,400万円"), R("1億円")],
    ], 0.3, 1.75, 5.75, [2.15, 1.2, 1.2, 1.2], { size: 10.5, rowH: 0.5, firstColFill: true });
    txt(s, "※契約率は全期間80%を基準として算出", 0.3, 5.3, 5.75, 0.28, { size: 8.5, color: GRAY, margin: 0 });
    // 右：売上の分解（3年目）
    rect(s, 6.3, 1.75, 3.4, 3.83, LGRAY, { round: true, r: 0.05 });
    txt(s, "売上の考え方（3年目）", 6.3, 1.82, 3.4, 0.32, { size: 10.5, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    const F = [["月間集客件数", "35件", NAVY], ["契約率", "80%", RED], ["月間契約件数", "28件", NAVY], ["単価", "30万円", NAVY]];
    const OP = ["×", "＝", "×"];
    F.forEach(([l, v, c], i) => {
      const y = 2.22 + i * 0.68;
      rect(s, 6.55, y, 2.9, 0.5, WHITE, { line: i === 1 ? RED : LINE, lw: i === 1 ? 1.5 : 1, round: true, r: 0.1 });
      txt(s, l, 6.65, y, 1.5, 0.5, { size: 9.5, bold: true, valign: "middle", margin: 0 });
      txt(s, v, 8.0, y, 1.35, 0.5, { size: 15, bold: true, color: c, align: "right", valign: "middle", margin: 0 });
      if (i < 3) txt(s, OP[i], 6.55, y + 0.47, 2.9, 0.22, { size: 11, bold: true, color: GRAY, align: "center", valign: "middle", margin: 0 });
    });
    arrowD(s, 7.8, 4.92, 0.4, 0.2, RED);
    txt(s, [B("年間売上 ", { fontSize: 11 }), B("約1億円", { fontSize: 17, color: RED })], 6.3, 5.12, 3.4, 0.4, { align: "center", valign: "middle", margin: 0 });
    // ポイント
    [["target", "契約率80%を基準に設定"], ["chart", "売上は3年で約2.8倍"], ["users", "一人当たり生産性を最大化"]].forEach(([k, t], i) => {
      const x = 0.3 + i * 3.17;
      rect(s, x, 5.78, 3.0, 0.5, WHITE, { line: NAVY, round: true, r: 0.12 });
      icon(s, k, x + 0.12, 5.88, 0.3, NAVY);
      txt(s, t, x + 0.5, 5.78, 2.45, 0.5, { size: 10.5, bold: true, color: NAVY, valign: "middle", margin: 0 });
    });
    key(s, "成果を左右するのは「契約率」。だからこそ、営業の型づくりが重要です。", { y: 6.42, size: 14 });
  }
  // ---- 営業 ②：全体フロー ----
  {
    const s = content(S7, "遺品整理営業の全体フロー", "【話し方】①流れは問い合わせ→現地確認→見積提出・契約→作業実施→②30万円未満はオンライン中心で効率化。訪問を省いたコストを価格に反映し、即日見積で契約率を上げる→③30万円以上は現地訪問で信頼と提案力を高め、追加提案で単価アップ→④共通の3つのポイント：スピード対応（5分以内架電）／ヒアリング力／信頼構築と提案力。");
    msg(s, "案件の規模で営業方法を使い分け、少ない人数でも契約率と単価を高めます");
    ["問い合わせ", "現地確認", "見積提出・契約", "作業実施"].forEach((t, i) => {
      chevron(s, 0.3 + i * 2.35, 1.72, 2.35, 0.42, [NAVY, NAVY2, BLUE, BLUE][i], t, { first: i === 0, size: 11 });
    });
    const COL = [
      ["30万円未満の案件", "オンライン中心で効率化", ORANGE, LORANGE, ["0次接客（電話ヒアリング）", "アンケート送付（SMS・LINE）", "写真のご提出（お客様）", "オンラインで見積提出", "電話でクロージング・契約"], ["訪問を省き、コストを価格に反映", "即日見積・即対応で契約率アップ"]],
      ["30万円以上の案件", "現地訪問で信頼と提案力を強化", BLUE, LBLUE, ["0次接客（電話ヒアリング）", "事前情報収集・日程調整", "現地確認（訪問）", "見積提示・追加提案", "その場でクロージング・契約"], ["正確な見積りで安心感を提供", "追加提案で単価アップ"]],
    ];
    COL.forEach(([h, sub, c, bg, steps, merit], i) => {
      const x = 0.3 + i * 4.75;
      rect(s, x, 2.32, 4.65, 3.3, WHITE, { line: c, lw: 1.25 });
      rect(s, x, 2.32, 4.65, 0.4, c);
      txt(s, h, x, 2.32, 4.65, 0.4, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      txt(s, sub, x, 2.76, 4.65, 0.3, { size: 10.5, bold: true, color: c, align: "center", valign: "middle", margin: 0 });
      steps.forEach((t, k) => {
        const y = 3.12 + k * 0.36;
        rect(s, x + 0.2, y, 4.25, 0.3, bg);
        txt(s, [B(`STEP${k + 1}　`, { color: c }), B(t)], x + 0.3, y, 4.1, 0.3, { size: 9.5, valign: "middle", margin: 0 });
      });
      merit.forEach((t, k) => {
        icon(s, "ok", x + 0.25, 4.98 + k * 0.28, 0.2, c);
        txt(s, t, x + 0.52, 4.94 + k * 0.28, 4.0, 0.28, { size: 9.5, valign: "middle", margin: 0 });
      });
    });
    // 3つのポイント
    txt(s, "共通する3つのポイント", 0.3, 5.72, 9.4, 0.28, { size: 10.5, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    [["clock", "スピード対応", "問い合わせ後5分以内に架電"], ["talk", "ヒアリング力", "状況・課題・希望を正確に把握"], ["heart", "信頼構築と提案力", "ご遺族に寄り添い、課題を解決"]].forEach(([k, t, d], i) => {
      const x = 0.3 + i * 3.17;
      rect(s, x, 6.05, 3.0, 0.75, WHITE, { line: LINE, round: true, r: 0.08 });
      icon(s, k, x + 0.12, 6.17, 0.5, WHITE, NAVY);
      txt(s, [B(t + "\n", { fontSize: 11, color: NAVY }), B(d, { fontSize: 8.5, bold: false })], x + 0.72, 6.05, 2.25, 0.75, { valign: "middle", margin: 0, ls: 1.15 });
    });
  }
  // ---- 営業 ③：0次接客とクロージング ----
  {
    const s = content(S7, "0次接客とクロージング", "【話し方】①お客様は最初の電話の印象で「この会社に任せたいか」を判断している→②電話では要望の把握・単価アップの情報収集・信頼の獲得の3つを行う（確認項目：故人様との関係・キーマン・希望時期・予算・他社状況／間取り・階数・駐車場・物量／捜索物・買取・供養／整理後の売却・解体・リフォーム予定）→③見積り説明では、よくある不安に先回りして答える→④キーフレーズ：共感「そのお気持ち、よく分かります」／安心「追加料金は一切ありません」／希少性「ご希望の日程は埋まりやすいので、早めのご決断がおすすめです」。\n【トーク例】他社と比較したい→「ご希望の日程が埋まってしまう可能性がございますので、まずは日程だけでも仮押さえさせていただきますと安心です」");
    msg(s, "契約率は、最初の電話と見積りのご説明でほぼ決まります");
    // 左：0次接客
    frame(s, 0.3, 1.72, 4.55, 4.5, "0次接客（最初の電話）", NAVY, { hsize: 12 });
    txt(s, [B("お客様は"), B("電話の印象", { color: RED }), B("で\n「この会社に任せたいか」を判断しています")], 0.4, 2.2, 4.35, 0.6, { size: 10.5, align: "center", valign: "middle", margin: 0, ls: 1.15 });
    [["信頼\nできる？"], ["高く\nない？"], ["対応は\n丁寧？"]].forEach(([t], i) => {
      bubble(s, t, 0.55 + i * 1.4, 2.9, 1.2, 0.75, { size: 10, color: ORANGE, line: ORANGE });
    });
    txt(s, "電話で行う3つのこと", 0.4, 3.9, 4.35, 0.3, { size: 10.5, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    [["要望を正確に把握", "最適なご提案につなげる"], ["単価アップの情報収集", "買取・供養・解体などの材料"], ["信頼の獲得", "安心して任せられる会社に"]].forEach(([t, d], i) => {
      const y = 4.27 + i * 0.6;
      num(s, i + 1, 0.5, y + 0.07, 0.36, NAVY);
      txt(s, [B(t + "\n", { fontSize: 10.5 }), B(d, { fontSize: 8.5, bold: false, color: GRAY })], 1.0, y, 3.75, 0.52, { valign: "middle", margin: 0, ls: 1.1 });
    });
    // 右：クロージング
    frame(s, 5.15, 1.72, 4.55, 4.5, "見積り説明〜クロージング", RED, { hsize: 12 });
    ["見積の説明", "安心のご提供", "日程のご提案", "契約"].forEach((t, i) => {
      chevron(s, 5.25 + i * 1.08, 2.25, 1.1, 0.4, i === 3 ? RED : MGRAY, t, { first: i === 0, size: 8.5 });
    });
    txt(s, "よくある不安に、先回りしてお答えします", 5.25, 2.8, 4.35, 0.3, { size: 10.5, bold: true, color: RED, align: "center", valign: "middle", margin: 0 });
    [["他社と比較したい", "まずは日程だけ仮押さえ"], ["金額が高いかも", "買取・供養の活用で費用を抑える"], ["追加料金が不安", "コミコミ価格で追加料金なし"], ["スタッフが不安", "自社スタッフ・マナー研修を徹底"]].forEach(([q, a], i) => {
      const y = 3.2 + i * 0.68;
      rect(s, 5.3, y, 1.55, 0.55, "FCE8E6", { round: true, r: 0.1 });
      txt(s, q, 5.3, y, 1.55, 0.55, { size: 9, bold: true, color: RED, align: "center", valign: "middle", margin: 0 });
      arrowR(s, 6.92, y + 0.17, 0.25, 0.22, MGRAY);
      txt(s, a, 7.25, y, 2.35, 0.55, { size: 9.5, bold: true, valign: "middle", margin: 0 });
      if (i < 3) line(s, 5.3, y + 0.62, 4.3, 0, LINE, 0.5, { dash: "dash" });
    });
    key(s, "電話で信頼をつくり、見積りで不安を消す。これが契約率80%の土台です。", { y: 6.35, size: 14 });
  }
  // ---- 見積作成 ①：段ボールのサイズ ----
  {
    const s = content(S7, "見積作成について", "【話し方】①見積は「段ボール何箱分か」で物量を数える→②基準は140サイズ段ボール。内寸で計算すると約0.09㎥、ほぼ0.1㎥→③つまり140サイズ10箱で1立米（リューベ）→④1立米あたり12,000〜15,000円が目安。紹介案件などは7,000〜10,000円。");
    s._pageTitle = "段ボール換算の基準";
    msg(s, "物量は「140サイズ段ボール何箱分か」で数え、立米単価で見積ります");
    const BX = [["140", "605", "435", "345", "1,070mm × 770mm × 230mm"], ["120", "475", "335", "310", "840mm × 680mm × 180mm"], ["100", "385", "275", "290", "680mm × 500mm × 160mm"]];
    BX.forEach(([sz, l, w, d, ok], i) => {
      const y = 1.72 + i * 1.6;
      rect(s, 0.3, y, 1.5, 1.38, LBROWN, { round: true, r: 0.06 });
      icon(s, "box", 0.62, y + 0.25, 0.85, BROWN);
      rect(s, 2.0, y + 0.06, 0.2, 0.2, BROWN);
      txt(s, `${sz}サイズダンボール箱`, 2.28, y, 3.8, 0.32, { size: 14, bold: true, color: BROWN, valign: "middle", margin: 0 });
      rect(s, 2.0, y + 0.42, 0.55, 0.24, BROWN, { round: true, r: 0.05 });
      txt(s, "内寸", 2.0, y + 0.42, 0.55, 0.24, { size: 9, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      [["長さ", l], ["幅", w], ["深さ", d]].forEach(([lb, v], k) => {
        const x = 2.65 + k * 1.15;
        txt(s, [B(lb + "\n", { fontSize: 9 }), B(v, { fontSize: 14 }), B("mm", { fontSize: 9 })], x, y + 0.38, 1.05, 0.6, { align: "center", valign: "middle", margin: 0, ls: 1.0 });
        if (k < 2) line(s, x + 1.1, y + 0.42, 0, 0.5, BROWN, 0.75, { dash: "sysDot" });
      });
      rect(s, 2.0, y + 1.06, 1.0, 0.26, BROWN, { round: true, r: 0.05 });
      txt(s, "送り荷寸法", 2.0, y + 1.06, 1.0, 0.26, { size: 9, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      txt(s, ok, 3.1, y + 1.06, 3.0, 0.26, { size: 10, bold: true, valign: "middle", margin: 0 });
      if (i < 2) line(s, 0.3, y + 1.48, 5.85, 0, BROWN, 1, { dash: "sysDot" });
    });
    // 右：計算
    rect(s, 6.35, 1.72, 3.35, 4.6, LGRAY, { round: true, r: 0.05 });
    txt(s, "140サイズ段ボール", 6.5, 1.85, 3.1, 0.38, { size: 14, bold: true, color: NAVY, valign: "middle", margin: 0 });
    txt(s, "＝ 0.605m × 0.435m × 0.345m\n≒ 約0.09㎥\n≒ 0.1㎥", 6.5, 2.25, 3.1, 0.95, { size: 11, valign: "middle", margin: 0, ls: 1.25 });
    line(s, 6.5, 3.3, 3.05, 0, LINE, 0.75, { dash: "dash" });
    txt(s, [B("1リューベ（立米）\n", { fontSize: 10, color: GRAY }), B("＝ 140サイズ ", { fontSize: 12 }), B("10箱", { fontSize: 16, color: RED })], 6.5, 3.4, 3.1, 0.75, { valign: "middle", margin: 0, ls: 1.1 });
    line(s, 6.5, 4.25, 3.05, 0, LINE, 0.75, { dash: "dash" });
    txt(s, [B("1リューベ単価\n", { fontSize: 10, color: GRAY }), B("12,000〜15,000円", { fontSize: 16, color: RED })], 6.5, 4.35, 3.1, 0.75, { valign: "middle", margin: 0, ls: 1.1 });
    rect(s, 6.5, 5.25, 3.05, 0.8, WHITE, { line: LINE, round: true, r: 0.08 });
    txt(s, [B("紹介等の場合\n", { fontSize: 9.5, color: GRAY }), B("7,000〜10,000円", { fontSize: 13, color: NAVY })], 6.5, 5.25, 3.05, 0.8, { align: "center", valign: "middle", margin: 0, ls: 1.1 });
    key(s, "140サイズ10箱 ＝ 1立米。これが見積の基本単位です。", { y: 6.45, size: 15 });
  }
  // ---- 見積作成 ②：現地調査・見積シート ----
  {
    const s = content(S7, "見積作成について", "【話し方】①現地調査は専用の見積シートで行う→②小物は1種類1箱、家具・家電は換算表で「何箱分か」に置き換える→③総物量（箱数）に箱単価を掛け、基本料金・特殊処分・オプションを足して、買取金額を差し引く→④誰が調査しても同じ金額になるのがポイント。");
    s._pageTitle = "現地調査・見積シート（段ボール換算方式）";
    msg(s, "段ボール何箱分に該当するかを確認し、箱単価を決めることで見積金額を算出します");
    // 左：シートの構成
    rect(s, 0.3, 1.72, 3.85, 4.15, WHITE, { line: NAVY, lw: 1.25 });
    rect(s, 0.3, 1.72, 3.85, 0.36, NAVY);
    txt(s, "遺品整理 現地調査・見積シート", 0.3, 1.72, 3.85, 0.36, { size: 10.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, "基準：140サイズ段ボール＝1.0箱（約0.1㎥）", 0.4, 2.12, 3.65, 0.24, { size: 8, bold: true, color: RED, align: "right", valign: "middle", margin: 0 });
    const cell = (t, x, y, w, h, c, hl) => {
      rect(s, x, y, w, h, hl ? "FFF8E8" : WHITE, { line: hl ? RED : c, lw: hl ? 1.75 : 0.75 });
      rect(s, x, y, w, 0.24, c);
      txt(s, t, x, y, w, 0.24, { size: 7.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      for (let k = 1; k * 0.16 + 0.24 < h - 0.05; k++) line(s, x + 0.08, y + 0.24 + k * 0.16, w - 0.16, 0, "E1E5EC", 0.5);
    };
    cell("基本情報", 0.4, 2.4, 3.65, 0.5, NAVY);
    cell("① 小物類", 0.4, 2.98, 1.2, 1.15, BLUE);
    cell("② 家具・家電", 1.68, 2.98, 2.37, 1.15, GREEN, true);
    cell("③ 特殊処分品", 0.4, 4.2, 1.2, 0.85, PURPLE);
    cell("④ 買取可能品", 1.68, 4.2, 1.15, 0.85, ORANGE);
    cell("⑤ オプション", 2.9, 4.2, 1.15, 0.85, RED);
    cell("⑥ 作業条件", 0.4, 5.12, 1.2, 0.65, BLUE);
    cell("⑦ 最終集計", 1.68, 5.12, 1.15, 0.65, PURPLE);
    cell("⑧ 見積算出", 2.9, 5.12, 1.15, 0.65, NAVY);
    arrowR(s, 4.18, 3.4, 0.3, 0.3, RED);
    // 右：家具・家電の換算表
    const H = { text: "", options: {} };
    const L1 = [["冷蔵庫（大型）", "", "12"], ["冷蔵庫（小型）", "", "8"], ["洗濯機", "", "6"], ["乾燥機", "", "4"], ["電子レンジ", "", "2"], ["炊飯器", "", "1"], ["テレビ32型以下", "", "2"], ["テレビ50型以上", "", "4"], ["テレビ台", "", "4"], ["ソファ1人掛", "脚", "4"], ["ソファ2人掛", "脚", "8"], ["本棚（大）", "台", "10"], ["机", "台", "5"], ["学習机", "台", "7"], ["椅子", "脚", "1"], ["仏壇（小）", "台", "6"], ["仏壇（大）", "台", "15"]];
    const R1 = [["ソファ3人掛", "脚", "12"], ["シングルベッド", "台", "10"], ["ダブルベッド", "台", "14"], ["マットレスS", "枚", "6"], ["マットレスD", "枚", "8"], ["婚礼タンス", "棹", "15"], ["整理タンス", "棹", "10"], ["チェスト", "台", "6"], ["食器棚（小）", "台", "8"], ["食器棚（大）", "台", "15"], ["本棚（小）", "台", "5"], ["金庫（小）", "台", "8"], ["金庫（大）", "台", "20"], ["その他家具①", "台", ""], ["その他家具②", "台", ""], ["その他家電", "台", ""], ["", "", ""]];
    const nL = (i) => (i < 11 ? i + 1 : i + 12), nR = (i) => (i < 11 ? i + 12 : (i < 16 ? i + 18 : ""));
    const G = { fill: { color: LGREEN } };
    const rows = [[{ text: "家電", options: { fill: { color: GREEN } } }, { text: "", options: { fill: { color: GREEN } } }, { text: "", options: { fill: { color: GREEN } } }, { text: "", options: { fill: { color: GREEN } } }, { text: "家具・その他", options: { fill: { color: GREEN } } }, { text: "", options: { fill: { color: GREEN } } }, { text: "", options: { fill: { color: GREEN } } }, { text: "", options: { fill: { color: GREEN } } }]];
    rows[0][0].options.colspan = 4; rows[0] = [rows[0][0], Object.assign({}, rows[0][4], { options: { fill: { color: GREEN }, colspan: 4 } })];
    rows.push(["No", "品目", "単位", "箱数", "No", "品目", "単位", "箱数"].map((t) => ({ text: t, options: { bold: true, color: GREEN, fill: { color: LGREEN }, align: "center" } })));
    for (let i = 0; i < 17; i++) {
      const a = L1[i], b = R1[i];
      rows.push([String(nL(i)), { text: a[0], options: { align: "left" } }, a[1], { text: a[2] ? a[2] + "箱" : "箱", options: { bold: true } }, String(nR(i)), { text: b[0], options: { align: "left" } }, b[1], { text: b[0] ? (b[2] ? b[2] + "箱" : "箱") : "", options: { bold: true } }]);
    }
    table(s, rows, 4.55, 1.72, 5.15, [0.36, 1.26, 0.4, 0.53, 0.36, 1.26, 0.4, 0.58], { size: 7.5, rowH: 0.2, head: GREEN, margin: [1, 3, 1, 3] });
    // 下：見積算出
    rect(s, 0.3, 5.98, 9.4, 0.42, LBLUE, { round: true, r: 0.1 });
    txt(s, [B("見積金額 ＝ ", { color: NAVY }), B("総物量（箱）× 箱単価", { color: RED }), B(" ＋ 基本料金 ＋ 特殊処分 ＋ オプション − 買取金額", { color: NAVY })], 0.3, 5.98, 9.4, 0.42, { size: 11.5, align: "center", valign: "middle", margin: 0 });
    key(s, "誰が調査しても、同じ金額が出せる仕組みにします。", { y: 6.48, size: 14 });
  }
  // ---- 見積作成 ③：立米・トラック・作業量 ----
  {
    const s = content(S7, "見積作成について", "【話し方】①140ダンボール＝3辺の合計140cm、10箱で約1立米→②2トントラック1台で10〜12立米＝約100〜120箱分→③作業量は1人あたり1.2立米/日が目安→④物量（立米）が分かれば、トラック台数と必要な人数・日数が決まる。例：2LDKで段ボール100箱程度＝約10立米→2トン車1台、約8人日。");
    s._pageTitle = "物量・トラック・作業量の目安";
    msg(s, "物量（立米）が分かれば、トラックの台数と必要な人員が決まります");
    const C = PURPLE, LC = "F2EEF8";
    const col = (i, head, lead, sub) => {
      const x = 0.3 + i * 3.17;
      rect(s, x, 1.72, 3.0, 4.95, WHITE, { line: "D9D0E8", round: true, r: 0.04 });
      rect(s, x + 0.1, 1.8, 2.8, 0.42, C, { round: true, r: 0.1 });
      txt(s, head, x + 0.1, 1.8, 2.8, 0.42, { size: 12, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      txt(s, [B(lead + "\n", { fontSize: 11 }), B(sub, { fontSize: 9, bold: false, color: GRAY })], x + 0.15, 2.3, 2.75, 0.95, { valign: "middle", margin: 0, ls: 1.2 });
      return x;
    };
    // ①
    let x = col(0, "ダンボールのサイズと立米換算", "140ダンボールとは？", "3辺の合計が140cmのダンボールです。\n縦×横×高さは複数のタイプがあります。");
    icon(s, "box", x + 0.2, 3.4, 1.1, BROWN);
    rect(s, x + 1.45, 3.3, 1.45, 1.3, LC, { round: true, r: 0.08 });
    txt(s, [B("3辺の合計\n", { fontSize: 9, color: C }), B("140cm\n", { fontSize: 17, color: C }), B("（例）50×40×50cm\n60×40×40cm など", { fontSize: 7.5, bold: false, color: GRAY })], x + 1.45, 3.3, 1.45, 1.3, { align: "center", valign: "middle", margin: 0, ls: 1.1 });
    txt(s, [B("10箱で約1立米（1㎥）\n", { fontSize: 11 }), B("140ダンボール10箱分の容量が\nおよそ1立米になります。", { fontSize: 9, bold: false, color: GRAY })], x + 0.15, 4.75, 2.75, 0.75, { valign: "middle", margin: 0, ls: 1.2 });
    oval(s, x + 0.6, 5.55, 1.8, 1.05, LC);
    txt(s, [B("140ダンボール10箱\n", { fontSize: 8.5, color: C }), B("＝約1㎥", { fontSize: 16, color: C })], x + 0.6, 5.55, 1.8, 1.05, { align: "center", valign: "middle", margin: 0, ls: 1.05 });
    // ②
    x = col(1, "2トントラックの積載目安", "最大積載量は約10〜12立米", "ダンボールに換算すると、\n約100〜120箱分を積むことができます。");
    icon(s, "truck", x + 0.85, 3.35, 1.3, MGRAY);
    txt(s, "2トントラック（標準サイズ）", x, 4.7, 3.0, 0.25, { size: 8.5, color: GRAY, align: "center", valign: "middle", margin: 0 });
    [["最大積載量", "10〜12", "立米"], ["ダンボール換算", "100〜120", "箱分"]].forEach(([h, v, u], k) => {
      const xx = x + 0.12 + k * 1.4;
      rect(s, xx, 5.05, 1.33, 1.2, LC, { round: true, r: 0.08 });
      rect(s, xx + 0.1, 5.12, 1.13, 0.24, C, { round: true, r: 0.12 });
      txt(s, h, xx + 0.1, 5.12, 1.13, 0.24, { size: 8, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      txt(s, [B(v + "\n", { fontSize: 15, color: C }), B(u, { fontSize: 10, color: C })], xx, 5.4, 1.33, 0.8, { align: "center", valign: "middle", margin: 0, ls: 1.05 });
    });
    txt(s, "※トラックのサイズや荷物の形状により前後します", x + 0.1, 6.3, 2.8, 0.3, { size: 7.5, color: GRAY, align: "center", valign: "middle", margin: 0 });
    // ③
    x = col(2, "作業量の目安（1人あたり）", "1人あたり1.2立米/日が目安", "荷物の種類や現場の状況により、\n多少前後します。");
    icon(s, "people", x + 0.9, 3.35, 1.2, NAVY);
    rect(s, x + 0.15, 4.8, 2.7, 1.45, LC, { round: true, r: 0.08 });
    rect(s, x + 0.45, 4.88, 2.1, 0.26, C, { round: true, r: 0.13 });
    txt(s, "1人あたりの作業量（目安）", x + 0.45, 4.88, 2.1, 0.26, { size: 8.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    txt(s, [B("1.2", { fontSize: 28, color: C }), B(" 立米/日", { fontSize: 13, color: C })], x + 0.15, 5.2, 2.7, 0.65, { align: "center", valign: "middle", margin: 0 });
    txt(s, "※搬出経路・作業環境により変動します", x + 0.15, 5.85, 2.7, 0.3, { size: 7.5, color: GRAY, align: "center", valign: "middle", margin: 0 });
  }
  {
    const s = content(S7, "現場品質の仕組み");
    msg(s, "誰が作業しても「頼んでよかった」と言われる品質を、ルールで標準化します", { size: 14 });
    const R3 = [["list", "捨てる前に必ず確認", "現金・通帳・印鑑・権利書・写真・手紙は必ず確認する"], ["box", "迷ったら捨てない", "判断に迷う物は「確認保留ケース」へ。ご遺族と一緒に確認する"], ["search", "全品手作業で確認", "ポケット・封筒・引き出し・箱の中まで一つずつ確認する"]];
    R3.forEach((r, i) => {
      const x = 0.4 + i * 3.13;
      rect(s, x, 1.6, 2.94, 2.3, WHITE, { line: RED, lw: 1.5 });
      rect(s, x, 1.6, 2.94, 0.38, RED);
      txt(s, "絶対遵守ルール " + (i + 1), x, 1.6, 2.94, 0.38, { size: 10.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      icon(s, r[0], x + 0.15, 2.12, 0.55, RED);
      txt(s, r[1], x + 0.8, 2.08, 2.05, 0.62, { size: 13, bold: true, color: RED, valign: "middle", margin: 0 });
      txt(s, r[2], x + 0.15, 2.85, 2.64, 0.95, { size: 10.5, ls: 1.25 });
    });
    txt(s, "作業の流れ", 0.4, 4.1, 3, 0.35, { size: 12, bold: true, color: NAVY, margin: 0 });
    [["作業前", "近隣挨拶・養生・作業前の写真撮影"], ["作業中", "時計回りの仕分け・貴重品の記録と保管"], ["引き渡し", "全室をご一緒に確認・貴重品をお渡し"], ["完了報告", "写真付きの報告書にご署名"]].forEach((r, i) => {
      chevron(s, 0.4 + i * 2.32, 4.5, 2.45, 0.55, NAVY, r[0], { first: i === 0, size: 11.5 });
      txt(s, r[1], 0.45 + i * 2.32, 5.1, 2.15, 0.7, { size: 9.5, ls: 1.2, align: "center" });
    });
    rect(s, 0.4, 5.85, 9.2, 0.5, LGREEN);
    txt(s, "現場マニュアル（60ページ以上）とチェックリストで、未経験のスタッフでも同じ品質に", 0.4, 5.85, 9.2, 0.5, { size: 11, bold: true, color: GREEN, align: "center", valign: "middle", margin: 0 });
    key(s, "品質の高さが紹介を生み、広告費を下げます。", { y: 6.45, size: 16 });
  }

  // =====================================================================
  // 成功事例
  // =====================================================================
  const S8 = "成功事例";
  pres.addSection({ title: S8 });
  divider(S8, "成功事例のご紹介", "事例企業名・数値は掲載許可を確認してから使う。");
  function caseSlide(name, area, n1, n2, l1, l2, detail, point, notes) {
    const s = content(S8, "成功事例紹介", notes);
    s._pageTitle = name;
    msg(s, `${area}／遺品整理専門店の立上げ事例`);
    photo(s, 0.4, 1.6, 5.2, 3.0, "HPのファーストビュー画像");
    txt(s, "立上げ2年目で", 5.85, 1.65, 3.7, 0.4, { size: 13, bold: true, margin: 0 });
    [[l1, n1], [l2, n2]].forEach((r, i) => {
      rect(s, 5.85, 2.15 + i * 1.2, 3.75, 1.05, "FCE8E6");
      txt(s, r[0], 5.95, 2.18 + i * 1.2, 3.5, 0.35, { size: 11, bold: true, color: GRAY, margin: 0 });
      txt(s, r[1], 5.95, 2.5 + i * 1.2, 3.5, 0.65, { size: 28, bold: true, color: RED, valign: "middle", margin: 0 });
    });
    txt(s, "を実現！", 5.85, 4.55, 3.7, 0.35, { size: 13, bold: true, margin: 0 });
    line(s, 0.4, 5.05, 9.2, 0, LINE, 1);
    txt(s, detail, 0.4, 5.2, 5.6, 1.1, { size: 12, ls: 1.35 });
    rect(s, 6.2, 5.2, 3.4, 1.1, LBLUE, { round: true });
    icon(s, "bulb", 6.3, 5.45, 0.5, NAVY);
    txt(s, point, 6.85, 5.2, 2.7, 1.1, { size: 10.5, bold: true, color: NAVY, valign: "middle" });
  }
  caseSlide("株式会社 蒼炎様", "東京都多摩市", "8,000万円", "3,000万円", "遺品整理事業の売上", "営業利益", "参入2年で初期投資を回収し、不用品買取・空き家買取など事業の裾野を拡大しながら成長中。\n販促費 月50万円／反響 月40件", "多摩エリアに特化した専門店の屋号・HPとWEB集客", "【要確認】掲載許可。HP画像を差し替える。");
  caseSlide("Mevius Moving Works様", "大阪府大阪市", "6,000万円", "30%超", "遺品整理事業の売上", "営業利益率", "不動産会社からの紹介を中心に集客し、成約率80%超を実現。引越し事業と連携し、遺品整理事業を強化。", "提携先からの紹介と、本業（引越し）との連携", "【要確認】掲載許可。HP画像を差し替える。");
  {
    const s = content(S8, "成功企業の共通点", "同業種の事例がある場合は、このページの前に追加する。");
    msg(s, "成功している企業は、共通して5つのことを実践しています");
    const C = [["store", "専門店として打ち出す", "多摩エリア特化の専門店", "遺品整理の専門ブランド", "貴社版の事業コンセプト"], ["shake", "紹介の仕組み", "事業の裾野を広げ紹介を獲得", "不動産会社からの紹介が中心", "提携先からの紹介"], ["clock", "スピード対応", "反響月40件に素早く対応", "成約率80%超", "5分以内の電話・現地即決"], ["link", "他事業との連携", "買取・空き家買取へ拡大", "引越し事業と連携", "基盤となる\n遺品整理事業に参入"], ["yen", "利益が残る事業にする", "営業利益率 約37%\n（3,000万円／8,000万円）", "営業利益率 30%超", "3年目 営業利益率\n20%超の計画"]];
    const XS = [2.95, 5.15, 7.6], CWD = 2.1;
    [["蒼炎様", "立上げ2年目 売上8,000万円"], ["Mevius様", "立上げ2年目 営業利益率30%超"], ["貴社での進め方", "本日ご説明した内容"]].forEach(([h, d], j) => {
      const X = XS[j];
      rect(s, X, 1.72, CWD, 0.62, j === 2 ? RED : NAVY, { round: true, r: 0.08 });
      txt(s, [B(h + "\n", { fontSize: 12, color: WHITE }), B(d, { fontSize: 8, bold: false, color: WHITE })], X, 1.72, CWD, 0.62, { align: "center", valign: "middle", margin: 0, ls: 1.1 });
    });
    C.forEach((c, i) => {
      const y = 2.48 + i * 0.78;
      rect(s, 0.3, y, 2.5, 0.68, LBLUE, { round: true, r: 0.08 });
      icon(s, c[0], 0.42, y + 0.09, 0.5, WHITE, NAVY);
      txt(s, c[1], 1.02, y, 1.75, 0.68, { size: 11, bold: true, color: NAVY, valign: "middle", margin: 0 });
      [c[2], c[3], c[4]].forEach((t, j) => {
        const X = XS[j];
        rect(s, X, y, CWD, 0.68, j === 2 ? "FCE8E6" : WHITE, { line: j === 2 ? "F2B8B3" : LINE, round: true, r: 0.08 });
        txt(s, t, X + 0.08, y, CWD - 0.16, 0.68, { size: 10, bold: j === 2, color: j === 2 ? RED : DARK, align: "center", valign: "middle", margin: 0 });
      });
      arrowR(s, 7.3, y + 0.22, 0.24, 0.24, RED);
    });
    line(s, 0.3, 6.32, 9.4, 0, LINE, 0.75, { dash: "dash" });
    key(s, "成功の型は、貴社でも再現できます。", { y: 6.42, size: 17 });
  }

  // =====================================================================
  // 船井総研がご提供できること
  // =====================================================================
  const S9 = "船井総研がご提供できること";
  pres.addSection({ title: S9 });
  {
    const s = pres.addSlide({ masterName: "BLANK", sectionTitle: S9 });
    rect(s, 5.4, 0, 4.6, 7.5, NAVY);
    s.addShape(pres.shapes.OVAL, { x: 6.2, y: 4.8, w: 6, h: 4, fill: { color: BLUE }, line: { type: "none" }, objectName: "div_wave" });
    txt(s, "船井総研が\nご提供できること", 0.4, 2.9, 5.0, 1.4, { size: 30, bold: true, color: NAVY, align: "center", valign: "middle" });
    txt(s, "ともに、地域に必要とされる遺品整理事業を。", 5.6, 3.3, 4.2, 0.8, { size: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
    s.addNotes("事後経営相談の最後に使う。");
  }
  {
    const s = content(S9, "3つの支援サービス");
    msg(s, "貴社の課題に合わせて、3つのサービスでトータルにご支援します");
    const SV = [["laptop", "01", "HPディレクション支援", ["商圏の競合調査", "強みの整理と打ち出し", "コンセプト設計", "キャッチコピーのご提案"], "※デザイン制作はご希望のHP制作会社へ依頼", NAVY], ["list", "02", "ツール＆マニュアル導入支援", ["事業概要マニュアル", "営業マニュアル", "現場マニュアル（60ページ以上）", "集客マニュアル（60ページ以上）", "見積システム・契約書 ほか"], "", "1F4E3D"], ["chart", "03", "月次コンサルティング", ["事業戦略（3ヵ年計画・財務分析）", "商品設計", "集客戦略（HP・チラシ・広告）", "営業戦略（ツール・研修）", "採用戦略"], "", "8A3B12"]];
    SV.forEach((v, i) => {
      const x = 0.4 + i * 3.13;
      rect(s, x, 1.6, 2.94, 4.7, WHITE, { line: v[5], lw: 1.5 });
      rect(s, x, 1.6, 2.94, 1.2, v[5]);
      txt(s, v[1], x + 0.12, 1.6, 0.8, 0.7, { size: 26, color: WHITE, valign: "middle", margin: 0 });
      icon(s, v[0], x + 2.25, 1.75, 0.5, WHITE);
      txt(s, v[2], x + 0.12, 2.25, 2.7, 0.5, { size: 13, bold: true, color: WHITE, valign: "middle", margin: 0 });
      v[3].forEach((t, k) => {
        s.addImage({ data: IC("ok", v[5] === NAVY ? NAVY : (i === 1 ? GREEN : BROWN)), x: x + 0.15, y: 3.06 + k * 0.5, w: 0.26, h: 0.26, objectName: nm("ok") });
        txt(s, t, x + 0.5, 3.0 + k * 0.5, 2.38, 0.38, { size: 10.5, valign: "middle", margin: 0 });
      });
      if (v[4]) txt(s, v[4], x + 0.15, 5.55, 2.65, 0.6, { size: 9, color: GRAY, ls: 1.2 });
    });
    key(s, "事業計画から商品設計・集客・営業・現場まで、一気通貫でご支援します。", { y: 6.45, size: 15 });
  }
  {
    const s = content(S9, "ご提供するツール（一部抜粋）");
    msg(s, "商品設計・集客・営業・現場作業まで、明日から使えるツールを貴社仕様でご提供します", { size: 13 });
    const F = [["pdf", RED, "A4エンド向け紹介資料"], ["doc", BLUE, "WEBサイト制作ディレクション資料"], ["ppt", ORANGE, "アライアンス提案書"], ["pdf", RED, "テンプレートチラシ"], ["xls", GREEN, "古物台帳（入力版）"], ["xls", GREEN, "現場準備物マスターリスト"], ["ppt", ORANGE, "口コミ収集POP"], ["pdf", RED, "遺品整理後アプローチブック"], ["ppt", ORANGE, "現場研修資料"], ["doc", BLUE, "業務提携契約書（汎用版）"], ["doc", BLUE, "遺品整理請負契約書"], ["doc", BLUE, "遺品買取同意書"], ["xls", GREEN, "提案書明細（7タブ連動）"], ["img", PURPLE, "見積書（140段ボール基準）"], ["xls", GREEN, "利益管理・見積システム"], ["video", RED, "会社信頼向上動画"], ["ppt", ORANGE, "業者提携用資料"], ["xls", GREEN, "アライアンス管理表"]];
    F.forEach((f, i) => {
      const x = 0.4 + (i % 3) * 3.1, y = 1.55 + Math.floor(i / 3) * 0.78;
      rect(s, x, y, 2.95, 0.66, WHITE, { line: LINE, shadow: true });
      s.addImage({ data: IC(f[0], f[1]), x: x + 0.12, y: y + 0.12, w: 0.42, h: 0.42, objectName: nm("ficon") });
      txt(s, f[2], x + 0.62, y, 2.28, 0.66, { size: 10, bold: true, valign: "middle", margin: 0 });
    });
    key(s, "マニュアル4種＋ツール類を、貴社の仕様に合わせて作成・納品します。", { y: 6.35, size: 15 });
  }
  {
    const s = content(S9, "月次支援の進め方");
    msg(s, "月1回の打ち合わせ（約4時間）と日常のサポートで、30日間の実行を最大化します", { size: 13 });
    txt(s, "1日の打ち合わせの例", 0.4, 1.55, 5.5, 0.35, { size: 13, bold: true, color: NAVY, margin: 0 });
    const TL = [["13:00", "事業全体の数値確認", "事業数値・営業KPIの進捗確認"], ["13:30", "営業", "進捗・追客案件・見込み案件の確認、営業ツールの修正"], ["14:40", "集客", "WEB・HPの改善提案、集客スケジュール、販促強化策"], ["16:30", "採用", "採用の進捗・人員体制の確認"]];
    line(s, 1.25, 2.1, 0, 3.9, NAVY, 3);
    TL.forEach((t, i) => {
      const y = 2.0 + i * 1.0;
      oval(s, 1.13, y + 0.08, 0.24, 0.24, i % 2 ? BLUE : NAVY);
      txt(s, t[0], 0.4, y, 0.7, 0.4, { size: 12, bold: true, color: RED, valign: "middle", margin: 0 });
      txt(s, t[1], 1.55, y, 4.2, 0.4, { size: 13, bold: true, color: NAVY, valign: "middle", margin: 0 });
      txt(s, t[2], 1.55, y + 0.4, 4.3, 0.5, { size: 10, color: GRAY, margin: 0 });
    });
    txt(s, "17:00 終了", 0.4, 6.0, 1.5, 0.35, { size: 10, bold: true, color: GRAY, margin: 0 });
    frame(s, 6.2, 1.6, 3.4, 4.75, "日常の実行管理サポート", NAVY);
    [["line", "LINEサポート"], ["phone", "TELサポート（9:00〜20:00）"], ["mail", "メールサポート"], ["talk", "ZOHO Cliqサポート"], ["laptop", "オンライン打合せ"], ["link", "ツール共有（クラウド）"]].forEach((r, i) => {
      icon(s, r[0], 6.4, 2.2 + i * 0.62, 0.36, NAVY);
      txt(s, r[1], 6.85, 2.17 + i * 0.62, 2.7, 0.42, { size: 10.5, bold: true, valign: "middle", margin: 0 });
    });
    txt(s, "進捗確認・課題相談にいつでも対応します", 6.3, 5.85, 3.2, 0.4, { size: 9.5, color: GRAY, align: "center", margin: 0 });
    key(s, "30日間で、実行を最大化させます。", { y: 6.45, size: 17 });
  }
  {
    const s = content(S9, "6ヶ月の立ち上げスケジュール", "【要カスタマイズ】開始月を実際の月名に差し替える。各項目で貴社にお願いすること（担当者の決定、協力業者との面談、写真素材の準備など）も伝える。");
    msg(s, "約6ヶ月で、集客・営業・現場の仕組みを立ち上げます");
    const X0 = 2.2, MW = 1.22, Y0 = 1.6;
    rect(s, 0.4, Y0, 9.2, 0.48, NAVY);
    txt(s, "項目", 0.4, Y0, 1.8, 0.48, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    for (let m = 0; m < 6; m++) txt(s, `${m + 1}ヶ月目`, X0 + m * MW, Y0, MW, 0.48, { size: 11, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    const G = [["業者探し", 0, 1.5, "協力業者探し・原価確認", "5B9BD5"], ["商品設計", 0, 1.5, "基本パック・オプション設計", "70AD47"], ["ブランド設計", 0, 1.5, "屋号・ロゴ・キャッチコピー", "4472C4"], ["HP制作", 0.5, 2.0, "TOP公開→全ページ公開", "ED7D31"], ["営業体制", 1.7, 4.3, "営業ツール制作・営業研修", "D4A017"], ["集客施策", 1.7, 4.3, "チラシ・WEB広告・SEO・MEO", "44546A"], ["提携", 1.2, 4.8, "リスト・資料作成→提携営業", "E0605D"]];
    G.forEach((g, i) => {
      const y = Y0 + 0.52 + i * 0.56;
      rect(s, 0.4, y, 9.2, 0.52, i % 2 ? WHITE : LGRAY);
      txt(s, g[0], 0.5, y, 1.7, 0.52, { size: 11, bold: true, color: NAVY, valign: "middle", margin: 0 });
    });
    for (let m = 1; m < 6; m++) line(s, X0 + m * MW, Y0 + 0.5, 0, 3.95, "D5DAE2", 0.75, { dash: "dash" });
    G.forEach((g, i) => {
      const y = Y0 + 0.52 + i * 0.56;
      rect(s, X0 + g[1] * MW, y + 0.08, g[2] * MW, 0.36, g[4], { round: true, r: 0.12 });
      txt(s, g[3], X0 + g[1] * MW, y + 0.08, g[2] * MW, 0.36, { size: 9.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    });
    rect(s, X0 + 2.0 * MW, Y0 + 4.5, 4.0 * MW, 0.4, "FCE8E6", { round: true });
    txt(s, "営業開始（反響獲得・現地確認・受注）", X0 + 2.0 * MW, Y0 + 4.5, 4.0 * MW, 0.4, { size: 10.5, bold: true, color: RED, align: "center", valign: "middle", margin: 0 });
    key(s, "○月スタートで、○月に営業開始を目指します。", { y: 6.45, size: 16 });
  }
  {
    const s = content(S9, "支援プランと費用", "【要確認】ツール＆マニュアル単体の価格（100万円か150万円か）、研究会の年会費。ここでは最新の企画書（96p）の価格を記載。");
    msg(s, "貴社の状況に合わせて、3つのプランからお選びいただけます（金額はすべて税抜）", { size: 15 });
    const PL = [["プラン1", "HP制作ディレクションのみ", [["商品設計＋HP制作ディレクション", "70万円"]], "※HP制作会社へのデザイン制作費は別途", NAVY], ["プラン2", "マニュアル・オリジナルツールのみ", [["マニュアル・オリジナルツール導入支援", "150万円"]], "", GREEN], ["プラン3", "フルサポート", [["HP制作ディレクション", "70万円"], ["マニュアル・ツール導入支援\n（プラン2より50万円お値引き）", "100万円"], ["業績アップ月次支援", "20万円／月"]], "※月次支援はオンライン＋必要に応じてご訪問", RED]];
    PL.forEach((p, i) => {
      const x = 0.4 + i * 3.13, hot = i === 2;
      rect(s, x, 1.65, 2.94, 4.5, WHITE, { line: p[4], lw: hot ? 2.5 : 1.25, shadow: hot });
      rect(s, x, 1.65, 2.94, 0.95, p[4]);
      txt(s, [B(p[0] + "\n", { fontSize: 11 }), B(p[1], { fontSize: 12.5 })], x, 1.65, 2.94, 0.95, { color: WHITE, align: "center", valign: "middle" });
      if (hot) { rect(s, x + 0.9, 1.56, 1.15, 0.22, YEL, { round: true }); txt(s, "おすすめ", x + 0.9, 1.56, 1.15, 0.22, { size: 9, bold: true, align: "center", valign: "middle", margin: 0 }); }
      p[2].forEach((r, k) => {
        const y = 2.75 + k * 0.92;
        txt(s, r[0], x + 0.15, y, 2.64, 0.42, { size: 9, color: GRAY, align: "center", valign: "middle", margin: 0 });
        txt(s, r[1], x + 0.15, y + 0.4, 2.64, 0.45, { size: 19, bold: true, color: p[4], align: "center", valign: "middle", margin: 0 });
      });
      if (p[3]) txt(s, p[3], x + 0.1, 5.6, 2.74, 0.5, { size: 8.5, color: GRAY, align: "center", ls: 1.15 });
    });
    txt(s, "・コンサルティングは月次支援と研究会の両方を含むサービスです　・PPC運用代行は広告運用費の20%の手数料が発生します", 0.4, 6.2, 9.2, 0.3, { size: 8.5, color: GRAY, margin: 0 });
    key(s, "おすすめは、立上げから定着まで伴走する「フルサポート」です。", { y: 6.5, size: 15 });
  }
  {
    const s = content(S9, "投資総額と回収の見通し", "【要カスタマイズ】初期投資（HP・許可）だけでなく、支援費用も含めた総額で示す。後から「聞いていた金額と違う」とならないようにする。");
    msg(s, "支援費用を含めた投資総額と、回収の見通しをお示しします");
    table(s, [["初年度の投資（プラン3の場合の例）", "金額（税抜）"], ["HP制作ディレクション", "70万円"], ["マニュアル・ツール導入支援", "100万円"], ["月次支援（20万円×12ヶ月）", "240万円"], ["HP制作費（制作会社へ）", "○○万円"], ["研究会 年会費", "○○万円"], ["古物商許可の取得", "約2〜8万円"], [{ text: "合計", options: { bold: true } }, { text: "○○○万円", options: { bold: true, color: RED } }]], 0.4, 1.6, 4.9, [3.3, 1.6], { size: 11, align: "right", rowH: 0.52 });
    frame(s, 5.55, 1.6, 4.05, 2.15, "回収の見通し", NAVY);
    txt(s, [B("遺品整理の営業利益\n", { fontSize: 10.5, color: GRAY }), B("1年目 ○○万円／2年目 ○○万円\n", { fontSize: 13 }), B("本業への波及利益\n", { fontSize: 10.5, color: GRAY }), B("年 ○○万円", { fontSize: 13 })], 5.7, 2.05, 3.8, 1.65, { valign: "middle", ls: 1.2 });
    arrowD(s, 7.3, 3.82, 0.55, 0.35, RED);
    rect(s, 5.55, 4.25, 4.05, 1.85, RED, { round: true });
    txt(s, [B("○年目", { fontSize: 30 }), B("に投資総額を回収\n", { fontSize: 14 }), B("（本業の利益を含めると ○ヶ月）", { fontSize: 11 })], 5.55, 4.25, 4.05, 1.85, { color: WHITE, align: "center", valign: "middle" });
    key(s, "支援費用を含めても、○年目に回収できる計画です。", { y: 6.4, size: 16 });
    foot(s, MODEL_NOTE);
  }
  {
    const s = content(S9, "コンサルティングに伴う契約概要", "【要確認】最新の契約条件に合わせて更新する。");
    msg(s, "ご契約にあたっての主な条件です");
    const items = [
      ["契約期間", "月次支援は原則1年契約です（調査期間を除く）。"],
      ["費用", "月次支援フィーとは別に、交通費を毎月固定額でご請求します。2年目以降のフィーはコンサルタントの役職により変動する場合があります。"],
      ["研究会", "月次支援のご契約には研究会へのご入会が必要です。月次支援は「クラウドサイン」、研究会は「船井総研マイページ」でのお手続きです。"],
      ["お支払い", "毎月のご支援終了後に請求書を送付し、銀行振込となります。"],
      ["別途費用", "ご提案した施策に別途費用がかかる場合があります（購入費・印刷費・サイト制作費など）。リスティング広告の運用手数料は広告費の20%〜です。"],
      ["成果について", "一定の成果を保証するものではありません。成果がないことを理由とした返金は行いません。数値シミュレーションや事例は一例です。"],
      ["その他", "個人情報を含む資料は原則お預かりしません。制作ツールの著作権は弊社に帰属します。営業時間は平日9:30〜18:00です。"],
    ];
    items.forEach((it, i) => {
      const y = 1.7 + i * 0.74, hot = i === 5;
      rect(s, 0.4, y, 1.8, 0.64, hot ? "FCE8E6" : LBLUE);
      txt(s, it[0], 0.4, y, 1.8, 0.64, { size: 11.5, bold: true, color: hot ? RED : NAVY, align: "center", valign: "middle", margin: 0 });
      txt(s, it[1], 2.35, y, 7.25, 0.64, { size: 10, color: hot ? RED : DARK, bold: hot, valign: "middle", ls: 1.2 });
      line(s, 2.35, y + 0.69, 7.25, 0, LINE, 0.5);
    });
  }

  await pres.writeFile({ fileName: OUT });
  const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
  await applyTheme(OUT, THEME);
  console.log("written:", OUT);
}
main().catch((e) => { console.error(e); process.exit(1); });
