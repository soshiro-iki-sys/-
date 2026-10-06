// 株式会社山岸様 塗装勉強会 空パッケージ（住まいるペイント様の資料デザインを踏襲）
const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require("/root/.claude/skills/synced/d4ca92c1-79b2-4135-88f3-5b208d472f62_ae682e09-d45b-40f9-9321-23b7b4a33115/pptx/scripts/apply_theme.js");

const A = (f) => path.join(__dirname, "assets", f);
const OUT = process.argv[2] || path.join(__dirname, "yamagishi_seminar.pptx");

const THEME = {
  name: "Yamagishi Seminar",
  headFontFace: "Meiryo",
  bodyFontFace: "Meiryo",
  colors: {
    dk1: "000000", lt1: "FFFFFF", dk2: "00395C", lt2: "FCE4D0",
    accent1: "F79646", // オレンジ（見出し帯の線）
    accent2: "C9161D", // 赤い四角
    accent3: "FFC000", // フッターの金線
    accent4: "00B0F0", // 締めの水色
    accent5: "2E75B6", // 工程カードの青
    accent6: "FF0000", // 強調の赤
    hlink: "0563C1", folHlink: "954F72",
  },
};

const pres = new pptxgen();
pres.defineLayout({ name: "SEMINAR_4x3", width: 10.833, height: 7.5 });
pres.layout = "SEMINAR_4x3";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "株式会社山岸様 塗装勉強会";
pres.company = "株式会社山岸";
const C = pres.SchemeColor;

const W = 10.833;
const GRAY_TXT = "7F7F7F";

// ---------- レイアウト ----------
// 本講座ページ：オレンジの見出し帯＋赤い四角＋金と紺のフッター線＋右下ロゴ
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { rect: { x: 0, y: 0.444, w: W, h: 0.625, fill: { color: C.background2 } } },
    { rect: { x: 0, y: 0.444, w: W, h: 0.083, fill: { color: C.accent1 } } },
    { rect: { x: 0, y: 1.069, w: W, h: 0.083, fill: { color: C.accent1 } } },
    { rect: { x: 0.11, y: 0.57, w: 0.23, h: 0.23, fill: { color: C.accent2 } } },
    { rect: { x: 0.2, y: 0.66, w: 0.2, h: 0.2, fill: { color: C.background1 } } },
    { rect: { x: 0, y: 6.875, w: W, h: 0.03, fill: { color: C.accent3 } } },
    { rect: { x: 0, y: 6.905, w: W, h: 0.016, fill: { color: C.text2 } } },
    { image: { x: 8.55, y: 7.0, w: 2.07, h: 0.352, path: A("logo_h.png") } },
    { placeholder: { options: { name: "title", type: "title", x: 0.55, y: 0.53, w: 10.1, h: 0.54, fontSize: 26, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0 }, text: "見出し" } },
  ],
});

// 中扉：家のアイコン＋番号＋細い線＋章タイトル
pres.defineSlideMaster({
  title: "DIVIDER",
  background: { color: C.background1 },
  objects: [
    { image: { x: 3.85, y: 2.3, w: 0.62, h: 0.585, path: A("house_icon.png") } },
    { rect: { x: 4.1, y: 2.86, w: 2.35, h: 0.012, fill: { color: "8C7B3E" } } },
    { placeholder: { options: { name: "num", type: "body", x: 4.6, y: 2.2, w: 1.8, h: 0.7, fontSize: 40, bold: true, color: C.text1, valign: "bottom", margin: 0 }, text: "00" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 3.25, w: 9.633, h: 1.6, fontSize: 36, bold: true, color: C.text1, align: "center", valign: "top", margin: 0 }, text: "章タイトル" } },
  ],
});

// 表紙：芝生の写真＋半透明の白い帯
pres.defineSlideMaster({
  title: "COVER",
  background: { path: A("cover_grass.jpg") },
  objects: [
    { rect: { x: 0, y: 1.78, w: W, h: 3.3, fill: { color: C.background1, transparency: 30 } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.4, y: 2.0, w: 10.033, h: 1.7, fontSize: 32, bold: true, color: C.text1, align: "center", valign: "middle", margin: 0 }, text: "タイトル" } },
  ],
});

// ---------- 部品 ----------
let curSection = "";
function section(t) { curSection = t; pres.addSection({ title: t }); }
function content(title, notes, titleSize) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: curSection });
  s.addText(titleSize ? [{ text: title, options: { fontSize: titleSize } }] : title, { placeholder: "title" });
  if (notes) s.addNotes(notes);
  return s;
}
function divider(num, title, notes) {
  const s = pres.addSlide({ masterName: "DIVIDER", sectionTitle: curSection });
  s.addText(num, { placeholder: "num" });
  s.addText(title, { placeholder: "title" });
  if (notes) s.addNotes(notes);
  return s;
}
// 小見出し「(1)外壁・屋根は劣化します」
function sub(s, text, o = {}) {
  s.addText(text, { x: o.x ?? 0.55, y: o.y ?? 1.25, w: o.w ?? 9.8, h: o.h ?? 0.45, fontSize: o.size ?? 20, color: C.text1, align: o.align ?? "left", margin: 0, isTextBox: true, objectName: "小見出し" });
}
// 本文
function txt(s, text, x, y, w, h, o = {}) {
  s.addText(text, Object.assign({ x, y, w, h, fontSize: 16, color: C.text1, margin: 0.04, valign: "top", isTextBox: true, paraSpaceAfter: 4 }, o));
}
// 素材の差し込み枠（写真・図・表）
function ph(s, x, y, w, h, label) {
  s.addText(label, {
    x, y, w, h, fontSize: 11, color: GRAY_TXT, align: "center", valign: "middle",
    fill: { color: "F2F2F2" }, line: { color: "BFBFBF", width: 1, dashType: "dash" },
    margin: 0.08, isTextBox: true, objectName: "差し込み枠",
  });
}
// 写真の下の「●」キャプション
function cap(s, text, x, y, w) {
  s.addText(text, { x, y, w, h: 0.28, fontSize: 11, color: C.text1, margin: 0, isTextBox: true });
}
// 青い丸の「特性」バッジ
function badge(s, x, y, label = "特性") {
  s.addText(label, { shape: pres.ShapeType.ellipse, x, y, w: 0.95, h: 0.95, fill: { color: "4A90E2" }, color: C.background1, fontSize: 20, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "特性バッジ" });
}
// 赤いラベル（写真に重ねる）
function redLabel(s, text, x, y, w = 1.1) {
  s.addText(text, { x, y, w, h: 0.48, fill: { color: "FF0000" }, color: C.background1, fontSize: 22, align: "center", valign: "middle", margin: 0, isTextBox: true });
}
// 吹き出し「〜〇〇〜」
function wave(s, text, y = 1.25) {
  s.addText(`〜${text}〜`, { x: 0.5, y, w: 9.833, h: 0.45, fontSize: 20, color: C.text1, align: "center", margin: 0, isTextBox: true });
}
const red = (t) => ({ text: t, options: { color: C.accent6, bold: true } });
const b = (t) => ({ text: t, options: { bold: true } });
const n = (t) => ({ text: t });
const br = (t, o = {}) => ({ text: t, options: Object.assign({ breakLine: true }, o) });

const T1 = "1.本勉強会の経緯と目的";
const T2 = "2.絶対に知っておきたい屋根・外壁の基礎知識";
const T3 = "3.意外な落とし穴？！失敗しない工事3つのポイント";
const T4 = "4.業者選びのポイント";

// =====================================================================
section("オープニング");
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: curSection });
  s.addText("失敗しない外壁・屋根塗装のための\n勉強会", { placeholder: "title" });
  s.addImage({ path: A("logo_h.png"), x: 4.0, y: 3.85, w: 2.83, h: 0.48, objectName: "ロゴ" });
  s.addText("主催：株式会社山岸（ヤマキシペイント）　20XX年X月X日（X）　会場：〇〇〇〇", { x: 0.5, y: 4.45, w: 9.833, h: 0.4, fontSize: 14, color: C.text1, align: "center", margin: 0, isTextBox: true });
  s.addNotes("表紙。タイトル／主催／日付／会場を記入。");
}
{
  const s = content("本日の流れ", "①〜④＋休憩＋まとめ・質疑応答の順に進めることを伝える。");
  const items = [
    ["1．本勉強会の経緯と目的", 1.55],
    ["2．絶対に知っておきたい屋根・外壁の基礎知識", 2.25],
    ["休憩", 2.95, true],
    ["3．意外な落とし穴？！失敗しない工事3つのポイント", 3.65],
    ["4．業者選びのポイント", 4.35],
    ["5．まとめ・質疑応答", 5.05],
  ];
  for (const [t, y, center] of items) {
    s.addText(t, { x: center ? 0.5 : 1.1, y, w: center ? 9.833 : 9.2, h: 0.5, fontSize: 22, bold: true, color: C.text1, align: center ? "center" : "left", margin: 0, isTextBox: true });
  }
}

// =====================================================================
section("① 本勉強会の経緯と目的");
divider("01", "本勉強会の経緯と目的");
{
  const s = content(T1, "講師の名前・肩書き・資格・経歴。創業120年以上の会社として地域の家を見てきたことを一言（会社の宣伝は④で）。");
  sub(s, "自己紹介", { x: 3.9, size: 22 });
  ph(s, 0.55, 1.45, 2.3, 3.6, "【写真】\n講師の写真");
  txt(s, [br("〇〇（資格名）", { fontSize: 18, align: "center" }), br("〇〇 〇〇", { fontSize: 22, bold: true, align: "center" }), n("（ふりがな）")], 0.35, 5.15, 2.7, 1.3, { fontSize: 10, align: "center" });
  ph(s, 3.9, 1.9, 6.4, 4.6, "【本文】\n・経歴（塗装・建築に携わって〇年）\n・保有資格\n・創業120年以上の会社として、地域の家を見てきたこと\n・この勉強会への想い");
}
{
  const s = content(T1, "リフォーム前のお客様の不安の上位4つ。統計は最新版（住宅リフォーム推進協議会）に差し替える。");
  sub(s, "なぜ、勉強会を開催するのか？", { align: "center", w: 9.833, x: 0.5, size: 22 });
  ph(s, 0.4, 1.85, 5.2, 4.75, "【グラフ】\nリフォーム前の不安の調査結果\n（住宅リフォーム推進協議会・最新版）\n上位4項目を赤い破線で囲む");
  txt(s, [br("塗装工事・リフォーム工事前における"), br("お客様のほとんどが、"), br(""),
    br("・適正価格がわからない", { color: C.accent6, bold: true }),
    br("・しっかりとした施工が行われるか", { color: C.accent6, bold: true }),
    br("・どんな業者を選べばいいかわからない", { color: C.accent6, bold: true }),
    br("・なにを比較したらいいかわからない", { color: C.accent6, bold: true }), br(""),
    n("という悩みがあります。")], 5.85, 2.3, 4.7, 3.5, { fontSize: 17 });
  txt(s, "引用：一般社団法人住宅リフォーム推進協議会（調査名・年度・URLを記載）", 5.85, 6.25, 4.7, 0.4, { fontSize: 10, color: GRAY_TXT });
}
{
  const s = content(T1, "住まいの相談で多いのは「ひび割れ」「雨漏り」＝外壁と屋根。訪問販売の点検商法トラブルも増えている。統計は最新版に差し替える。");
  sub(s, "なぜ、勉強会を開催するのか？", { align: "center", w: 9.833, x: 0.5, size: 22 });
  ph(s, 0.4, 1.85, 5.6, 3.3, "【表】\n住宅の不具合の相談内容（住宅相談統計年報・最新版）\nひび割れ・雨漏りなど上位を赤い破線で囲む");
  ph(s, 0.4, 5.3, 5.6, 1.35, "【数字】\n訪問販売の点検商法の相談件数（国民生活センター）");
  txt(s, [br("住まいの不具合の相談で多いのは"), br(""),
    br("・ひび割れ"), br("・雨漏り"), br("・はがれ、変形"), br(""),
    br("といった外壁や屋根のことです。"), br(""),
    br("訪問販売の点検商法によるトラブルも増えています。"), br(""),
    br("⇒外壁・屋根のメンテナンスでは", { color: C.accent6, bold: true }),
    n("　業者選びがとても大切！", { color: C.accent6, bold: true })], 6.25, 1.9, 4.3, 4.7, { fontSize: 16 });
}
{
  const s = content(T1, "地域のお客様から実際にいただくお悩み。「今日はこの4つに順番にお答えします」と予告する。");
  sub(s, "地域のお客様から頂いているお悩み", { align: "center", w: 9.833, x: 0.5, size: 20 });
  const q = ["◆どこに頼めばいいか分からない", "◆工事価格はどれくらいが適正なの？", "◆業者選びのポイントは？", "◆自分の家にどんな工事をすればいいか分からない"];
  q.forEach((t, i) => s.addText(t, { x: 0.45, y: 1.95 + i * 0.95, w: 10, h: 0.6, fontSize: 26, bold: true, color: C.text1, margin: 0, isTextBox: true }));
  s.addText("→ 今日はこの4つに順番にお答えします", { x: 0.45, y: 5.95, w: 10, h: 0.5, fontSize: 20, bold: true, color: C.accent6, margin: 0, isTextBox: true });
}

// =====================================================================
section("② 屋根・外壁の基礎知識");
divider("02", "絶対に知っておきたい\n屋根・外壁の基礎知識");
{
  const s = content(T2, "塗装の目的は2つ。車検のような決まりはないので、自分で家の状態を知ることが大切。2,000万〜3,000万円の資産を守る。");
  sub(s, "(1)外壁・屋根は劣化します");
  sub(s, "なぜ塗装が必要なのか？", { y: 1.8, size: 24, w: 9.8 });
  s.addText([br("目的①", { fontSize: 16 }), n("家を長持ちさせる")], { x: 0.6, y: 2.55, w: 4.6, h: 1.1, fontSize: 22, bold: true, color: C.text1, fill: { color: C.background2 }, align: "center", valign: "middle", isTextBox: true });
  s.addText([br("目的②", { fontSize: 16 }), n("美しさを保つ")], { x: 5.6, y: 2.55, w: 4.6, h: 1.1, fontSize: 22, bold: true, color: C.text1, fill: { color: C.background2 }, align: "center", valign: "middle", isTextBox: true });
  txt(s, [br("・「新築から何年で塗装」という決まりや、車検のような制度はありません"), br("　→ 自分で家の状態を知ることが大切です"), br("・2,000万〜3,000万円かけて建てた", { }), n("　大切な資産を守りましょう", { color: C.accent6, bold: true })], 0.6, 3.95, 5.6, 2.6, { fontSize: 16 });
  ph(s, 6.5, 3.95, 3.7, 2.6, "【写真・イラスト】\n家のイメージ");
}
{
  const s = content(T2, "塗膜は樹脂・顔料・添加剤でできている。紫外線で樹脂が傷むと手に白い粉がつく（チョーキング）＝守る力が落ちてきたサイン。");
  sub(s, "(1)外壁・屋根は劣化します");
  ph(s, 0.9, 1.8, 3.3, 4.9, "【図】\n塗膜のしくみ\n（外壁＋専用防水塗膜：\n樹脂・顔料・添加剤）");
  ph(s, 4.8, 1.8, 2.0, 0.8, "【アイコン】\n太陽・雨・CO₂");
  ph(s, 4.8, 2.75, 4.2, 2.6, "【写真】\nチョーキング（手に白い粉がついた写真）");
  txt(s, [br("【チョーキング現象】", { bold: true }), n("　が防水切れの目安です。", { bold: true })], 4.8, 5.45, 4.6, 0.8, { fontSize: 15 });
}
{
  const s = content(T2, "北陸は雨・雪が多く湿気が高い。凍ったりとけたりで表面がはがれる（凍害）。カビ・コケも出やすい。写真は自社の現場のもの。");
  sub(s, "(1)外壁・屋根は劣化します");
  txt(s, [b("北陸の家は傷みやすい！"), n("　雨・雪が多く湿気が高い／凍ったりとけたりをくり返す")], 0.55, 1.75, 9.8, 0.45, { fontSize: 18 });
  const xs = [0.4, 3.85, 7.3];
  const labels = ["【写真】凍害で表面がはがれた外壁", "【写真】カビ・コケ", "【写真】雪・雨による傷み"];
  const caps = ["●凍害（表面のはがれ）", "●カビやコケが発生する", "●雪・雨による傷み"];
  xs.forEach((x, i) => { ph(s, x, 2.4, 3.15, 3.8, labels[i]); cap(s, caps[i], x, 6.25, 3.15); });
}
{
  const s = content(T2, "ひび割れ（クラック）は毛細管現象で水が入りやすい。");
  sub(s, "(1)外壁・屋根は劣化します");
  badge(s, 0.6, 1.85);
  txt(s, [{ text: "・ 経年劣化で", options: { bold: true } }, br("クラック（ひび割れ）", { color: "4A90E2" }), n("・ "), { text: "毛細管現象", options: { color: "4A90E2" } }, { text: "により水が浸入し易い", options: { bold: true } }], 1.75, 1.9, 8.5, 1.0, { fontSize: 22 });
  ph(s, 0.9, 3.15, 4.4, 3.2, "【写真】\n外壁のひび割れ（赤い楕円で囲む）");
  cap(s, "クラック（ひび割れ）", 2.2, 6.4, 2.5);
  ph(s, 5.8, 3.15, 3.9, 3.5, "【図】\nひび割れの入った家のイラスト");
}
{
  const s = content(T2, "サイディングは経年劣化で割れやすく、反りが出やすい。");
  sub(s, "(1)外壁・屋根は劣化します");
  badge(s, 0.85, 1.85);
  txt(s, [br("・ 経年劣化で割れやすい", { bold: true }), { text: "・ サイディング壁は[反り]が出やすい", options: { bold: true } }], 2.0, 1.9, 8.0, 1.0, { fontSize: 22 });
  ph(s, 0.9, 3.15, 4.3, 3.3, "【写真】\nサイディングの反り");
  redLabel(s, "反り", 3.95, 3.3);
  ph(s, 5.6, 3.15, 4.3, 3.3, "【写真】\nサイディングの割れ");
  redLabel(s, "割れ", 8.65, 3.3);
}
{
  const s = content(T2, "サイディングをつなぐ目地（シーリング）は経年劣化で割れ・ひび・痩せが起こる。");
  sub(s, "(1)外壁・屋根は劣化します");
  badge(s, 0.85, 1.85);
  txt(s, [br("・ サイディング壁を繋ぐ【目地】は経年劣化で", { bold: true }), { text: "　割れ・ひび・痩せ", options: { color: "4A90E2", bold: true } }, { text: "が起こる！", options: { bold: true } }], 2.0, 1.9, 8.3, 1.0, { fontSize: 22 });
  ph(s, 0.45, 3.1, 4.8, 3.5, "【写真】\n目地の割れ・痩せ①");
  ph(s, 5.55, 3.1, 4.8, 3.5, "【写真】\n目地の割れ・痩せ②");
}
{
  const s = content(T2, "屋根の劣化（色あせ・錆・割れ・ずれ）。地域で多い屋根（金属・スレート・瓦）の例を自社の現場写真で。");
  sub(s, "(1)外壁・屋根は劣化します", { y: 1.2 });
  const caps = ["●屋根塗膜の劣化・色あせ", "●錆", "●割れ", "●ずれ", "●棟板金の傷み", "●雪止め・板金まわり"];
  for (let i = 0; i < 6; i++) {
    const x = 0.25 + (i % 3) * 3.5, y = 1.75 + Math.floor(i / 3) * 2.55;
    ph(s, x, y, 3.3, 2.1, "【写真】屋根の劣化");
    cap(s, caps[i], x + 0.6, y + 2.12, 2.7);
  }
}
{
  const s = content(T2, "放っておくと雨漏り→柱や下地が腐る→塗装では済まない大がかりな改修に。");
  sub(s, "(1)外壁・屋根は劣化します");
  txt(s, [br("外壁や屋根の劣化を放置してしまうと・・・・"), br("最悪の場合、お家の中に雨水が侵入して雨漏りの原因になるだけでなく、"), br("お家の中が腐ってしまったりなど、", { color: C.accent6 }), n("塗装工事どころではなく、家自体の改修工事が必要な状態になってしまうことも少なくありません。")], 0.55, 1.75, 9.8, 1.6, { fontSize: 17 });
  ph(s, 0.55, 3.5, 4.6, 2.75, "【写真】\n壁の中の腐食・雨漏りの跡");
  ph(s, 5.4, 3.5, 2.3, 2.75, "【写真】\n目地の割れ");
  ph(s, 7.95, 3.5, 2.3, 2.75, "【写真】\nサイディングの欠け");
  s.addText("外壁や屋根の劣化を放置することはやめましょう！", { x: 0.5, y: 6.3, w: 9.833, h: 0.5, fontSize: 22, bold: true, color: C.accent6, align: "center", margin: 0, isTextBox: true });
}
{
  const s = content(T2, "初回は築8〜10年、2回目以降は使った塗料の耐久年数しだい。ただし年数より家の状態で判断する。");
  s.addText("ベストな塗り替えの時期とは？", { x: 0.55, y: 1.45, w: 6.0, h: 0.8, fontSize: 28, bold: true, color: C.text1, margin: 0, isTextBox: true });
  s.addText("塗り替え目安 築8〜10年", { shape: pres.ShapeType.roundRect, rectRadius: 0.1, x: 6.75, y: 1.5, w: 3.7, h: 0.7, fill: { color: C.accent6 }, color: C.background1, fontSize: 20, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
  txt(s, "（例：サイディングの外壁の場合）", 0.7, 2.4, 8, 0.5, { fontSize: 20, bold: true });
  txt(s, [b("初めて塗り替えの方で　"), { text: "8〜10年", options: { color: C.accent6, fontSize: 32 } }, b("程度")], 0.55, 3.2, 9.5, 0.75, { fontSize: 22 });
  txt(s, [b("2回目、3回目の方で　"), { text: "使った塗料の耐久年数しだい", options: { color: C.accent6, fontSize: 26 } }], 0.55, 4.15, 9.8, 0.75, { fontSize: 22 });
  txt(s, [br("年数よりも「家の状態」で判断することが大切です。", { bold: true }), n("ご自分で住まいの状態を把握し、塗り替えのタイミングを計画しましょう", { bold: true })], 0.55, 5.2, 9.8, 1.3, { fontSize: 20 });
}
{
  const s = content(T2, "外壁10工程・屋根8工程の全体像。上から順番に進むことと、準備→下地づくり→塗装→仕上げの段階を見せる。この手順を省く会社は要注意。");
  sub(s, "(2)塗装工事の流れ", { y: 1.2, w: 4.5 });
  s.addText("上から順番に進みます。この手順を省く会社は要注意！", { x: 4.2, y: 1.22, w: 6.2, h: 0.42, fontSize: 15, bold: true, color: C.accent6, align: "right", margin: 0, isTextBox: true });
  // 段階ごとの色（濃いほど工事が進む）
  const PH = { "準備": "BFBFBF", "下地づくり": "F4B183", "塗装": "F79646", "仕上げ": "C55A11" };
  const wall = [["足場組立", "準備"], ["ネット養生", "準備"], ["高圧洗浄", "下地づくり"], ["下地処理", "下地づくり"], ["養生", "下地づくり"], ["下塗り", "塗装"], ["中塗り", "塗装"], ["上塗り", "塗装"], ["完工チェック", "仕上げ"], ["足場解体・清掃", "仕上げ"]];
  const roof = [["高圧洗浄", "下地づくり"], ["鉄部下地調整", "下地づくり"], ["鉄部さび止め塗装", "下地づくり"], ["屋根下塗り", "塗装"], ["屋根中塗り", "塗装"], ["屋根上塗り", "塗装"], ["確認作業", "仕上げ"], ["清掃", "仕上げ"]];
  const ROW = 0.4, TOP = 2.28;
  const flow = (x0, head, steps) => {
    s.addText(head, { x: x0, y: 1.75, w: 4.7, h: 0.45, fill: { color: "1F3864" }, color: C.background1, fontSize: 17, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
    // 段階ラベル（左の縦帯）
    let k = 0;
    while (k < steps.length) {
      let e = k; while (e + 1 < steps.length && steps[e + 1][1] === steps[k][1]) e++;
      const ph = steps[k][1];
      s.addText(ph, { x: x0, y: TOP + k * ROW + 0.03, w: 1.05, h: (e - k + 1) * ROW - 0.06, fill: { color: PH[ph] }, color: ph === "準備" || ph === "下地づくり" ? "000000" : "FFFFFF", fontSize: 12, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "段階" });
      k = e + 1;
    }
    // 番号をつなぐ縦の矢印線
    const cx = x0 + 1.42;
    s.addShape(pres.ShapeType.line, { x: cx, y: TOP + ROW / 2, w: 0, h: (steps.length - 1) * ROW, line: { color: C.accent1, width: 3.5 }, objectName: "流れの線" });
    steps.forEach(([t, ph], i) => {
      const y = TOP + i * ROW;
      s.addText(String(i + 1), { shape: pres.ShapeType.ellipse, x: cx - 0.18, y: y + 0.035, w: 0.36, h: 0.36, fill: { color: C.accent1 }, line: { color: C.background1, width: 1.5 }, color: C.background1, fontSize: 13, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(t, { x: cx + 0.3, y: y + 0.03, w: x0 + 4.7 - (cx + 0.3), h: ROW - 0.06, fill: { color: "FFF4EA" }, color: C.text1, fontSize: 15, bold: true, valign: "middle", margin: [0, 0, 0, 8], isTextBox: true });
    });
    const endY = TOP + steps.length * ROW + 0.06;
    s.addText("完成！", { x: cx - 0.45, y: endY, w: 0.9, h: 0.34, fill: { color: "C00000" }, color: C.background1, fontSize: 13, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
  };
  flow(0.45, "外壁工事（全10工程）", wall);
  flow(5.7, "屋根工事（全8工程）", roof);
}
function koutei(title, body, notes, photos) {
  const s = content(T2, notes);
  sub(s, "(2)塗装工事の流れ", { y: 1.2 });
  txt(s, title, 0.55, 1.65, 9.8, 0.5, { fontSize: 22 });
  txt(s, body, 0.65, 2.2, 9.7, 0.95, { fontSize: 15 });
  return s;
}
{
  const s = koutei("【足場の架設・養生】", [br("正しい施工を行うためには、しっかりとした足場が必要です！"), br("2024年4月から、幅1m以上の場所では「本足場」が原則義務になりました。", { color: C.accent6 }), n("塗料の飛散を防ぐネット張り・養生もきれいに行います。")], "足場・養生：2024年4月から本足場が原則義務／踏板のない足場は危険。近隣への飛散を防ぐ養生。");
  ph(s, 0.6, 3.3, 4.5, 3.3, "【写真】\n自社の足場（本足場・踏板付き）");
  ph(s, 5.6, 3.3, 4.5, 3.3, "【写真】\n養生・飛散防止ネット");
}
{
  const s = koutei("【高圧洗浄・下地処理】", [br("長年たまった汚れを高圧洗浄で落とし、しっかり乾かします。"), n("塗装後の「はがれ」の多くは下地処理不足が原因です。", { color: C.accent6 })], "高圧洗浄・下地処理：塗装後の「はがれ」の多くは下地処理不足が原因。洗浄後はしっかり乾かす。");
  ph(s, 0.6, 3.3, 4.5, 3.3, "【写真】\n高圧洗浄");
  ph(s, 5.6, 3.3, 4.5, 3.3, "【写真】\n下地処理（補修・ケレン）");
}
{
  const s = content(T2, "サイディングの目地は打ち替え（古いものを全部撤去）、サッシまわりは増し打ち。2面接着の図は打ち替えのページに。");
  sub(s, "(2)塗装工事の流れ", { y: 1.2 });
  txt(s, "【シーリング】", 0.55, 1.65, 4, 0.5, { fontSize: 22 });
  const rows = [
    ["打ち替え工法（サイディングの目地）", ["① 撤去", "② プライマー", "③ 充填", "④ 仕上げ"], 2.2],
    ["増し打ち工法（サッシまわり）", ["① プライマー", "② 充填", "③ ヘラならし", "④ 仕上げ"], 4.45],
  ];
  for (const [label, steps, y] of rows) {
    txt(s, label, 0.6, y, 6, 0.38, { fontSize: 14, bold: true });
    steps.forEach((t, i) => {
      const x = 0.6 + i * 2.5;
      s.addText(t, { x, y: y + 0.4, w: 2.15, h: 0.36, fill: { color: C.accent5 }, color: C.background1, fontSize: 13, margin: 0.06, isTextBox: true });
      ph(s, x, y + 0.76, 2.15, 1.3, "【写真】");
      if (i < 3) s.addText("❯", { x: x + 2.15, y: y + 1.15, w: 0.35, h: 0.5, fontSize: 20, color: "A6A6A6", align: "center", margin: 0, isTextBox: true });
    });
  }
}
{
  const s = koutei("【3回塗り（下塗り・中塗り・上塗り）】", [br("メーカーが決めた塗る量と乾燥時間を守ります。"), n("秤で調合し、決められた希釈率を守ります。", { color: C.accent6 })], "3回塗り：メーカーが決めた塗る量と乾燥時間を守る。秤で調合し、決められた希釈率を守る。");
  ["【下塗り】", "【中塗り】", "【上塗り】"].forEach((t, i) => { ph(s, 0.45 + i * 3.4, 3.3, 3.15, 2.8, "【写真】" + t.replace(/[【】]/g, "")); txt(s, t, 0.45 + i * 3.4, 6.15, 3.15, 0.4, { fontSize: 15, align: "center" }); });
}
{
  const s = koutei("【付帯部・完工チェック・お引き渡し】", [br("雨樋・破風・鉄部の錆止めなど、細かい部分も丁寧に仕上げます。"), n("社内の完工チェック → お客様立ち会いの検査 → 保証書のお渡し")], "付帯部（雨樋・破風・鉄部の錆止め）。完工チェック→立ち会い検査→保証書のお渡し。");
  ph(s, 0.6, 3.3, 4.2, 3.3, "【写真】\n付帯部の塗装（雨樋・破風・鉄部）");
  s.addText("Before", { x: 5.2, y: 3.3, w: 1.0, h: 0.35, fill: { color: C.accent6 }, color: C.background1, fontSize: 13, margin: 0.05, isTextBox: true });
  ph(s, 5.2, 3.65, 2.2, 1.7, "【写真】施工前");
  s.addText("After", { x: 7.95, y: 4.35, w: 1.0, h: 0.35, fill: { color: C.accent6 }, color: C.background1, fontSize: 13, margin: 0.05, isTextBox: true });
  ph(s, 7.6, 4.7, 2.6, 1.95, "【写真】施工後");
}
{
  const s = content("休憩", "休憩（〇分）。");
  s.addText("休憩（〇分）", { x: 0.5, y: 3.0, w: 9.833, h: 1.0, fontSize: 40, bold: true, color: C.text1, align: "center", margin: 0, isTextBox: true });
}

// =====================================================================
section("③ 失敗しない工事3つのポイント");
divider("03", "意外な落とし穴？！\n失敗しない工事3つのポイント");
{
  const s = content(T3, "お問合せから引き渡し・アフターまでの流れ。いつまでに工事を終えたいかを決めて、逆算して計画する。各ステップに日付の記入欄。", 22);
  sub(s, "ポイント1　流れを知って、逆算して計画する");
  const steps = ["お問い合わせ", "健康診断\n（現場調査）", "ご要望の確認", "診断結果・\nお見積りのご説明", "ご契約", "色の打合せ", "近隣への\nご挨拶", "着工", "完工チェック", "お引き渡し・\n保証書", "定期点検・\nアフターサービス"];
  steps.forEach((t, i) => {
    const row = i < 6 ? 0 : 1, col = row ? i - 6 : i;
    const x = 0.25 + col * 1.73, y = 1.85 + row * 2.05;
    s.addShape(pres.ShapeType.rect, { x, y, w: 1.62, h: 1.85, fill: { color: C.background1 }, line: { color: C.accent3, width: 1.5 } });
    s.addText(String(i + 1), { x, y, w: 0.38, h: 0.5, fill: { color: C.accent3 }, color: C.background1, fontSize: 18, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(t, { x: x + 0.4, y, w: 1.2, h: 0.62, fontSize: 10.5, bold: true, color: C.text1, valign: "middle", margin: 0.02, isTextBox: true });
    s.addText("月　　日", { x, y: y + 0.65, w: 1.62, h: 0.3, fontSize: 11, color: C.text1, align: "center", margin: 0, isTextBox: true });
    ph(s, x + 0.08, y + 0.98, 1.46, 0.8, "【説明・イラスト】");
  });
  txt(s, "いつまでに工事を終えたいかを決めて、そこから逆算してスケジュールを組みましょう！", 0.55, 6.1, 9.8, 0.6, { fontSize: 17, bold: true });
}
{
  const s = content(T3, "工事中も安心のお約束：近隣挨拶の代行／施工管理の責任者（丸投げしない）／禁煙／お茶菓子不要／毎日の作業報告／工程写真の報告書／工期厳守。", 22);
  sub(s, "ポイント1　工事中も安心のお約束");
  const items = [["①近隣挨拶は当社が代行", "ご契約者様に代わって、工事の専門家がご挨拶します"], ["②施工管理の責任者がつく", "丸投げはしません"], ["③施工中禁煙", "現場内での喫煙は行いません"], ["④お茶菓子不要", "職人へのお気づかいは不要です"], ["⑤毎日の作業報告", "作業の前後に内容をご報告します"], ["⑥工程を写真で記録", "工事後に報告書にまとめてご提出します"], ["⑦整理整頓・清掃", "毎日の後片付けを徹底します"], ["⑧工期厳守", "予定期間内の完了に努めます"]];
  items.forEach(([h, d], i) => {
    const x = 0.5 + (i % 2) * 5.0, y = 1.8 + Math.floor(i / 2) * 1.2;
    s.addText(h, { x, y, w: 4.7, h: 0.42, fill: { color: "C00000" }, color: C.background1, fontSize: 15, bold: true, margin: 0.08, isTextBox: true });
    s.addText(d, { x, y: y + 0.42, w: 4.7, h: 0.62, fill: { color: "F2F2F2" }, color: C.text1, fontSize: 13, margin: 0.08, valign: "middle", isTextBox: true });
  });
}
{
  const s = content(T3, "健康診断（現場調査）で見るところ。機械を使わない点検：目で見る・手で触る・写真で記録して説明。約1時間30分。", 22);
  sub(s, "ポイント1　健康診断（現場調査）で見るところ");
  txt(s, "次のことをしているかチェックしましょう　※機械を使わない点検", 0.55, 1.75, 9.8, 0.45, { fontSize: 18, bold: true });
  s.addText([
    br("□目で見る：ひび割れ・はがれ・色あせ・カビコケ・目地・錆・屋根"),
    br("□手で触る：チョーキング・浮き・反り"),
    br("□写真で記録し、写真を見せながら説明してくれるか"),
    n("□約1時間30分かけて、家のまわりを一周して確認してくれるか"),
  ], { x: 0.55, y: 2.3, w: 9.7, h: 1.9, fontSize: 16, bold: true, color: C.accent6, line: { color: "1F3864", width: 2 }, margin: 0.12, valign: "middle", paraSpaceAfter: 6, isTextBox: true });
  ph(s, 0.55, 4.4, 3.0, 2.3, "【写真】目で見る点検");
  ph(s, 3.9, 4.4, 3.0, 2.3, "【写真】手で触る点検（チョーキング）");
  ph(s, 7.25, 4.4, 3.0, 2.3, "【写真】写真付きの診断結果");
}
{
  const s = content(T3, "傷みすぎると塗装できない。外壁材がはがれる・腐る・変形していると塗っても密着しない→張替え・カバー工事が必要になる。", 22);
  sub(s, "ポイント2　劣化度合いによって工事を決める");
  txt(s, [br("劣化が進みすぎると、塗装しても塗料がすぐにはがれてしまい、", { color: C.accent6 }), br("塗装自体ができなくなってしまいます。", { color: C.accent6 }), br(""), br("・外壁材自体がはがれ落ちている"), br("・カビや藻で外壁が傷み、塗料が密着しない"), br("・外壁が水を吸って変形している"), br(""), n("→張替え・カバー工事が必要になります", { bold: true })], 0.55, 1.8, 5.4, 4.6, { fontSize: 16 });
  ph(s, 6.2, 1.85, 4.1, 2.25, "【写真】塗装できないほど傷んだ外壁①");
  ph(s, 6.2, 4.3, 4.1, 2.25, "【写真】塗装できないほど傷んだ外壁②");
}
{
  const s = content(T3, "放っておくと結局高くつく。外壁100㎡・足場代込みで、塗装パック54.8万円 vs サイディングパック148万円〜（約2.7倍）。「あと何年住みたいか」で工事を選ぶ。価格は最新のものに。", 22);
  sub(s, "ポイント2　放っておくと結局高くつく");
  const card = (x, head, name, price, dur, color) => {
    s.addText(head, { x, y: 1.85, w: 4.6, h: 0.5, fill: { color }, color: C.background1, fontSize: 16, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText([br(name, { fontSize: 15, bold: true }), br(price, { fontSize: 32, bold: true, color }), n(dur, { fontSize: 13 })], { x, y: 2.35, w: 4.6, h: 2.1, align: "center", valign: "middle", line: { color, width: 1.5 }, color: C.text1, margin: 0.1, isTextBox: true });
  };
  card(0.55, "塗装パック", "プレミアムシリコン（ラジカル）", "54.8万円", "（税込）外壁100㎡・足場代込・10年保証\n耐久年数〇〜〇年", "1F9BA8");
  card(5.7, "外壁サイディングパック", "金属サイディング", "148万円〜", "（税込）外壁100㎡・足場代込・10年保証\n耐久年数〇〜〇年", "2E9E44");
  s.addText("サイディング工事は塗装工事の約2.7倍！", { x: 0.5, y: 4.7, w: 9.833, h: 0.6, fontSize: 26, bold: true, color: C.accent6, align: "center", margin: 0, isTextBox: true });
  txt(s, [br("劣化を放置すると結局高くつきます。", { bold: true }), n("「あと何年住みたいか」で、塗装か張替えかを選びましょう。", { bold: true })], 0.55, 5.45, 9.8, 1.0, { fontSize: 18, align: "center" });
}
{
  const s = content(T3, "同じ「シリコン」「ラジカル」でもメーカーによって品質や実績が違う。見積書にメーカー名と商品名が書いてあるかを確認。", 22);
  sub(s, "ポイント3　家に合った塗料を選ぶ");
  s.addText("塗料は「メーカー」で選ぶ時代", { x: 0.5, y: 2.1, w: 9.833, h: 0.9, fontSize: 34, bold: true, color: C.text1, align: "center", margin: 0, isTextBox: true });
  txt(s, [br("同じ「シリコン」「ラジカル」でも、"), br("メーカーによって品質や実績が違います。"), br(""), n("見積書にメーカー名と商品名が書いてあるかを確認しましょう！", { color: C.accent6, bold: true })], 0.8, 3.3, 9.2, 2.2, { fontSize: 20, align: "center" });
  ph(s, 3.4, 5.4, 4.0, 1.2, "【イラスト・写真】\n塗料缶");
}
{
  const s = content(T3, "主な塗料メーカーの比較表。他社を悪く言わず、事実だけを並べる。各社の公式情報で最終確認。", 22);
  sub(s, "ポイント3　主な塗料メーカーの比較");
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: "F79646" }, align: "center", valign: "middle" } });
  const ek = (t, o = {}) => ({ text: t, options: Object.assign({ fill: { color: "FCE4D0" }, bold: true }, o) });
  const rows = [
    [hdr("メーカー"), hdr("創業"), hdr("得意分野"), hdr("外壁用の代表商品"), hdr("備考")],
    ["日本ペイント", "1881年", "自動車・工業用も手がける総合メーカー", "パーフェクトトップ（ラジカル制御形）", "日本で最初の塗料メーカー"],
    ["関西ペイント", "1918年", "自動車・工業用も手がける総合メーカー", "アレスダイナミックTOP（ラジカル制御形）", ""],
    [ek("エスケー化研"), ek("1955年"), ek("建築用の仕上塗材が専門"), ek("エスケープレミアムシリコン（ラジカル制御形）"), ek("建築仕上塗材 国内シェア53％でNo.1", { color: "FF0000" })],
    ["菊水化学工業", "1959年", "建築用の仕上塗材", "グラナダフレッシュ（シリコン）", ""],
    ["アステックペイント", "日本では2000年から販売", "オーストラリアのメーカー", "超低汚染リファイン", "加盟店しか扱えない"],
    ["オリジナル塗料", "—", "塗装店の自社ブランド", "—", "性能の根拠がわかりにくい／比べにくい"],
  ];
  s.addTable(rows, { x: 0.35, y: 1.8, w: 10.13, colW: [1.7, 1.35, 2.4, 2.6, 2.08], fontSize: 12, color: "000000", border: { type: "solid", pt: 0.75, color: "BFBFBF" }, valign: "middle", rowH: 0.62 });
  txt(s, "※各社の公式情報をもとに作成（掲載前に最新情報を確認）", 0.35, 6.35, 9, 0.35, { fontSize: 10, color: GRAY_TXT });
}
{
  const s = content(T3, "なぜエスケー化研の塗料を選ぶのか：①ものがいい ②トラブルが少ない ③日本一（建築仕上塗材 国内シェア53％）。", 22);
  sub(s, "ポイント3　なぜエスケー化研の塗料を選ぶのか");
  const pillars = [
    ["①ものがいい", ["家の外壁・屋根に使う建築用の塗料が専門", "プレミアムシリコン（ラジカル制御・期待耐用年数15年）", "クールテクトSi（遮熱・汚れにくい）", "性能が試験データで確認できる"]],
    ["②トラブルが少ない", ["全国の多くの現場で長年使われてきた実績", "製品情報がすべて公開されていて、お客様自身でも調べられる", "どの塗装店でも同じ品質の材料が手に入る"]],
    ["③日本一", ["建築仕上塗材の国内シェア", "53％", "でNo.1", "六本木ヒルズや甲子園球場でも使われている"]],
  ];
  pillars.forEach(([h, lines], i) => {
    const x = 0.4 + i * 3.4;
    s.addText(h, { x, y: 1.85, w: 3.2, h: 0.6, fill: { color: C.accent1 }, color: C.background1, fontSize: 20, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
    const runs = i === 2
      ? [br(lines[0], { fontSize: 14 }), br(lines[1], { fontSize: 48, bold: true, color: C.accent6 }), br(lines[2], { fontSize: 16, bold: true }), br(""), n(lines[3], { fontSize: 12 })]
      : lines.map((t, k) => ({ text: "・" + t, options: { breakLine: k < lines.length - 1 } }));
    s.addText(runs, { x, y: 2.45, w: 3.2, h: 3.4, fontSize: 14, color: C.text1, fill: { color: "FFF7F0" }, line: { color: C.accent1, width: 1 }, align: i === 2 ? "center" : "left", valign: i === 2 ? "middle" : "top", margin: 0.12, paraSpaceAfter: 6, isTextBox: true });
  });
  ph(s, 0.4, 6.0, 10.0, 0.7, "【ロゴ・写真】エスケー化研のロゴ／代表商品の缶／施工物件の写真");
}
{
  const s = content(T3, "当社のおすすめプランと施工実績（Before/After）。", 22);
  sub(s, "ポイント3　当社のおすすめプランと施工実績");
  ph(s, 0.45, 1.85, 4.6, 2.2, "【プラン】\n塗装パック（プレミアムシリコン）");
  ph(s, 0.45, 4.25, 4.6, 2.2, "【プラン】\n遮熱プラン（クールテクトSi）");
  s.addText("施工前 Before", { x: 5.5, y: 1.85, w: 2.0, h: 0.35, fill: { color: "006B61" }, color: C.background1, fontSize: 12, margin: 0.05, isTextBox: true });
  ph(s, 5.5, 2.2, 2.3, 1.9, "【写真】施工前");
  s.addText("施工後 After", { x: 8.0, y: 1.85, w: 2.0, h: 0.35, fill: { color: "006B61" }, color: C.background1, fontSize: 12, margin: 0.05, isTextBox: true });
  ph(s, 8.0, 2.2, 2.3, 1.9, "【写真】施工後");
  ph(s, 5.5, 4.25, 4.8, 2.2, "【お住まいと工事の概要】\n所在地／部位／工事期間／使用塗料／コメント");
}
{
  const s = content(T3, "3つのポイントのまとめ。そのためには、正しく診断できる会社を選ぶことが大切（→④へ）。", 22);
  wave(s, "3つのポイントのまとめ");
  ["ポイント1\n流れを知って、逆算して計画する", "ポイント2\n劣化度合いによって工事を決める", "ポイント3\n家に合った塗料を選ぶ"].forEach((t, i) => {
    s.addText(t, { x: 0.6 + i * 3.35, y: 2.1, w: 3.05, h: 1.9, fill: { color: C.background2 }, color: C.text1, fontSize: 18, bold: true, align: "center", valign: "middle", margin: 0.1, isTextBox: true });
  });
  s.addText([br("そのためには、"), n("正しく診断できる会社を選ぶことが大切です！", { color: C.accent6 })], { x: 0.5, y: 4.5, w: 9.833, h: 1.4, fontSize: 26, bold: true, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true });
}

// =====================================================================
section("④ 業者選びのポイント");
divider("04", "業者選びのポイント");
{
  const s = content(T4, "価格だけで選ばない。安いには安いなりの理由がある。費用＝足場代・材料費＋職人の人件費＋会社の経費。適正な金額を職人に払うから品質が守れる。");
  wave(s, "業者選びのポイント");
  s.addText("価格だけで業者を選ぶのはやめましょう", { x: 0.55, y: 1.85, w: 9.8, h: 0.65, fontSize: 28, bold: true, color: C.text1, margin: 0, isTextBox: true });
  txt(s, [br("「塗装工事は費用がかかるので、一番安い業者に頼みたい・・・。」", { bold: true }), br(""), br("「安すぎると手抜き工事されるのかと不安になる…」", { bold: true }), br(""), n("安いには安いなりの理由があります。適正な金額を職人に払うから、品質が守れます。", { bold: true })], 0.55, 2.65, 9.8, 2.0, { fontSize: 16 });
  ["【アイコン】\n足場代・材料費", "【アイコン】\n職人の人件費", "【アイコン】\n会社の経費"].forEach((t, i) => ph(s, 0.8 + i * 3.4, 4.9, 2.4, 1.5, t));
  s.addText("＋", { x: 3.25, y: 5.3, w: 0.9, h: 0.7, fontSize: 36, color: "7F7F7F", align: "center", margin: 0, isTextBox: true });
  s.addText("＋", { x: 6.65, y: 5.3, w: 0.9, h: 0.7, fontSize: 36, color: "7F7F7F", align: "center", margin: 0, isTextBox: true });
}
{
  const s = content(T4, "塗装面積を小さくする会社に要注意。見積の面積を実際より小さく書いて安く見せ、そのぶん塗料を減らしたり薄めたりする。坪数が同じでも外壁の面積は家ごとに違う。");
  wave(s, "塗装面積を小さくする会社がいるので要注意");
  txt(s, [br("見積の面積を実際より小さく書いて安く見せ、"), n("そのぶん塗料を減らしたり、薄めたりする会社があります。", { color: C.accent6, bold: true })], 0.55, 1.85, 9.8, 0.9, { fontSize: 18 });
  ph(s, 0.8, 2.9, 4.2, 2.4, "【図】\n坪数が違っても外壁の面積は同じ／\n坪数が同じでも外壁の面積は違う");
  ph(s, 5.6, 2.9, 4.2, 2.4, "【写真】\n形の違う2軒の家（窓・バルコニー・高さ）");
  s.addText("窓の大きさ・数　バルコニーの有・無　軒高", { x: 0.5, y: 5.4, w: 9.833, h: 0.55, fontSize: 24, color: C.accent6, align: "center", margin: 0, isTextBox: true });
  txt(s, "→見積書の㎡が、実測や図面にもとづいているか確認しましょう", 0.55, 6.05, 9.8, 0.5, { fontSize: 18, bold: true, align: "center" });
}
{
  const s = content(T4, "納得のいく見積・提案。悪い例：一式ばかり。良い例：部位ごと／㎡などの単位／メーカー名と製品名／塗る回数。見本は塗装工事用に作る。");
  wave(s, "納得のいく見積・提案");
  s.addText("悪い例", { shape: pres.ShapeType.roundRect, rectRadius: 0.1, x: 0.5, y: 1.85, w: 1.2, h: 0.6, fill: { color: "EDEAE5" }, color: C.text1, fontSize: 18, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
  txt(s, "「一式」ばかりで中身がわからない", 1.8, 1.9, 3.65, 0.6, { fontSize: 14, bold: true });
  ph(s, 0.5, 2.6, 4.75, 3.4, "【見本】\n一式ばかりの見積書\n（塗装工事用に作成）");
  s.addText("良い例", { shape: pres.ShapeType.roundRect, rectRadius: 0.1, x: 5.6, y: 1.85, w: 1.2, h: 0.6, fill: { color: "4A90E2" }, color: C.background1, fontSize: 18, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
  txt(s, "部位ごと／㎡などの単位／メーカー名と製品名／塗る回数", 6.95, 1.85, 3.5, 0.7, { fontSize: 13, bold: true });
  ph(s, 5.6, 2.6, 4.75, 3.4, "【見本】\n詳しい見積書\n（当社の見積書）");
  txt(s, "見積の内容に注目せずに、工事金額だけに目を奪われている人ほど不良工事に遭遇します。", 0.5, 6.15, 9.9, 0.55, { fontSize: 15, bold: true, align: "center" });
}
{
  const s = content(T4, "地域密着・会社の歴史。「地域密着」でも創業数年の会社もある。歴史と、家からの距離（アフターの早さ）で見る。当社は創業120年以上。");
  wave(s, "地域密着・会社の歴史");
  txt(s, [br("「地域密着」と言っても、創業して数年の会社もあります。"), br("会社の歴史と、家からの距離（アフターの早さ）で見ましょう。"), n("")], 0.55, 1.85, 9.8, 1.2, { fontSize: 18 });
  s.addText([br("当社は", { fontSize: 22 }), br("創業120年以上", { fontSize: 38, color: C.accent6 }), n("地域の皆様と一緒に歩んできました", { fontSize: 18 })], { x: 0.55, y: 3.1, w: 5.0, h: 3.3, bold: true, color: C.text1, align: "center", valign: "middle", fill: { color: C.background2 }, margin: 0.1, isTextBox: true });
  ph(s, 5.85, 3.1, 4.45, 3.3, "【写真】\n創業当時・現在の会社の写真／対応エリアの地図");
}
{
  const s = content(T4, "ショールームがある：所在地がはっきりしている会社は手抜きや見積以上の請求をしにくい。相談しやすく、色見本や塗料の実物を確かめられる。");
  wave(s, "ショールームがある");
  txt(s, [br("ショールームがあるとなぜ良いのか？", { bold: true, fontSize: 20 }), br(""), br("・会社の所在地がはっきりしている"), br("　→手抜き工事や見積以上の請求をしにくい"), br("・気軽に相談できる"), n("・色見本や塗料の実物を確かめられる")], 0.55, 1.85, 4.8, 4.5, { fontSize: 17 });
  ph(s, 5.6, 1.85, 4.7, 3.0, "【写真】\nショールームの外観・内観");
  ph(s, 5.6, 5.0, 4.7, 1.6, "【地図】\nショールームの場所・営業時間");
}
{
  const s = content(T4, "安心の保証制度（連帯工事保証）。保証書は書面で。誰が保証するのかを確認。元請けと施工した会社が連帯して責任を負う。責任が重いため多くの会社はリスクを嫌って出していない。定期点検も確認。");
  wave(s, "安心の保証制度（連帯工事保証）");
  txt(s, [br("・保証書は口約束ではなく、書面で発行されているか"), br("・「誰が保証するのか」を確認しましょう", { color: C.accent6, bold: true }), n("・定期点検（何年目に来てくれるか）も確認しましょう")], 0.55, 1.85, 9.8, 1.2, { fontSize: 17 });
  s.addText("お客様", { x: 4.15, y: 3.15, w: 2.5, h: 0.65, fill: { color: C.background2 }, color: C.text1, fontSize: 18, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addText("元請け", { x: 1.6, y: 4.6, w: 2.6, h: 0.75, fill: { color: C.accent1 }, color: C.background1, fontSize: 18, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addText("実際に施工した会社", { x: 6.6, y: 4.6, w: 2.6, h: 0.75, fill: { color: C.accent1 }, color: C.background1, fontSize: 16, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addText("連帯して責任を負う", { x: 4.05, y: 4.3, w: 2.7, h: 0.45, fontSize: 14, bold: true, color: C.accent6, align: "center", margin: 0, isTextBox: true });
  s.addShape(pres.ShapeType.line, { x: 4.2, y: 4.97, w: 2.4, h: 0, line: { color: C.accent6, width: 2, beginArrowType: "triangle", endArrowType: "triangle" } });
  s.addShape(pres.ShapeType.line, { x: 2.9, y: 3.8, w: 1.25, h: 0.8, flipV: true, line: { color: "7F7F7F", width: 1.5, beginArrowType: "triangle" } });
  s.addShape(pres.ShapeType.line, { x: 6.65, y: 3.8, w: 1.25, h: 0.8, line: { color: "7F7F7F", width: 1.5, beginArrowType: "triangle" } });
  txt(s, [br("→どちらかが対応できなくても、お客様は守られます", { bold: true }), n("責任が重いため、多くの会社はリスクを嫌って出していません")], 0.55, 5.6, 9.8, 1.0, { fontSize: 16, align: "center" });
}
{
  const s = content(T4, "経営基盤はHPで確認しましょう。10年保証でも、10年後に会社がなければ意味がない。");
  wave(s, "経営基盤はHPで確認しましょう");
  s.addText("10年保証でも、10年後に会社がなければ意味がありません", { x: 0.55, y: 1.85, w: 9.8, h: 0.55, fontSize: 20, bold: true, color: C.accent6, margin: 0, isTextBox: true });
  s.addText([br("□創業年数"), br("□売上の規模"), br("□決算・利益の公開"), br("□施工実績の数"), br("□資格者"), n("□所在地とショールーム")], { x: 0.55, y: 2.6, w: 4.6, h: 3.9, fontSize: 20, bold: true, color: C.text1, line: { color: "1F3864", width: 2 }, margin: 0.2, valign: "middle", paraSpaceAfter: 8, isTextBox: true });
  ph(s, 5.5, 2.6, 4.8, 3.9, "【画面】\n当社ホームページの会社概要ページ");
}
{
  const s = content(T4, "法令順守：契約書（建設業法で書面が義務）／訪問販売は8日以内ならクーリング・オフ／工程表／保証書／本足場・アスベストの事前調査。");
  wave(s, "法令を守った契約書・工程表・保証書を使っている");
  const boxes = [["契約書", "工事の大小にかかわらず、書面での契約が義務（建設業法）"], ["工程表", "何日に何をするかがわかり、手抜きを防げる"], ["保証書", "保証の範囲と年数を確認"], ["クーリング・オフ", "訪問販売での契約は、8日以内なら解約できる"]];
  boxes.forEach(([h, d], i) => {
    const x = 0.5 + (i % 2) * 5.0, y = 1.9 + Math.floor(i / 2) * 1.55;
    s.addText(h, { x, y, w: 4.7, h: 0.5, fill: { color: C.accent1 }, color: C.background1, fontSize: 18, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(d, { x, y: y + 0.5, w: 4.7, h: 0.85, fill: { color: "FFF7F0" }, color: C.text1, fontSize: 14, valign: "middle", margin: 0.1, isTextBox: true });
  });
  txt(s, "※足場（本足場）やアスベストの事前調査も法令で決まっています", 0.5, 5.1, 9.9, 0.4, { fontSize: 14, bold: true });
  ph(s, 0.5, 5.6, 9.9, 1.05, "【写真】当社の契約書・工程表・保証書");
}
{
  const s = content(T4, "良い会社のチェックリスト（自社施工は外した版）。");
  wave(s, "良い会社ってどんな会社？");
  const items = ["健康診断を丁寧にしてくれる（写真で説明）", "見積が詳しい（㎡・製品名・塗る回数）", "塗装面積が正確", "ショールームがある", "会社の歴史がある・地域密着", "連帯工事保証・定期点検がある", "HPで経営基盤がわかる", "施工管理の責任者がいる（丸投げしない）", "近隣挨拶・職人のマナーがしっかりしている", "契約書・工程表・保証書がそろっている"];
  items.forEach((t, i) => s.addText("□ " + t, { x: 0.5 + (i % 2) * 5.0, y: 1.95 + Math.floor(i / 2) * 0.88, w: 4.8, h: 0.7, fontSize: 15, bold: true, color: C.text1, fill: { color: "F2F2F2" }, valign: "middle", margin: 0.12, isTextBox: true }));
}
{
  const s = content(T4, "さいごに：塗装工事を行う理由は家を守るため。資産価値を長持ちさせるために「良い業者」を選びましょう。");
  wave(s, "さいごに");
  s.addText([br("塗装工事を行う理由は家を守るためです！"), br("皆様のお住まいの資産価値を長持ちさせる"), { text: "ためにも", options: {} }, { text: "「良い業者」", options: { color: C.accent6 } }, br("を選定し、"), n("工事を行うことが大切です！")], { x: 0.4, y: 2.2, w: 10.033, h: 3.4, fontSize: 30, bold: true, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true });
}

// =====================================================================
section("まとめ・質疑応答");
divider("05", "まとめ・質疑応答");
{
  const s = content("まとめ", "本日のまとめ：家は必ず劣化する（北陸は特に）／3つのポイント／業者は価格ではなく見積・保証・会社の中身で選ぶ。");
  wave(s, "本日のまとめ");
  const rows = [["1", "家は必ず劣化します（北陸は特に）"], ["2", "失敗しない工事3つのポイント\n流れを知って逆算／劣化度合いで工事を決める／家に合った塗料を選ぶ"], ["3", "業者は価格ではなく、見積・保証・会社の中身で選びましょう"]];
  rows.forEach(([no, t], i) => {
    const y = 1.95 + i * 1.5;
    s.addText(no, { shape: pres.ShapeType.ellipse, x: 0.7, y: y + 0.15, w: 0.9, h: 0.9, fill: { color: C.accent1 }, color: C.background1, fontSize: 28, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(t, { x: 1.85, y, w: 8.4, h: 1.2, fontSize: 20, bold: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  });
}
{
  const s = content("さいごに", "まずは外壁・屋根の健康診断を。無料／機械を使わない点検（目視・触診・写真で記録）／写真付きの診断結果でご説明／約1時間30分／申込み方法。");
  s.addText([br("まず、一度は外壁・屋根の"), n("“健康診断”をしてみて下さい。")], { x: 0.4, y: 1.5, w: 10.033, h: 1.9, fontSize: 32, bold: true, color: "595959", align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addText([br("無料の健康診断", { bold: true, fontSize: 20, color: C.accent6 }), br("・機械を使わない点検（目視・触診・写真で記録）"), br("・写真付きの診断結果でご説明"), br("・所要時間 約1時間30分"), n("・お申込み：申込用紙／電話／QRコード　ショールームでもご相談いただけます")], { x: 0.6, y: 3.7, w: 7.2, h: 2.9, fontSize: 16, color: C.text1, line: { color: C.accent1, width: 1.5 }, margin: 0.15, valign: "middle", isTextBox: true });
  ph(s, 8.1, 3.7, 2.2, 2.9, "【QRコード】\n申込み・電話番号");
}
{
  const s = content("さいごに", "お礼。質疑応答へ。");
  s.addText([br("本勉強会が少しでもお役に立てば幸いです。"), br(""), br("ご清聴いただき、"), n("誠にありがとうございます。")], { x: 0.4, y: 2.0, w: 10.033, h: 3.4, fontSize: 30, bold: true, color: C.accent4, align: "center", valign: "middle", margin: 0, isTextBox: true });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
