// 外装工事ガイドブック（A4縦・12ページ）を PowerPoint で作る
// 使い方: NODE_PATH=<pptxgenjsのnode_modules> node build_pptx.js [出力先.pptx]
const pptxgen = require("pptxgenjs");
const path = require("path");

const A = (f) => path.join(__dirname, "assets", f);
const OUT = process.argv[2] || path.join(__dirname, "guidebook.pptx");
const mm = (v) => v / 25.4;

// アプローチブック（2025年1月版）の配色：緑の見出し帯・山吹色の帯・赤の強調・メイリオ
const GREEN = "00B050", GREEN_D = "00873E", YELLOW = "FED86C";
const NAVY = GREEN_D, NAVY2 = GREEN, NAVY_SOFT = "F2F2F2";   // 見出し文字・表の見出し・カード地
const ORANGE = GREEN, ORANGE_SOFT = "FFF6D6";                // 番号・ラベル・豆知識の地（薄い山吹）
const NG = "C00000", NG_SOFT = "FDECEC", OK = GREEN_D, OK_SOFT = "EAF6EE";
const INK = "2D2D2D", MUTED = "595959", LINE = "D9D9D9", PH_BG = "F4F4F4";
const FONT = "メイリオ", FONT_UI = "Meiryo UI";
const SLOGAN = "地域の屋根・外壁を安心・安全に塗装して、長持ちをさせる為に全力を尽くします";
// 本文は緑の帯の下（y=62mm）から。旧レイアウト（y=58mm 起点）の座標を4mm下げる
let SHIFT = true;
const Y = (y) => (SHIFT && y >= 55 && y < 280 ? y + 4 : y);

const pres = new pptxgen();
pres.defineLayout({ name: "A4_PORTRAIT", width: mm(210), height: mm(297) });
pres.layout = "A4_PORTRAIT";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = "失敗しない外装工事ガイドブック";
pres.company = "株式会社山岸";

const L = 14, W = 182; // 左余白と本文の幅（mm）

// ---------- 基本部品 ----------
function T(s, text, x, y, w, h, o = {}) {
  s.addText(text, Object.assign({ x: mm(x), y: mm(Y(y)), w: mm(w), h: mm(h), fontFace: FONT, fontSize: 10, color: INK, margin: 0, valign: "top", isTextBox: true }, o));
}
function R(s, x, y, w, h, fill, o = {}) {
  s.addShape(o.round ? pres.ShapeType.roundRect : pres.ShapeType.rect, Object.assign({ x: mm(x), y: mm(Y(y)), w: mm(w), h: mm(h), fill: { color: fill }, line: o.line || { type: "none" } }, o.round ? { rectRadius: mm(o.round) } : {}));
}
function ph(s, label, x, y, w, h) {
  T(s, label, x, y, w, h, { fontSize: 8.5, color: "7A8594", align: "center", valign: "middle", fill: { color: PH_BG }, line: { color: "A9B3C1", width: 0.75, dashType: "dash" }, objectName: "写真枠" });
}
const b = (t, o = {}) => ({ text: t, options: Object.assign({ bold: true }, o) });
const n = (t, o = {}) => ({ text: t, options: o });
const nl = (t, o = {}) => ({ text: t, options: Object.assign({ breakLine: true }, o) });

function page(no) {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  SHIFT = !!no;
  if (no) {
    T(s, SLOGAN, 10, 284, 170, 7, { fontFace: "HGP教科書体", fontSize: 11, color: INK, valign: "middle" });
    T(s, String(no), 182, 284, 14, 7, { fontSize: 11, bold: true, color: GREEN_D, align: "right", valign: "middle" });
  }
  return s;
}
function head(s, scene, title, sub) {
  // 上：山吹色の帯（左に白い縦すじ）／下：緑の見出し帯
  R(s, 0, 0, 210, 25, YELLOW);
  R(s, 5, 0, 3, 25, "FFFFFF");
  R(s, 0, 30, 210, 27, GREEN);
  T(s, scene, L, 5, Math.max(22, scene.length * 4.6 + 8), 7.5, { fontFace: FONT_UI, fontSize: 11, bold: true, color: "FFFFFF", fill: { color: GREEN }, align: "center", valign: "middle" });
  if (sub) T(s, sub, L, 14.5, W, 7, { fontSize: 10.5, bold: true, color: INK, valign: "middle" });
  T(s, title, L, 31.5, W, 24, { fontSize: 21, bold: true, color: "FFFFFF", valign: "middle", lineSpacingMultiple: 1.0 });
}
function sec(s, text, y, x = L, w = W) {
  T(s, text, x, y, Math.min(w, text.length * 4.75 + 8), 7, { fontFace: FONT_UI, fontSize: 12.5, bold: true, color: "FFFFFF", fill: { color: GREEN }, valign: "middle", margin: [0, mm(3), 0, mm(3)] });
}
// 「見るところ」カード
function look(s, items, y, h, cols = items.length, x = L, w = W) {
  const gap = 3, cw = (w - gap * (cols - 1)) / cols;
  items.forEach(([t, d], i) => {
    const cx = x + (i % cols) * (cw + gap), cy = y + Math.floor(i / cols) * (h + gap);
    R(s, cx, cy, cw, h, NAVY_SOFT, { round: 1.5 });
    T(s, [nl(t, { bold: true, color: NAVY, fontSize: 11 }), n(d, { fontSize: 9.3 })], cx + 3, cy + 2.5, cw - 6, h - 4, { paraSpaceAfter: 2, lineSpacingMultiple: 1.1 });
  });
}
// ×要注意／○良い業者
function ngok(s, y, h, ng, ok, titles = ["要注意", "良い業者"], kinds = ["ng", "ok"]) {
  const cw = (W - 4) / 2;
  [ng, ok].forEach((items, i) => {
    const isNg = kinds[i] === "ng";
    const x = L + i * (cw + 4), col = isNg ? NG : OK;
    R(s, x, y, cw, h, isNg ? NG_SOFT : OK_SOFT, { round: 1.5, line: { color: col, width: 1.25 } });
    s.addText(isNg ? "×" : "○", { shape: pres.ShapeType.ellipse, x: mm(x + 4), y: mm(Y(y + 3.2)), w: mm(7), h: mm(7), fill: { color: col }, color: "FFFFFF", fontFace: FONT, fontSize: 12, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
    T(s, titles[i], x + 13, y + 3.2, cw - 16, 7, { fontSize: 11.5, bold: true, color: col, valign: "middle" });
    T(s, items.map((t, k) => ({ text: "・" + t, options: { breakLine: k < items.length - 1 } })), x + 4, y + 12, cw - 8, h - 14, { fontSize: 9.6, paraSpaceAfter: 2, lineSpacingMultiple: 1.1 });
  });
}
// チェック欄
function check(s, title, note, items, y, cols = 1, rowH = 7.2) {
  const rows = Math.ceil(items.length / cols), h = 9 + rows * rowH + 3;
  R(s, L, y, W, h, "FFFFFF", { round: 1.5, line: { color: NAVY, width: 1.5 } });
  R(s, L, y, W, 8.5, NAVY);
  T(s, title, L + 4, y, 100, 8.5, { fontSize: 11, bold: true, color: "FFFFFF", valign: "middle" });
  if (note) T(s, note, L + W - 84, y, 80, 8.5, { fontSize: 8.5, bold: true, color: "DCE6F5", align: "right", valign: "middle" });
  const cw = (W - 8 - (cols - 1) * 6) / cols;
  items.forEach((t, i) => {
    const c = Math.floor(i / rows), r = i % rows;
    const x = L + 4 + c * (cw + 6), yy = y + 10.5 + r * rowH;
    R(s, x, yy + 0.9, 4.2, 4.2, "FFFFFF", { line: { color: NAVY, width: 1 } });
    T(s, t, x + 6.5, yy, cw - 7, rowH - 0.5, { fontSize: 9.6, valign: "top", lineSpacingMultiple: 1.0 });
  });
  return h;
}
// 豆知識などのオレンジの囲み
function tip(s, label, x, y, w, h, body) {
  R(s, x, y, w, h, ORANGE_SOFT, { round: 1.5 });
  T(s, label, x + 4, y + 3, label.length * 3.4 + 6, 5, { fontSize: 8.5, bold: true, color: "FFFFFF", fill: { color: ORANGE }, align: "center", valign: "middle", shape: pres.ShapeType.roundRect, rectRadius: mm(2.5) });
  if (body) T(s, body, x + 4, y + 10, w - 8, h - 12, { fontSize: 9.6, lineSpacingMultiple: 1.15, paraSpaceAfter: 2 });
}

// =============================== P1 表紙 ===============================
{
  const s = page(0);
  s.addImage({ path: A("cover.png"), x: 0, y: 0, w: mm(210), h: mm(297), objectName: "表紙画像" });
}

// =============================== P2 はじめに ===============================
{
  const s = page(2);
  head(s, "はじめに", "業者選びで、外装工事の\n9割が決まります");
  T(s, [nl("外壁や屋根の工事は、一生に何度もない大きな買い物です。"), n("ところが、工事の失敗の多くは「悪徳業者」ではなく、"), b("普通の業者の診断ミスや知識不足", { color: NG }), n("から起きています。")], L, 58, W, 15, { fontSize: 11, bold: true, lineSpacingMultiple: 1.25 });
  T(s, [n("このガイドは、業者と出会ってから工事が終わるまでの"), b("7つの場面", { color: NAVY }), n("ごとに、「ここを見れば良い業者かどうかわかる」ポイントをまとめました。見積書と並べて、チェックしながらお使いください。")], L, 75, W, 15, { fontSize: 11, bold: true, lineSpacingMultiple: 1.25 });
  sec(s, "このガイドの道のり", 94);
  const road = [["場面 0", "問い合わせる前に", "わが家の状態を知っておく", 3], ["場面 1", "会社を探す", "会社の中身を確かめる", 4], ["場面 2", "健康診断", "家をどこまで見てくれるか", 5], ["場面 3", "提案を聞く・塗料を選ぶ", "わが家に合った工事か", 6], ["場面 4", "見積書を受け取る", "ここを見れば中身がわかる", 8], ["場面 5", "契約する", "書類と保証を確かめる", 9], ["場面 6", "工事中", "手順どおりに進んでいるか", 10], ["場面 7", "工事のあと", "長く付き合える会社か", 11]];
  road.forEach(([sc, t, d, p], i) => {
    const y = 103 + i * 11;
    T(s, sc, L, y, 22, 8.5, { fontSize: 10, bold: true, color: "FFFFFF", fill: { color: i % 2 ? NAVY2 : NAVY }, align: "center", valign: "middle", shape: pres.ShapeType.roundRect, rectRadius: mm(1.2) });
    T(s, [b(t, { fontSize: 11.5 }), n("　" + d, { fontSize: 9.5, color: MUTED })], L + 26, y, W - 46, 8.5, { valign: "middle" });
    T(s, "P." + p, L + W - 18, y, 18, 8.5, { fontSize: 10.5, bold: true, color: ORANGE, align: "right", valign: "middle" });
    s.addShape(pres.ShapeType.line, { x: mm(L + 26), y: mm(Y(y + 9.6)), w: mm(W - 26), h: 0, line: { color: LINE, width: 0.75, dashType: "dash" } });
  });
  tip(s, "各ページの見方", L, 197, W, 24);
  s.addImage({ path: A("mascot.png"), x: mm(160), y: mm(Y(230)), w: mm(34), h: mm(34.5), objectName: "キャラクター" });
  [["見るところ", "この場面で確認すること", NAVY], ["× 要注意", "こんな業者は気をつけて", NG], ["○ 良い業者", "信頼できる業者の対応", OK], ["豆知識", "知っておくと役立つ情報", ORANGE]].forEach(([h, d, c], i) => {
    T(s, [nl(h, { bold: true, color: c }), n(d)], L + 4 + i * 44.5, 207, 42, 12, { fontSize: 9.3 });
  });
}

// =============================== P3 場面0 ===============================
{
  const s = page(3);
  head(s, "場面 0", "問い合わせる前に\nわが家の状態を知っておく", "自分の家のことを知っていれば、業者の説明が正しいかどうか判断できます");
  const h = check(s, "わが家のセルフチェック", "当てはまるものに✓を付けましょう", ["外壁をこすると白い粉がつく（チョーキング）", "外壁にひび割れがある", "目地（シーリング）が割れている・痩せている", "外壁の板が反っている、表面がはがれている", "北側などにカビ・コケが生えている", "屋根の色があせている、錆が出ている", "雨樋が傾いている、詰まっている", "前回の塗装から10年以上たっている"], 58, 2, 10);
  T(s, "✓が1つでもあれば、一度プロの「健康診断」を受けておくと安心です。", L, 58 + h + 1.5, W, 5, { fontSize: 8.5, color: MUTED });
  const y2 = 58 + h + 8.5;
  ph(s, "【写真】チョーキング（手に白い粉）", L, y2, 89, 55);
  ph(s, "【写真】凍害で表面がはがれた外壁", L + 93, y2, 89, 55);
  const y3 = y2 + 59;
  tip(s, "豆知識", L, y3, 89, 34, [nl("北陸の家は傷みやすい", { bold: true, fontSize: 11.5 }), n("雨や雪が多く、湿気も高い地域です。冬に凍ったりとけたりをくり返すと、外壁の表面がはがれる「凍害」も起きます。")]);
  tip(s, "豆知識", L + 93, y3, 89, 34, [nl("塗り替えの目安", { bold: true, fontSize: 11.5 }), n("初めての塗り替えは"), b("築8〜10年"), n("。2回目以降は前回の塗料の耐久年数しだいです。ただし年数より「家の状態」で判断しましょう。")]);
  const y4 = y3 + 38;
  R(s, L, y4, W, 33, "FFFFFF", { round: 1.5, line: { color: NG, width: 1.5 } });
  T(s, "こんな訪問には要注意", L + 4, y4 + 3, 100, 6, { fontSize: 11, bold: true, color: NG });
  T(s, [nl("▶ 突然やってきて「屋根がずれていますよ」「近くで工事をしているので見てあげます」"), nl("▶ 「火災保険を使えば無料で直せます」（経年劣化は火災保険の対象外です）"), n("▶ 「今日決めてくれたら大幅に値引きします」とその場で契約を迫る")], L + 4, y4 + 10.5, W - 8, 21, { fontSize: 9.6, paraSpaceAfter: 3 });
}

// =============================== P4 場面1 ===============================
{
  const s = page(4);
  head(s, "場面 1", "会社を探す\nまずは「会社の中身」を確かめる", "10年保証でも、10年後に会社がなければ意味がありません");
  sec(s, "見るところ", 58);
  look(s, [["会社の歴史", "「地域密着」と言っても、創業して数年の会社もあります。どれだけ長くこの地域で仕事をしてきたかを確認しましょう。"], ["家からの距離", "近い会社ほど、困ったときにすぐ来てもらえます。アフターの早さに差が出ます。"], ["ショールームがあるか", "所在地がはっきりしている会社は逃げられません。色見本や塗料の実物を見ながら相談できます。"], ["ホームページの会社情報", "創業年数・売上の規模・決算の公開・施工実績・資格者・所在地を確認しましょう。"]], 67, 27, 2);
  sec(s, "こんな業者は要注意／良い業者はこうする", 128);
  ngok(s, 137, 42, ["ホームページや会社案内に、住所や会社の情報がほとんどない", "チラシや訪問だけで、事務所や店舗の場所がわからない", "施工事例が少ない、他社の写真を使っている"], ["ショールームで気軽に相談できる", "会社の歴史や実績、スタッフの顔がわかる", "地元で長く営業し、施工した家の近くに事務所がある"]);
  check(s, "場面1のチェック", "", ["会社の所在地がはっきりしていて、ショールームや事務所がある", "地域で長く営業している（創業年数を確認した）", "ホームページで会社の情報や施工実績を確認できた"], 187);
}

// =============================== P5 場面2 ===============================
{
  const s = page(5);
  head(s, "場面 2", "健康診断（現地調査）\n家をどこまで見てくれるか", "工事の失敗は、最初の診断ミスから始まります");
  sec(s, "見るところ", 58);
  look(s, [["目で見る", "ひび割れ・はがれ・色あせ・カビやコケ・目地・錆・屋根"], ["手で触る", "チョーキング（白い粉）・外壁の浮きや反り"], ["写真で説明する", "撮った写真を見せながら、状態と原因を説明してくれるか"]], 67, 24);
  sec(s, "こんな業者は要注意／良い業者はこうする", 97);
  ngok(s, 106, 36, ["車から降りて外から少し眺めるだけ", "口頭の説明だけで、すぐに見積もりを出す", "図面や坪数だけで金額を決める"], ["家のまわりを確認し、劣化の場所を写真に撮る", "写真付きの診断結果で、原因まで説明する", "劣化に合った工事方法と塗料を提案する"]);
  tip(s, "豆知識", L, 149, W, 88);
  T(s, "傷みすぎると、塗装ではもう直せません", L + 4, 158, W - 8, 6, { fontSize: 11.5, bold: true });
  const cw = (W - 12) / 2;
  T(s, [nl("まだ塗装で直せる", { bold: true, color: NAVY }), nl("・色あせ・チョーキング"), nl("・細かいひび割れ"), n("・目地の痩せ・割れ")], L + 4, 166, cw, 22, { fontSize: 9.6 });
  T(s, [nl("塗装ではもう直せない", { bold: true, color: NG }), nl("・外壁材そのものがはがれ落ちている"), nl("・カビや藻で傷み、塗料が密着しない"), n("・水を吸って外壁が変形している")], L + 8 + cw, 166, cw, 22, { fontSize: 9.6 });
  ph(s, "【写真】塗装で直せる劣化", L + 4, 192, cw, 40);
  ph(s, "【写真】塗装できないほど傷んだ外壁", L + 8 + cw, 192, cw, 40);
}

// =============================== P6 場面3 ===============================
{
  const s = page(6);
  head(s, "場面 3", "提案を聞く\nわが家に合った工事か", "「何をするか」の前に「あと何年住みたいか」を聞いてくれる業者を");
  sec(s, "見るところ", 58);
  look(s, [["暮らしの予定", "「あと何年住みたいか」「将来どうしたいか」を聞いてくれるか"], ["工事の理由", "塗装か張替えか、なぜその工事なのかを説明できるか"], ["塗料の説明", "メーカー名と商品名、その塗料を選んだ理由を言えるか"]], 67, 24);
  tip(s, "豆知識", L, 96, W, 64);
  T(s, "放っておくと、結局高くつきます", L + 4, 104.5, W - 8, 6, { fontSize: 11.5, bold: true });
  const card = (x, head_, price, unit, note, col) => {
    R(s, x, 113, 80, 30, "FFFFFF", { round: 1.5, line: { color: col, width: 1.5 } });
    T(s, head_, x, 115, 80, 5, { fontSize: 10, bold: true, color: col, align: "center" });
    T(s, [n(price, { fontSize: 28, bold: true }), n(unit, { fontSize: 13, bold: true })], x, 121, 80, 12, { color: col, align: "center", valign: "middle" });
    T(s, note, x, 134.5, 80, 5, { fontSize: 8.5, color: MUTED, align: "center" });
  };
  card(L + 4, "早めに塗装した場合", "54.8", "万円", "塗装パック（プレミアムシリコン）", NAVY);
  T(s, "vs", L + 86, 123, 10, 8, { fontSize: 12, bold: true, color: MUTED, align: "center", valign: "middle" });
  card(L + 98, "傷みすぎて張替えた場合", "148", "万円〜", "外壁サイディングパック", NG);
  T(s, "その差 約2.7倍！", L, 145, W, 7, { fontSize: 13, bold: true, color: NG, align: "center" });
  T(s, "※税込・外壁100㎡・足場代込の当社参考価格。家の状態により異なります。", L, 152.5, W, 5, { fontSize: 8.5, color: MUTED, align: "center" });
  tip(s, "豆知識", L, 165, W, 72);
  T(s, [nl("塗料は「メーカー」で選ぶ時代", { bold: true, fontSize: 11.5 }), n("同じ「シリコン」「ラジカル」という名前でも、メーカーや商品によって品質も実績も違います。", { fontSize: 9.6 })], L + 4, 173, W - 8, 12);
  const hd = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY } } });
  const hl = (t) => ({ text: t, options: { bold: true, fill: { color: "FFE2BF" }, color: NAVY } });
  s.addTable([
    [hd("メーカー"), hd("創業"), hd("特徴")],
    ["日本ペイント", "1881年", "日本で最初の塗料メーカー。自動車・工業用も手がける総合メーカー"],
    ["関西ペイント", "1918年", "自動車・工業用も手がける総合メーカー"],
    [hl("エスケー化研"), hl("1955年"), hl("建築用の仕上塗材が専門。国内シェア53％でNo.1")],
    ["菊水化学工業", "1959年", "建築用の仕上塗材のメーカー"],
    ["オリジナル塗料", "—", "塗装店の自社ブランド。性能の根拠がわかりにくく、比べにくい"],
  ], { x: mm(L + 4), y: mm(Y(187)), w: mm(W - 8), colW: [mm(36), mm(22), mm(W - 8 - 58)], fontFace: FONT, fontSize: 8.8, color: INK, fill: { color: "FFFFFF" }, border: { type: "solid", pt: 0.5, color: LINE }, rowH: mm(7.3), valign: "middle", margin: [0, mm(1.5), 0, mm(1.5)] });
}

// =============================== P7 場面3つづき ===============================
{
  const s = page(7);
  head(s, "場面 3 つづき", "塗料を選ぶ\n当社がエスケー化研を選ぶ理由", "中身がわかる塗料だから、お客様も安心して選べます");
  const reasons = [
    ["ものがいい", "家の外壁や屋根に使う「建築用の塗料」が専門のメーカーです。プレミアムシリコン（ラジカル制御・期待耐用年数15年）、クールテクトSi（遮熱・汚れにくい）など、性能を試験データで確かめられます。"],
    ["トラブルが少ない", "全国の多くの現場で長年使われてきた実績があります。製品の情報がすべて公開されているので、お客様ご自身でも調べられます。どの塗装店でも同じ品質の材料が手に入ります。"],
    ["日本一", "建築仕上塗材の国内シェアNo.1。六本木ヒルズや甲子園球場など、多くの建物で使われています。"],
  ];
  reasons.forEach(([h, d], i) => {
    const y = 58 + i * 31;
    T(s, [nl("理由", { fontSize: 12 }), n(String(i + 1), { fontSize: 24 })], L, y, 30, 27, { bold: true, color: "FFFFFF", fill: { color: ORANGE }, align: "center", valign: "middle", shape: pres.ShapeType.roundRect, rectRadius: mm(1.5) });
    R(s, L + 34, y, W - 34, 27, NAVY_SOFT, { round: 1.5 });
    T(s, [nl(h, { bold: true, color: NAVY, fontSize: 12 }), n(d, { fontSize: 9.6 })], L + 38, y + 3, i === 2 ? W - 80 : W - 42, 22, { lineSpacingMultiple: 1.12 });
    if (i === 2) T(s, [nl("国内シェア", { fontSize: 8.5, color: MUTED, bold: false }), n("53", { fontSize: 32 }), n("％", { fontSize: 15 })], L + W - 40, y + 2, 36, 23, { bold: true, color: NG, align: "center", valign: "middle" });
  });
  ngok(s, 154, 32, ["「当社オリジナル塗料」と言うが、どこのメーカーの何かわからない", "塗料の名前を聞いても「シリコンです」としか答えない"], ["メーカー名・商品名・耐久年数を説明できる", "カタログや色見本を見せてくれる"]);
  ph(s, "【写真】エスケー化研の塗料缶・色見本／当社の施工事例", L, 192, W, 72);
  T(s, "※国内シェアは日本建築仕上工業会のデータ（2024年5月）による。期待耐用年数はメーカー公表値で、保証年数ではありません。", L, 266, W, 8, { fontSize: 8, color: MUTED });
}

// =============================== P8 場面4 ===============================
{
  const s = page(8);
  head(s, "場面 4", "見積書を受け取る\nここを見れば中身がわかる", "金額だけを見ず、「何をどれだけやるか」を見ましょう");
  sec(s, "見積書のここを見る", 58);
  const hd = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY } } });
  const hl = (t) => ({ text: t, options: { bold: true, fill: { color: ORANGE_SOFT }, color: NAVY } });
  s.addTable([
    [hd("項目"), hd("数量"), hd("単位")],
    ["足場（本足場）・飛散防止ネット", "〇〇", "㎡"],
    ["高圧洗浄", "〇〇", "㎡"],
    ["シーリング打ち替え（サイディング目地）", "〇〇", "m"],
    ["外壁 下塗り／エスケー化研 〇〇〇", "〇〇", "㎡"],
    [hl("外壁 中塗り・上塗り／エスケー化研 プレミアムシリコン"), hl("〇〇"), hl("㎡")],
    ["付帯部（軒天・破風・雨樋 など部位ごと）", "〇〇", "m・㎡"],
    ["屋根 さび止め・下塗り・上塗り／商品名", "〇〇", "㎡"],
  ], { x: mm(L), y: mm(Y(67)), w: mm(122), colW: [mm(92), mm(15), mm(15)], fontFace: FONT, fontSize: 8.8, color: INK, border: { type: "solid", pt: 0.5, color: LINE }, rowH: mm(7.6), valign: "middle", margin: [0, mm(1.5), 0, mm(1.5)] });
  [["① 部位ごとに分かれている", "外壁・屋根・付帯部が別々に書いてある"], ["② ㎡・mなど単位がある", "「一式」ばかりになっていない"], ["③ メーカー名と商品名", "どの塗料を何回塗るかがわかる"]].forEach(([h, d], i) => {
    const y = 67 + i * 20.5;
    R(s, L + 126, y, 56, 18, NAVY_SOFT, { round: 1.5 });
    T(s, [nl(h, { bold: true, color: NAVY }), n(d)], L + 129, y + 2.5, 51, 14, { fontSize: 9.3 });
  });
  sec(s, "こんな見積書は要注意", 135);
  ngok(s, 144, 38, ["「外壁塗装　一式」では、何をどれだけするのかわからない", "あとから追加費用を請求されることも"], ["面積を小さく書いて安く見せ、そのぶん塗料を減らしたり薄めたりする", "坪数が同じでも、窓・バルコニー・高さで外壁の面積は変わる"], ["「一式」ばかり", "塗装面積が小さすぎる"], ["ng", "ng"]);
  check(s, "場面4のチェック", "見積書を横に置いて確認しましょう", ["部位ごとに項目が分かれていて、「一式」ばかりではない", "外壁の面積（㎡）が、実際に測った数字か図面にもとづいている", "塗料のメーカー名・商品名・塗る回数が書いてある", "足場（本足場）・高圧洗浄・シーリングが含まれている"], 189);
}

// =============================== P9 場面5 ===============================
{
  const s = page(9);
  head(s, "場面 5", "契約する\n書類と保証を確かめる", "保証は「あるかどうか」より「誰が保証するのか」が大切です");
  sec(s, "そろっているか確認したい3つの書類", 58);
  look(s, [["契約書", "工事の大小にかかわらず、書面での契約が法律で決まっています"], ["工程表", "何日に何をするかがわかり、手抜きを防げます"], ["保証書", "保証の範囲と年数、定期点検の時期を確認しましょう"]], 67, 24);
  sec(s, "山岸の連帯保証は、エスケー化研を使うからこそ", 97);
  T(s, [nl("株式会社山岸", { fontSize: 12 }), n("施工", { fontSize: 9 })], L, 106, 52, 15, { bold: true, color: "FFFFFF", fill: { color: ORANGE }, align: "center", valign: "middle", shape: pres.ShapeType.roundRect, rectRadius: mm(1.5) });
  T(s, "×", L + 52, 106, 10, 15, { fontSize: 18, bold: true, color: NG, align: "center", valign: "middle" });
  T(s, [nl("エスケー化研", { fontSize: 12 }), n("塗料メーカー", { fontSize: 9 })], L + 62, 106, 52, 15, { bold: true, color: "FFFFFF", fill: { color: NAVY }, align: "center", valign: "middle", shape: pres.ShapeType.roundRect, rectRadius: mm(1.5) });
  T(s, "▶", L + 116, 106, 12, 15, { fontSize: 15, color: MUTED, align: "center", valign: "middle" });
  T(s, "お客様", L + 130, 106, 52, 15, { fontSize: 12, bold: true, color: NAVY, fill: { color: NAVY_SOFT }, align: "center", valign: "middle", shape: pres.ShapeType.roundRect, rectRadius: mm(1.5) });
  T(s, "施工も材料も、2社が連名で責任を持ちます", L, 124, W, 7, { fontSize: 11.5, bold: true, align: "center" });
  tip(s, "なぜできるのか", L, 133, 89, 28, [nl("・出荷証明書で、本物の材料を使ったと証明できる"), n("・メーカーが決めた量・回数・乾燥時間を守って施工する")]);
  tip(s, "他ではなぜできないのか", L + 93, 133, 89, 28, [nl("・中身のわからないオリジナル塗料では、メーカーが保証に加われない"), n("・責任が重く、多くの会社は出したがらない")]);
  T(s, "保証期間：〇年／対象：〇〇〇（詳しくは担当者にお尋ねください）", L, 163, W, 5, { fontSize: 8.5, color: MUTED });
  ph(s, "【写真】当社の契約書・工程表", L, 171, 89, 54);
  ph(s, "【写真】山岸×エスケー化研の連名保証書", L + 93, 171, 89, 54);
  R(s, L, 230, W, 24, "FFFFFF", { round: 1.5, line: { color: NG, width: 1.5 } });
  T(s, "豆知識：クーリング・オフ", L + 4, 233, 100, 6, { fontSize: 11, bold: true, color: NG });
  T(s, [n("訪問販売で契約した場合、契約書面を受け取った日から"), b("8日以内"), n("なら、理由を問わず契約を解除できます（書面のほか、メールなどでも通知できます）。")], L + 4, 240, W - 8, 12, { fontSize: 9.6 });
}

// =============================== P10 場面6 ===============================
{
  const s = page(10);
  head(s, "場面 6", "工事中\n手順どおりに進んでいるか", "正しい工事には決まった順番があります。省かれていないか見てみましょう");
  const PH = { "準備": ["BFC7D3", INK], "下地づくり": ["FBC88A", INK], "塗装": [ORANGE, "FFFFFF"], "仕上げ": [NAVY, "FFFFFF"] };
  const flow = (x, title, steps) => {
    sec(s, title, 58, x, 89);
    const ROW = 8.2, TOP = 67;
    let k = 0;
    while (k < steps.length) {
      let e = k; while (e + 1 < steps.length && steps[e + 1][1] === steps[k][1]) e++;
      const [bg, fg] = PH[steps[k][1]];
      T(s, steps[k][1], x, TOP + k * ROW, 17, (e - k + 1) * ROW - 1.2, { fontSize: 8.5, bold: true, color: fg, fill: { color: bg }, align: "center", valign: "middle" });
      k = e + 1;
    }
    steps.forEach(([t], i) => {
      const y = TOP + i * ROW;
      R(s, x + 19, y, 70, ROW - 1.2, ORANGE_SOFT, { round: 1 });
      s.addText(String(i + 1), { shape: pres.ShapeType.ellipse, x: mm(x + 20.5), y: mm(Y(y + 0.9)), w: mm(5.4), h: mm(5.4), fill: { color: ORANGE }, color: "FFFFFF", fontFace: FONT, fontSize: 8.5, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
      T(s, t, x + 28, y, 60, ROW - 1.2, { fontSize: 9.6, bold: true, valign: "middle" });
    });
  };
  flow(L, "外壁工事（全10工程）", [["足場組立（本足場）", "準備"], ["ネット養生", "準備"], ["高圧洗浄", "下地づくり"], ["下地処理", "下地づくり"], ["養生", "下地づくり"], ["下塗り", "塗装"], ["中塗り", "塗装"], ["上塗り", "塗装"], ["完工チェック", "仕上げ"], ["足場解体・清掃", "仕上げ"]]);
  flow(L + 93, "屋根工事（全8工程）", [["高圧洗浄", "下地づくり"], ["鉄部下地調整", "下地づくり"], ["鉄部さび止め塗装", "下地づくり"], ["屋根下塗り", "塗装"], ["屋根中塗り", "塗装"], ["屋根上塗り", "塗装"], ["確認作業", "仕上げ"], ["清掃", "仕上げ"]]);
  sec(s, "ここが手抜きされやすい", 154);
  look(s, [["乾燥時間", "塗った後しっかり乾かさないと、数年ではがれます"], ["塗る回数", "下塗り・中塗り・上塗りの3回塗りが基本です"], ["薄めすぎ", "決められた量を守り、秤で調合しているか"]], 163, 24);
  tip(s, "当社の工事中のお約束", L, 194, W, 38);
  const pr = ["近隣へのご挨拶は当社が代行します", "施工管理の責任者がつきます（丸投げしません）", "毎日の作業内容をご報告します", "工程を写真で記録し、報告書でお渡しします", "施工中は禁煙・整理整頓を徹底します", "お茶菓子などのお気づかいは不要です"];
  pr.forEach((t, i) => T(s, "・" + t, L + 4 + (i % 2) * 89, 204 + Math.floor(i / 2) * 8, 87, 7, { fontSize: 9.6 }));
}

// =============================== P11 場面7 ===============================
{
  const s = page(11);
  head(s, "場面 7", "工事のあと\n長く付き合える会社か", "塗装は「終わってから」が本当のお付き合いの始まりです");
  look(s, [["完工チェック", "塗り残しや汚れを社内で確認"], ["立ち会い検査", "お客様と一緒に仕上がりを確認"], ["保証書", "その場で保証書をお渡し"], ["定期点検", "何年目に来てくれるか確認"]], 58, 21);
  sec(s, "業者選び 総まとめチェックシート", 85);
  T(s, "見積もりを取った会社ごとに ○・△・× を書き込んで比べてみましょう。", L, 93.5, W, 5, { fontSize: 8.5, color: MUTED });
  const hd = (t, c = NAVY) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: c }, align: "center" } });
  const cat = (t) => [{ text: t, options: { colspan: 4, bold: true, color: NAVY, fill: { color: NAVY_SOFT } } }];
  const row = (t) => [t, "", "", ""];
  s.addTable([
    [{ text: "チェック項目", options: { bold: true, color: "FFFFFF", fill: { color: NAVY } } }, { text: "ヤマキシ\nペイント", options: { bold: true, color: INK, fill: { color: YELLOW }, align: "center" } }, hd("A社"), hd("B社")],
    cat("会社（場面1）"),
    row("所在地がはっきりしていて、ショールームや事務所がある"),
    row("地域で長く営業し、HPで会社の中身がわかる"),
    cat("診断・提案（場面2・3）"),
    row("写真を見せながら、劣化の原因まで説明してくれた"),
    row("「あと何年住みたいか」を聞いて工事を提案してくれた"),
    row("塗料のメーカー名・商品名を説明できる"),
    cat("見積書（場面4）"),
    row("「一式」ばかりでなく、部位・㎡・塗る回数が書いてある"),
    row("塗装面積が実測や図面にもとづいている"),
    cat("契約・保証（場面5）"),
    row("契約書・工程表・保証書がそろっている"),
    row("メーカーとの連帯保証・定期点検がある"),
    cat("工事中・工事後（場面6・7）"),
    row("施工管理の責任者がいる（丸投げしない）"),
    row("近隣挨拶・工程写真の報告・職人のマナーがしっかりしている"),
    [{ text: "○の数", options: { bold: true, fill: { color: PH_BG } } }, { text: "", options: { fill: { color: PH_BG } } }, { text: "", options: { fill: { color: PH_BG } } }, { text: "", options: { fill: { color: PH_BG } } }],
  ], { x: mm(L), y: mm(Y(99)), w: mm(W), colW: [mm(W - 63), mm(21), mm(21), mm(21)], fontFace: FONT, fontSize: 9, color: INK, border: { type: "solid", pt: 0.5, color: LINE }, rowH: mm(9.2), valign: "middle", margin: [0, mm(2), 0, mm(2)] });
}

// =============================== P12 裏表紙（アプローチブックの表紙と同じ枠：緑の帯・山吹色の帯）
{
  const s = page(0);
  R(s, 0, 0, 210, 25, GREEN);
  R(s, 0, 25, 5, 247, GREEN);
  R(s, 6.5, 30, 3, 242, YELLOW);
  R(s, 0, 272, 210, 25, YELLOW);
  T(s, "勉強会ご参加の皆様へ", 18, 32, 50, 8, { fontFace: FONT_UI, fontSize: 11, bold: true, color: "FFFFFF", fill: { color: GREEN }, align: "center", valign: "middle" });
  T(s, [nl("まずは一度、外壁・屋根の"), n("“健康診断”", { color: "FF0000" }), n("を受けてみてください")], 18, 43, 180, 24, { fontSize: 22, bold: true, color: INK, valign: "middle", lineSpacingMultiple: 1.1 });
  R(s, 18, 72, 178, 72, "FFFFFF", { line: { color: GREEN, width: 2 } });
  T(s, "無料 外壁・屋根の健康診断", 18, 72, 178, 10, { fontFace: FONT_UI, fontSize: 14, bold: true, color: "FFFFFF", fill: { color: GREEN }, valign: "middle", margin: [0, mm(5), 0, mm(5)] });
  T(s, [nl("● 機械を使わず、目で見て・手で触って確認します"), nl("● 写真付きの診断結果で、わかりやすくご説明します"), n("● 診断を受けたからといって、契約の必要はありません")], 24, 86, 120, 22, { fontSize: 10.5, paraSpaceAfter: 3 });
  s.addShape(pres.ShapeType.line, { x: mm(24), y: mm(111), w: mm(120), h: 0, line: { color: LINE, width: 0.75 } });
  T(s, [nl("お電話でのお申込み", { fontSize: 8.5, color: MUTED }), nl("0000-00-0000", { fontSize: 22, bold: true, color: GREEN_D }), n("受付時間 〇:〇〇〜〇:〇〇（〇曜定休）", { fontSize: 8.5, color: MUTED })], 24, 113, 120, 27);
  ph(s, "【QRコード】\nWEB申込み", 150, 88, 40, 40);
  T(s, "ヤマキシペイントについて", 18, 152, 98, 8, { fontFace: FONT_UI, fontSize: 12.5, bold: true, color: "FFFFFF", fill: { color: GREEN }, valign: "middle", margin: [0, mm(3), 0, mm(3)] });
  T(s, [nl("・創業120年以上、地域の住まいを見守ってきました"), nl("・ショールームで色見本や塗料を見ながらご相談いただけます"), nl("・エスケー化研の塗料と連帯保証で安心を"), n("・対応エリア：〇〇市・〇〇市・〇〇町")], 18, 163, 98, 34, { fontSize: 10, paraSpaceAfter: 4 });
  ph(s, "【地図・写真】\nショールームの外観／地図\n住所・営業時間", 122, 152, 74, 48);
  s.addImage({ path: A("logo.png"), x: mm(78), y: mm(206), w: mm(54), h: mm(34.4), objectName: "ロゴ" });
  T(s, "株式会社山岸　〒000-0000 〇〇県〇〇市〇〇町0-0", 18, 244, 178, 8, { fontSize: 10, align: "center", valign: "middle" });
  s.addImage({ path: A("mascot.png"), x: mm(14), y: mm(246), w: mm(38), h: mm(38.5), objectName: "キャラクター" });
}

pres.writeFile({ fileName: OUT }).then(() => console.log("wrote", OUT));
