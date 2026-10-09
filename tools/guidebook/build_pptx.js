// 外装工事ガイドブック（A4縦・12ページ）PowerPoint 生成
// 体裁はヤマキシペイント様アプローチブック（2025年1月版）に合わせる
//   上部：山吹色の斜め帯（長方形＋直角三角形）／緑の帯に白抜き中央タイトル（メイリオ太字）
//   見出し：緑の角丸ピル（Meiryo UI 白太字）／囲み：薄いグレー（#EFEFEF）／強調：赤
//   ページ末：キャラクター＋緑の吹き出し／裏表紙：緑と山吹の縦帯＋斜めの三角
// 使い方: NODE_PATH=<pptxgenjsのnode_modules> node build_pptx.js [出力先.pptx]
const pptxgen = require("pptxgenjs");
const path = require("path");

const A = (f) => path.join(__dirname, "assets", f);
const OUT = process.argv[2] || path.join(__dirname, "guidebook.pptx");
const mm = (v) => v / 25.4;

const GREEN = "00B050", YELLOW = "FED86C", GRAY = "EFEFEF", INK = "2D2D2D", MUTED = "595959";
const RED = "FF0000", DRED = "C00000", BLUE = "2573D1", BADGE = "4091F2", ORANGE = "FF6600", TBL = "F79646";
const FONT = "メイリオ", UI = "Meiryo UI";

const pres = new pptxgen();
pres.defineLayout({ name: "A4_PORTRAIT", width: mm(210), height: mm(297) });
pres.layout = "A4_PORTRAIT";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = "失敗しない外装工事ガイドブック";
pres.company = "株式会社山岸";

const L = 10, W = 190;

// ---------- 部品 ----------
function T(s, text, x, y, w, h, o = {}) {
  s.addText(text, Object.assign({ x: mm(x), y: mm(y), w: mm(w), h: mm(h), fontFace: FONT, fontSize: 11, color: INK, margin: 0, valign: "top", isTextBox: true }, o));
}
function box(s, x, y, w, h, fill, o = {}) {
  s.addShape(o.r ? pres.ShapeType.roundRect : pres.ShapeType.rect, Object.assign({ x: mm(x), y: mm(y), w: mm(w), h: mm(h), fill: { color: fill }, line: o.line || { type: "none" } }, o.r ? { rectRadius: mm(o.r) } : {}));
}
const b = (t, o = {}) => ({ text: t, options: Object.assign({ bold: true }, o) });
const n = (t, o = {}) => ({ text: t, options: o });
const nl = (t, o = {}) => ({ text: t, options: Object.assign({ breakLine: true }, o) });
const red = (t) => b(t, { color: RED });

// 本文ページの枠
function page(title, titleSize = 26) {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  box(s, 0, 0, 4.9, 24.8, YELLOW);
  s.addShape(pres.ShapeType.rtTriangle, { x: mm(7.6), y: 0, w: mm(202.4), h: mm(24.8), fill: { color: YELLOW }, line: { type: "none" } });
  box(s, 0, 29.6, 210, 27.4, GREEN);
  T(s, title, 8.8, 30.6, 192.4, 25.4, { fontSize: titleSize, bold: true, color: "FFFFFF", align: "center", valign: "middle", lineSpacingMultiple: 0.95 });
  return s;
}
// 緑の角丸ピル見出し
function pill(s, text, x, y, w) {
  s.addText(text, { shape: pres.ShapeType.roundRect, rectRadius: mm(5.5), x: mm(x), y: mm(y), w: mm(w || text.length * 5.2 + 16), h: mm(11), fill: { color: GREEN }, line: { type: "none" }, fontFace: UI, fontSize: 14, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "見出し" });
}
function panel(s, x, y, w, h) { box(s, x, y, w, h, GRAY, { r: 2.5 }); }
function ph(s, label, x, y, w, h) {
  T(s, label, x, y, w, h, { fontSize: 9, color: "7F7F7F", align: "center", valign: "middle", fill: { color: "F7F7F7" }, line: { color: "A6A6A6", width: 0.75, dashType: "dash" }, objectName: "写真枠" });
}
// キャラクター＋緑の吹き出し（ページの締め）
function bubble(s, text, y, h = 22) {
  s.addText(text, { shape: pres.ShapeType.roundRect, rectRadius: mm(4), x: mm(L), y: mm(y), w: mm(146), h: mm(h), fill: { color: GREEN }, line: { type: "none" }, fontFace: FONT, fontSize: 13, bold: true, color: "FFFFFF", valign: "middle", margin: [0, mm(5), 0, mm(5)], isTextBox: true, objectName: "吹き出し" });
  s.addShape(pres.ShapeType.triangle, { x: mm(L + 145), y: mm(y + h / 2 - 3.5), w: mm(7), h: mm(8), rotate: 90, fill: { color: GREEN }, line: { type: "none" } });
  s.addImage({ path: A("mascot.png"), x: mm(166), y: mm(y + h / 2 - 15.2), w: mm(30), h: mm(30.4), objectName: "キャラクター" });
}
function badge(s, label, x, y, d = 13) {
  s.addText(label, { shape: pres.ShapeType.ellipse, x: mm(x), y: mm(y), w: mm(d), h: mm(d), fill: { color: BADGE }, line: { type: "none" }, fontFace: FONT, fontSize: 11, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true });
}
function redCard(s, head, body, x, y, w, h) {
  T(s, head, x, y, w, 8, { fontSize: 12, bold: true, color: "FFFFFF", fill: { color: DRED }, align: "center", valign: "middle" });
  T(s, body, x, y + 8, w, h - 8, { fontSize: 10, fill: { color: GRAY }, margin: [mm(2), mm(3), mm(2), mm(3)], lineSpacingMultiple: 1.1 });
}
const thO = (t) => ({ text: t, options: { bold: true, fill: { color: TBL }, color: "000000", align: "center" } });
const TBLOPT = { fontFace: FONT, fontSize: 9.5, color: INK, border: { type: "solid", pt: 0.75, color: "A6A6A6" }, valign: "middle", margin: [mm(1), mm(2), mm(1), mm(2)] };

// =============================== P1 表紙 ===============================
{
  const s = pres.addSlide();
  s.addImage({ path: A("cover.png"), x: 0, y: 0, w: mm(210), h: mm(297), objectName: "表紙画像" });
}

// =============================== P2 はじめに ===============================
{
  const s = page("このガイドブックについて");
  T(s, [nl("外壁や屋根の塗装は、10年以上に一度の大きな工事です。"), nl("そのため「適正な価格がわからない」「どの業者を選べばいいかわからない」「何を比べればいいかわからない」という声が多く聞かれます。"), n("このガイドブックでは、工事の前に知っておきたい基礎知識と、業者を比べるときのポイントを、国や公的機関の情報をもとにまとめました。")], L, 64, W, 40, { fontSize: 12, lineSpacingMultiple: 1.3, paraSpaceAfter: 4 });
  pill(s, "目次", L, 108, 50);
  const toc = [["外壁・屋根は必ず劣化します", 3], ["塗り替えの時期と、塗装できなくなる状態", 4], ["塗料の種類と耐用年数", 5], ["塗料メーカーを知っておこう", 6], ["正しい塗装工事の流れ", 7], ["業者を選ぶときに確かめること", 8], ["見積書の見方", 9], ["契約と保証のポイント", 10], ["トラブルを防ぐために", 11]];
  panel(s, L, 123, W, 117);
  toc.forEach(([t, p], i) => {
    const y = 127 + i * 12.5;
    s.addText(String(i + 1), { shape: pres.ShapeType.ellipse, x: mm(L + 6), y: mm(y + 0.5), w: mm(9), h: mm(9), fill: { color: GREEN }, line: { type: "none" }, fontFace: UI, fontSize: 12, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true });
    T(s, t, L + 20, y, 140, 10, { fontSize: 13, bold: true, valign: "middle" });
    T(s, "P." + p, L + W - 30, y, 24, 10, { fontSize: 12, bold: true, color: GREEN, align: "right", valign: "middle" });
  });
  bubble(s, "見積書を比べるときに、手元に置いて使ってください！", 252);
}

// =============================== P3 劣化 ===============================
{
  const s = page("外壁・屋根は必ず劣化します");
  pill(s, "塗装の2つの役割", L, 63);
  T(s, [n("①雨・紫外線から"), red("建物を守る"), n("　②"), red("美しい外観を保つ"), nl(""), n("塗装の膜（塗膜）は、雨や紫外線で少しずつ傷みます。車検のような決まりはないため、住まいの状態を自分で知っておくことが大切です。")], L, 77, W, 20, { fontSize: 11.5, lineSpacingMultiple: 1.25 });
  pill(s, "劣化のサインをチェック", L, 99);
  panel(s, L, 113, W, 58);
  const items = ["外壁をこすると白い粉がつく（チョーキング）", "外壁にひび割れがある", "目地（シーリング）が割れている・痩せている", "外壁の板が反っている、表面がはがれている", "カビ・コケが生えている", "屋根の色あせ・錆がある", "雨樋が傾いている、詰まっている", "前回の塗装から10年以上たっている"];
  items.forEach((t, i) => {
    const x = L + 5 + (i % 2) * 93, y = 117 + Math.floor(i / 2) * 13;
    box(s, x, y + 1.8, 5, 5, "FFFFFF", { line: { color: GREEN, width: 1.5 } });
    T(s, t, x + 7.5, y, 84, 11, { fontSize: 10.5, valign: "middle" });
  });
  badge(s, "特性", L, 177);
  T(s, [nl("ひび割れは「毛細管現象」で水が入りやすい", { bold: true }), n("細いすき間ほど水を吸い込み、外壁の内側に雨水が入る原因になります。")], L + 16, 177, 80, 22, { fontSize: 10 });
  badge(s, "特性", L + 98, 177);
  T(s, [nl("北陸は湿気・雨・雪が多い地域", { bold: true }), n("凍ったりとけたりをくり返すと、外壁の表面がはがれる「凍害」が起きることがあります。")], L + 114, 177, 76, 22, { fontSize: 10 });
  ph(s, "【写真】チョーキング", L, 203, 61, 40);
  ph(s, "【写真】ひび割れ・目地の痩せ", L + 64.5, 203, 61, 40);
  ph(s, "【写真】凍害・カビコケ", L + 129, 203, 61, 40);
  bubble(s, "劣化を放っておくと雨漏りや下地の腐れにつながります。早めのメンテナンスを考えましょう！", 252, 24);
}

// =============================== P4 時期・塗装できない状態 ===============================
{
  const s = page("塗り替えの時期と、\n塗装できなくなる状態", 24);
  pill(s, "塗り替えの目安", L, 63);
  panel(s, L, 77, W, 33);
  T(s, [n("初めての塗り替え　"), b("築8〜10年", { color: RED, fontSize: 18 }), nl("　程度（新築時の塗料による）"), n("2回目以降　　　　"), b("前回の塗料の耐用年数", { color: RED, fontSize: 15 }), nl("　が目安"), n("※年数はあくまで目安です。実際の劣化の状態で判断しましょう。", { fontSize: 9.5, color: MUTED })], L + 6, 80, W - 12, 28, { fontSize: 12, lineSpacingMultiple: 1.25 });
  pill(s, "塗装で直せる？　直せない？", L, 116);
  panel(s, L, 130, W, 74);
  T(s, "まだ塗装で直せる状態", L + 4, 133, 80, 8, { fontSize: 13, bold: true, color: GREEN, align: "center" });
  T(s, "塗装ではもう直せない状態", L + 106, 133, 80, 8, { fontSize: 13, bold: true, color: ORANGE, align: "center" });
  T(s, "VS", L + 85, 160, 20, 10, { fontSize: 20, bold: true, color: ORANGE, align: "center", valign: "middle" });
  ph(s, "【写真】色あせ・チョーキング", L + 4, 142, 80, 30);
  ph(s, "【写真】外壁材のはがれ・欠け", L + 106, 142, 80, 30);
  T(s, [nl("・色あせ・チョーキング"), nl("・細かいひび割れ"), n("・目地の痩せ・割れ")], L + 6, 175, 78, 27, { fontSize: 10.5 });
  T(s, [nl("・外壁材そのものがはがれ・欠けている"), nl("・大きな反り・浮きがある"), n("・下地の腐れ・雨漏りが起きている")], L + 108, 175, 78, 27, { fontSize: 10.5 });
  T(s, [n("塗装で直せない場合は、"), red("張替えやカバー工法（重ね張り）"), n("が必要になり、塗装より費用も工期も大きくなります。")], L, 208, W, 14, { fontSize: 11.5, lineSpacingMultiple: 1.2 });
  T(s, [n("工事を選ぶときは「"), b("あと何年この家に住むか"), n("」も大切な目安です。")], L, 223, W, 10, { fontSize: 11.5 });
  bubble(s, "傷みすぎる前の塗り替えが、結果的にいちばん負担の少ない方法です！", 252);
}

// =============================== P5 塗料の種類 ===============================
{
  const s = page("塗料の種類と耐用年数");
  T(s, [n("塗料の耐用年数を左右するのは、主な成分の"), b("「樹脂」"), n("です。")], L, 63, W, 8, { fontSize: 12 });
  s.addTable([
    [thO("種類（樹脂）"), thO("耐用年数の目安"), thO("特徴")],
    ["アクリル", "5〜7年", "価格は安いが耐久性が低く、現在の外壁塗装ではあまり使われない"],
    ["ウレタン", "8〜10年", "やわらかく密着しやすい。雨樋などの付帯部に使われることが多い"],
    ["シリコン", "10〜15年", "価格と耐久性のバランスがよく、外壁塗装でよく使われる"],
    ["ラジカル制御形", "12〜15年", "塗膜の劣化（チョーキング）の原因を抑える技術を使ったもの"],
    ["フッ素", "15〜20年", "耐久性・汚れにくさに優れる。価格は高め"],
    ["無機", "20〜25年", "劣化しにくい無機成分を含み、最も長持ちする。価格は最も高い"],
  ], Object.assign({ x: mm(L), y: mm(73), w: mm(W), colW: [mm(36), mm(34), mm(120)], rowH: mm(10.5) }, TBLOPT, { fontSize: 10 }));
  T(s, "※一般的な目安です。実際の年数は製品・立地・下地の状態・施工によって変わります。", L, 148, W, 6, { fontSize: 9, color: MUTED });
  pill(s, "よくある誤解", L, 157);
  const mis = [["「耐用年数」＝「保証年数」ではない", "耐用年数はメーカーの試験などによる目安です。保証の年数・内容は、保証書で別に確認しましょう。"], ["「遮熱」と「無機」は別のもの", "遮熱は日差しを反射して温度上昇を抑える「機能」、無機は劣化しにくい「成分」のことです。"], ["屋根と外壁では適した塗料が違う", "屋根は外壁より日差しや雨を強く受けます。屋根用の製品が使われているか確認しましょう。"]];
  mis.forEach(([h, d], i) => {
    const y = 171 + i * 25;
    panel(s, L, y, W, 22);
    badge(s, String(i + 1), L + 4, y + 4.5, 13);
    T(s, [nl(h, { bold: true, color: BLUE, fontSize: 12 }), n(d)], L + 21, y + 2.5, W - 25, 18, { fontSize: 10.5 });
  });
  bubble(s, "大切なのは「何年もたせたいか」に合った塗料を選ぶことです！", 252);
}

// =============================== P6 メーカー ===============================
{
  const s = page("塗料メーカーを知っておこう");
  T(s, [n("同じ「シリコン塗料」でも、メーカーや商品によって性能は違います。"), red("見積書にメーカー名と商品名が書いてあれば"), n("、カタログで性能を自分で調べることができます。")], L, 63, W, 16, { fontSize: 11.5, lineSpacingMultiple: 1.25 });
  pill(s, "主な塗料メーカー", L, 82);
  s.addTable([
    [thO("メーカー"), thO("創業"), thO("特徴"), thO("外壁用の代表的な商品")],
    ["日本ペイント", "1881年", "日本で最初の塗料メーカー。自動車・工業用なども手がける総合メーカー", "パーフェクトトップ"],
    ["関西ペイント", "1918年", "自動車・工業用なども手がける総合メーカー", "アレスダイナミックTOP"],
    ["エスケー化研", "1955年", "建築用の仕上塗材が専門。建築仕上塗材の国内シェア53％（2024年）", "エスケープレミアムシリコン"],
    ["菊水化学工業", "1959年", "建築用の仕上塗材のメーカー", "グラナダフレッシュ"],
    ["アステックペイント", "日本では\n2000年から", "オーストラリアのメーカー。加盟店（登録した施工店）のみが扱う", "超低汚染リファイン"],
  ], Object.assign({ x: mm(L), y: mm(96), w: mm(W), colW: [mm(34), mm(24), mm(82), mm(50)], rowH: mm(13) }, TBLOPT));
  T(s, "※各社の公表情報より。シェアは日本建築仕上工業会のデータ（2024年5月）による。", L, 176, W, 6, { fontSize: 9, color: MUTED });
  pill(s, "オリジナル塗料とは", L, 185);
  panel(s, L, 199, W, 22);
  T(s, "塗装店が自社ブランドとして販売する塗料です。良い製品もありますが、製造元や性能データが公開されていないと、他の製品と比べることができません。製造元と性能の根拠を確認しましょう。", L + 5, 201, W - 10, 18, { fontSize: 10.5, lineSpacingMultiple: 1.2, valign: "middle" });
  pill(s, "見積書で確認すること", L, 226);
  T(s, [red("メーカー名"), n("　・　"), red("商品名"), n("　・　"), red("塗る回数"), n("　・　"), red("塗る面積（㎡）")], L, 240, W, 9, { fontSize: 13, align: "center", valign: "middle" });
  bubble(s, "わからない塗料名は、その場で業者に聞いてみましょう！", 254, 20);
}

// =============================== P7 工事の流れ ===============================
{
  const s = page("正しい塗装工事の流れ");
  T(s, "外壁・屋根の塗装は、決まった順番で進みます。工程が省かれると、数年ではがれる原因になります。", L, 63, W, 11, { fontSize: 11.5 });
  const flow = (x, title, steps) => {
    pill(s, title, x, 77, 92);
    steps.forEach((t, i) => {
      const y = 91 + i * 10.6;
      s.addText(String(i + 1), { shape: pres.ShapeType.ellipse, x: mm(x), y: mm(y + 0.8), w: mm(8), h: mm(8), fill: { color: GREEN }, line: { type: "none" }, fontFace: UI, fontSize: 11, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true });
      T(s, t[0], x + 10, y, 82, 9.6, { fontSize: 10.5, bold: true, fill: { color: GRAY }, valign: "middle", margin: [0, mm(3), 0, mm(3)] });
      if (t[1]) T(s, t[1], x + 52, y, 39, 9.6, { fontSize: 8.5, color: MUTED, align: "right", valign: "middle" });
    });
  };
  flow(L, "外壁（全10工程）", [["足場の組立", "本足場"], ["飛散防止ネット"], ["高圧洗浄", "洗った後は乾燥"], ["下地処理", "ひび補修・ケレン"], ["養生", "窓・車などを保護"], ["下塗り", "密着をよくする"], ["中塗り"], ["上塗り"], ["完了検査", "塗り残し確認"], ["足場解体・清掃"]]);
  flow(L + 98, "屋根（全8工程）", [["高圧洗浄"], ["鉄部の下地調整", "錆を落とす"], ["鉄部の錆止め塗装"], ["下塗り"], ["中塗り"], ["上塗り"], ["完了検査"], ["清掃"]]);
  pill(s, "手抜きされやすいポイント", L, 200);
  const pts = [["乾燥時間", "塗り重ねの間に、塗料ごとに決められた乾燥時間をとる"], ["塗る回数", "下塗り・中塗り・上塗りの3回塗りが基本"], ["塗料の量・薄め方", "決められた量を塗り、薄めすぎない"], ["下地処理", "ひび割れの補修や錆落としを丁寧に行う"]];
  pts.forEach(([h, d], i) => {
    const x = L + (i % 2) * 96, y = 214 + Math.floor(i / 2) * 20;
    panel(s, x, y, 94, 18);
    T(s, [nl(h, { bold: true, color: RED, fontSize: 11.5 }), n(d)], x + 4, y + 2, 86, 14, { fontSize: 10 });
  });
  T(s, "工事中の写真を撮って、報告してもらえるか確認しておくと安心です。", L, 256, W, 9, { fontSize: 11, bold: true, color: GREEN });
}

// =============================== P8 業者の確かめ方 ===============================
{
  const s = page("業者を選ぶときに\n確かめること", 24);
  T(s, [n("価格だけで選ぶと失敗のもとです。工事の失敗は、悪質な業者だけでなく"), red("診断の間違いや知識不足"), n("からも起こります。")], L, 63, W, 15, { fontSize: 11.5, lineSpacingMultiple: 1.25 });
  pill(s, "会社について確かめること", L, 80);
  const c = [
    ["建設業許可", "税込500万円以上の工事には建設業許可が必要です。許可があれば、一定の経営・技術の基準を満たしています。"],
    ["資格", "「一級・二級塗装技能士」は国の技能検定による国家資格です。「外壁診断士」などの民間資格もあります。"],
    ["所在地・店舗", "事務所や店舗の場所がはっきりしているか。家から近いと、工事後の対応が早くなります。"],
    ["実績", "施工事例・創業年数・地域での実績を、ホームページなどで確認しましょう。"],
    ["保険への加入", "工事中の事故や、近隣への塗料の飛散などに備える賠償責任保険に入っているか。"],
    ["施工の体制", "誰が工事を管理するのか。下請けに任せきりにしていないか。"],
  ];
  c.forEach(([h, d], i) => {
    const x = L + (i % 2) * 96, y = 94 + Math.floor(i / 2) * 31;
    panel(s, x, y, 94, 28);
    T(s, [nl(h, { bold: true, color: GREEN, fontSize: 12.5 }), n(d)], x + 4, y + 2.5, 86, 24, { fontSize: 10, lineSpacingMultiple: 1.15 });
  });
  pill(s, "現地調査で確かめること", L, 191);
  panel(s, L, 205, W, 40);
  T(s, [nl("・外壁や屋根を、目で見て・手で触って確認しているか"), nl("・劣化の場所を写真に撮り、写真を見せながら説明してくれるか"), nl("・劣化の原因と、それに合った工事方法を説明してくれるか"), n("・図面や坪数だけで、すぐに金額を出していないか")], L + 5, 208, W - 10, 35, { fontSize: 11, paraSpaceAfter: 3, valign: "middle" });
  bubble(s, "説明がわかりやすく、質問に正直に答えてくれるかも大切です！", 254, 20);
}

// =============================== P9 見積書 ===============================
{
  const s = page("見積書の見方");
  pill(s, "わかりやすい見積書の3つの条件", L, 63);
  T(s, [b("①", { color: GREEN }), n("部位ごとに分かれている　"), b("②", { color: GREEN }), n("数量と単位（㎡・m）がある　"), b("③", { color: GREEN }), n("メーカー名・商品名・塗る回数がある")], L, 77, W, 9, { fontSize: 11, bold: true, valign: "middle" });
  T(s, "悪い例", L, 90, 40, 8, { fontSize: 13, bold: true, color: ORANGE });
  T(s, "良い例", L + 62, 90, 40, 8, { fontSize: 13, bold: true, color: GREEN });
  s.addTable([[thO("項目"), thO("数量")], ["外壁塗装工事", "一式"], ["屋根塗装工事", "一式"], ["付帯部塗装", "一式"], ["諸経費", "一式"]], Object.assign({ x: mm(L), y: mm(99), w: mm(55), colW: [mm(38), mm(17)], rowH: mm(9) }, TBLOPT));
  T(s, "VS", L + 55, 112, 7, 10, { fontSize: 14, bold: true, color: ORANGE, align: "center", valign: "middle" });
  s.addTable([[thO("項目"), thO("数量"), thO("単位")],
    ["足場（本足場）・飛散防止ネット", "〇〇", "㎡"], ["高圧洗浄", "〇〇", "㎡"], ["シーリング打ち替え（目地）", "〇〇", "m"],
    ["外壁 下塗り／〇〇社 商品名", "〇〇", "㎡"], ["外壁 中塗り・上塗り／〇〇社 商品名", "〇〇", "㎡"], ["付帯部（軒天・破風・雨樋 など）", "〇〇", "m・㎡"], ["屋根 錆止め・下塗り・上塗り／商品名", "〇〇", "㎡"]],
    Object.assign({ x: mm(L + 62), y: mm(99), w: mm(128), colW: [mm(96), mm(16), mm(16)], rowH: mm(9) }, TBLOPT, { fontSize: 9 }));
  T(s, "「一式」ばかりだと、何をどれだけするのかわからず、あとから追加費用を求められることもあります。", L, 175, W, 12, { fontSize: 11, lineSpacingMultiple: 1.2 });
  pill(s, "塗装面積に注意", L, 189);
  panel(s, L, 203, W, 37);
  T(s, [nl("外壁の面積は「延床面積×約1.2」がおおよその目安ですが、"), n("窓の大きさ・数、バルコニーの有無、家の高さで変わります。"), red("同じ坪数でも外壁の面積は家ごとに違う"), nl("のです。"), n("見積書の面積が実際に測った数字や図面にもとづいているか確認しましょう。面積が小さすぎると、塗料が足りなくなる原因になります。")], L + 5, 205, W - 10, 33, { fontSize: 10.5, lineSpacingMultiple: 1.25, valign: "middle" });
  bubble(s, "見積もりは2〜3社から取り、同じ条件で比べてみましょう！", 254, 20);
}

// =============================== P10 契約・保証 ===============================
{
  const s = page("契約と保証のポイント");
  pill(s, "契約前にそろえたい3つの書類", L, 63);
  [["工事請負契約書", "建設業法で、工事の大小にかかわらず書面での契約が決められています（工事内容・代金・工期など）"], ["工程表", "いつ何をするかがわかり、工程が省かれていないか確認できます"], ["保証書", "誰が・何を・何年保証するかを確認しましょう"]].forEach(([h, d], i) => {
    const x = L + i * 64;
    panel(s, x, 77, 62, 38);
    T(s, [nl(h, { bold: true, color: GREEN, fontSize: 12 }), n(d)], x + 3.5, 79.5, 55, 34, { fontSize: 9.8, lineSpacingMultiple: 1.15 });
  });
  pill(s, "保証の種類", L, 120);
  s.addTable([
    [thO("種類"), thO("内容")],
    [{ text: "施工店の保証", options: { bold: true } }, "工事をした会社が、施工の不具合を保証するもの。保証の範囲・年数は会社ごとに違います"],
    [{ text: "メーカーとの連名保証", options: { bold: true } }, "条件を満たした工事について、塗料メーカーと施工店が連名で出す保証。塗料メーカー単独の保証は一般的ではありません"],
    [{ text: "リフォーム瑕疵保険", options: { bold: true } }, "登録事業者が工事ごとに加入し、第三者の検査を受ける保険。事業者が倒産した場合は、発注者が保険法人に直接請求できます"],
  ], Object.assign({ x: mm(L), y: mm(134), w: mm(W), colW: [mm(44), mm(146)], rowH: mm(13) }, TBLOPT, { fontSize: 9.8 }));
  T(s, "保証期間中の定期点検があるかどうかも確認しましょう。", L, 189, W, 7, { fontSize: 11, bold: true, color: RED });
  pill(s, "法律で決まっていること", L, 198);
  [["足場", "2024年4月から、幅1m以上の場所では原則として「本足場」（支柱が2列の足場）を使うことが義務になりました"], ["石綿", "改修工事の前には、有資格者による石綿（アスベスト）の事前調査が必要です。税込100万円以上の改修工事は、調査結果の報告も必要です"]].forEach(([h, d], i) => {
    const y = 212 + i * 22;
    badge(s, h, L, y + 2, 15);
    T(s, d, L + 19, y, W - 19, 19, { fontSize: 10.5, lineSpacingMultiple: 1.2, valign: "middle" });
  });
  T(s, "※出典：建設業法第19条、労働安全衛生規則、大気汚染防止法、住宅瑕疵担保責任保険協会の案内", L, 258, W, 6, { fontSize: 8.5, color: MUTED });
}

// =============================== P11 トラブル対策 ===============================
{
  const s = page("トラブルを防ぐために");
  pill(s, "こんなときは要注意", L, 63);
  const cards = [
    ["突然の訪問「屋根がずれている」", "点検を口実にした訪問販売のトラブルが増えています。その場で屋根に上らせたり、契約したりしないようにしましょう"],
    ["「火災保険で無料で直せる」", "経年劣化は火災保険の対象外です。保険の請求は損害が起きてから3年以内が原則。まず保険会社に確認しましょう"],
    ["「今日決めれば大幅値引き」", "その場で契約を急がせる業者には注意。必ず複数の見積もりを比べてから決めましょう"],
    ["訪問販売で契約してしまったら", "契約書面を受け取った日から8日以内なら、クーリング・オフで契約を解除できます（書面やメールで通知）"],
  ];
  cards.forEach(([h, d], i) => redCard(s, h, d, L + (i % 2) * 96, 77 + Math.floor(i / 2) * 33, 94, 31));
  pill(s, "困ったときの相談窓口", L, 145);
  panel(s, L, 159, W, 30);
  T(s, [nl("住まいるダイヤル　0570-016-100", { bold: true, color: GREEN, fontSize: 13 }), nl("国土交通大臣指定の住まいの相談窓口。リフォームの見積書の無料チェックもあります", { fontSize: 9.5 }), nl("消費者ホットライン　188", { bold: true, color: GREEN, fontSize: 13 }), n("最寄りの消費生活センターにつながります", { fontSize: 9.5 })], L + 5, 161, W - 10, 26, { lineSpacingMultiple: 1.1 });
  pill(s, "業者比較チェックシート", L, 194);
  const r = (t) => [t, "", "", ""];
  s.addTable([
    [thO("チェック項目"), thO("A社"), thO("B社"), thO("C社")],
    r("所在地・建設業許可・資格が確認できる"), r("写真を見せながら劣化を説明してくれた"), r("見積書に部位・㎡・メーカー名・商品名がある"),
    r("契約書・工程表・保証書がそろっている"), r("保証の内容（誰が・何を・何年）が明確"), r("その場で契約を急がせない"),
  ], Object.assign({ x: mm(L), y: mm(208), w: mm(W), colW: [mm(118), mm(24), mm(24), mm(24)], rowH: mm(9.4) }, TBLOPT, { fontSize: 10 }));
}

// =============================== P12 裏表紙（アプローチブックの表紙と同じ枠） ===============================
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  box(s, 11.2, 0.3, 4.7, 271, YELLOW);
  box(s, -0.8, 0, 9.6, 271.3, GREEN);
  s.addShape(pres.ShapeType.rtTriangle, { x: 0, y: 0, w: mm(210), h: mm(39.4), rotate: 180, fill: { color: GREEN }, line: { type: "none" } });
  s.addShape(pres.ShapeType.rtTriangle, { x: mm(-0.8), y: mm(264.4), w: mm(210.8), h: mm(32.6), fill: { color: YELLOW }, line: { type: "none" } });
  s.addImage({ path: A("logo.png"), x: mm(57.5), y: mm(105), w: mm(94.9), h: mm(60.5), objectName: "ロゴ" });
  T(s, "発行：株式会社山岸（ヤマキシペイント）", 30, 172, 160, 8, { fontSize: 12, bold: true, align: "center" });
  T(s, [nl("【参考】国土交通省／厚生労働省／環境省／消費者庁／"), nl("公益財団法人 住宅リフォーム・紛争処理支援センター／住宅瑕疵担保責任保険協会／"), n("日本建築仕上工業会／各塗料メーカーの公表情報")], 30, 186, 160, 16, { fontSize: 8.5, color: MUTED, align: "center" });
  T(s, "本書の内容は〇〇年〇月時点の情報です。", 30, 204, 160, 6, { fontSize: 8.5, color: MUTED, align: "center" });
  s.addImage({ path: A("mascot.png"), x: mm(8.8), y: mm(232.8), w: mm(63.3), h: mm(64.2), objectName: "キャラクター" });
}

pres.writeFile({ fileName: OUT }).then(() => console.log("wrote", OUT));
