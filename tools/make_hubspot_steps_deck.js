const pptxgen = require("pptxgenjs");

const BROWN = "5F3F33";   // ブランドの濃茶
const TERRA = "D84A2E";   // ブランドの朱色
const GOLD  = "F5B31C";
const CREAM = "F7F1EA";
const MUTED = "8A7670";
const INK   = "33241F";
const WHITE = "FFFFFF";

const FB = "Meiryo";      // 本文
const FH = "Meiryo";      // 見出し

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";          // 13.333 x 7.5
pres.author = "イエオモイ";
pres.title = "HubSpot側の作業手順";

const W = 13.333, H = 7.5, M = 0.6;

function shadow() {
  return { type: "outer", color: "5F3F33", opacity: 0.18, blur: 10, offset: 3, angle: 90 };
}

// 各ステップ共通のヘッダー（丸バッジ＋見出し）
function stepHead(slide, num, title, lead) {
  slide.addShape(pres.ShapeType.ellipse, {
    x: M, y: 0.42, w: 0.82, h: 0.82, fill: { color: TERRA }, line: { color: TERRA },
  });
  slide.addText(String(num), {
    x: M, y: 0.42, w: 0.82, h: 0.82, align: "center", valign: "middle",
    fontFace: FH, fontSize: 34, bold: true, color: WHITE, margin: 0,
  });
  slide.addText(title, {
    x: M + 1.08, y: 0.42, w: W - M * 2 - 1.08, h: 0.82, valign: "middle",
    fontFace: FH, fontSize: 30, bold: true, color: BROWN, margin: 0,
  });
  if (lead) {
    slide.addText(lead, {
      x: M, y: 1.38, w: W - M * 2, h: 0.42, valign: "middle",
      fontFace: FB, fontSize: 15, color: MUTED, margin: 0,
    });
  }
}

// 手順の1行（番号つき）
function stepRow(slide, y, n, head, body) {
  slide.addShape(pres.ShapeType.ellipse, {
    x: M + 0.06, y: y + 0.06, w: 0.42, h: 0.42, fill: { color: CREAM }, line: { color: GOLD, width: 1.5 },
  });
  slide.addText(String(n), {
    x: M + 0.06, y: y + 0.06, w: 0.42, h: 0.42, align: "center", valign: "middle",
    fontFace: FB, fontSize: 14, bold: true, color: BROWN, margin: 0,
  });
  slide.addText(head, {
    x: M + 0.72, y: y, w: W - M * 2 - 0.72, h: 0.36, valign: "middle",
    fontFace: FH, fontSize: 16, bold: true, color: INK, margin: 0,
  });
  if (body) {
    slide.addText(body, {
      x: M + 0.72, y: y + 0.34, w: W - M * 2 - 0.72, h: 0.34, valign: "middle",
      fontFace: FB, fontSize: 13, color: MUTED, margin: 0,
    });
  }
}

// コード表示用の箱
function codeBox(slide, x, y, w, h, lines, opt) {
  opt = opt || {};
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: opt.fill || "2E2320" }, line: { color: opt.line || "2E2320" },
  });
  slide.addText(lines, {
    x: x + 0.22, y: y + 0.14, w: w - 0.44, h: h - 0.28, valign: "top",
    fontFace: "Courier New", fontSize: opt.size || 13, color: opt.color || "F2E8E2", margin: 0,
  });
}

/* ---------------------------------------------------------- 1 表紙 */
{
  const s = pres.addSlide();
  s.background = { color: BROWN };

  s.addShape(pres.ShapeType.ellipse, {
    x: 9.6, y: -1.7, w: 5.6, h: 5.6, fill: { color: "6E4A3C" }, line: { color: "6E4A3C" },
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 11.2, y: 4.5, w: 3.4, h: 3.4, fill: { color: "70493B" }, line: { color: "70493B" },
  });

  s.addText("イエオモイ 画像LP", {
    x: M + 0.2, y: 1.55, w: 8.6, h: 0.5, fontFace: FB, fontSize: 17, color: GOLD, charSpacing: 2, margin: 0,
  });
  s.addText("HubSpot側で\nやるべきこと", {
    x: M + 0.2, y: 2.05, w: 8.6, h: 2.1, fontFace: FH, fontSize: 48, bold: true, color: WHITE,
    lineSpacing: 58, margin: 0,
  });
  s.addText("公開までの5ステップ", {
    x: M + 0.2, y: 4.35, w: 8.6, h: 0.45, fontFace: FB, fontSize: 18, color: "E6D5CB", margin: 0,
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: M + 0.2, y: 5.15, w: 3.5, h: 0.72, rectRadius: 0.36,
    fill: { color: TERRA }, line: { color: TERRA },
  });
  s.addText("所要時間 およそ30分", {
    x: M + 0.2, y: 5.15, w: 3.5, h: 0.72, align: "center", valign: "middle",
    fontFace: FB, fontSize: 15, bold: true, color: WHITE, margin: 0,
  });

  s.addNotes("LP本体のHTMLは完成済み。ここからはHubSpot側の設定作業のみ。");
}

/* ---------------------------------------------------------- 2 全体像 */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addText("全体の流れ", {
    x: M, y: 0.5, w: 8, h: 0.7, fontFace: FH, fontSize: 32, bold: true, color: BROWN, margin: 0,
  });
  s.addText("① 〜 ③ はHubSpotの管理画面、④ ⑤ はコードを貼る作業です", {
    x: M, y: 1.22, w: W - M * 2, h: 0.4, fontFace: FB, fontSize: 15, color: MUTED, margin: 0,
  });

  const cards = [
    ["1", "アカウントID\nを控える", "PORTAL_ID", "2分"],
    ["2", "プロパティ\nを作る", "desired_plan", "5分"],
    ["3", "フォーム\nを作る", "FORM_GUID", "10分"],
    ["4", "コードを\n貼る", "テンプレート", "5分"],
    ["5", "2行だけ\n書き換える", "公開", "3分"],
  ];
  const cw = 2.3, gap = 0.26;
  const x0 = (W - (cw * 5 + gap * 4)) / 2;

  cards.forEach((c, i) => {
    const x = x0 + i * (cw + gap);
    const y = 2.05;
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w: cw, h: 3.0, rectRadius: 0.1,
      fill: { color: i === 4 ? TERRA : CREAM }, line: { color: i === 4 ? TERRA : "EADFD5" },
      shadow: shadow(),
    });
    s.addShape(pres.ShapeType.ellipse, {
      x: x + cw / 2 - 0.26, y: y + 0.3, w: 0.52, h: 0.52,
      fill: { color: i === 4 ? WHITE : BROWN }, line: { color: i === 4 ? WHITE : BROWN },
    });
    s.addText(c[0], {
      x: x + cw / 2 - 0.26, y: y + 0.3, w: 0.52, h: 0.52, align: "center", valign: "middle",
      fontFace: FB, fontSize: 18, bold: true, color: i === 4 ? TERRA : WHITE, margin: 0,
    });
    s.addText(c[1], {
      x: x + 0.12, y: y + 1.0, w: cw - 0.24, h: 1.0, align: "center", valign: "top",
      fontFace: FH, fontSize: 17, bold: true, color: i === 4 ? WHITE : INK, lineSpacing: 24, margin: 0,
    });
    s.addText(c[2], {
      x: x + 0.12, y: y + 2.02, w: cw - 0.24, h: 0.38, align: "center", valign: "middle",
      fontFace: "Courier New", fontSize: 11, color: i === 4 ? "FBE3DB" : MUTED, margin: 0,
    });
    s.addText(c[3], {
      x: x + 0.12, y: y + 2.44, w: cw - 0.24, h: 0.36, align: "center", valign: "middle",
      fontFace: FB, fontSize: 12, bold: true, color: i === 4 ? WHITE : TERRA, margin: 0,
    });
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.5, w: W - M * 2, h: 1.05, rectRadius: 0.1,
    fill: { color: "FDF6EC" }, line: { color: GOLD, width: 1.5 },
  });
  s.addText([
    { text: "LP側の作業はありません。", options: { bold: true, color: BROWN } },
    { text: "  画像9枚・CSS・JavaScriptはすべてHTMLに埋め込み済みです。画像のアップロードもURLの書き換えも不要です。", options: { color: INK } },
  ], {
    x: M + 0.3, y: 5.5, w: W - M * 2 - 0.6, h: 1.05, valign: "middle",
    fontFace: FB, fontSize: 14, margin: 0,
  });

  s.addNotes("①〜③がHubSpot管理画面での準備、④⑤がコード側。合計30分程度。");
}

/* ---------------------------------------------------------- 3 STEP1 */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  stepHead(s, 1, "アカウントID（PORTAL_ID）を控える", "HubSpotにログインして、アドレスバーの数字を見るだけです");

  codeBox(s, M, 2.0, W - M * 2, 1.15, [
    { text: "https://app.hubspot.com/contacts/", options: { color: "A79A94" } },
    { text: "12345678", options: { color: GOLD, bold: true } },
    { text: "/objects/0-1/views/all/list", options: { color: "A79A94" } },
    { text: "\n", options: { breakLine: true } },
    { text: "                                 ↑ この8桁ほどの数字がPORTAL_ID", options: { color: "F2E8E2" } },
  ], { size: 14 });

  stepRow(s, 3.5, 1, "HubSpotにログインする", "どの画面でも構いません");
  stepRow(s, 4.35, 2, "ブラウザのアドレスバーを見る", "app.hubspot.com の次に出てくる数字がPORTAL_IDです");
  stepRow(s, 5.2, 3, "その数字をメモしておく", "右上のアカウントメニュー内「アカウントID」でも同じ値が確認できます");

  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 6.2, w: W - M * 2, h: 0.72, rectRadius: 0.1,
    fill: { color: CREAM }, line: { color: "EADFD5" },
  });
  s.addText("会社ごとに1つだけの番号です。パスワードのような機密情報ではなく、公開ページのHTMLにも載る値です。", {
    x: M + 0.3, y: 6.2, w: W - M * 2 - 0.6, h: 0.72, valign: "middle",
    fontFace: FB, fontSize: 13, color: MUTED, margin: 0,
  });
}

/* ---------------------------------------------------------- 4 STEP2 */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  stepHead(s, 2, "カスタムプロパティを1つ作る", "「希望プラン」を保存する入れ物がHubSpotの標準にはないため、先に作ります");

  stepRow(s, 2.0, 1, "設定（歯車アイコン）→ データ管理 → プロパティ", "画面右上の歯車から入ります");
  stepRow(s, 2.85, 2, "「プロパティを作成」をクリック", "オブジェクトの種類は「コンタクト」を選びます");
  stepRow(s, 3.7, 3, "内部名を desired_plan にする", "ラベルは「希望プラン」など日本語で構いません。内部名だけ厳密に合わせます");
  stepRow(s, 4.55, 4, "フィールドタイプは「1行テキスト」", "ドロップダウンにする場合は選択肢をプラン名と完全一致させます");

  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.5, w: 6.0, h: 1.45, rectRadius: 0.1,
    fill: { color: "FDF6EC" }, line: { color: GOLD, width: 1.5 },
  });
  s.addText([
    { text: "作りたくない場合\n", options: { bold: true, color: BROWN, fontSize: 14, breakLine: true } },
    { text: "514行目の plan: 'desired_plan' を削除すればOK。希望プランが記録されなくなるだけで、他の4項目は届きます。", options: { color: INK, fontSize: 13 } },
  ], {
    x: M + 0.28, y: 5.62, w: 5.44, h: 1.2, valign: "top", fontFace: FB, margin: 0,
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: M + 6.4, y: 5.5, w: W - M * 2 - 6.4, h: 1.45, rectRadius: 0.1,
    fill: { color: "FCEDE9" }, line: { color: TERRA, width: 1.5 },
  });
  s.addText([
    { text: "作らずに送信すると\n", options: { bold: true, color: TERRA, fontSize: 14, breakLine: true } },
    { text: "希望プランの項目だけエラーで弾かれます。他の4項目も含めて送信自体が失敗する場合があります。", options: { color: INK, fontSize: 13 } },
  ], {
    x: M + 6.68, y: 5.62, w: W - M * 2 - 6.96, h: 1.2, valign: "top", fontFace: FB, margin: 0,
  });
}

/* ---------------------------------------------------------- 5 STEP3 */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  stepHead(s, 3, "フォームを作って FORM_GUID を控える", "マーケティング → リードの獲得 → フォーム →「フォームを作成」");

  s.addText("LPの入力欄5つと、HubSpot側の項目の対応", {
    x: M, y: 1.92, w: 7.3, h: 0.36, fontFace: FH, fontSize: 15, bold: true, color: BROWN, margin: 0,
  });

  const rows = [
    ["LPの入力欄", "HubSpotの内部名", true],
    ["お名前", "lastname", false],
    ["電話番号", "phone", false],
    ["メールアドレス", "email", false],
    ["ご住所", "city", false],
    ["希望プラン", "desired_plan", false],
  ];
  rows.forEach((r, i) => {
    const y = 2.36 + i * 0.5;
    s.addShape(pres.ShapeType.rect, {
      x: M, y, w: 7.3, h: 0.5,
      fill: { color: r[2] ? BROWN : (i % 2 ? WHITE : CREAM) },
      line: { color: "EADFD5", width: 0.75 },
    });
    s.addText(r[0], {
      x: M + 0.25, y, w: 3.3, h: 0.5, valign: "middle",
      fontFace: FB, fontSize: 14, bold: r[2], color: r[2] ? WHITE : INK, margin: 0,
    });
    s.addText(r[1], {
      x: M + 3.7, y, w: 3.4, h: 0.5, valign: "middle",
      fontFace: r[2] ? FB : "Courier New", fontSize: r[2] ? 14 : 13,
      bold: r[2], color: r[2] ? WHITE : TERRA, margin: 0,
    });
  });

  s.addText("フォームを開いてアドレスバーを見る", {
    x: M + 7.75, y: 1.92, w: W - M * 2 - 7.75, h: 0.36,
    fontFace: FH, fontSize: 15, bold: true, color: BROWN, margin: 0,
  });
  codeBox(s, M + 7.75, 2.36, W - M * 2 - 7.75, 1.75, [
    { text: "https://app.hubspot.com/forms/\n", options: { color: "A79A94", breakLine: true } },
    { text: "12345678", options: { color: GOLD, bold: true } },
    { text: "/", options: { color: "A79A94" } },
    { text: "a1b2c3d4-0000-\n1111-2222-333344445555", options: { color: "7FD1B9", bold: true, breakLine: true } },
    { text: "/edit", options: { color: "A79A94" } },
  ], { size: 12 });

  s.addText([
    { text: "■ ", options: { color: GOLD } },
    { text: "短い数字 = PORTAL_ID\n", options: { color: INK, breakLine: true } },
    { text: "■ ", options: { color: "7FD1B9" } },
    { text: "ハイフン入りの36文字 = FORM_GUID", options: { color: INK } },
  ], {
    x: M + 7.75, y: 4.25, w: W - M * 2 - 7.75, h: 0.9, valign: "top",
    fontFace: FB, fontSize: 13, lineSpacing: 21, margin: 0,
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: M + 7.75, y: 5.3, w: W - M * 2 - 7.75, h: 1.62, rectRadius: 0.1,
    fill: { color: "FDF6EC" }, line: { color: GOLD, width: 1.5 },
  });
  s.addText([
    { text: "このURL 1本に両方写っています\n", options: { bold: true, color: BROWN, breakLine: true } },
    { text: "「共有 → 埋め込みコード」の portalId / formId でも同じ値が取れます。", options: { color: INK } },
  ], {
    x: M + 8.03, y: 5.44, w: W - M * 2 - 8.31, h: 1.34, valign: "top",
    fontFace: FB, fontSize: 13, margin: 0,
  });
}

/* ---------------------------------------------------------- 6 STEP4 */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  stepHead(s, 4, "デザインマネージャーにコードを貼る", "受け取ったHTML（814KB・575行）を、まるごと貼り付けます");

  stepRow(s, 2.0, 1, "設定 → コンテンツ → デザインマネージャー を開く", "画面右上の歯車から入ります");
  stepRow(s, 2.85, 2, "ファイル → 新規ファイル → HTML + HubL テンプレート", "ページテンプレートとして作成します（例：ieomoi.html）");
  stepRow(s, 3.7, 3, "エディタの中身を全部消す", "Ctrl + A →  Delete。既存のコードは1行も残しません");
  stepRow(s, 4.55, 4, "受け取ったコードを貼り付ける", "Ctrl + V。画像のアップロードもURLの書き換えも不要です");

  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.5, w: W - M * 2, h: 1.45, rectRadius: 0.1,
    fill: { color: "FCEDE9" }, line: { color: TERRA, width: 1.5 },
  });
  s.addText([
    { text: "先頭の4行を消さないでください\n", options: { bold: true, color: TERRA, fontSize: 15, breakLine: true } },
    { text: "templateType: page  というコメントがあることで、HubSpotがページテンプレートとして認識します。消すとテンプレート一覧に出てこなくなります。", options: { color: INK, fontSize: 13.5 } },
  ], {
    x: M + 0.3, y: 5.64, w: W - M * 2 - 0.6, h: 1.17, valign: "top", fontFace: FB, margin: 0,
  });
}

/* ---------------------------------------------------------- 7 STEP5 */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  stepHead(s, 5, "2行だけ書き換えて公開する", "貼ったコードの21行目と22行目。控えた値をクォートの中に入れます");

  s.addText("書き換え前", {
    x: M, y: 1.95, w: 5.9, h: 0.34, fontFace: FH, fontSize: 14, bold: true, color: MUTED, margin: 0,
  });
  codeBox(s, M, 2.32, 5.9, 1.15, [
    { text: "var IEOMOI_PORTAL_ID = ", options: { color: "C7BAB4" } },
    { text: "''", options: { color: "F08A6E" } },
    { text: ";\n", options: { color: "C7BAB4", breakLine: true } },
    { text: "var IEOMOI_FORM_GUID = ", options: { color: "C7BAB4" } },
    { text: "''", options: { color: "F08A6E" } },
    { text: ";", options: { color: "C7BAB4" } },
  ], { size: 13, fill: "463733", line: "463733" });

  s.addShape(pres.ShapeType.chevron, {
    x: 6.28, y: 2.62, w: 0.62, h: 0.55, fill: { color: TERRA }, line: { color: TERRA },
  });

  s.addText("書き換え後", {
    x: 7.15, y: 1.95, w: 5.6, h: 0.34, fontFace: FH, fontSize: 14, bold: true, color: TERRA, margin: 0,
  });
  codeBox(s, 7.15, 2.32, W - M - 7.15, 1.15, [
    { text: "var IEOMOI_PORTAL_ID = ", options: { color: "C7BAB4" } },
    { text: "'12345678'", options: { color: GOLD, bold: true } },
    { text: ";\n", options: { color: "C7BAB4", breakLine: true } },
    { text: "var IEOMOI_FORM_GUID = ", options: { color: "C7BAB4" } },
    { text: "'a1b2c3d4-…'", options: { color: "7FD1B9", bold: true } },
    { text: ";", options: { color: "C7BAB4" } },
  ], { size: 13 });

  s.addText("よくある間違い", {
    x: M, y: 3.75, w: 6.0, h: 0.36, fontFace: FH, fontSize: 15, bold: true, color: BROWN, margin: 0,
  });
  const ng = [
    ["= 12345678;", "クォートを消してしまった。数字でも ' ' で囲んだままにします"],
    ["= '\"12345678\"';", "HubSpotの \" ごとコピーした。中の文字だけを入れます"],
    ["2つを入れ違えた", "短い数字がPORTAL_ID、長いハイフン入りがFORM_GUID"],
  ];
  ng.forEach((r, i) => {
    const y = 4.2 + i * 0.62;
    s.addText("✕", {
      x: M, y, w: 0.36, h: 0.5, align: "center", valign: "middle",
      fontFace: FB, fontSize: 15, bold: true, color: TERRA, margin: 0,
    });
    s.addText(r[0], {
      x: M + 0.4, y, w: 2.5, h: 0.5, valign: "middle",
      fontFace: "Courier New", fontSize: 12, color: INK, margin: 0,
    });
    s.addText(r[1], {
      x: M + 3.0, y, w: 3.4, h: 0.5, valign: "middle",
      fontFace: FB, fontSize: 11.5, color: MUTED, margin: 0,
    });
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: 7.15, y: 3.75, w: W - M - 7.15, h: 3.05, rectRadius: 0.1,
    fill: { color: CREAM }, line: { color: "EADFD5" },
  });
  s.addText("最後に", {
    x: 7.43, y: 3.92, w: W - M - 7.71, h: 0.36, fontFace: FH, fontSize: 15, bold: true, color: BROWN, margin: 0,
  });
  s.addText([
    { text: "1.  右上の「変更を公開」を押す", options: { breakLine: true } },
    { text: "2.  マーケティング → ウェブサイト →", options: { breakLine: true } },
    { text: "      ランディングページ で新規作成", options: { breakLine: true } },
    { text: "3.  テンプレートに ieomoi.html を選ぶ", options: { breakLine: true } },
    { text: "4.  URLとページタイトルを設定して公開", options: {} },
  ], {
    x: 7.43, y: 4.36, w: W - M - 7.71, h: 2.3, valign: "top",
    fontFace: FB, fontSize: 13, color: INK, lineSpacing: 22, paraSpaceAfter: 8, margin: 0,
  });
}

/* ---------------------------------------------------------- 8 チェック */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addText("公開前の確認リスト", {
    x: M, y: 0.5, w: 8, h: 0.7, fontFace: FH, fontSize: 32, bold: true, color: BROWN, margin: 0,
  });
  s.addText("実際に自分でテスト送信して、コンタクトに登録されるところまで確認します", {
    x: M, y: 1.22, w: W - M * 2, h: 0.4, fontFace: FB, fontSize: 15, color: MUTED, margin: 0,
  });

  const checks = [
    ["表示", "パソコンとスマホの両方で開き、画像が3枚とも出るか"],
    ["表示", "申し込みフォームの入力欄5つが、画像の枠にぴったり重なっているか"],
    ["送信", "実際に自分の情報を入れて送信し、完了メッセージが出るか"],
    ["送信", "HubSpotのコンタクト一覧に、その内容が登録されているか"],
    ["送信", "希望プランが desired_plan に入っているか"],
    ["導線", "プラン別のボタンから進むと、希望プランが選択済みになっているか"],
  ];
  checks.forEach((c, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = M + col * 6.15, y = 1.95 + row * 1.13;
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w: 5.95, h: 0.95, rectRadius: 0.09,
      fill: { color: WHITE }, line: { color: "EADFD5", width: 1.25 }, shadow: shadow(),
    });
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.24, y: y + 0.26, w: 0.72, h: 0.43, rectRadius: 0.09,
      fill: { color: c[0] === "送信" ? TERRA : (c[0] === "表示" ? BROWN : GOLD) },
      line: { color: c[0] === "送信" ? TERRA : (c[0] === "表示" ? BROWN : GOLD) },
    });
    s.addText(c[0], {
      x: x + 0.24, y: y + 0.26, w: 0.72, h: 0.43, align: "center", valign: "middle",
      fontFace: FB, fontSize: 11, bold: true, color: WHITE, margin: 0,
    });
    s.addText(c[1], {
      x: x + 1.1, y, w: 4.7, h: 0.95, valign: "middle",
      fontFace: FB, fontSize: 13, color: INK, margin: 0,
    });
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.6, w: W - M * 2, h: 1.3, rectRadius: 0.1,
    fill: { color: "FDF6EC" }, line: { color: GOLD, width: 1.5 },
  });
  s.addText([
    { text: "「送信先が未設定です」と赤字が出たら\n", options: { bold: true, color: BROWN, breakLine: true } },
    { text: "21行目・22行目のどちらかが空のままです。公開後のページでも、ブラウザで表示 → ソースを見れば入力済みか確認できます。", options: { color: INK } },
  ], {
    x: M + 0.3, y: 5.74, w: W - M * 2 - 0.6, h: 1.02, valign: "top",
    fontFace: FB, fontSize: 13.5, margin: 0,
  });
}

/* ---------------------------------------------------------- 9 まとめ */
{
  const s = pres.addSlide();
  s.background = { color: BROWN };

  s.addShape(pres.ShapeType.ellipse, {
    x: -1.9, y: 4.6, w: 5.2, h: 5.2, fill: { color: "6E4A3C" }, line: { color: "6E4A3C" },
  });

  s.addText("結局、用意するのは3つだけ", {
    x: M + 0.2, y: 0.85, w: 11.5, h: 0.75, fontFace: FH, fontSize: 34, bold: true, color: WHITE, margin: 0,
  });

  const items = [
    ["PORTAL_ID", "アカウントの番号", "アドレスバーの数字"],
    ["FORM_GUID", "フォームの番号", "フォームを作ると発行される"],
    ["desired_plan", "希望プランの保存先", "カスタムプロパティを作成"],
  ];
  items.forEach((it, i) => {
    const y = 2.0 + i * 1.25;
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 0.2, y, w: 11.5, h: 1.05, rectRadius: 0.1,
      fill: { color: "6B473A" }, line: { color: "7C564A", width: 1.25 },
    });
    s.addText(it[0], {
      x: M + 0.55, y, w: 3.2, h: 1.05, valign: "middle",
      fontFace: "Courier New", fontSize: 17, bold: true, color: GOLD, margin: 0,
    });
    s.addText(it[1], {
      x: M + 3.9, y, w: 3.3, h: 1.05, valign: "middle",
      fontFace: FB, fontSize: 15, bold: true, color: WHITE, margin: 0,
    });
    s.addText(it[2], {
      x: M + 7.3, y, w: 4.2, h: 1.05, valign: "middle",
      fontFace: FB, fontSize: 13.5, color: "DCC8BD", margin: 0,
    });
  });

  s.addText("この3つが揃えば、あとはコードを貼って2行を書き換えるだけで公開できます。", {
    x: M + 0.2, y: 6.05, w: 11.5, h: 0.5, valign: "middle",
    fontFace: FB, fontSize: 15, color: "E6D5CB", margin: 0,
  });
}

pres.writeFile({ fileName: "/tmp/claude-0/-home-user--/a5cbf139-252c-54a1-922d-451785f728b5/scratchpad/hubspot-steps.pptx" })
  .then(f => console.log("wrote", f));
