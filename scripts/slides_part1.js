/* P1〜P16 */
module.exports = function (K) {
  const { pres, note, deco, NAVY, NAVY2, RED, HRED, YEL, WHITE, BLACK, PALE, PINK, GRAY2, INK, SOFT, F,
    SW, SH, ML, CW, contentSlide, card, band, bulletRow, statCard, table, th, lbl, val, blank,
    compareTable, agendaSlide, fitSize } = K;
  const CH1 = "1.長岡市のエネルギー事情と災害対策";

  /* ===== P1 タイトル ===== */
  {
    const s = pres.addSlide();
    s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: SW, h: SH, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    s.addShape(pres.ShapeType.rect, { x: 0, y: 1.98, w: SW, h: 4.16, fill: { color: NAVY2 }, line: { color: NAVY2, width: 0 } });
    K.deco(s, "title");
    s.addText("雪国でも大丈夫？　電気代が上がり続ける今こそ知りたい", {
      x: 0.4, y: 1.28, w: SW - 0.8, h: 0.5, fontFace: F, fontSize: 22, bold: true, color: YEL,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText([
      { text: "太陽光・蓄電池", options: { fontFace: F, fontSize: 64, bold: true, color: WHITE, breakLine: true } },
      { text: "補助金活用講座", options: { fontFace: F, fontSize: 64, bold: true, color: WHITE } },
    ], { x: 0.4, y: 2.18, w: SW - 0.8, h: 2.5, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.12 });
    s.addShape(pres.ShapeType.line, { x: 4.1, y: 4.82, w: 3.5, h: 0, line: { color: WHITE, width: 1.5 } });
    s.addText("2026年9月5日（土）", {
      x: 0.4, y: 4.96, w: SW - 0.8, h: 0.62, fontFace: F, fontSize: 30, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("～誰でもわかる！補助金活用法と業者の見極め方～", {
      x: 0.4, y: 6.32, w: SW - 0.8, h: 0.46, fontFace: F, fontSize: 20, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("主催：株式会社ヤシロ　／　会場：ハイブ長岡　会議室D", {
      x: 0.4, y: 6.8, w: SW - 0.8, h: 0.42, fontFace: F, fontSize: 17, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0,
    });
    s.addNotes("差し替え：背景写真（自社施工事例・長岡の街並みなど）を全面に敷き、現状のネイビー地と入れ替える。");
    note(1, "背景", "タイトル背景写真（現状はネイビー単色）");
  }

  /* ===== P2 会社概要 ===== */
  {
    const s = contentSlide({
      chapter: "はじめに", chip: "弊社のご紹介",
      lead: "長岡市を拠点とする工務店として、地域のお客様の住まいを支えてきた。",
      source: "出典：登記情報・自社公開情報",
      footer: "雪国の太陽光を、実体験からお伝えします",
      note: "差し替え：従業員数／年商／事業所・店舗／太陽光施工実績。会社名・代表者・設立・資本金・本社所在地・事業内容は記入済み。",
    });
    const rows = [
      [lbl("会社名"), val("株式会社ヤシロ")],
      [lbl("代表者"), val("代表取締役社長　屋代　健")],
      [lbl("設立"), val("1964年（昭和39年）1月10日")],
      [lbl("資本金"), val("1,000万円")],
      [lbl("本社所在地"), val("〒940-0004　新潟県長岡市高見町1050番地")],
      [lbl("事業内容"), val("建築（住宅の設計・施工）／不動産\n除雪機・農業資材の販売／太陽光発電")],
      [lbl("従業員数"), blank()],
      [lbl("年商"), blank()],
      [lbl("事業所・店舗"), blank()],
      [lbl("太陽光施工実績"), blank()],
    ];
    table(s, rows, ML, 1.78, 5.6, [1.66, 3.94], { rowH: 0.47, fontSize: 12.5 });
    s.addShape(pres.ShapeType.line, { x: 5.94, y: 1.78, w: 0, h: 4.9, line: { color: NAVY, width: 1.5 } });
    card(s, 6.16, 1.86, 5.37, 4.72, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("自給自足の時代を、技術と信頼で支える", {
      x: 6.34, y: 2.04, w: 5.01, h: 0.62, fontFace: F, fontSize: fitSize("自給自足の時代を、技術と信頼で支える", 4.9, 22, 15),
      bold: true, color: NAVY, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pres.ShapeType.line, { x: 6.34, y: 2.72, w: 5.01, h: 0, line: { color: NAVY, width: 1 } });
    s.addText(
      "2014年、柏崎・長岡で自社の太陽光発電所を立ち上げました。\n\n" +
      "積雪による破損、冬期の発電量低下。雪国ならではの課題に、実際の運用を通じて向き合ってきました。\n\n" +
      "そこで得た知見を、お客様の住まいに還元します。",
      { x: 6.34, y: 2.9, w: 5.01, h: 3.5, fontFace: F, fontSize: 16, bold: true, color: BLACK,
        align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.3 }
    );
    note(2, "表", "従業員数／年商／事業所・店舗／太陽光施工実績");
  }

  /* ===== P3 講師紹介 ===== */
  {
    const s = contentSlide({
      chapter: "はじめに", chip: "自己紹介",
      footer: "雪国での実践を重ねてきた立場から、正直にお話しします",
      note: "差し替え：経歴年表の空欄3段（年・内容）。2014年の1段は記入済み。",
    });
    card(s, ML, 1.42, 4.5, 5.2, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("株式会社ヤシロ", { x: 0.44, y: 1.78, w: 4.0, h: 0.5, fontFace: F, fontSize: 22, bold: true, color: NAVY, margin: 0, valign: "middle" });
    s.addText("代表取締役社長", { x: 0.44, y: 2.26, w: 4.0, h: 0.44, fontFace: F, fontSize: 18, bold: true, color: BLACK, margin: 0, valign: "middle" });
    s.addText("屋代　健", { x: 0.44, y: 2.78, w: 4.0, h: 0.9, fontFace: F, fontSize: 42, bold: true, color: BLACK, margin: 0, valign: "middle" });
    s.addShape(pres.ShapeType.line, { x: 0.44, y: 3.82, w: 4.0, h: 0, line: { color: NAVY, width: 1 } });
    s.addText(
      "2014年より柏崎・長岡で自社の太陽光発電所を運営。\n\n" +
      "積雪による破損や冬期の発電量低下といった、雪国特有の課題に現場で向き合ってきました。\n\n" +
      "本日は、その経験からお話しします。",
      { x: 0.44, y: 3.98, w: 4.0, h: 2.5, fontFace: F, fontSize: 14, bold: true, color: BLACK, margin: 0, valign: "top", lineSpacingMultiple: 1.3 }
    );
    band(s, 5.0, 1.42, 6.53, "経歴", { size: 20 });
    s.addShape(pres.ShapeType.line, { x: 5.42, y: 2.06, w: 0, h: 4.52, line: { color: NAVY, width: 2 } });
    const tl = [
      { y: 2.14, year: "", body: "" },
      { y: 3.24, year: "2014年", body: "柏崎・長岡で自社の太陽光発電所を立ち上げ。積雪による破損・冬期の発電量低下という課題に取り組む" },
      { y: 4.34, year: "", body: "" },
      { y: 5.44, year: "", body: "" },
    ];
    tl.forEach((r) => {
      s.addShape(pres.ShapeType.ellipse, { x: 5.28, y: r.y + 0.16, w: 0.28, h: 0.28, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
      s.addText(r.year, { x: 5.62, y: r.y - 0.04, w: 1.24, h: 0.5, fontFace: F, fontSize: 18, bold: true, color: NAVY, align: "left", valign: "middle", margin: 0 });
      s.addText(r.body, { x: 6.9, y: r.y - 0.12, w: 4.6, h: 0.94, fontFace: F, fontSize: 14, bold: true, color: BLACK, align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
      s.addShape(pres.ShapeType.line, { x: 5.62, y: r.y + 0.84, w: 5.88, h: 0, line: { color: "D9D9D9", width: 1 } });
    });
    note(3, "年表", "経歴（年・内容）×3段");
  }

  /* ===== P4 知る場 ===== */
  {
    const s = contentSlide({
      chapter: "はじめに",
      footer: [{ t: "この10年で、雪国の太陽光は" }, { t: "大きく変わりました", hl: true }],
    });
    band(s, ML, 0.82, CW, "本日は「売る」ための場ではなく、「知る」ための勉強会です", { h: 0.7, size: 28 });
    s.addText("長岡で、こんな声をよく耳にします", {
      x: ML, y: 1.68, w: CW, h: 0.6, fontFace: F, fontSize: 30, bold: true, color: BLACK,
      align: "center", valign: "middle", margin: 0,
    });
    const voices = [
      "「長岡の雪じゃ\n発電しないでしょ？」", "「屋根が雪で\n傷むのでは？」", "「訪問販売が怖い」",
      "「業者によって\n言うことが違う」", "「どのメーカーが\nいいの？」", "「いつ導入するのが\nお得なの？」",
    ];
    const bw = 3.45, bh = 1.76, gx = 0.36, gy = 0.36, x0 = (SW - (bw * 3 + gx * 2)) / 2, y0 = 2.48;
    voices.forEach((t, i) => {
      const x = x0 + (i % 3) * (bw + gx), y = y0 + Math.floor(i / 3) * (bh + gy);
      card(s, x, y, bw, bh, { fill: PALE, line: NAVY, lw: 1.5 });
      s.addText(t, { x: x + 0.12, y, w: bw - 0.24, h: bh, fontFace: F, fontSize: 18, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.18 });
    });
  }

  /* ===== P5 ゴール4つ ===== */
  {
    const s = contentSlide({ chapter: "はじめに", footer: "本日は何卒よろしくお願い申し上げます" });
    band(s, ML, 0.82, CW, "講座を聞き終えた後、こうなっていただきたい", { h: 0.7, size: 28 });
    const goals = [
      "長岡市のエネルギー事情を\n正確に把握している",
      "メリット・デメリットを\n正確に把握している",
      "信頼できる業者の\n見極め方を知っている",
      "「我が家の場合はどうなのか」\nを知りたいと思っている",
    ];
    const cw = 4.9, ch = 2.06, gx = 0.5, gy = 0.66, x0 = (SW - (cw * 2 + gx)) / 2, y0 = 2.04;
    goals.forEach((t, i) => {
      const x = x0 + (i % 2) * (cw + gx), y = y0 + Math.floor(i / 2) * (ch + gy);
      card(s, x, y, cw, ch, { fill: WHITE, line: NAVY, lw: 2 });
      s.addShape(pres.ShapeType.ellipse, { x: x + cw / 2 - 0.36, y: y - 0.34, w: 0.72, h: 0.72, fill: { color: NAVY }, line: { color: WHITE, width: 2 } });
      s.addText(String(i + 1), { x: x + cw / 2 - 0.36, y: y - 0.34, w: 0.72, h: 0.72, fontFace: F, fontSize: 28, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      s.addText(t, { x: x + 0.2, y: y + 0.34, w: cw - 0.4, h: ch - 0.5, fontFace: F, fontSize: 20, bold: true, color: BLACK, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
    });
  }

  /* ===== P6 目次(1) ===== */
  agendaSlide(0);

  /* ===== P7 章イメージ ===== */
  {
    const s = contentSlide({ chapter: CH1, footer: "この章では、長岡の「今」をデータで確認します" });
    s.addText("電気代高騰で家計圧迫…！", {
      x: 0.4, y: 1.44, w: SW - 0.8, h: 1.16, fontFace: F, fontSize: 52, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    const items = [
      ["上がり続ける電気代", "東北電力は2023年6月、規制料金を平均25.47％値上げ。モデル世帯で月2,110円の増加。"],
      ["復旧が長引く停電", "中越地震では約11日。2022年12月の大雪では佐渡で9日間。雪国は復旧に時間がかかる。"],
      ["補助金頼みという構造", "国の負担軽減措置は期限つき。終われば請求額はそのまま跳ね上がる。"],
    ];
    const cw = 3.55, ch = 3.0, gx = 0.32, x0 = (SW - (cw * 3 + gx * 2)) / 2, y0 = 3.1;
    items.forEach((it, i) => {
      const x = x0 + i * (cw + gx);
      card(s, x, y0, cw, ch, { fill: WHITE, line: NAVY, lw: 2 });
      band(s, x, y0, cw, "①②③".charAt(i) + "　" + it[0], { h: 0.62, size: 16 });
      s.addText(it[1], { x: x + 0.22, y: y0 + 0.78, w: cw - 0.44, h: ch - 0.98, fontFace: F, fontSize: 15, bold: true, color: BLACK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.3 });
    });
  }

  /* ===== P8 東北電力の2023年値上げ ===== */
  {
    const s = contentSlide({
      chapter: CH1, chip: "東北電力の電気代",
      lead: "東北電力は2023年6月、規制料金を平均25.47％値上げした。",
      source: "出典：東北電力「小売規制料金見直しの概要」（2023年6月1日実施）／モデルケースは従量電灯B・30A・260kWh/月",
      footer: "使い方は変わっていないのに、請求額だけが増えました",
    });
    band(s, ML, 1.76, CW, "モデルケース（従量電灯B・30A・月260kWh）の月額", { h: 0.5, size: 19 });
    statCard(s, 0.9, 2.46, 2.86, 1.9, "値上げ前", "8,032", "円/月", { valueSize: 34, line: GRAY2, valueColor: INK, labelColor: INK, fill: SOFT });
    s.addShape(pres.ShapeType.rightArrow, { x: 3.98, y: 3.12, w: 0.7, h: 0.58, fill: { color: RED }, line: { color: RED, width: 0 } });
    statCard(s, 4.9, 2.46, 2.86, 1.9, "値上げ後", "10,142", "円/月", { valueSize: 34, line: RED, valueColor: RED, labelColor: RED, fill: PINK });
    statCard(s, 7.98, 2.46, 2.86, 1.9, "1か月あたりの増加", "＋2,110", "円", { valueSize: 34, line: NAVY });
    card(s, 0.9, 4.62, 9.94, 1.98, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("年間に置きかえると", { x: 1.14, y: 4.8, w: 3.2, h: 0.46, fontFace: F, fontSize: 18, bold: true, color: NAVY, align: "left", valign: "middle", margin: 0 });
    s.addText([
      { text: "＋2,110円 × 12か月 ＝ ", options: { fontFace: F, fontSize: 22, bold: true, color: BLACK } },
      { text: "年間 約25,000円", options: { fontFace: F, fontSize: 34, bold: true, color: RED } },
      { text: " の負担増", options: { fontFace: F, fontSize: 22, bold: true, color: BLACK } },
    ], { x: 1.14, y: 5.26, w: 9.46, h: 0.7, align: "left", valign: "middle", margin: 0 });
    s.addText("これは規制料金の改定分のみ。燃料費調整額や再エネ賦課金は、これとは別に毎月加算されます。", {
      x: 1.14, y: 5.98, w: 9.46, h: 0.44, fontFace: F, fontSize: 15, bold: true, color: INK, align: "left", valign: "middle", margin: 0,
    });
  }

  /* ===== P9 再エネ賦課金の推移（実データ） ===== */
  {
    const s = contentSlide({
      chapter: CH1, chip: "電気代の中身",
      lead: "電気代は「使った分だけ」ではない。再エネ賦課金は制度開始から19倍になっている。",
      source: "出典：経済産業省 資源エネルギー庁「再生可能エネルギー発電促進賦課金単価」各年度公表値",
      footer: "電気代は「上がり続ける」前提で考える時代です",
    });
    const labels = ["2012", "2013", "2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026"];
    const values = [0.22, 0.35, 0.75, 1.58, 2.25, 2.64, 2.90, 2.95, 2.98, 3.36, 3.45, 1.40, 3.49, 3.98, 4.18];
    s.addChart(pres.ChartType.bar, [{ name: "再エネ賦課金", labels, values }], {
      x: 0.42, y: 2.12, w: 7.9, h: 4.48, barDir: "col", barGapWidthPct: 45,
      showTitle: true, title: "再エネ賦課金単価の推移（円/kWh・年度）", titleFontFace: F, titleFontSize: 15, titleColor: NAVY,
      chartColors: [NAVY], showLegend: false,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontFace: F, dataLabelFontSize: 9, dataLabelColor: INK, dataLabelFormatCode: "0.00",
      catAxisLabelFontFace: F, catAxisLabelFontSize: 10, catAxisLabelColor: INK,
      valAxisLabelFontFace: F, valAxisLabelFontSize: 11, valAxisLabelColor: INK, valAxisMaxVal: 5,
      valGridLine: { color: "D9D9D9", size: 1 }, catGridLine: { style: "none" },
    });
    card(s, 8.5, 2.12, 3.03, 4.48, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("賦課金は\n電気を使う限りかかる", {
      x: 8.66, y: 2.3, w: 2.71, h: 0.86, fontFace: F, fontSize: 18, bold: true, color: NAVY,
      align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.15,
    });
    s.addShape(pres.ShapeType.line, { x: 8.66, y: 3.24, w: 2.71, h: 0, line: { color: NAVY, width: 1 } });
    s.addText([
      { text: "2012年度　", options: { fontFace: F, fontSize: 14, bold: true, color: INK } },
      { text: "0.22円", options: { fontFace: F, fontSize: 20, bold: true, color: INK, breakLine: true } },
      { text: "2026年度　", options: { fontFace: F, fontSize: 14, bold: true, color: RED } },
      { text: "4.18円", options: { fontFace: F, fontSize: 28, bold: true, color: RED, breakLine: true } },
      { text: "\n月400kWh使うご家庭なら、賦課金だけで月1,672円。年間で約20,000円になります。", options: { fontFace: F, fontSize: 14, bold: true, color: BLACK } },
    ], { x: 8.66, y: 3.4, w: 2.71, h: 3.0, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.25 });
  }

  /* ===== P10 補助金頼みの構造リスク ===== */
  {
    const s = contentSlide({
      chapter: CH1, chip: "電気代を取り巻く構造",
      lead: "電気料金の負担は、国の負担軽減措置に支えられてきた。しかし措置は期限つきである。",
      source: "※国の電気・ガス料金負担軽減支援事業の実施状況は変動します。最新情報をご確認ください",
      footer: [{ t: "補助金が前提の家計が、" }, { t: "いちばん危ない", hl: true }, { t: "かもしれません" }],
    });
    card(s, 0.42, 2.16, 5.36, 4.38, { fill: PALE, line: NAVY, lw: 2 });
    band(s, 0.42, 2.16, 5.36, "補助があるうちは見えないこと", { h: 0.58, size: 18 });
    [["補助は「期限つき」", "延長されるかどうかは、そのつど政策判断で決まる"],
     ["終了・縮小のたびに請求額が跳ね上がる", "使い方は変えていないのに、請求書だけが増える"],
     ["補助が前提の家計は、いちばん揺れやすい", "支出の前提を、外部要因に握られている状態"]]
      .forEach((it, i) => {
        const y = 2.94 + i * 1.18;
        s.addText("●　" + it[0], { x: 0.66, y, w: 4.9, h: 0.5, fontFace: F, fontSize: fitSize("●　" + it[0], 4.8, 17, 13), bold: true, color: NAVY, align: "left", valign: "middle", margin: 0 });
        s.addText(it[1], { x: 0.94, y: y + 0.48, w: 4.62, h: 0.6, fontFace: F, fontSize: 14, bold: true, color: BLACK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
      });
    card(s, 5.92, 2.16, 5.36, 4.38, { fill: WHITE, line: RED, lw: 2 });
    band(s, 5.92, 2.16, 5.36, "補助が終わったとき、どうなるか", { h: 0.58, size: 18, fill: RED });
    s.addText("補助あり", { x: 6.3, y: 3.06, w: 2.0, h: 0.44, fontFace: F, fontSize: 17, bold: true, color: INK, align: "center", valign: "middle", margin: 0 });
    card(s, 6.3, 3.54, 2.0, 1.2, { fill: SOFT, line: GRAY2, lw: 1.5 });
    s.addText("請求額", { x: 6.3, y: 3.54, w: 2.0, h: 1.2, fontFace: F, fontSize: 18, bold: true, color: INK, align: "center", valign: "middle", margin: 0 });
    s.addShape(pres.ShapeType.rightArrow, { x: 8.44, y: 3.92, w: 0.6, h: 0.46, fill: { color: RED }, line: { color: RED, width: 0 } });
    s.addText("補助が終了", { x: 9.1, y: 3.06, w: 2.0, h: 0.44, fontFace: F, fontSize: 17, bold: true, color: RED, align: "center", valign: "middle", margin: 0 });
    card(s, 9.1, 3.06, 2.0, 1.68, { fill: PINK, line: RED, lw: 2 });
    s.addText("請求額", { x: 9.1, y: 3.06, w: 2.0, h: 1.68, fontFace: F, fontSize: 22, bold: true, color: RED, align: "center", valign: "middle", margin: 0 });
    s.addText("使い方は変えていないのに、支払いだけが増える。\n電気代そのものを減らす手段を持っておくことが、いちばん確実な備えになります。", {
      x: 6.16, y: 5.0, w: 4.9, h: 1.4, fontFace: F, fontSize: 15, bold: true, color: BLACK,
      align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.3,
    });
  }

  /* ===== P11 柏崎刈羽6号機 ===== */
  {
    const s = contentSlide({
      chapter: CH1, chip: "電力供給の前提",
      lead: "柏崎刈羽原子力発電所6号機は、2026年4月16日に営業運転を開始した。",
      source: "出典：資源エネルギー庁・東京電力の公表資料（2026年1月21日再稼働／同年4月16日営業運転開始）",
      footer: "再稼働＝長岡の電気代が下がる、とは限りません",
    });
    card(s, 0.42, 2.16, 5.36, 4.38, { fill: WHITE, line: NAVY, lw: 2 });
    band(s, 0.42, 2.16, 5.36, "【事実の整理】", { h: 0.58, size: 18 });
    ["柏崎刈羽原子力発電所は、東京電力の発電所である",
     "立地は柏崎市・刈羽村（長岡市の隣接地域）",
     "6号機は2026年4月16日に営業運転を開始した"]
      .forEach((t, i) => bulletRow(s, 0.66, 2.96 + i * 1.18, 4.88, 1.04, i + 1, t, NAVY, 15));
    card(s, 5.92, 2.16, 5.36, 4.38, { fill: PALE, line: NAVY, lw: 2 });
    band(s, 5.92, 2.16, 5.36, "【長岡市への影響】", { h: 0.58, size: 18 });
    ["長岡市は東北電力の供給エリアである",
     "東京電力の再稼働が、長岡の電気料金を直接下げるとは限らない",
     "原発の賛否は、本講座では扱いません"]
      .forEach((t, i) => bulletRow(s, 6.16, 2.96 + i * 1.18, 4.88, 1.04, i + 1, t, NAVY, 15));
  }

  /* ===== P12 中越地震 ===== */
  {
    const s = contentSlide({
      chapter: CH1, chip: "災害・停電対策",
      lead: "2004年の新潟県中越地震では、長岡市を含む34市町村で停電が発生した。",
      source: "出典：内閣府「平成16年（2004年）新潟県中越地震について」／新潟県公表資料",
      footer: "停電はいつ・どこで起きるか分かりません",
    });
    band(s, ML, 1.76, CW, "新潟県中越地震（2004年10月23日発生・最大震度7）", { h: 0.5, size: 20 });
    const stats = [
      ["停電した戸数", "約30.1", "万戸"],
      ["停電した市町村", "34", "市町村"],
      ["ほぼ全域の復旧まで", "約11", "日"],
    ];
    const cw = 3.4, ch = 2.02, gx = 0.42, x0 = (SW - (cw * 3 + gx * 2)) / 2, y0 = 2.48;
    stats.forEach((st, i) => statCard(s, x0 + i * (cw + gx), y0, cw, ch, st[0], st[1], st[2], { valueSize: 46, valueColor: i === 2 ? RED : NAVY, line: i === 2 ? RED : NAVY }));
    card(s, x0, 4.86, cw * 3 + gx * 2, 1.7, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("長岡市は、この地震の震源域の中心にありました。\n「うちの地域は大丈夫」と言える場所は、どこにもありません。", {
      x: x0 + 0.24, y: 4.86, w: cw * 3 + gx * 2 - 0.48, h: 1.7, fontFace: F, fontSize: 19, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.3,
    });
  }

  /* ===== P13 停電実績表 ===== */
  {
    const s = contentSlide({
      chapter: CH1, chip: "災害・停電対策",
      lead: "過去の災害では、復旧までに1週間以上を要したケースも少なくない。",
      source: "出典：内閣府・各電力会社の公表資料をもとに作成",
      footer: "災害大国だからこそ、もしもの備えが必要です",
    });
    band(s, ML, 1.76, 5.7, "主な災害時の停電規模と復旧日数", { h: 0.46, size: 18 });
    const hi = { fill: { color: PINK }, color: RED, bold: true };
    const rows = [
      [th("発生年"), th("災害名"), th("停電（約）"), th("復旧（約）")],
      [val("1995年", { align: "center" }), val("阪神淡路大震災"), val("260万戸", { align: "right" }), val("5日", { align: "right" })],
      [val("2003年", { align: "center" }), val("十勝沖地震"), val("38万戸", { align: "right" }), val("1日", { align: "right" })],
      [val("2004年", Object.assign({ align: "center" }, hi)), val("新潟県中越地震", hi), val("30万戸", Object.assign({ align: "right" }, hi)), val("11日", Object.assign({ align: "right" }, hi))],
      [val("2011年", { align: "center" }), val("東日本大震災"), val("466万戸", { align: "right" }), val("8日", { align: "right" })],
      [val("2016年", { align: "center" }), val("熊本地震"), val("48万戸", { align: "right" }), val("5日", { align: "right" })],
      [val("2018年", { align: "center" }), val("台風21号"), val("261万戸", { align: "right" }), val("4日", { align: "right" })],
      [val("2018年", { align: "center" }), val("胆振東部地震"), val("295万戸", { align: "right" }), val("2日", { align: "right" })],
      [val("2024年", { align: "center" }), val("能登半島地震"), val("4万戸", { align: "right" }), val("30日", { align: "right" })],
    ];
    table(s, rows, ML, 2.3, 5.7, [1.1, 2.16, 1.26, 1.18], { rowH: 0.47, fontSize: 14 });
    card(s, 6.16, 2.3, 5.37, 4.24, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("件数の多さより、長さを見てください", {
      x: 6.34, y: 2.48, w: 5.01, h: 0.56, fontFace: F, fontSize: 20, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    statCard(s, 6.5, 3.2, 2.34, 1.5, "中越地震（2004）", "11", "日", { valueSize: 34, valueColor: RED, line: RED, labelColor: RED });
    statCard(s, 8.94, 3.2, 2.34, 1.5, "能登半島地震（2024）", "30", "日", { valueSize: 34, valueColor: RED, line: RED, labelColor: RED });
    s.addText(
      "能登半島地震は停電4万戸に対し、復旧は30日。\n" +
      "戸数が少なくても、道路が寸断されれば復旧は進みません。\n\n" +
      "雪国では、そこに「除雪」という条件が加わります。",
      { x: 6.5, y: 4.92, w: 4.78, h: 1.5, fontFace: F, fontSize: 14, bold: true, color: BLACK,
        align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.3 }
    );
  }

  /* ===== P14 2022年12月の大雪 ===== */
  {
    const s = contentSlide({
      chapter: CH1, chip: "災害・停電対策",
      lead: "停電の原因は地震だけではない。2022年12月18日の大雪では、県内各地で停電が発生した。",
      source: "出典：新潟県「今後の停電対策に向けた検討会」資料（令和4年12月18日の大雪による停電）",
      footer: "雪国の停電は、復旧に時間がかかります",
    });
    const items = [
      ["湿った雪が電線に着雪", "気温が高い時期の湿った雪は電線に付きやすく、重みで倒木・断線を招いた"],
      ["佐渡市の一部で9日間の停電", "離島という条件に加え、現場へのアクセス自体が雪で断たれた"],
      ["原因箇所が広く分散する", "地震と違い、被害が県内各地に散らばるため復旧に手間がかかる"],
    ];
    const cw = 3.55, ch = 2.4, gx = 0.32, x0 = (SW - (cw * 3 + gx * 2)) / 2, y0 = 2.2;
    items.forEach((it, i) => {
      const x = x0 + i * (cw + gx);
      card(s, x, y0, cw, ch, { fill: WHITE, line: NAVY, lw: 2 });
      band(s, x, y0, cw, it[0], { h: 0.62, size: 15 });
      s.addText(it[1], { x: x + 0.22, y: y0 + 0.76, w: cw - 0.44, h: ch - 0.94, fontFace: F, fontSize: 15, bold: true, color: BLACK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.3 });
    });
    card(s, x0, 4.86, cw * 3 + gx * 2, 1.66, { fill: PINK, line: RED, lw: 2 });
    s.addText("雪国特有の「3つの重なり」", {
      x: x0, y: 4.98, w: cw * 3 + gx * 2, h: 0.44, fontFace: F, fontSize: 18, bold: true, color: RED,
      align: "center", valign: "middle", margin: 0,
    });
    const three = ["停電", "低気温", "除雪の遅れ"];
    const tw = 2.2, tgx = 0.86, tx0 = x0 + (cw * 3 + gx * 2 - (tw * 3 + tgx * 2)) / 2;
    three.forEach((t, i) => {
      const x = tx0 + i * (tw + tgx);
      s.addText(t, { x, y: 5.48, w: tw, h: 0.76, fontFace: F, fontSize: 22, bold: true, color: WHITE, fill: { color: RED }, align: "center", valign: "middle", margin: 0 });
      if (i < 2) s.addText("＋", { x: x + tw, y: 5.48, w: tgx, h: 0.76, fontFace: F, fontSize: 24, bold: true, color: RED, align: "center", valign: "middle", margin: 0 });
    });
  }

  /* ===== P15 厳冬期に暖房が止まる ===== */
  {
    const s = contentSlide({
      chapter: CH1, chip: "災害・停電対策",
      lead: "厳冬期に停電が起これば、暖房をはじめとする生活インフラが同時に止まる。",
      footer: [{ t: "雪国の停電は、" }, { t: "命に直結します", hl: true }],
    });
    s.addText("厳冬期に停電したら、家の中で何が起きるか", {
      x: ML, y: 1.74, w: CW, h: 0.54, fontFace: F, fontSize: 26, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    const items = [
      ["暖房", "エアコンも\nファンヒーターも停止\n室温が一気に下がる"],
      ["給湯", "お湯が出ない\n入浴も洗い物も\nできない"],
      ["給水", "ポンプ停止と配管凍結\n水そのものが\n止まる"],
      ["トイレ", "電動洗浄・ポンプ停止\n衛生環境が\n急激に悪化"],
      ["情報", "Wi-Fiも充電も不可\n復旧の情報が\n届かない"],
    ];
    const cw = 2.12, ch = 3.5, gx = 0.2, x0 = (SW - (cw * 5 + gx * 4)) / 2, y0 = 2.44;
    items.forEach((it, i) => {
      const x = x0 + i * (cw + gx);
      card(s, x, y0, cw, ch, { fill: WHITE, line: NAVY, lw: 2 });
      band(s, x, y0, cw, it[0], { h: 0.64, size: 20 });
      s.addText("停止", { x: x + 0.36, y: y0 + 0.82, w: cw - 0.72, h: 0.5, fontFace: F, fontSize: 18, bold: true, color: RED, fill: { color: PINK }, align: "center", valign: "middle", margin: 0 });
      s.addText(it[1], { x: x + 0.12, y: y0 + 1.46, w: cw - 0.24, h: ch - 1.62, fontFace: F, fontSize: 13.5, bold: true, color: BLACK, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.28 });
    });
    s.addText("氷点下の夜に、これらが同時に止まります。", {
      x: ML, y: 6.14, w: CW, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: RED,
      align: "center", valign: "middle", margin: 0,
    });
  }

  /* ===== P16 太陽光・蓄電池は当たり前 ===== */
  {
    const s = contentSlide({
      chapter: CH1, chip: "太陽光・蓄電池は当たり前",
      lead: "電気代高騰と災害への不安から、太陽光・蓄電池を備えた住宅の標準化が進んでいる。",
      footer: "太陽光・蓄電池は、贅沢品ではなく「備え」です",
    });
    compareTable(s, 0.42, 2.16, 10.85, "未導入のお家", "太陽光＋蓄電池のお家", [
      ["平常時の電気代", "使った分をすべて買う", "つくった電気から使う"],
      ["停電したとき", "家じゅうの電気が止まる", "蓄電池から電気を使える"],
      ["厳冬期の停電", "暖房も給湯も止まる", "最低限の暖をとれる"],
      ["電気代が上がったら", "そのまま負担が増える", "買う量が少ないほど影響が小さい"],
    ], { rowH: 0.82, labelW: 2.4 });
  }
};
