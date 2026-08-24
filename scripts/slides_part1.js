/* P1〜P15 */
module.exports = function (K) {
  const { pres, note, NAVY, RED, HRED, YEL, WHITE, BLACK, PALE, F, SW, SH, ML, CW,
    chapterTitle, chip, lead, source, footerBar, contentSlide, card, bulletRow,
    statCard, table, th, lbl, val, blank, agendaSlide } = K;

  /* ===== P1 タイトル ===== */
  {
    const s = pres.addSlide();
    s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: SW, h: SH, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    s.addShape(pres.ShapeType.rect, { x: 0, y: 2.02, w: SW, h: 4.22, fill: { color: "12386E" }, line: { color: "12386E", width: 0 } });
    s.addText("雪国でも大丈夫？　電気代が上がり続ける今こそ知りたい", {
      x: 0.4, y: 1.34, w: SW - 0.8, h: 0.5, fontFace: F, fontSize: 22, bold: true, color: YEL,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText([
      { text: "太陽光・蓄電池", options: { fontFace: F, fontSize: 66, bold: true, color: WHITE, breakLine: true } },
      { text: "補助金活用講座", options: { fontFace: F, fontSize: 66, bold: true, color: WHITE, breakLine: true } },
      { text: "2026年9月5日（土）", options: { fontFace: F, fontSize: 30, bold: true, color: WHITE } },
    ], { x: 0.4, y: 2.20, w: SW - 0.8, h: 3.86, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.12 });
    s.addText("～誰でもわかる！補助金活用法と業者の見極め方～", {
      x: 0.4, y: 6.34, w: SW - 0.8, h: 0.46, fontFace: F, fontSize: 20, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("主催：株式会社ヤシロ　／　会場：ハイブ長岡　会議室D", {
      x: 0.4, y: 6.78, w: SW - 0.8, h: 0.44, fontFace: F, fontSize: 18, bold: true, color: WHITE,
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
      footer: "雪国の太陽光を、机上の理論ではなく実体験からお伝えします",
      note: "差し替え：会社概要の各値（創業／設立／従業員数／年商／資本金／事業所・店舗／太陽光施工実績）。会社名・代表者・本社所在地は記入済み。",
    });
    const rows = [
      [lbl("会社名"), val("株式会社ヤシロ")],
      [lbl("代表者"), val("代表取締役社長　屋代　健")],
      [lbl("創業"), blank()],
      [lbl("設立"), blank()],
      [lbl("本社所在地"), val("〒940-0004　新潟県長岡市高見町1050番地")],
      [lbl("従業員数"), blank()],
      [lbl("年商"), blank()],
      [lbl("資本金"), blank()],
      [lbl("事業所・店舗"), blank()],
      [lbl("太陽光施工実績"), blank()],
    ];
    table(s, rows, ML, 1.78, 5.58, [1.62, 3.96], { rowH: 0.46, fontSize: 13 });
    s.addShape(pres.ShapeType.line, { x: 5.94, y: 1.78, w: 0, h: 4.9, line: { color: NAVY, width: 1.5 } });
    K.card(s, 6.16, 1.86, 5.37, 4.74, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("自給自足の時代を、技術と信頼で支える", {
      x: 6.36, y: 2.06, w: 4.97, h: 0.9, fontFace: F, fontSize: 24, bold: true, color: NAVY,
      align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.1,
    });
    s.addShape(pres.ShapeType.line, { x: 6.36, y: 3.02, w: 4.97, h: 0, line: { color: NAVY, width: 1 } });
    s.addText(
      "2014年、柏崎・長岡で自社の太陽光発電所を立ち上げました。\n\n" +
      "積雪による破損、冬期の発電量低下。雪国ならではの課題に、実際の運用を通じて向き合ってきました。\n\n" +
      "そこで得た知見を、お客様の住まいに還元します。",
      { x: 6.36, y: 3.18, w: 4.97, h: 3.3, fontFace: F, fontSize: 16, bold: true, color: BLACK,
        align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.25 }
    );
    note(2, "表", "会社概要（創業／設立／従業員数／年商／資本金／事業所／施工実績）");
  }

  /* ===== P3 講師紹介 ===== */
  {
    const s = contentSlide({
      chapter: "はじめに", chip: "自己紹介",
      footer: "雪国での実践を重ねてきた立場から、正直にお話しします",
      note: "差し替え：経歴年表の空欄3段（年・内容）。2014年の1段は記入済み。",
    });
    K.card(s, ML, 1.42, 4.5, 5.2, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("株式会社ヤシロ", {
      x: 0.42, y: 3.5, w: 4.0, h: 0.5, fontFace: F, fontSize: 22, bold: true, color: NAVY, margin: 0, valign: "middle",
    });
    s.addText("代表取締役社長", {
      x: 0.42, y: 3.98, w: 4.0, h: 0.44, fontFace: F, fontSize: 18, bold: true, color: BLACK, margin: 0, valign: "middle",
    });
    s.addText("屋代　健", {
      x: 0.42, y: 4.42, w: 4.0, h: 0.8, fontFace: F, fontSize: 40, bold: true, color: BLACK, margin: 0, valign: "middle",
    });
    s.addShape(pres.ShapeType.line, { x: 0.42, y: 5.32, w: 4.0, h: 0, line: { color: NAVY, width: 1 } });
    s.addText("2014年より柏崎・長岡で自社太陽光発電所を運営。\n雪国での発電・積雪対策の知見を蓄積。", {
      x: 0.42, y: 5.44, w: 4.0, h: 1.0, fontFace: F, fontSize: 14, bold: true, color: BLACK,
      margin: 0, valign: "top", lineSpacingMultiple: 1.2,
    });
    s.addText("経歴", {
      x: 5.0, y: 1.42, w: 1.2, h: 0.46, fontFace: F, fontSize: 22, bold: true, color: NAVY, margin: 0, valign: "middle",
    });
    s.addShape(pres.ShapeType.line, { x: 5.36, y: 2.0, w: 0, h: 4.6, line: { color: NAVY, width: 2 } });
    const tl = [
      { y: 2.16, year: "", body: "" },
      { y: 3.28, year: "2014年", body: "柏崎・長岡で自社の太陽光発電所を立ち上げ。積雪による破損・冬期の発電量低下という課題に取り組む" },
      { y: 4.40, year: "", body: "" },
      { y: 5.52, year: "", body: "" },
    ];
    tl.forEach((r) => {
      s.addShape(pres.ShapeType.ellipse, { x: 5.22, y: r.y + 0.14, w: 0.28, h: 0.28, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
      s.addText(r.year, { x: 3.86, y: r.y, w: 1.3, h: 0.5, fontFace: F, fontSize: 19, bold: true, color: NAVY, align: "right", valign: "middle", margin: 0 });
      s.addText(r.body, { x: 5.66, y: r.y - 0.1, w: 5.85, h: 0.9, fontFace: F, fontSize: 16, bold: true, color: BLACK, align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.15 });
    });
    note(3, "年表", "経歴（年・内容）×3段");
  }

  /* ===== P4 知る場 ===== */
  {
    const s = contentSlide({
      chapter: "はじめに",
      footer: [{ t: "この10年で、雪国の太陽光を取り巻く環境は" }, { t: "大きく変わりました", hl: true }, { t: "！" }],
    });
    s.addShape(pres.ShapeType.rect, { x: ML, y: 0.82, w: CW, h: 0.7, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    s.addText("本日は「売る」ための場ではなく、「知る」ための勉強会です", {
      x: ML, y: 0.82, w: CW, h: 0.7, fontFace: F, fontSize: 28, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("長岡で、こんな声をよく耳にします", {
      x: ML, y: 1.66, w: CW, h: 0.62, fontFace: F, fontSize: 32, bold: true, color: BLACK,
      align: "center", valign: "middle", margin: 0,
    });
    const voices = [
      "長岡の雪じゃ\n発電しないでしょ？",
      "屋根が雪で\n傷むのでは？",
      "訪問販売が怖い",
      "業者によって\n言うことが違う",
      "どのメーカーが\nいいの？",
      "いつ導入するのが\nお得なの？",
    ];
    const bw = 3.45, bh = 1.72, gx = 0.36, gy = 0.36, x0 = (SW - (bw * 3 + gx * 2)) / 2, y0 = 2.5;
    voices.forEach((t, i) => {
      const x = x0 + (i % 3) * (bw + gx), y = y0 + Math.floor(i / 3) * (bh + gy);
      K.card(s, x, y, bw, bh, { fill: PALE, line: NAVY, lw: 1.5 });
      s.addText("「" + t + "」", {
        x: x + 0.12, y, w: bw - 0.24, h: bh, fontFace: F, fontSize: 18, bold: true, color: NAVY,
        align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.15,
      });
    });
  }

  /* ===== P5 ゴール4つ ===== */
  {
    const s = contentSlide({ chapter: "はじめに", footer: "本日は何卒よろしくお願い申し上げます。" });
    s.addShape(pres.ShapeType.rect, { x: ML, y: 0.82, w: CW, h: 0.7, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    s.addText("講座を聞き終えた後、こうなっていただきたい", {
      x: ML, y: 0.82, w: CW, h: 0.7, fontFace: F, fontSize: 28, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0,
    });
    const goals = [
      "長岡市のエネルギー事情を\n正確に把握している",
      "メリット・デメリットを\n正確に把握している",
      "信頼できる業者の\n見極め方を知っている",
      "「我が家の場合はどうなのか」\nを知りたいと思っている",
    ];
    const cw = 4.9, ch = 2.06, gx = 0.5, gy = 0.62, x0 = (SW - (cw * 2 + gx)) / 2, y0 = 2.06;
    goals.forEach((t, i) => {
      const x = x0 + (i % 2) * (cw + gx), y = y0 + Math.floor(i / 2) * (ch + gy);
      K.card(s, x, y, cw, ch, { fill: WHITE, line: NAVY, lw: 2 });
      s.addShape(pres.ShapeType.ellipse, { x: x + cw / 2 - 0.36, y: y - 0.34, w: 0.72, h: 0.72, fill: { color: NAVY }, line: { color: WHITE, width: 2 } });
      s.addText(String(i + 1), { x: x + cw / 2 - 0.36, y: y - 0.34, w: 0.72, h: 0.72, fontFace: F, fontSize: 28, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      s.addText(t, { x: x + 0.2, y: y + 0.34, w: cw - 0.4, h: ch - 0.5, fontFace: F, fontSize: 20, bold: true, color: BLACK, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
    });
  }

  /* ===== P6 目次(1) ===== */
  agendaSlide(0);

  /* ===== P7 章イメージ ===== */
  {
    const s = contentSlide({
      chapter: "1.長岡市のエネルギー事情と災害対策",
      footer: "この章では、長岡の「今」をデータで確認します",
    });
    s.addText("電気代高騰で家計圧迫…！", {
      x: 0.4, y: 1.5, w: SW - 0.8, h: 1.2, fontFace: F, fontSize: 54, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    const items = [
      ["上がり続ける電気代", "2023年、東北電力は規制料金を大幅に値上げ。全国的にも上昇が続いている。"],
      ["いつ起きるか分からない停電", "中越地震では約11日。大雪では倒木・断線。雪国は復旧に時間がかかる。"],
      ["補助金頼みという構造", "国の負担軽減措置は期限つき。終われば請求額はそのまま跳ね上がる。"],
    ];
    const cw = 3.55, ch = 3.0, gx = 0.32, x0 = (SW - (cw * 3 + gx * 2)) / 2, y0 = 3.16;
    items.forEach((it, i) => {
      const x = x0 + i * (cw + gx);
      K.card(s, x, y0, cw, ch, { fill: WHITE, line: NAVY, lw: 2 });
      s.addShape(pres.ShapeType.rect, { x, y: y0, w: cw, h: 0.62, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
      s.addText("①②③".charAt(i) + "　" + it[0], {
        x: x + 0.1, y: y0, w: cw - 0.2, h: 0.62, fontFace: F, fontSize: 16, bold: true, color: WHITE,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(it[1], {
        x: x + 0.22, y: y0 + 0.78, w: cw - 0.44, h: ch - 0.98, fontFace: F, fontSize: 16, bold: true, color: BLACK,
        align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.25,
      });
    });
  }

  /* ===== P8 東北電力の電気代（グラフ空） ===== */
  {
    const s = contentSlide({
      chapter: "1.長岡市のエネルギー事情と災害対策", chip: "東北電力の電気代",
      lead: "東北電力は2023年、規制料金の大幅な値上げを実施している。",
      source: "出典：東北電力「電気料金単価表」より作成　※登壇前に最新単価を確認",
      footer: "電気代は「上がり続ける」前提で考える時代です",
      note: "差し替え：2015〜2026年の電気料金推移（規制料金／自由料金）。グラフのデータシートに入力する。2023年の補助線位置はデータ入力後に微調整すること。",
    });
    const labels = ["2015", "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026"];
    const empty = labels.map(() => null);
    s.addChart(pres.ChartType.line,
      [{ name: "規制料金", labels, values: empty.slice() }, { name: "自由料金", labels, values: empty.slice() }],
      {
        x: 0.5, y: 1.74, w: 10.69, h: 4.9,
        showTitle: true, title: "東北電力　電気料金単価の推移（円/kWh）", titleFontFace: F, titleFontSize: 16, titleColor: NAVY,
        chartColors: [NAVY, "8FAADC"], lineSize: 3, lineDataSymbolSize: 8,
        showLegend: true, legendPos: "b", legendFontFace: F, legendFontSize: 13,
        catAxisLabelFontFace: F, catAxisLabelFontSize: 13, catAxisLabelColor: "404040",
        valAxisLabelFontFace: F, valAxisLabelFontSize: 13, valAxisLabelColor: "404040",
        valAxisTitle: "円/kWh", showValAxisTitle: true, valAxisTitleFontFace: F, valAxisTitleFontSize: 12,
        valGridLine: { color: "D9D9D9", size: 1 }, catGridLine: { style: "none" },
      });
    s.addShape(pres.ShapeType.line, { x: 8.48, y: 2.24, w: 0, h: 3.66, line: { color: HRED, width: 2, dashType: "dash" } });
    s.addText("2023年\n大幅値上げ", {
      x: 7.72, y: 1.82, w: 1.6, h: 0.5, fontFace: F, fontSize: 13, bold: true, color: HRED,
      align: "center", valign: "middle", margin: 0,
    });
    note(8, "グラフ", "電気料金推移 2015〜2026（規制料金／自由料金）");
  }

  /* ===== P9 補助金頼みの構造リスク ===== */
  {
    const s = contentSlide({
      chapter: "1.長岡市のエネルギー事情と災害対策", chip: "電気代を取り巻く構造",
      lead: "電気料金の負担は、国の負担軽減措置に支えられてきた。しかし措置は期限つきで、いつ終わるかは読めない。",
      leadLines: 2,
      source: "※国の電気・ガス料金負担軽減支援事業の実施状況は変動します。最新情報をご確認ください",
      footer: [{ t: "補助金が前提の家計が、" }, { t: "いちばん危ない", hl: true }, { t: "かもしれません" }],
    });
    K.card(s, 0.42, 2.14, 5.36, 4.4, { fill: PALE, line: NAVY, lw: 2 });
    s.addShape(pres.ShapeType.rect, { x: 0.42, y: 2.14, w: 5.36, h: 0.6, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    s.addText("補助があるうちは見えないこと", { x: 0.52, y: 2.14, w: 5.16, h: 0.6, fontFace: F, fontSize: 19, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    [["補助は「期限つき」", "延長されるかどうかは、そのつど政治判断で決まる"],
     ["終了・縮小のたびに請求額が跳ね上がる", "使い方は変えていないのに、請求書だけが増える"],
     ["補助が前提の家計は、いちばん揺れやすい", "支出の前提が外部要因に握られている状態"]]
      .forEach((it, i) => {
        const y = 2.92 + i * 1.16;
        s.addText(it[0], { x: 0.66, y, w: 4.9, h: 0.5, fontFace: F, fontSize: 17, bold: true, color: NAVY, align: "left", valign: "middle", margin: 0 });
        s.addText(it[1], { x: 0.66, y: y + 0.46, w: 4.9, h: 0.6, fontFace: F, fontSize: 14, bold: true, color: BLACK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.15 });
      });
    K.card(s, 5.92, 2.14, 5.36, 4.4, { fill: WHITE, line: RED, lw: 2 });
    s.addShape(pres.ShapeType.rect, { x: 5.92, y: 2.14, w: 5.36, h: 0.6, fill: { color: RED }, line: { color: RED, width: 0 } });
    s.addText("補助が終わったとき、どうなるか", { x: 6.02, y: 2.14, w: 5.16, h: 0.6, fontFace: F, fontSize: 19, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    s.addText("補助あり", { x: 6.3, y: 3.1, w: 2.0, h: 0.5, fontFace: F, fontSize: 18, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    K.card(s, 6.3, 3.62, 2.0, 1.1, { fill: PALE, line: NAVY, lw: 1.5 });
    s.addText("請求額", { x: 6.3, y: 3.62, w: 2.0, h: 1.1, fontFace: F, fontSize: 18, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    s.addShape(pres.ShapeType.rightArrow, { x: 8.42, y: 3.92, w: 0.62, h: 0.5, fill: { color: RED }, line: { color: RED, width: 0 } });
    s.addText("補助終了", { x: 9.1, y: 3.1, w: 2.0, h: 0.5, fontFace: F, fontSize: 18, bold: true, color: RED, align: "center", valign: "middle", margin: 0 });
    K.card(s, 9.1, 3.14, 2.0, 1.58, { fill: "FDE9E7", line: RED, lw: 1.5 });
    s.addText("請求額", { x: 9.1, y: 3.14, w: 2.0, h: 1.58, fontFace: F, fontSize: 22, bold: true, color: RED, align: "center", valign: "middle", margin: 0 });
    s.addText("使い方は変えていないのに、支払いだけが増える。\n電気代そのものを減らす手段を持っておくことが、いちばん確実な備えになる。", {
      x: 6.16, y: 5.0, w: 4.9, h: 1.34, fontFace: F, fontSize: 15, bold: true, color: BLACK,
      align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.25,
    });
  }

  /* ===== P10 柏崎刈羽6号機 ===== */
  {
    const s = contentSlide({
      chapter: "1.長岡市のエネルギー事情と災害対策", chip: "電力供給の前提",
      lead: "柏崎刈羽原子力発電所6号機は2026年の再稼働が見込まれている。ただし同発電所は東京電力の設備であり、長岡市は東北電力の供給エリアである。",
      leadLines: 2,
      source: "※再稼働時期は本資料作成時点の見込みです。登壇直前に最新状況をご確認ください",
      footer: "再稼働＝長岡の電気代が下がる、とは限りません",
    });
    K.card(s, 0.42, 2.2, 5.36, 4.34, { fill: WHITE, line: NAVY, lw: 2 });
    s.addShape(pres.ShapeType.rect, { x: 0.42, y: 2.2, w: 5.36, h: 0.6, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    s.addText("【事実の整理】", { x: 0.52, y: 2.2, w: 5.16, h: 0.6, fontFace: F, fontSize: 19, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    ["柏崎刈羽原子力発電所は、東京電力の発電所である",
     "立地は柏崎市・刈羽村（長岡市の隣接地域）",
     "6号機は2026年の再稼働が見込まれている"]
      .forEach((t, i) => bulletRow(s, 0.66, 3.0 + i * 1.16, 4.9, 1.02, i + 1, t, NAVY, 16));
    K.card(s, 5.92, 2.2, 5.36, 4.34, { fill: PALE, line: NAVY, lw: 2 });
    s.addShape(pres.ShapeType.rect, { x: 5.92, y: 2.2, w: 5.36, h: 0.6, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    s.addText("【長岡市への影響】", { x: 6.02, y: 2.2, w: 5.16, h: 0.6, fontFace: F, fontSize: 19, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    ["長岡市は東北電力の供給エリアである",
     "再稼働が長岡の電気料金を直接下げるとは限らない",
     "原発の賛否は、本講座では扱いません"]
      .forEach((t, i) => bulletRow(s, 6.16, 3.0 + i * 1.16, 4.9, 1.02, i + 1, t, NAVY, 16));
  }

  /* ===== P11 中越地震 ===== */
  {
    const s = contentSlide({
      chapter: "1.長岡市のエネルギー事情と災害対策", chip: "災害・停電対策",
      lead: "2004年の新潟県中越地震では、長岡市を含む広い範囲で停電が発生した。",
      source: "出典：各種調査レポートより　※数値の一次出典は要確認",
      footer: "停電はいつ・どこで起きるか分かりません",
    });
    s.addText("新潟県中越地震（2004年10月23日）", {
      x: ML, y: 1.78, w: CW, h: 0.56, fontFace: F, fontSize: 24, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    const stats = [
      ["停電の発生件数", "約30", "万件"],
      ["復旧までの日数", "11", "日"],
      ["最大震度", "7", ""],
    ];
    const cw = 3.4, ch = 2.1, gx = 0.42, x0 = (SW - (cw * 3 + gx * 2)) / 2, y0 = 2.52;
    stats.forEach((st, i) => statCard(s, x0 + i * (cw + gx), y0, cw, ch, st[0], st[1], st[2], { valueSize: 52 }));
    K.card(s, x0, 5.06, cw * 3 + gx * 2, 1.5, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("長岡市は、この地震の震源域の中心にありました。\n「うちの地域は大丈夫」と言える場所は、どこにもありません。", {
      x: x0 + 0.24, y: 5.06, w: cw * 3 + gx * 2 - 0.48, h: 1.5, fontFace: F, fontSize: 19, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.25,
    });
  }

  /* ===== P12 停電実績表 ===== */
  {
    const s = contentSlide({
      chapter: "1.長岡市のエネルギー事情と災害対策", chip: "災害・停電対策",
      lead: "過去の災害では、復旧までに1週間以上を要したケースも少なくない。",
      source: "出典：ヤシロ調べ。各調査レポートより停電件数・復旧日数を記載　※一次出典は要確認",
      footer: "災害大国だからこそ、もしもの備えが必要です",
    });
    s.addText("主な災害時の停電発生件数と日数", {
      x: ML, y: 1.76, w: 5.7, h: 0.44, fontFace: F, fontSize: 18, bold: true, color: WHITE,
      fill: { color: NAVY }, align: "center", valign: "middle", margin: 0,
    });
    const hi = { fill: { color: "FDE9E7" }, color: RED, bold: true };
    const rows = [
      [th("発生年"), th("災害名"), th("発生件数（約）"), th("復旧日数（約）")],
      [val("1995年", { align: "center" }), val("阪神淡路大震災"), val("260万件", { align: "right" }), val("5日", { align: "right" })],
      [val("2003年", { align: "center" }), val("十勝沖地震"), val("38万件", { align: "right" }), val("1日", { align: "right" })],
      [val("2004年", Object.assign({ align: "center" }, hi)), val("新潟県中越地震", hi), val("30万件", Object.assign({ align: "right" }, hi)), val("11日", Object.assign({ align: "right" }, hi))],
      [val("2011年", { align: "center" }), val("東日本大震災"), val("466万件", { align: "right" }), val("8日", { align: "right" })],
      [val("2016年", { align: "center" }), val("熊本地震"), val("48万件", { align: "right" }), val("5日", { align: "right" })],
      [val("2018年", { align: "center" }), val("台風21号"), val("261万件", { align: "right" }), val("4日", { align: "right" })],
      [val("2018年", { align: "center" }), val("胆振東部地震"), val("295万件", { align: "right" }), val("2日", { align: "right" })],
      [val("2024年", { align: "center" }), val("能登半島地震"), val("4万件", { align: "right" }), val("30日", { align: "right" })],
    ];
    table(s, rows, ML, 2.3, 5.7, [1.1, 2.14, 1.28, 1.18], { rowH: 0.47, fontSize: 14 });
    K.card(s, 6.16, 2.3, 5.37, 4.24, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("長岡は、すでに経験しています", {
      x: 6.36, y: 2.5, w: 4.97, h: 0.6, fontFace: F, fontSize: 22, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    statCard(s, 6.5, 3.24, 2.34, 1.5, "停電発生件数", "約30", "万件", { valueSize: 36 });
    statCard(s, 8.94, 3.24, 2.34, 1.5, "復旧までの日数", "11", "日", { valueSize: 36, valueColor: RED, line: RED, labelColor: RED });
    s.addText("件数が少なくても、復旧に時間がかかります。\n雪国では除雪が進まず、作業そのものが遅れます。\n\n2024年の能登半島地震も、4万件に対して復旧は30日。", {
      x: 6.5, y: 4.92, w: 4.78, h: 1.56, fontFace: F, fontSize: 14, bold: true, color: BLACK,
      align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.28,
    });
  }

  /* ===== P13 2022年12月の大雪 ===== */
  {
    const s = contentSlide({
      chapter: "1.長岡市のエネルギー事情と災害対策", chip: "災害・停電対策",
      lead: "停電の原因は地震だけではない。2022年12月の大雪では、倒木や断線による停電が県内各地で発生している。",
      leadLines: 2,
      source: "※停電規模・復旧日数の一次出典は登壇前に確認してください",
      footer: "雪国の停電は、復旧に時間がかかります",
    });
    const items = [
      ["大雪による倒木・断線", "雪の重みで木が倒れ、電線を切る。原因箇所が多数に分散する。"],
      ["佐渡では約9日間の停電", "island の孤立に加え、アクセス自体が雪で断たれた。"],
      ["除雪が進まず復旧が遅れる", "現場にたどり着けない。地震とは別の「遅さ」がある。"],
    ];
    items[1][1] = "離島という条件に加え、現場へのアクセス自体が雪で断たれた。";
    const cw = 3.55, ch = 2.34, gx = 0.32, x0 = (SW - (cw * 3 + gx * 2)) / 2, y0 = 2.24;
    items.forEach((it, i) => {
      const x = x0 + i * (cw + gx);
      K.card(s, x, y0, cw, ch, { fill: WHITE, line: NAVY, lw: 2 });
      s.addShape(pres.ShapeType.rect, { x, y: y0, w: cw, h: 0.62, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
      s.addText(it[0], { x: x + 0.08, y: y0, w: cw - 0.16, h: 0.62, fontFace: F, fontSize: 16, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      s.addText(it[1], { x: x + 0.22, y: y0 + 0.76, w: cw - 0.44, h: ch - 0.94, fontFace: F, fontSize: 15, bold: true, color: BLACK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.25 });
    });
    K.card(s, x0, 4.86, cw * 3 + gx * 2, 1.68, { fill: "FDE9E7", line: RED, lw: 2 });
    s.addText("雪国特有の「3つの重なり」", {
      x: x0, y: 4.98, w: cw * 3 + gx * 2, h: 0.44, fontFace: F, fontSize: 18, bold: true, color: RED,
      align: "center", valign: "middle", margin: 0,
    });
    const three = ["停電", "低気温", "除雪の遅れ"];
    const tw = 2.2, tgx = 0.86, tx0 = x0 + (cw * 3 + gx * 2 - (tw * 3 + tgx * 2)) / 2;
    three.forEach((t, i) => {
      const x = tx0 + i * (tw + tgx);
      s.addText(t, { x, y: 5.5, w: tw, h: 0.78, fontFace: F, fontSize: 22, bold: true, color: WHITE, fill: { color: RED }, align: "center", valign: "middle", margin: 0 });
      if (i < 2) s.addText("＋", { x: x + tw, y: 5.5, w: tgx, h: 0.78, fontFace: F, fontSize: 24, bold: true, color: RED, align: "center", valign: "middle", margin: 0 });
    });
  }

  /* ===== P14 厳冬期に暖房が止まる ===== */
  {
    const s = contentSlide({
      chapter: "1.長岡市のエネルギー事情と災害対策", chip: "災害・停電対策",
      lead: "厳冬期に停電が起これば、暖房をはじめとする生活インフラが同時に止まる。",
      footer: [{ t: "雪国の停電は、" }, { t: "命に直結します", hl: true }],
    });
    s.addText("厳冬期に停電したら、家の中で何が起きるか", {
      x: ML, y: 1.76, w: CW, h: 0.56, fontFace: F, fontSize: 26, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    const items = [
      ["暖房", "エアコン・ファンヒーターが停止。\n室温が一気に下がる"],
      ["給湯", "お湯が出ない。\n入浴も洗い物もできない"],
      ["給水", "ポンプ停止・配管の凍結。\n水そのものが止まる"],
      ["トイレ", "電動洗浄・ポンプが停止。\n衛生環境が急激に悪化"],
      ["情報", "Wi-Fi・充電が不可。\n復旧情報が届かない"],
    ];
    const cw = 2.12, ch = 3.6, gx = 0.2, x0 = (SW - (cw * 5 + gx * 4)) / 2, y0 = 2.5;
    items.forEach((it, i) => {
      const x = x0 + i * (cw + gx);
      K.card(s, x, y0, cw, ch, { fill: WHITE, line: NAVY, lw: 2 });
      s.addShape(pres.ShapeType.rect, { x, y: y0, w: cw, h: 0.66, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
      s.addText(it[0], { x, y: y0, w: cw, h: 0.66, fontFace: F, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      s.addText("停止", { x: x + 0.36, y: y0 + 0.86, w: cw - 0.72, h: 0.54, fontFace: F, fontSize: 18, bold: true, color: RED, fill: { color: "FDE9E7" }, align: "center", valign: "middle", margin: 0 });
      s.addText(it[1], { x: x + 0.14, y: y0 + 1.54, w: cw - 0.28, h: ch - 1.72, fontFace: F, fontSize: 14, bold: true, color: BLACK, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.22 });
    });
    s.addText("氷点下の夜に、これらが同時に止まります。", {
      x: ML, y: 6.24, w: CW, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: RED,
      align: "center", valign: "middle", margin: 0,
    });
  }

  /* ===== P15 太陽光・蓄電池は当たり前 ===== */
  {
    const s = contentSlide({
      chapter: "1.長岡市のエネルギー事情と災害対策", chip: "太陽光・蓄電池は当たり前",
      lead: "電気代高騰と災害への不安から電気の自給自足への関心が高まり、太陽光・蓄電池を備えた住宅の標準化が進んでいる。",
      leadLines: 2,
      footer: "太陽光・蓄電池は、贅沢品ではなく「備え」です",
    });
    const cw = 5.36, lx = 0.42, rx = SW - 0.42 - cw, y0 = 2.2, ch = 4.34;
    K.card(s, lx, y0, cw, ch, { fill: WHITE, line: "808080", lw: 2 });
    s.addShape(pres.ShapeType.rect, { x: lx, y: y0, w: cw, h: 0.64, fill: { color: "808080" }, line: { color: "808080", width: 0 } });
    s.addText("未導入のお家", { x: lx, y: y0, w: cw, h: 0.64, fontFace: F, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    K.card(s, rx, y0, cw, ch, { fill: WHITE, line: NAVY, lw: 2 });
    s.addShape(pres.ShapeType.rect, { x: rx, y: y0, w: cw, h: 0.64, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    s.addText("太陽光＋蓄電池のお家", { x: rx, y: y0, w: cw, h: 0.64, fontFace: F, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    const compare = [
      ["平常時の電気代", "使った分をすべて買う", "つくった電気から使う"],
      ["停電したとき", "家じゅうの電気が止まる", "蓄電池から電気を使える"],
      ["厳冬期の停電", "暖房も給湯も止まる", "最低限の暖をとれる"],
    ];
    compare.forEach((c, i) => {
      const y = y0 + 0.88 + i * 1.16;
      s.addText(c[0], { x: 0.42, y: y - 0.3, w: SW - 0.84, h: 0.32, fontFace: F, fontSize: 13, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
      s.addText(c[1], { x: lx + 0.24, y, w: cw - 0.48, h: 0.72, fontFace: F, fontSize: 17, bold: true, color: "404040", fill: { color: "F2F2F2" }, align: "center", valign: "middle", margin: 0 });
      s.addText(c[2], { x: rx + 0.24, y, w: cw - 0.48, h: 0.72, fontFace: F, fontSize: 17, bold: true, color: NAVY, fill: { color: PALE }, align: "center", valign: "middle", margin: 0 });
    });
  }
};
