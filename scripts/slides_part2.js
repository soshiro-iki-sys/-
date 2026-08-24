/* P17〜P32 */
module.exports = function (K) {
  const { pres, note, NAVY, RED, HRED, YEL, WHITE, BLACK, PALE, PINK, GRAY2, INK, SOFT, F,
    SW, SH, ML, CW, contentSlide, card, band, bulletRow, statCard, table, th, lbl, val, blank,
    compareTable, agendaSlide, dividerSlide, fitSize } = K;
  const CH2 = "2.長岡市でも太陽光はお得なのか";

  /* ===== P17 目次(2) ===== */
  agendaSlide(1);

  /* ===== P18 章イメージ ===== */
  {
    const s = contentSlide({ chapter: CH2, footer: "イメージではなく、データで判断しましょう" });
    s.addText("「長岡の雪じゃ、発電しないでしょ？」", {
      x: 0.4, y: 1.26, w: SW - 0.8, h: 0.94, fontFace: F, fontSize: 38, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("いちばん多くいただくご質問です。ひとつずつ、データで確かめていきます。", {
      x: 0.4, y: 2.24, w: SW - 0.8, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: BLACK,
      align: "center", valign: "middle", margin: 0,
    });
    const qs = [
      ["発電量は足りるのか", "新潟県の発電量は、全国平均と比べてどの程度なのか"],
      ["冬はどれくらい落ちるか", "落ちるのは事実。ではゼロになるのか"],
      ["積雪で壊れないのか", "荷重・落雪に、どう設計で対応するのか"],
    ];
    const cw = 3.55, ch = 3.3, gx = 0.32, x0 = (SW - (cw * 3 + gx * 2)) / 2, y0 = 3.16;
    qs.forEach((q, i) => {
      const x = x0 + i * (cw + gx);
      card(s, x, y0, cw, ch, { fill: WHITE, line: NAVY, lw: 2 });
      s.addShape(pres.ShapeType.ellipse, { x: x + cw / 2 - 0.4, y: y0 - 0.38, w: 0.8, h: 0.8, fill: { color: NAVY }, line: { color: WHITE, width: 2 } });
      s.addText("Q" + (i + 1), { x: x + cw / 2 - 0.4, y: y0 - 0.38, w: 0.8, h: 0.8, fontFace: F, fontSize: 22, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      s.addText(q[0], { x: x + 0.14, y: y0 + 0.6, w: cw - 0.28, h: 0.66, fontFace: F, fontSize: fitSize(q[0], cw - 0.4, 20, 15), bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
      s.addShape(pres.ShapeType.line, { x: x + 0.4, y: y0 + 1.38, w: cw - 0.8, h: 0, line: { color: NAVY, width: 1 } });
      s.addText(q[1], { x: x + 0.24, y: y0 + 1.56, w: cw - 0.48, h: 1.5, fontFace: F, fontSize: 15, bold: true, color: BLACK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.3 });
    });
  }

  /* ===== P19 新潟県の発電量（実データ） ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "新潟県の発電量",
      lead: "新潟県は全国平均をやや下回る。ただし最も多い地域との差は、約2割にとどまる。",
      source: "出典：NEDO日射量データベースをもとにした都道府県別の年間発電量試算値より作成　※前提条件は要確認",
      footer: "「雪国だから発電しない」は、事実ではありません",
    });
    const labels = ["山梨・長野など\n（最多）", "全国平均", "新潟県", "秋田県\n（最少）"];
    const values = [1330, 1237, 1105, 1095];
    s.addChart(pres.ChartType.bar, [{ name: "1kWあたり年間発電量", labels, values }], {
      x: 0.42, y: 2.12, w: 7.3, h: 4.5, barDir: "col", barGapWidthPct: 70,
      showTitle: true, title: "太陽光1kWあたりの年間発電量の比較（kWh/年）", titleFontFace: F, titleFontSize: 15, titleColor: NAVY,
      chartColors: ["A6A6A6", "A6A6A6", NAVY, "A6A6A6"], varyColors: true, showLegend: false,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontFace: F, dataLabelFontSize: 13, dataLabelColor: INK, dataLabelFormatCode: "#,##0",
      catAxisLabelFontFace: F, catAxisLabelFontSize: 12, catAxisLabelColor: INK,
      valAxisLabelFontFace: F, valAxisLabelFontSize: 11, valAxisLabelColor: INK,
      valAxisMinVal: 0, valAxisMaxVal: 1500,
      valGridLine: { color: "D9D9D9", size: 1 }, catGridLine: { style: "none" },
    });
    card(s, 7.9, 2.12, 3.63, 4.5, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("差は思ったより小さい", { x: 7.9, y: 2.3, w: 3.63, h: 0.5, fontFace: F, fontSize: 19, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    s.addShape(pres.ShapeType.line, { x: 8.12, y: 2.88, w: 3.19, h: 0, line: { color: NAVY, width: 1 } });
    s.addText([
      { text: "新潟県　", options: { fontFace: F, fontSize: 15, bold: true, color: NAVY } },
      { text: "約1,105", options: { fontFace: F, fontSize: 30, bold: true, color: NAVY } },
      { text: " kWh／kW・年", options: { fontFace: F, fontSize: 13, bold: true, color: NAVY, breakLine: true } },
      { text: "\n全国平均との差　約1割", options: { fontFace: F, fontSize: 16, bold: true, color: BLACK, breakLine: true } },
      { text: "最多地域との差　約2割", options: { fontFace: F, fontSize: 16, bold: true, color: BLACK } },
    ], { x: 8.12, y: 3.06, w: 3.19, h: 2.0, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.25 });
    s.addText("「半分しか発電しない」わけではありません。長岡でも十分に成立する水準です。", {
      x: 8.12, y: 5.16, w: 3.19, h: 1.3, fontFace: F, fontSize: 14, bold: true, color: BLACK,
      align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.3,
    });
  }

  /* ===== P20 冬の発電量と積雪対策 ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "冬の発電量と積雪対策",
      lead: "冬の発電量が落ちるのは事実である。ただし、ゼロになるわけではない。",
      source: "出典：長岡市「雪国対応の太陽光発電実証実験」／自社発電所（柏崎・長岡）の実測値",
      footer: "冬は確実に落ちますが、ゼロにはなりません",
      note: "差し替え：自社発電所（柏崎・長岡）の月別実測値（発電量／最深積雪／備考）。月の列は記入済み。",
    });
    band(s, ML, 1.76, 5.6, "自社発電所の積雪期 実測データ", { h: 0.46, size: 18 });
    const rows = [
      [th("月"), th("発電量（kWh）"), th("最深積雪（cm）"), th("備考")],
      [val("11月", { align: "center", bold: true }), blank(), blank(), blank()],
      [val("12月", { align: "center", bold: true }), blank(), blank(), blank()],
      [val("1月", { align: "center", bold: true }), blank(), blank(), blank()],
      [val("2月", { align: "center", bold: true }), blank(), blank(), blank()],
      [val("3月", { align: "center", bold: true }), blank(), blank(), blank()],
    ];
    table(s, rows, ML, 2.28, 5.6, [0.92, 1.6, 1.6, 1.48], { rowH: 0.43, fontSize: 14 });
    card(s, ML, 5.0, 5.6, 1.54, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("長岡市も、雪国対応の実証を進めています", { x: ML + 0.16, y: 5.1, w: 5.28, h: 0.4, fontFace: F, fontSize: 15, bold: true, color: NAVY, align: "left", valign: "middle", margin: 0 });
    s.addText("・ハイブ長岡の南壁面にパネル26枚（約9kW）を設置\n・降雪や強風に合わせて角度を変える可変架台（19.5kW）で最適角度を検証", {
      x: ML + 0.16, y: 5.5, w: 5.28, h: 0.96, fontFace: F, fontSize: 13, bold: true, color: BLACK,
      align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.28,
    });
    band(s, 6.16, 1.76, 5.37, "雪国での積雪対策", { h: 0.46, size: 18 });
    card(s, 6.16, 2.28, 5.37, 4.26, { fill: WHITE, line: NAVY, lw: 2 });
    [["屋根の傾斜角を活かした自然落雪", "雪が自然に滑り落ちる角度と配置を、現地調査で見極める"],
     ["雪下ろしを前提としない設計", "屋根に上がらなくても運用できる形にする。安全がいちばん"],
     ["架台・荷重の設計", "積雪荷重に耐える架台と固定方法を選定する"]]
      .forEach((it, i) => {
        const y = 2.52 + i * 1.34;
        bulletRow(s, 6.42, y, 4.86, 0.54, i + 1, it[0], NAVY, 16);
        s.addText(it[1], { x: 7.08, y: y + 0.56, w: 4.2, h: 0.66, fontFace: F, fontSize: 13.5, bold: true, color: INK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.25 });
      });
    note(20, "表", "自社発電所の積雪期実測（発電量／最深積雪／備考）");
  }

  /* ===== P21 長岡市の発電量試算 ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "長岡市の発電量",
      lead: "長岡市で5kWの太陽光を設置した場合、家庭の消費電力の多くを賄える。",
      source: "※1kWあたりの発電量1,100kWhは、日本の国家規格「JIS C 8907（太陽光発電システムの発電電力量推定方法）」をもとに算出",
      footer: "消費電力の84％をカバーすることが可能に",
    });
    band(s, ML, 1.9, CW, "長岡市での発電量試算：5kWシステムの場合", { h: 0.5, size: 20 });
    card(s, 0.8, 2.62, 4.62, 1.74, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("年間の発電量", { x: 0.8, y: 2.74, w: 4.62, h: 0.36, fontFace: F, fontSize: 15, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    s.addText("5kW × 1,100kWh", { x: 0.8, y: 3.1, w: 4.62, h: 0.42, fontFace: F, fontSize: 19, bold: true, color: BLACK, align: "center", valign: "middle", margin: 0 });
    s.addText([
      { text: "約5,500", options: { fontFace: F, fontSize: 34, bold: true, color: NAVY } },
      { text: " kWh／年", options: { fontFace: F, fontSize: 16, bold: true, color: NAVY } },
    ], { x: 0.8, y: 3.52, w: 4.62, h: 0.72, align: "center", valign: "middle", margin: 0 });
    card(s, 0.8, 4.56, 4.62, 1.74, { fill: SOFT, line: GRAY2, lw: 2 });
    s.addText("1年間の電気使用量（4人家族・オール電化）", { x: 0.8, y: 4.68, w: 4.62, h: 0.36, fontFace: F, fontSize: 14, bold: true, color: INK, align: "center", valign: "middle", margin: 0 });
    s.addText([
      { text: "6,500", options: { fontFace: F, fontSize: 34, bold: true, color: INK } },
      { text: " kWh／年", options: { fontFace: F, fontSize: 16, bold: true, color: INK } },
    ], { x: 0.8, y: 5.2, w: 4.62, h: 0.9, align: "center", valign: "middle", margin: 0 });
    s.addShape(pres.ShapeType.rightArrow, { x: 5.62, y: 4.1, w: 0.72, h: 0.72, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    card(s, 6.54, 2.62, 4.72, 3.68, { fill: WHITE, line: NAVY, lw: 3 });
    s.addText("消費電力のカバー率", { x: 6.54, y: 2.86, w: 4.72, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    s.addText([
      { text: "84", options: { fontFace: F, fontSize: 88, bold: true, color: NAVY } },
      { text: "％", options: { fontFace: F, fontSize: 34, bold: true, color: NAVY } },
    ], { x: 6.7, y: 3.42, w: 4.4, h: 1.6, align: "center", valign: "middle", margin: 0 });
    s.addText("5,500kWh ÷ 6,500kWh", { x: 6.54, y: 5.08, w: 4.72, h: 0.4, fontFace: F, fontSize: 16, bold: true, color: INK, align: "center", valign: "middle", margin: 0 });
    s.addText("雪国の長岡でも、この水準です。", { x: 6.54, y: 5.6, w: 4.72, h: 0.44, fontFace: F, fontSize: 17, bold: true, color: RED, align: "center", valign: "middle", margin: 0 });
  }

  /* ===== P22 電気代削減の試算（実数） ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "電気代削減の試算",
      lead: "5kWシステムで年間どれだけ得になるか。前提を置いて試算する。",
      source: "前提：発電5,500kWh／買電単価30円/kWh（東北電力 従量電灯B の段階単価より）／売電単価24円/kWh（2026年度FIT・当初4年）",
      footer: "蓄電池を足すと、経済メリットはさらに伸びます",
    });
    band(s, ML, 1.9, CW, "年間の経済メリット試算（5kWシステム・4人家族オール電化）", { h: 0.5, size: 20 });
    const rows = [
      [th("項目"), th("太陽光のみ"), th("太陽光＋蓄電池")],
      [lbl("年間発電量"), val("5,500 kWh", { align: "right" }), val("5,500 kWh", { align: "right" })],
      [lbl("自家消費率"), val("約30％", { align: "right" }), val("約65％", { align: "right", color: NAVY, bold: true })],
      [lbl("買電の削減額"), val("約49,000円", { align: "right" }), val("約107,000円", { align: "right", color: NAVY, bold: true })],
      [lbl("売電収入"), val("約92,000円", { align: "right" }), val("約46,000円", { align: "right" })],
      [{ text: "年間の経済メリット合計", options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "left" } },
       { text: "約141,000円", options: { fill: { color: PALE }, color: BLACK, bold: true, align: "right", fontSize: 17 } },
       { text: "約153,000円", options: { fill: { color: PINK }, color: RED, bold: true, align: "right", fontSize: 17 } }],
    ];
    table(s, rows, ML, 2.6, 7.2, [2.7, 2.25, 2.25], { rowH: 0.58, fontSize: 15 });
    card(s, 7.6, 2.6, 3.93, 3.98, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("単価はご契約プランで変わります", { x: 7.78, y: 2.78, w: 3.57, h: 0.8, fontFace: F, fontSize: 18, bold: true, color: NAVY, align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.15 });
    s.addShape(pres.ShapeType.line, { x: 7.78, y: 3.64, w: 3.57, h: 0, line: { color: NAVY, width: 1 } });
    s.addText(
      "・契約プラン（従量電灯／時間帯別）\n" +
      "・ご家族の人数と在宅時間\n" +
      "・オール電化かどうか\n" +
      "・屋根の向きと形状\n\n" +
      "これらで削減額は大きく変わります。\n一般論ではなく、お住まいの条件で計算します。",
      { x: 7.78, y: 3.8, w: 3.57, h: 2.6, fontFace: F, fontSize: 14, bold: true, color: BLACK,
        align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.3 }
    );
  }

  /* ===== P23 お客様データ①（空欄） ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "実際のお客様のデータ",
      lead: "新潟県内にお住まいのお客様の、導入前後の実測データである。",
      source: "※お客様の実測値。掲載にあたっては事前に許諾を得てください",
      footer: "実際のお客様のデータで確かめてみましょう",
      note: "差し替え：お客様概要（所在地・家族構成・システム容量・蓄電池）／年間買電量・年間電気料金の導入前後と削減率／グラフのデータシート。",
    });
    table(s, [
      [th("所在地"), th("家族構成"), th("システム容量"), th("蓄電池")],
      [blank(), blank(), blank(), blank()],
    ], ML, 1.8, 5.7, [1.42, 1.42, 1.5, 1.36], { rowH: 0.44, fontSize: 13 });
    table(s, [
      [th("項目"), th("導入前"), th("導入後"), th("削減率")],
      [lbl("年間買電量（kWh）"), blank(), blank(), blank()],
      [lbl("年間電気料金（円）"), blank(), blank(), blank()],
    ], ML, 3.0, 5.7, [1.86, 1.32, 1.32, 1.2], { rowH: 0.62, fontSize: 14 });
    card(s, ML, 5.14, 5.7, 1.4, { fill: PALE, line: NAVY, lw: 1.5 });
    s.addText("同じ暮らし方のまま、買う電気の量そのものが減ります。\n節約をがまんで実現するのではありません。", {
      x: ML + 0.2, y: 5.14, w: 5.3, h: 1.4, fontFace: F, fontSize: 16, bold: true, color: NAVY,
      align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.3,
    });
    s.addChart(pres.ChartType.bar,
      [{ name: "年間買電量（kWh）", labels: ["導入前", "導入後"], values: [null, null] },
       { name: "年間電気料金（円）", labels: ["導入前", "導入後"], values: [null, null] }],
      {
        x: 6.16, y: 1.8, w: 5.37, h: 4.74, barDir: "col", barGapWidthPct: 90,
        showTitle: true, title: "導入前後の比較", titleFontFace: F, titleFontSize: 15, titleColor: NAVY,
        chartColors: ["A6A6A6", NAVY],
        showLegend: true, legendPos: "b", legendFontFace: F, legendFontSize: 12,
        catAxisLabelFontFace: F, catAxisLabelFontSize: 13, catAxisLabelColor: INK,
        valAxisLabelFontFace: F, valAxisLabelFontSize: 12, valAxisLabelColor: INK,
        valGridLine: { color: "D9D9D9", size: 1 }, catGridLine: { style: "none" },
      });
    note(23, "表", "お客様概要（所在地／家族構成／システム容量／蓄電池）");
    note(23, "表", "年間買電量・年間電気料金の導入前後／削減率");
    note(23, "グラフ", "導入前後の買電量・電気料金");
  }

  /* ===== P24 お客様データ②（空欄） ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "実際のお客様のデータ",
      lead: "売電収入も含めた経済メリットの合計を算出する。",
      source: "※お客様の実測値。掲載にあたっては事前に許諾を得てください",
      footer: "売電も含めた「合計」で見ることが大切です",
      note: "差し替え：電気代の削減額／売電収入／合計の経済メリット（月あたり・年あたり）。",
    });
    table(s, [
      [th("項目"), th("月あたり"), th("年あたり")],
      [lbl("電気代の削減額"), blank(), blank()],
      [lbl("売電収入"), blank(), blank()],
      [{ text: "合計の経済メリット", options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "left" } },
       { text: "", options: { fill: { color: PALE } } }, { text: "", options: { fill: { color: PALE } } }],
    ], ML, 2.06, 6.5, [2.5, 2.0, 2.0], { rowH: 0.78, fontSize: 16 });
    card(s, 7.0, 2.06, 4.53, 3.1, { fill: WHITE, line: NAVY, lw: 3 });
    s.addText("1か月あたりの経済メリット合計", { x: 7.0, y: 2.26, w: 4.53, h: 0.46, fontFace: F, fontSize: 17, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    card(s, 7.5, 2.9, 3.53, 1.28, { fill: PALE, line: NAVY, lw: 1.5 });
    s.addText("円", { x: 10.24, y: 2.9, w: 0.6, h: 1.28, fontFace: F, fontSize: 22, bold: true, color: NAVY, align: "left", valign: "middle", margin: 0 });
    s.addText("電気代の削減額 ＋ 売電収入", { x: 7.0, y: 4.3, w: 4.53, h: 0.4, fontFace: F, fontSize: 14, bold: true, color: INK, align: "center", valign: "middle", margin: 0 });
    card(s, 7.0, 5.3, 4.53, 1.24, { fill: PALE, line: NAVY, lw: 1.5 });
    s.addText("電気代の削減だけで判断すると、実際の効果を小さく見積もることになります。", {
      x: 7.18, y: 5.3, w: 4.17, h: 1.24, fontFace: F, fontSize: 14, bold: true, color: NAVY,
      align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.3,
    });
    note(24, "表", "経済メリット（電気代削減額／売電収入／合計、月・年）");
  }

  /* ===== P25 蓄電池の必要性 ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "蓄電池の必要性",
      lead: "太陽光だけでも電気代は削減できるが、蓄電池を組み合わせると効果はさらに高まる。",
      footer: "蓄電池で、電気代削減と停電対策の効果が高まります",
    });
    compareTable(s, 0.42, 2.16, 10.85, "太陽光のみ", "太陽光＋蓄電池", [
      ["昼間の電気", "つくった電気をその場で使う", "使い、余りは貯める"],
      ["夜間の電気", "電力会社から買う", "昼に貯めた電気を使う"],
      ["自家消費率", "約30％", "約65％"],
      ["停電したとき", "日中しか使えない", "夜間も蓄電池から使える"],
      ["厳冬期の停電", "夜は暖をとれない", "最低限の暖房・照明を確保"],
    ], { rowH: 0.7, labelW: 2.4 });
  }

  /* ===== P26 補助金の金額 ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "補助金情報",
      lead: "長岡市から、太陽光・蓄電池の導入に対して手厚い補助金が出ている。",
      source: "出典：長岡市「令和8年度 雪国長岡での再エネ導入促進補助金」　※ZEH／ZEH＋の対象要件・金額は要確認",
      footer: "補助金をフル活用すれば、初期費用を大きく減らせます",
    });
    card(s, 0.24, 1.9, 11.21, 3.42, { fill: WHITE, line: NAVY, lw: 3 });
    s.addText("令和8年度　雪国長岡での再エネ導入促進補助金", {
      x: 0.4, y: 2.02, w: 10.89, h: 0.5, fontFace: F, fontSize: 21, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    s.addShape(pres.ShapeType.line, { x: 0.5, y: 2.58, w: 10.69, h: 0, line: { color: NAVY, width: 2.25 } });
    const items = [
      { x: 0.5, t: "太陽光発電", v: "7万円", u: "／kW", n: "上限35万円" },
      { x: 3.24, t: "蓄電池", v: "費用の1/3", u: "", n: "上限56.4万円" },
      { x: 5.98, t: "ZEH住宅", v: "55万円", u: "", n: "一律（新築）" },
      { x: 8.72, t: "ZEH＋住宅", v: "100万円", u: "", n: "一律（新築）" },
    ];
    items.forEach((it, i) => {
      band(s, it.x, 2.74, 2.5, it.t, { h: 0.44, size: 16 });
      s.addText([
        { text: it.v, options: { fontFace: F, fontSize: it.v.length > 4 ? 26 : 30, bold: true, color: NAVY } },
        { text: it.u, options: { fontFace: F, fontSize: 15, bold: true, color: NAVY } },
      ], { x: it.x, y: 3.26, w: 2.5, h: 0.72, align: "center", valign: "middle", margin: 0 });
      s.addText(it.n, { x: it.x, y: 3.98, w: 2.5, h: 0.36, fontFace: F, fontSize: 13, bold: true, color: INK, align: "center", valign: "middle", margin: 0 });
      if (i < 3) s.addShape(pres.ShapeType.line, { x: it.x + 2.62, y: 2.8, w: 0, h: 1.5, line: { color: "D9D9D9", width: 1 } });
    });
    s.addShape(pres.ShapeType.line, { x: 0.5, y: 4.46, w: 10.69, h: 0, line: { color: NAVY, width: 2.25 } });
    s.addText("補助対象は自家消費を目的とした設備。交付決定後に着手した事業が対象です。", {
      x: 0.5, y: 4.6, w: 10.69, h: 0.44, fontFace: F, fontSize: 15, bold: true, color: INK,
      align: "center", valign: "middle", margin: 0,
    });
    card(s, 0.6, 5.52, 5.2, 1.0, { fill: PALE, line: NAVY, lw: 2 });
    s.addText([
      { text: "太陽光＋蓄電池（既築もOK）", options: { fontFace: F, fontSize: 14, bold: true, color: NAVY, breakLine: true } },
      { text: "最大91.4万円", options: { fontFace: F, fontSize: 26, bold: true, color: NAVY } },
    ], { x: 0.7, y: 5.52, w: 5.0, h: 1.0, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.05 });
    card(s, 5.9, 5.52, 5.2, 1.0, { fill: PINK, line: RED, lw: 2 });
    s.addText([
      { text: "ZEH＋と併用できる新築なら", options: { fontFace: F, fontSize: 14, bold: true, color: RED, breakLine: true } },
      { text: "最大191.4万円", options: { fontFace: F, fontSize: 26, bold: true, color: RED } },
    ], { x: 6.0, y: 5.52, w: 5.0, h: 1.0, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.05 });
  }

  /* ===== P27 補助金の注意点 ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "補助金情報",
      lead: "補助金には申請の順序と期限がある。ここを外すと、受けられない。",
      source: "※受付状況・残予算は変動します。登壇直前に最新情報をご確認ください（受付開始日は要確認）",
      footer: "「決まってから工事」が鉄則です",
    });
    const warns = [
      ["申請受付は5月18日から", "受付開始前の申請は受け付けられません"],
      ["交付決定前の着工は対象外", "先に工事を始めると、補助を受けられません"],
      ["予算上限に達し次第、終了", "年度内でも、早期に締め切られることがあります"],
    ];
    const cw = 3.55, ch = 1.9, gx = 0.32, x0 = (SW - (cw * 3 + gx * 2)) / 2, y0 = 1.86;
    warns.forEach((w, i) => {
      const x = x0 + i * (cw + gx);
      card(s, x, y0, cw, ch, { fill: WHITE, line: RED, lw: 2 });
      band(s, x, y0, cw, "①②③".charAt(i) + "　" + w[0], { h: 0.6, size: 15, fill: RED });
      s.addText(w[1], { x: x + 0.2, y: y0 + 0.72, w: cw - 0.4, h: ch - 0.88, fontFace: F, fontSize: 15, bold: true, color: BLACK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.28 });
    });
    band(s, ML, 4.02, CW, "申請から工事までの正しい順序", { h: 0.5, size: 20 });
    const flow = ["申請", "交付決定", "着工", "完了報告", "補助金の交付"];
    const fw = 1.9, fgx = 0.42, fx0 = (SW - (fw * 5 + fgx * 4)) / 2, fy = 4.88;
    flow.forEach((t, i) => {
      const x = fx0 + i * (fw + fgx);
      card(s, x, fy, fw, 0.9, { fill: i === 1 || i === 2 ? PALE : WHITE, line: NAVY, lw: 2 });
      s.addText(t, { x, y: fy, w: fw, h: 0.9, fontFace: F, fontSize: 17, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
      if (i < 4) s.addShape(pres.ShapeType.rightArrow, { x: x + fw + 0.05, y: fy + 0.28, w: 0.32, h: 0.34, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    });
    s.addShape(pres.ShapeType.line, { x: fx0 + fw + fgx, y: 5.96, w: fw * 2 + fgx, h: 0, line: { color: RED, width: 3 } });
    s.addText("この順序が逆になると補助対象外", {
      x: fx0 + fw + fgx - 0.6, y: 6.02, w: fw * 2 + fgx + 1.2, h: 0.44, fontFace: F, fontSize: 16, bold: true, color: RED,
      align: "center", valign: "middle", margin: 0,
    });
  }

  /* ===== P28〜P31 章扉 ===== */
  dividerSlide(["太陽光発電", "良い商品ですよね！？"], 50);
  dividerSlide(["ですがもちろん", "デメリット・注意点", "もあります。。"], 46);
  dividerSlide(["皆様にはデメリットを知ったうえで", "判断してほしい。", "", "買って後悔してほしくない。"], 38);
  dividerSlide(["今回は", "すべてお伝えします！", "", "それに対する弊社の対策も", "お伝えします！"], 36);

  /* ===== P32 目次(3) ===== */
  agendaSlide(2);
};
