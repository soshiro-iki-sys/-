/* P16〜P31 */
module.exports = function (K) {
  const { pres, note, NAVY, RED, HRED, YEL, WHITE, BLACK, PALE, F, SW, SH, ML, CW,
    contentSlide, bulletRow, statCard, table, th, lbl, val, blank,
    agendaSlide, dividerSlide } = K;
  const CH2 = "2.長岡市でも太陽光はお得なのか";

  /* ===== P16 目次(2) ===== */
  agendaSlide(1);

  /* ===== P17 章イメージ ===== */
  {
    const s = contentSlide({ chapter: CH2, footer: "イメージではなく、データで判断しましょう！" });
    s.addText("「長岡の雪じゃ、発電しないでしょ？」", {
      x: 0.4, y: 1.28, w: SW - 0.8, h: 1.0, fontFace: F, fontSize: 40, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("いちばん多くいただくご質問です。ひとつずつ、データで確かめていきます。", {
      x: 0.4, y: 2.32, w: SW - 0.8, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: BLACK,
      align: "center", valign: "middle", margin: 0,
    });
    const qs = [
      ["日射量は足りるのか", "新潟県の日射量は、全国平均と比べてどうなのか"],
      ["冬はどれくらい落ちるのか", "落ちるのは事実。ではゼロになるのか"],
      ["積雪で壊れないのか", "荷重・落雪に、どう設計で対応するのか"],
    ];
    const cw = 3.55, ch = 3.36, gx = 0.32, x0 = (SW - (cw * 3 + gx * 2)) / 2, y0 = 3.1;
    qs.forEach((q, i) => {
      const x = x0 + i * (cw + gx);
      K.card(s, x, y0, cw, ch, { fill: WHITE, line: NAVY, lw: 2 });
      s.addShape(pres.ShapeType.ellipse, { x: x + cw / 2 - 0.4, y: y0 - 0.38, w: 0.8, h: 0.8, fill: { color: NAVY }, line: { color: WHITE, width: 2 } });
      s.addText("Q" + (i + 1), { x: x + cw / 2 - 0.4, y: y0 - 0.38, w: 0.8, h: 0.8, fontFace: F, fontSize: 22, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      s.addText(q[0], { x: x + 0.18, y: y0 + 0.62, w: cw - 0.36, h: 0.9, fontFace: F, fontSize: 20, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.15 });
      s.addShape(pres.ShapeType.line, { x: x + 0.4, y: y0 + 1.6, w: cw - 0.8, h: 0, line: { color: NAVY, width: 1 } });
      s.addText(q[1], { x: x + 0.24, y: y0 + 1.76, w: cw - 0.48, h: 1.4, fontFace: F, fontSize: 15, bold: true, color: BLACK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.25 });
    });
  }

  /* ===== P18 新潟県の日射量（グラフ空） ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "新潟県の日射量",
      lead: "NEDOのデータでは、新潟県の日射量は全国平均とほぼ同水準である。",
      source: "出典：NEDO 日射量データベース（MONSOLA）より作成",
      footer: "「日射量が少ないから無理」は思い込みです",
      note: "差し替え：月別日射量（新潟県（長岡）／全国平均）。グラフのデータシートに入力する。",
    });
    const labels = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"];
    const empty = labels.map(() => null);
    s.addChart(pres.ChartType.bar,
      [{ name: "新潟県（長岡）", labels, values: empty.slice() }, { name: "全国平均", labels, values: empty.slice() }],
      {
        x: 0.5, y: 1.74, w: 10.69, h: 4.9, barDir: "col", barGapWidthPct: 60,
        showTitle: true, title: "月別 日射量の比較（kWh/㎡/日）", titleFontFace: F, titleFontSize: 16, titleColor: NAVY,
        chartColors: [NAVY, "A6A6A6"],
        showLegend: true, legendPos: "b", legendFontFace: F, legendFontSize: 13,
        catAxisLabelFontFace: F, catAxisLabelFontSize: 13, catAxisLabelColor: "404040",
        valAxisLabelFontFace: F, valAxisLabelFontSize: 13, valAxisLabelColor: "404040",
        valAxisTitle: "kWh/㎡/日", showValAxisTitle: true, valAxisTitleFontFace: F, valAxisTitleFontSize: 12,
        valGridLine: { color: "D9D9D9", size: 1 }, catGridLine: { style: "none" },
      });
    note(18, "グラフ", "月別日射量（新潟県（長岡）／全国平均）");
  }

  /* ===== P19 冬の発電量と積雪対策（表空） ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "冬の発電量と積雪対策",
      lead: "冬の発電量が落ちるのは事実である。ただし、ゼロになるわけではない。",
      source: "※自社発電所（柏崎・長岡）の実測値。計測期間・条件は要記載",
      footer: "冬は確実に落ちますが、ゼロにはなりません",
      note: "差し替え：自社発電所（柏崎・長岡）の月別実測値（発電量／最深積雪／備考）。月の列は記入済み。",
    });
    s.addText("自社発電所の積雪期 実測データ", {
      x: ML, y: 1.76, w: 5.7, h: 0.44, fontFace: F, fontSize: 18, bold: true, color: WHITE,
      fill: { color: NAVY }, align: "center", valign: "middle", margin: 0,
    });
    const rows = [
      [th("月"), th("発電量（kWh）"), th("最深積雪（cm）"), th("備考")],
      [val("11月", { align: "center", bold: true }), blank(), blank(), blank()],
      [val("12月", { align: "center", bold: true }), blank(), blank(), blank()],
      [val("1月", { align: "center", bold: true }), blank(), blank(), blank()],
      [val("2月", { align: "center", bold: true }), blank(), blank(), blank()],
      [val("3月", { align: "center", bold: true }), blank(), blank(), blank()],
    ];
    table(s, rows, ML, 2.3, 5.7, [0.94, 1.62, 1.62, 1.52], { rowH: 0.62, fontSize: 14 });
    s.addText("雪国での積雪対策", {
      x: 6.16, y: 1.76, w: 5.37, h: 0.44, fontFace: F, fontSize: 18, bold: true, color: WHITE,
      fill: { color: NAVY }, align: "center", valign: "middle", margin: 0,
    });
    K.card(s, 6.16, 2.3, 5.37, 4.24, { fill: WHITE, line: NAVY, lw: 2 });
    [["屋根の傾斜角を活かした自然落雪", "雪が自然に滑り落ちる角度と配置を、現地調査で見極める"],
     ["雪下ろしを前提としない設計", "屋根に上がらなくても運用できる形にする。安全がいちばん"],
     ["架台・荷重の設計", "積雪荷重に耐える架台と固定方法を選定する"]]
      .forEach((it, i) => {
        const y = 2.54 + i * 1.32;
        bulletRow(s, 6.42, y, 4.9, 0.54, i + 1, it[0], NAVY, 17);
        s.addText(it[1], { x: 7.1, y: y + 0.54, w: 4.22, h: 0.62, fontFace: F, fontSize: 14, bold: true, color: "404040", align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
      });
    note(19, "表", "自社発電所の積雪期実測（発電量／最深積雪／備考）");
  }

  /* ===== P20 長岡市の発電量 ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "長岡市の発電量",
      lead: "長岡市で5kWの太陽光を設置した場合の発電量を試算すると、家庭の消費電力の多くを賄えることが分かる。",
      leadLines: 2,
      source: "※1kWあたりの発電量1,100kWhは、日本の国家規格「JIS C 8907（太陽光発電システムの発電電力量推定方法）」をもとに算出",
      footer: "消費電力の84％をカバーすることが可能に！",
    });
    s.addText("長岡市での発電量試算：5kWシステムの場合", {
      x: ML, y: 2.06, w: CW, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: WHITE,
      fill: { color: NAVY }, align: "center", valign: "middle", margin: 0,
    });
    K.card(s, 0.8, 2.76, 4.62, 1.66, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("年間の発電量", { x: 0.8, y: 2.88, w: 4.62, h: 0.36, fontFace: F, fontSize: 15, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    s.addText([
      { text: "5kW × 1,100kWh ＝ ", options: { fontFace: F, fontSize: 17, bold: true, color: BLACK } },
      { text: "約5,500", options: { fontFace: F, fontSize: 30, bold: true, color: NAVY } },
      { text: " kWh", options: { fontFace: F, fontSize: 16, bold: true, color: NAVY } },
    ], { x: 0.86, y: 3.26, w: 4.50, h: 1.0, align: "center", valign: "middle", margin: 0 });
    K.card(s, 0.8, 4.62, 4.62, 1.66, { fill: "F2F2F2", line: "808080", lw: 2 });
    s.addText("1年間の電気使用量（4人家族・オール電化）", { x: 0.8, y: 4.74, w: 4.62, h: 0.36, fontFace: F, fontSize: 15, bold: true, color: "404040", align: "center", valign: "middle", margin: 0 });
    s.addText([
      { text: "6,500", options: { fontFace: F, fontSize: 32, bold: true, color: "404040" } },
      { text: " kWh", options: { fontFace: F, fontSize: 18, bold: true, color: "404040" } },
    ], { x: 0.94, y: 5.12, w: 4.34, h: 1.0, align: "center", valign: "middle", margin: 0 });
    s.addShape(pres.ShapeType.rightArrow, { x: 5.62, y: 4.16, w: 0.72, h: 0.72, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    K.card(s, 6.54, 2.76, 4.72, 3.52, { fill: WHITE, line: NAVY, lw: 3 });
    s.addText("消費電力のカバー率", { x: 6.54, y: 3.0, w: 4.72, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    s.addText([
      { text: "84", options: { fontFace: F, fontSize: 92, bold: true, color: NAVY } },
      { text: "％", options: { fontFace: F, fontSize: 36, bold: true, color: NAVY } },
    ], { x: 6.7, y: 3.5, w: 4.4, h: 1.82, align: "center", valign: "middle", margin: 0 });
    s.addText("5,500kWh ÷ 6,500kWh", { x: 6.54, y: 5.32, w: 4.72, h: 0.4, fontFace: F, fontSize: 16, bold: true, color: "404040", align: "center", valign: "middle", margin: 0 });
    s.addText("雪国の長岡でも、この水準です。", { x: 6.54, y: 5.74, w: 4.72, h: 0.42, fontFace: F, fontSize: 17, bold: true, color: RED, align: "center", valign: "middle", margin: 0 });
  }

  /* ===== P21 電気代削減の試算（表空） ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "電気代削減の試算",
      lead: "発電量から電気代の削減額を試算する。前提となる単価は、ご家庭の契約プランによって異なる。",
      leadLines: 2,
      footer: "我が家の場合はいくらか、個別に試算できます",
      note: "差し替え：自家消費率／買電単価／売電単価／年間の買電削減額／年間の売電収入／年間の経済メリット合計。年間発電量は記入済み。",
    });
    s.addText("電気代削減の試算（前提と結果）", {
      x: ML, y: 2.06, w: 6.5, h: 0.46, fontFace: F, fontSize: 18, bold: true, color: WHITE,
      fill: { color: NAVY }, align: "center", valign: "middle", margin: 0,
    });
    const rows = [
      [lbl("年間発電量（kWh）"), val("約5,500kWh", { align: "right", bold: true })],
      [lbl("自家消費率（%）"), blank()],
      [lbl("買電単価（円/kWh）"), blank()],
      [lbl("売電単価（円/kWh）"), blank()],
      [lbl("年間の買電削減額（円）"), blank()],
      [lbl("年間の売電収入（円）"), blank()],
      [{ text: "年間の経済メリット合計（円）", options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "left" } },
       { text: "", options: { fill: { color: PALE } } }],
    ];
    table(s, rows, ML, 2.62, 6.5, [3.7, 2.8], { rowH: 0.56, fontSize: 15 });
    K.card(s, 7.0, 2.62, 4.53, 3.92, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("単価はご契約プランで変わります", {
      x: 7.2, y: 2.86, w: 4.13, h: 0.8, fontFace: F, fontSize: 20, bold: true, color: NAVY,
      align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.15,
    });
    s.addShape(pres.ShapeType.line, { x: 7.2, y: 3.74, w: 4.13, h: 0, line: { color: NAVY, width: 1 } });
    s.addText(
      "・契約プラン（従量電灯／時間帯別）\n" +
      "・ご家族の人数と在宅時間\n" +
      "・オール電化かどうか\n" +
      "・屋根の向きと形状\n\n" +
      "これらで削減額は大きく変わります。\n一般論ではなく、お住まいの\n条件で計算します。",
      { x: 7.2, y: 3.9, w: 4.13, h: 2.5, fontFace: F, fontSize: 14, bold: true, color: BLACK,
        align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.3 }
    );
    note(21, "表", "電気代削減の試算（自家消費率／買電・売電単価／削減額／売電収入／合計）");
  }

  /* ===== P22 お客様データ①（表＋グラフ空） ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "実際のお客様のデータ",
      lead: "新潟県内にお住まいのお客様の、導入前後の実測データである。",
      source: "※お客様の実測値。掲載にあたっては事前に許諾を得てください",
      footer: "実際のお客様のデータで確かめてみましょう",
      note: "差し替え：お客様概要（所在地・家族構成・システム容量・蓄電池）／年間買電量・年間電気料金の導入前後と削減率／グラフのデータシート。",
    });
    const head = [
      [th("所在地"), th("家族構成"), th("システム容量"), th("蓄電池")],
      [blank(), blank(), blank(), blank()],
    ];
    table(s, head, ML, 1.8, 5.7, [1.42, 1.42, 1.5, 1.36], { rowH: 0.44, fontSize: 13 });
    const rows = [
      [th("項目"), th("導入前"), th("導入後"), th("削減率")],
      [lbl("年間買電量（kWh）"), blank(), blank(), blank()],
      [lbl("年間電気料金（円）"), blank(), blank(), blank()],
    ];
    table(s, rows, ML, 3.0, 5.7, [1.86, 1.32, 1.32, 1.2], { rowH: 0.62, fontSize: 14 });
    K.card(s, ML, 5.14, 5.7, 1.4, { fill: PALE, line: NAVY, lw: 1.5 });
    s.addText("同じ暮らし方のまま、買う電気の量そのものが減ります。\n節約をがまんで実現するのではありません。", {
      x: ML + 0.2, y: 5.14, w: 5.3, h: 1.4, fontFace: F, fontSize: 16, bold: true, color: NAVY,
      align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.25,
    });
    const labels = ["導入前", "導入後"];
    s.addChart(pres.ChartType.bar,
      [{ name: "年間買電量（kWh）", labels, values: [null, null] }, { name: "年間電気料金（円）", labels, values: [null, null] }],
      {
        x: 6.16, y: 1.8, w: 5.37, h: 4.74, barDir: "col", barGapWidthPct: 90,
        showTitle: true, title: "導入前後の比較", titleFontFace: F, titleFontSize: 15, titleColor: NAVY,
        chartColors: ["A6A6A6", NAVY],
        showLegend: true, legendPos: "b", legendFontFace: F, legendFontSize: 12,
        catAxisLabelFontFace: F, catAxisLabelFontSize: 13, catAxisLabelColor: "404040",
        valAxisLabelFontFace: F, valAxisLabelFontSize: 12, valAxisLabelColor: "404040",
        valGridLine: { color: "D9D9D9", size: 1 }, catGridLine: { style: "none" },
      });
    note(22, "表", "お客様概要（所在地／家族構成／システム容量／蓄電池）");
    note(22, "表", "年間買電量・年間電気料金の導入前後／削減率");
    note(22, "グラフ", "導入前後の買電量・電気料金");
  }

  /* ===== P23 お客様データ②（表空） ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "実際のお客様のデータ",
      lead: "売電収入も含めた経済メリットの合計を算出する。",
      source: "※お客様の実測値。掲載にあたっては事前に許諾を得てください",
      footer: "売電も含めた「合計」で見ることが大切です",
      note: "差し替え：電気代の削減額／売電収入／合計の経済メリット（月あたり・年あたり）。",
    });
    const rows = [
      [th("項目"), th("月あたり"), th("年あたり")],
      [lbl("電気代の削減額"), blank(), blank()],
      [lbl("売電収入"), blank(), blank()],
      [{ text: "合計の経済メリット", options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "left" } },
       { text: "", options: { fill: { color: PALE } } }, { text: "", options: { fill: { color: PALE } } }],
    ];
    table(s, rows, ML, 2.06, 6.5, [2.5, 2.0, 2.0], { rowH: 0.78, fontSize: 16 });
    K.card(s, 7.0, 2.06, 4.53, 3.12, { fill: WHITE, line: NAVY, lw: 3 });
    s.addText("1か月あたりの経済メリット合計", {
      x: 7.0, y: 2.3, w: 4.53, h: 0.5, fontFace: F, fontSize: 17, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    s.addShape(pres.ShapeType.line, { x: 7.6, y: 4.14, w: 3.33, h: 0, line: { color: NAVY, width: 2 } });
    s.addText("円", { x: 10.5, y: 3.62, w: 0.6, h: 0.5, fontFace: F, fontSize: 24, bold: true, color: NAVY, align: "left", valign: "middle", margin: 0 });
    s.addText("電気代の削減額 ＋ 売電収入", { x: 7.0, y: 4.28, w: 4.53, h: 0.4, fontFace: F, fontSize: 14, bold: true, color: "404040", align: "center", valign: "middle", margin: 0 });
    K.card(s, 7.0, 5.4, 4.53, 1.14, { fill: PALE, line: NAVY, lw: 1.5 });
    s.addText("電気代の削減だけで判断すると、実際の効果を\n小さく見積もることになります。", {
      x: 7.16, y: 5.4, w: 4.21, h: 1.14, fontFace: F, fontSize: 14, bold: true, color: NAVY,
      align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.25,
    });
    note(23, "表", "経済メリット（電気代削減額／売電収入／合計、月・年）");
  }

  /* ===== P24 蓄電池の必要性 ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "蓄電池の必要性",
      lead: "太陽光だけでも電気代は削減できるが、蓄電池を組み合わせると効果はさらに高まる。",
      leadLines: 2,
      footer: "蓄電池で、電気代削減と停電対策の効果がさらに高まります",
    });
    const cw = 5.36, lx = 0.42, rx = SW - 0.42 - cw, y0 = 2.06, ch = 4.48;
    K.card(s, lx, y0, cw, ch, { fill: WHITE, line: "808080", lw: 2 });
    s.addShape(pres.ShapeType.rect, { x: lx, y: y0, w: cw, h: 0.64, fill: { color: "808080" }, line: { color: "808080", width: 0 } });
    s.addText("太陽光のみ", { x: lx, y: y0, w: cw, h: 0.64, fontFace: F, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    K.card(s, rx, y0, cw, ch, { fill: WHITE, line: NAVY, lw: 2 });
    s.addShape(pres.ShapeType.rect, { x: rx, y: y0, w: cw, h: 0.64, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    s.addText("太陽光＋蓄電池", { x: rx, y: y0, w: cw, h: 0.64, fontFace: F, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    const rowsC = [
      ["昼間の電気", "つくった電気をその場で使う", "つくった電気を使い、余りは貯める"],
      ["夜間の電気", "電力会社から買う", "昼に貯めた電気を使う"],
      ["停電したとき", "日中しか使えない", "夜間も蓄電池から使える"],
      ["厳冬期の停電", "夜は暖をとれない", "最低限の暖房・照明を確保できる"],
    ];
    rowsC.forEach((c, i) => {
      const y = y0 + 0.98 + i * 0.9;
      s.addText(c[0], { x: 0.42, y: y - 0.28, w: SW - 0.84, h: 0.26, fontFace: F, fontSize: 12, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
      s.addText(c[1], { x: lx + 0.24, y, w: cw - 0.48, h: 0.58, fontFace: F, fontSize: 16, bold: true, color: "404040", fill: { color: "F2F2F2" }, align: "center", valign: "middle", margin: 0 });
      s.addText(c[2], { x: rx + 0.24, y, w: cw - 0.48, h: 0.58, fontFace: F, fontSize: 16, bold: true, color: NAVY, fill: { color: PALE }, align: "center", valign: "middle", margin: 0 });
    });
  }

  /* ===== P25 補助金の金額 ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "補助金情報",
      lead: "長岡市から、太陽光・蓄電池の導入に対して手厚い補助金が出ている。",
      source: "出典：令和8年度 雪国長岡での再エネ導入促進補助金　※ZEH／ZEH+の対象要件は要確認",
      footer: [{ t: "太陽光＋蓄電池なら最大91.4万円。ZEH＋と併用できる新築なら" }, { t: "最大191.4万円", hl: true }],
      footerSize: 22,
    });
    K.card(s, 0.24, 1.74, 11.21, 4.9, { fill: WHITE, line: NAVY, lw: 3 });
    s.addText("令和8年度　雪国長岡での再エネ導入促進補助金", {
      x: 0.4, y: 1.9, w: 10.89, h: 0.56, fontFace: F, fontSize: 22, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    s.addShape(pres.ShapeType.line, { x: 0.5, y: 2.54, w: 10.69, h: 0, line: { color: NAVY, width: 2.25 } });
    s.addShape(pres.ShapeType.line, { x: 5.84, y: 2.68, w: 0, h: 2.0, line: { color: NAVY, width: 2.25 } });
    const items = [
      { x: 0.6, y: 2.72, w: 4.9, t: "太陽光発電", v: "7万円", u: "/kW", n: "（上限35万円）" },
      { x: 6.2, y: 2.72, w: 4.9, t: "蓄電池", v: "費用の1/3", u: "", n: "（上限56.4万円）" },
      { x: 0.6, y: 4.72, w: 4.9, t: "ZEH住宅", v: "55万円", u: "", n: "（一律）" },
      { x: 6.2, y: 4.72, w: 4.9, t: "ZEH＋住宅", v: "100万円", u: "", n: "（一律）" },
    ];
    items.forEach((it) => {
      s.addText(it.t, { x: it.x, y: it.y, w: it.w, h: 0.44, fontFace: F, fontSize: 17, bold: true, color: WHITE, fill: { color: NAVY }, align: "center", valign: "middle", margin: 0 });
      s.addText([
        { text: it.v, options: { fontFace: F, fontSize: 36, bold: true, color: NAVY } },
        { text: it.u, options: { fontFace: F, fontSize: 18, bold: true, color: NAVY } },
        { text: "　" + it.n, options: { fontFace: F, fontSize: 15, bold: true, color: "404040" } },
      ], { x: it.x, y: it.y + 0.5, w: it.w, h: 0.86, align: "center", valign: "middle", margin: 0 });
    });
    s.addShape(pres.ShapeType.line, { x: 0.5, y: 4.6, w: 10.69, h: 0, line: { color: NAVY, width: 2.25 } });
    K.card(s, 0.6, 5.66, 5.3, 0.82, { fill: PALE, line: NAVY, lw: 2 });
    s.addText([
      { text: "太陽光＋蓄電池（既築もOK）", options: { fontFace: F, fontSize: 14, bold: true, color: NAVY, breakLine: true } },
      { text: "最大91.4万円", options: { fontFace: F, fontSize: 24, bold: true, color: NAVY } },
    ], { x: 0.7, y: 5.66, w: 5.1, h: 0.82, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.05 });
    K.card(s, 6.2, 5.66, 5.1, 0.82, { fill: "FDE9E7", line: RED, lw: 2 });
    s.addText([
      { text: "ZEH＋と併用できる新築", options: { fontFace: F, fontSize: 14, bold: true, color: RED, breakLine: true } },
      { text: "最大191.4万円", options: { fontFace: F, fontSize: 24, bold: true, color: RED } },
    ], { x: 6.3, y: 5.66, w: 4.9, h: 0.82, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.05 });
  }

  /* ===== P26 補助金の注意点 ===== */
  {
    const s = contentSlide({
      chapter: CH2, chip: "補助金情報",
      lead: "補助金には申請の順序と期限がある。ここを外すと、受けられない。",
      source: "※受付状況・残予算は変動します。登壇直前に最新情報をご確認ください",
      footer: "「決まってから工事」が鉄則です",
    });
    const warns = [
      ["申請受付は5月18日から", "受付開始前の申請は受け付けられません"],
      ["交付決定前の着工は対象外", "先に工事を始めると、補助を受けられません"],
      ["予算上限に達し次第、終了", "年度内でも、早期に締め切られることがあります"],
    ];
    const cw = 3.55, ch = 1.92, gx = 0.32, x0 = (SW - (cw * 3 + gx * 2)) / 2, y0 = 1.86;
    warns.forEach((w, i) => {
      const x = x0 + i * (cw + gx);
      K.card(s, x, y0, cw, ch, { fill: WHITE, line: RED, lw: 2 });
      s.addShape(pres.ShapeType.rect, { x, y: y0, w: cw, h: 0.6, fill: { color: RED }, line: { color: RED, width: 0 } });
      s.addText("①②③".charAt(i) + "　" + w[0], { x: x + 0.06, y: y0, w: cw - 0.12, h: 0.6, fontFace: F, fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      s.addText(w[1], { x: x + 0.2, y: y0 + 0.72, w: cw - 0.4, h: ch - 0.9, fontFace: F, fontSize: 15, bold: true, color: BLACK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.25 });
    });
    s.addText("申請から工事までの正しい順序", {
      x: ML, y: 4.06, w: CW, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: WHITE,
      fill: { color: NAVY }, align: "center", valign: "middle", margin: 0,
    });
    const flow = ["申請", "交付決定", "着工", "完了報告", "補助金の交付"];
    const fw = 1.9, fgx = 0.42, fx0 = (SW - (fw * 5 + fgx * 4)) / 2, fy = 4.94;
    flow.forEach((t, i) => {
      const x = fx0 + i * (fw + fgx);
      K.card(s, x, fy, fw, 0.92, { fill: i === 1 || i === 2 ? PALE : WHITE, line: NAVY, lw: 2 });
      s.addText(t, { x, y: fy, w: fw, h: 0.92, fontFace: F, fontSize: 18, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
      if (i < 4) s.addShape(pres.ShapeType.rightArrow, { x: x + fw + 0.05, y: fy + 0.28, w: 0.32, h: 0.36, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    });
    s.addShape(pres.ShapeType.line, { x: fx0 + fw + fgx, y: 6.06, w: fw * 2 + fgx, h: 0, line: { color: RED, width: 3 } });
    s.addText("この順序が逆になると補助対象外", {
      x: fx0 + fw + fgx - 0.5, y: 6.12, w: fw * 2 + fgx + 1.0, h: 0.44, fontFace: F, fontSize: 16, bold: true, color: RED,
      align: "center", valign: "middle", margin: 0,
    });
  }

  /* ===== P27〜P30 章扉 ===== */
  dividerSlide(["太陽光発電", "良い商品ですよね！？"], 52);
  dividerSlide(["ですがもちろん", "デメリット・注意点", "もあります。。"], 48);
  dividerSlide(["皆様にはデメリットを知ったうえで", "判断してほしい。", "", "買って後悔してほしくない。"], 40);
  dividerSlide(["今回は", "すべてお伝えします！", "", "それに対する弊社の対策も", "お伝えします！"], 38);

  /* ===== P31 目次(3) ===== */
  agendaSlide(2);
};
