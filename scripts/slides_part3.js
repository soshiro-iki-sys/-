/* P33〜P46 */
module.exports = function (K) {
  const { pres, note, deco, NAVY, RED, HRED, YEL, WHITE, BLACK, PALE, PINK, GRAY2, INK, SOFT, F,
    SW, SH, ML, CW, contentSlide, card, band, bulletRow, statCard, table, th, lbl, val, blank,
    agendaSlide, twoColSlide, fitSize } = K;
  const CH3 = "3.太陽光・蓄電池のデメリットと弊社の対策";
  const CH4 = "4.正しい業者の選び方";

  /* ===== P33〜P37 デメリット5項目 ===== */
  twoColSlide({
    chapter: CH3, chip: "①初期費用の高さ",
    lead: "太陽光・蓄電池は安い商品ではなく、初期費用はどうしても掛かってくる。",
    leftTitle: "【デメリット】", rightTitle: "【ヤシロの対応】",
    left: ["太陽光5kW＋蓄電池のセットで、まとまった初期費用が必要になる",
           "補助金を使っても、自己資金または借入の準備は必要になる",
           "投資回収には年数がかかる（一般的には8〜12年）"],
    right: ["補助金の申請は全件サポート。申請漏れをつくらない",
            "ご予算とご状況に合わせた、無理のないプランをご提案",
            "「無理な販売はしない」ことをお約束します"],
    footer: "決して安くはありません。だからこそ最適なご提案を",
  });

  twoColSlide({
    chapter: CH3, chip: "②すべての屋根に向いているわけではない",
    lead: "残念ながら、すべての住宅の屋根に太陽光が向いているわけではない。",
    leftTitle: "【設置に向かない屋根】", rightTitle: "【ヤシロの対応】",
    left: ["北向きで、日射を十分に受けられない屋根",
           "周囲の建物や樹木の影が多くかかる屋根",
           "築年数が古く、屋根の下地や強度に不安がある家"],
    right: ["工務店としての経験を活かし、現地調査で屋根の状態を正確に診断",
            "向かない場合は、はっきりと「向きません」とお伝えします",
            "屋根の補修が先か、太陽光が先か。順序からご提案します"],
    footer: "向かない場合は、無理な設置はせず正直にお伝えします",
  });

  twoColSlide({
    chapter: CH3, chip: "③冬の時期の発電量は落ちる",
    lead: "技術の進歩で対策は可能だが、冬の発電量が落ちることは避けられない。",
    leftTitle: "【デメリット】", rightTitle: "【ヤシロの対応】",
    left: ["長岡の冬は曇天と積雪が多く、発電量は夏に比べて低下する",
           "冬の暖房費を発電だけで完全にカバーするのは難しい場合がある",
           "積雪による荷重や、落雪のリスクへの配慮が必要になる"],
    right: ["冬季の低下を織り込んだ「年間トータル」でシミュレーションを提示",
            "柏崎・長岡の自社発電所で蓄積した積雪対策のノウハウを反映",
            "蓄電池を組み合わせ、少ない発電量を効率よく使う設計にする"],
    footer: "冬の低下も見込んだ「年間トータル」で判断します",
  });

  twoColSlide({
    chapter: CH3, chip: "④機器の寿命とメンテナンスコスト",
    lead: "太陽光・蓄電池は長く使える設備だが、定期的なメンテナンスが必要になる。",
    leftTitle: "【デメリット】", rightTitle: "【ヤシロの対応】",
    left: ["パワーコンディショナの寿命は約10〜15年。交換費用が発生する",
           "長く使っていくためには、定期的な点検・清掃が必要になる",
           "積雪地では、架台や固定部の状態確認も欠かせない"],
    right: ["定期メンテナンスを含めた、長期のアフターフォロー体制",
            "地域密着だからこそ、不具合があればすぐに駆けつけられる",
            "メーカー保証に加え、弊社独自の施工保証をお付けします"],
    footer: "充実したアフターフォローと保証制度が重要です",
  });

  twoColSlide({
    chapter: CH3, chip: "⑤売電収入の低下",
    lead: "2026年度のFIT単価は当初4年24円、5年目以降は8.3円。売電頼みは成り立たない。",
    source: "出典：資源エネルギー庁「2026年度以降の調達価格等」（10kW未満・初期投資支援スキーム）",
    leftTitle: "【デメリット】", rightTitle: "【ヤシロの対応】",
    left: ["売電単価は当初4年24円／kWh、5〜10年目は8.3円／kWhまで下がる",
           "「売電で儲ける」という考え方で導入すると、失敗するリスクが高い",
           "売電を前提にした収支計画は、制度変更の影響を受けやすい"],
    right: ["売電ではなく「自家消費プラン」をご提案します",
            "蓄電池との組み合わせで、高い電気を買わない仕組みを構築",
            "売電収入に依存しない、堅実なシミュレーションを作成します"],
    footer: [{ t: "これからは、作って貯めて使う" }, { t: "「自家消費」", hl: true }, { t: "の時代です" }],
  });

  /* ===== P38 目次(4) ===== */
  agendaSlide(3);

  /* ===== P39 太陽光のトラブル（実データ） ===== */
  {
    const s = contentSlide({
      chapter: CH4, chip: "太陽光のトラブル",
      lead: "太陽光の点検商法に関する相談は、この7年で約11倍に増えている。",
      source: "出典：国民生活センター「太陽光発電システムの点検商法が急増！」（2025年6月4日公表・PIO-NET登録件数）",
      footer: "だからこそ、業者選びが非常に重要です",
    });
    const labels = ["2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024"];
    const values = [57, 59, 53, 62, 90, 154, 304, 613];
    s.addChart(pres.ChartType.bar, [{ name: "相談件数", labels, values }], {
      x: 0.42, y: 2.02, w: 6.5, h: 4.56, barDir: "col", barGapWidthPct: 50,
      showTitle: true, title: "太陽光発電システムの点検商法に関する相談件数（年度・件）", titleFontFace: F, titleFontSize: 14, titleColor: NAVY,
      chartColors: [RED], showLegend: false,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontFace: F, dataLabelFontSize: 12, dataLabelColor: INK, dataLabelFormatCode: "#,##0",
      catAxisLabelFontFace: F, catAxisLabelFontSize: 12, catAxisLabelColor: INK,
      valAxisLabelFontFace: F, valAxisLabelFontSize: 11, valAxisLabelColor: INK, valAxisMaxVal: 700,
      valGridLine: { color: "D9D9D9", size: 1 }, catGridLine: { style: "none" },
    });
    const items = [
      ["訪問販売による契約トラブル", "その場で契約を迫られ、冷静に比較する時間を与えられない"],
      ["「点検が義務化された」", "無料点検と称して訪問し、高額な洗浄・部材交換を契約させる"],
      ["説明と違う発電量", "有利な条件だけを並べたシミュレーション。実際は届かない"],
      ["連絡が取れない・倒産", "売ったあとは音信不通。保証書が使えない"],
    ];
    const cx = 7.16, cw = 4.37, chh = 1.0, gy = 0.1, y0 = 2.02;
    items.forEach((it, i) => {
      const y = y0 + i * (chh + gy);
      card(s, cx, y, cw, chh, { fill: WHITE, line: RED, lw: 1.5 });
      s.addText("●　" + it[0], { x: cx + 0.16, y: y + 0.08, w: cw - 0.32, h: 0.38, fontFace: F, fontSize: fitSize("●　" + it[0], cw - 0.4, 16, 12), bold: true, color: RED, align: "left", valign: "middle", margin: 0 });
      s.addText(it[1], { x: cx + 0.42, y: y + 0.46, w: cw - 0.58, h: 0.54, fontFace: F, fontSize: 12.5, bold: true, color: BLACK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
    });
    s.addText("誰から買うかで、結果が変わります", {
      x: cx, y: 6.4, w: cw, h: 0.36, fontFace: F, fontSize: 13, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
  }

  /* ===== P40 ①地域密着 ===== */
  {
    const s = contentSlide({
      chapter: CH4, chip: "①地域密着",
      lead: "長岡の地に根を張り、雪国での実証を重ねてきた会社かどうかを見てほしい。",
      source: "出典：登記情報・自社公開情報",
      footer: "県外の会社ではなく、地元で実績のある会社を",
      note: "差し替え：対応エリア／太陽光施工実績／長岡市との協業（名称・期間・成果）／講演会の登壇実績。設立・事業年数・太陽光の取り組み開始は記入済み。",
    });
    const rows = [
      [th("項目"), th("内容")],
      [lbl("設立"), val("1964年（昭和39年）1月10日")],
      [lbl("事業年数"), val("62年（2026年時点）")],
      [lbl("太陽光の取り組み開始"), val("2014年（柏崎・長岡で自社発電所）", { bold: true })],
      [lbl("対応エリア"), blank()],
      [lbl("太陽光施工実績"), blank()],
      [lbl("長岡市との協業（名称）"), blank()],
      [lbl("協業の期間"), blank()],
      [lbl("実験・実証の成果"), blank()],
      [lbl("講演会の登壇実績"), blank()],
    ];
    table(s, rows, ML, 1.82, 6.9, [2.5, 4.4], { rowH: 0.47, fontSize: 14 });
    card(s, 7.28, 1.82, 4.25, 4.72, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("地元の会社を選ぶ、3つの理由", {
      x: 7.28, y: 2.0, w: 4.25, h: 0.5, fontFace: F, fontSize: 19, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    [["売り逃げができない", "同じ地域で商売を続ける以上、逃げるという選択肢がない"],
     ["雪国での実証がある", "他県の一般論ではなく、長岡の気候で確かめた知見で話せる"],
     ["何かあればすぐ行ける", "距離が近いことは、それだけで保証の一部になる"]]
      .forEach((it, i) => {
        const y = 2.66 + i * 1.28;
        bulletRow(s, 7.46, y, 3.9, 0.5, i + 1, it[0], NAVY, 16);
        s.addText(it[1], { x: 8.12, y: y + 0.52, w: 3.24, h: 0.68, fontFace: F, fontSize: 12.5, bold: true, color: INK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.22 });
      });
    note(40, "表", "対応エリア／太陽光施工実績／長岡市との協業（名称・期間・成果）／講演会の登壇実績");
  }

  /* ===== P41 ②適正価格 ===== */
  twoColSlide({
    chapter: CH4, chip: "②適正価格",
    lead: "価格は安いほど良い、とは限らない。安すぎる見積りには理由がある。",
    leftTitle: "【安すぎる見積りの危うさ】", rightTitle: "【ヤシロの考え方】",
    left: ["工事費を削るために施工の質が落ちる。数年後に雨漏りで出てくる",
           "訪問販売の会社は、売ったあとの不具合に対応しないことがある",
           "「今日決めれば特別価格」は、比較させないための手法"],
    right: ["見積りの内訳を開示し、何にいくらかかるかを説明します",
            "相場から極端に外れた提案はしません。理由を添えて価格を出します",
            "他社と比較していただいて構いません。その時間をお取りください"],
    footer: "安さの理由を説明できる会社かどうかで判断を",
  });

  /* ===== P42 ③国産メーカー（保証内容 実数） ===== */
  {
    const s = contentSlide({
      chapter: CH4, chip: "③国産メーカーの取り扱い",
      lead: "弊社が国産メーカーの長州産業を扱う理由は、不具合時の対応にある。",
      source: "出典：長州産業の公表保証内容（2026年時点）　※適用条件・最新内容はメーカー資料をご確認ください",
      footer: "困ったときに動いてもらえるかで選びます",
    });
    card(s, ML, 1.94, 5.02, 4.6, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("長州産業を取り扱う理由", {
      x: ML, y: 2.12, w: 5.02, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    [["国内生産のメーカーである", "生産から供給まで国内で完結している"],
     ["不具合時の対応がしっかりしている", "連絡がつき、話が前に進む。これがいちばん大きい"],
     ["雪害が自然災害補償の対象", "雪国で使ううえで、これは大きな差になる"]]
      .forEach((it, i) => {
        const y = 2.8 + i * 1.24;
        bulletRow(s, 0.4, y, 4.6, 0.5, i + 1, it[0], NAVY, 16);
        s.addText(it[1], { x: 1.06, y: y + 0.52, w: 3.94, h: 0.62, fontFace: F, fontSize: 13, bold: true, color: INK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.22 });
      });
    band(s, 5.5, 1.94, 6.03, "保証内容", { h: 0.46, size: 18 });
    const rows = [
      [th("保証項目"), th("期間・内容")],
      [lbl("機器保証"), val("15年（無償）", { bold: true })],
      [lbl("出力保証"), val("25年（無償）", { bold: true })],
      [lbl("自然災害補償"), val("15年　台風・落雷・雪害など\n※地震・津波は対象外")],
      [lbl("施工保証（雨漏り）"), val("10年（メーカー保証）")],
      [lbl("弊社独自の施工保証"), blank()],
    ];
    table(s, rows, 5.5, 2.5, 6.03, [2.3, 3.73], { rowH: 0.66, fontSize: 14 });
    note(42, "表", "弊社独自の施工保証の年数・内容");
  }

  /* ===== P43 ④現地調査 ===== */
  {
    const s = contentSlide({
      chapter: CH4, chip: "④現地調査の丁寧さ",
      lead: "工務店として住宅を扱ってきたからこそ、屋根の下まで見て判断できる。",
      footer: "屋根に載せる前に、屋根そのものを見ます",
    });
    band(s, ML, 1.86, CW, "現地調査でヤシロが確認する項目", { h: 0.5, size: 20 });
    const checks = [
      ["屋根材の種類と状態", "瓦・スレート・金属。材質ごとに固定方法が変わる"],
      ["下地・野地板の健全性", "表面ではなく、留め付ける下地が持つかどうか"],
      ["雨仕舞い", "穴を開ける以上、水の処理をどうするか"],
      ["積雪荷重", "パネルと雪の重さに、構造が耐えられるか"],
      ["落雪の方向と着地点", "隣家・道路・カーポートに落とさない配置か"],
      ["電気容量・分電盤", "蓄電池・V2Hを入れられる余地があるか"],
    ];
    const cw = 3.55, chh = 1.42, gx = 0.32, gy = 0.26, x0 = (SW - (cw * 3 + gx * 2)) / 2, y0 = 2.6;
    checks.forEach((c, i) => {
      const x = x0 + (i % 3) * (cw + gx), y = y0 + Math.floor(i / 3) * (chh + gy);
      card(s, x, y, cw, chh, { fill: WHITE, line: NAVY, lw: 1.5 });
      s.addShape(pres.ShapeType.ellipse, { x: x + 0.18, y: y + 0.22, w: 0.42, h: 0.42, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
      s.addText(String(i + 1), { x: x + 0.18, y: y + 0.22, w: 0.42, h: 0.42, fontFace: F, fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      s.addText(c[0], { x: x + 0.7, y: y + 0.16, w: cw - 0.86, h: 0.5, fontFace: F, fontSize: 16, bold: true, color: NAVY, align: "left", valign: "middle", margin: 0 });
      s.addText(c[1], { x: x + 0.22, y: y + 0.72, w: cw - 0.44, h: 0.58, fontFace: F, fontSize: 12.5, bold: true, color: INK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
    });
    s.addText("屋根の状態によっては「今はやめておきましょう」とお伝えすることもあります。", {
      x: ML, y: 5.98, w: CW, h: 0.5, fontFace: F, fontSize: 18, bold: true, color: RED,
      align: "center", valign: "middle", margin: 0,
    });
  }

  /* ===== P44 まとめ ===== */
  {
    const s = contentSlide({ chapter: "まとめ", footer: "本日の内容を、5つに整理します" });
    const items = [
      "太陽光・蓄電池は、電気代高騰と停電への「備え」になる",
      "長岡・雪国でも、太陽光は十分に有効。データがそれを示している",
      "デメリットはある。ただし正しく業者を選べば、対策はできる",
      "補助金は交付決定前の着工がNG。動くなら早めに",
      "業者選びは「地域密着・適正価格・国産メーカー・丁寧な現地調査」",
    ];
    const y0 = 1.3, h = 1.02, gap = 0.16;
    items.forEach((t, i) => {
      const y = y0 + i * (h + gap);
      card(s, 0.6, y, SW - 1.2, h, { fill: i % 2 === 0 ? PALE : WHITE, line: NAVY, lw: 1.5 });
      s.addShape(pres.ShapeType.ellipse, { x: 0.86, y: y + (h - 0.58) / 2, w: 0.58, h: 0.58, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
      s.addText(String(i + 1), { x: 0.86, y: y + (h - 0.58) / 2, w: 0.58, h: 0.58, fontFace: F, fontSize: 22, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      s.addText(t, { x: 1.62, y, w: SW - 2.4, h, fontFace: F, fontSize: 20, bold: true, color: BLACK, align: "left", valign: "middle", margin: 0 });
    });
  }

  /* ===== P45 個別相談のご案内 ===== */
  {
    const s = contentSlide({
      chapter: "まとめ", chip: "個別相談のご案内",
      lead: "本講座は、まずは正しい知識を身につけていただくことを目的としている。",
      footer: "無理な売り込みは一切いたしません",
    });
    card(s, 0.42, 1.86, 5.36, 4.68, { fill: WHITE, line: NAVY, lw: 2 });
    band(s, 0.42, 1.86, 5.36, "参加をご検討の方へ", { h: 0.62, size: 20 });
    s.addText(
      "無理な売り込みは一切いたしませんので、\nお気軽にご参加ください。\n\n" +
      "お見積り・現地調査をご希望の方は、\n恐れ入りますが講座終了後にスタッフへ\nお声がけください。\n\n" +
      "その場でご契約をお願いすることは\nございません。",
      { x: 0.72, y: 2.72, w: 4.76, h: 3.6, fontFace: F, fontSize: 17, bold: true, color: BLACK,
        align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.4 }
    );
    card(s, 5.92, 1.86, 5.36, 4.68, { fill: PALE, line: NAVY, lw: 2 });
    band(s, 5.92, 1.86, 5.36, "参加者特典", { h: 0.62, size: 20 });
    s.addText("『太陽光・蓄電池の\n　すべてが分かる小冊子』", {
      x: 6.22, y: 2.78, w: 4.76, h: 1.2, fontFace: F, fontSize: 24, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2,
    });
    s.addShape(pres.ShapeType.line, { x: 6.42, y: 4.06, w: 4.36, h: 0, line: { color: NAVY, width: 1 } });
    s.addText("太陽光の仕組み・よくある失敗から\nお得な使い方まで、すべて掲載しています。", {
      x: 6.22, y: 4.22, w: 4.76, h: 0.96, fontFace: F, fontSize: 15, bold: true, color: BLACK,
      align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.25,
    });
    card(s, 6.42, 5.3, 4.36, 0.96, { fill: WHITE, line: NAVY, lw: 1.5 });
    s.addText("株式会社ヤシロ　TEL 0120-17-5521\n営業時間 8:00〜17:00／定休日 日曜・祝日", {
      x: 6.42, y: 5.3, w: 4.36, h: 0.96, fontFace: F, fontSize: 14, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.25,
    });
  }

  /* ===== P46 御礼 ===== */
  {
    const s = pres.addSlide();
    s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: SW, h: SH, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    K.deco(s, "divider");
    s.addShape(pres.ShapeType.line, { x: 3.6, y: 2.3, w: 4.5, h: 0, line: { color: WHITE, width: 1.5 } });
    s.addText("本日はお忙しいところご来場いただき\n誠にありがとうございました", {
      x: 0.6, y: 2.62, w: SW - 1.2, h: 1.8, fontFace: F, fontSize: 34, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.3,
    });
    s.addText("今後とも株式会社ヤシロを\n何卒よろしくお願いいたします", {
      x: 0.6, y: 4.5, w: SW - 1.2, h: 1.4, fontFace: F, fontSize: 24, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.3,
    });
    s.addShape(pres.ShapeType.line, { x: 3.6, y: 6.02, w: 4.5, h: 0, line: { color: WHITE, width: 1.5 } });
    s.addText("詳しくは、講座終了後に個別でご相談ください", {
      x: 0.35, y: 6.4, w: SW - 0.7, h: 0.6, fontFace: F, fontSize: 22, bold: true, color: YEL,
      align: "center", valign: "middle", margin: 0,
    });
    s.addNotes("差し替え：背景写真（自社施工事例など）を全面に敷き、現状のネイビー地と入れ替える。");
    note(46, "背景", "御礼ページ背景写真（現状はネイビー単色）");
  }
};
