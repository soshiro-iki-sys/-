/* P32〜P45 */
module.exports = function (K) {
  const { pres, note, NAVY, RED, HRED, YEL, WHITE, BLACK, PALE, F, SW, SH, ML, CW,
    contentSlide, bulletRow, statCard, table, th, lbl, val, blank,
    agendaSlide, twoColSlide } = K;
  const CH3 = "3.太陽光・蓄電池のデメリットと弊社の対策";
  const CH4 = "4.正しい業者の選び方";

  /* ===== P32〜P36 デメリット5項目 ===== */
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
    footer: "決して安くはありません。だからこそ、最適なご提案をさせていただきます",
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
    footer: "冬の発電低下も見込んだ「年間トータル」での判断が大切です",
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
    lead: "売電単価は年々下がり、大きな売電収入は見込めなくなってきている。",
    leftTitle: "【デメリット】", rightTitle: "【ヤシロの対応】",
    left: ["太陽光の売電価格（FIT単価）は年々下がってきている",
           "「売電で儲ける」という考え方で導入すると、失敗するリスクが高い",
           "売電を前提にした収支計画は、制度変更の影響を受けやすい"],
    right: ["売電ではなく「自家消費プラン」をご提案します",
            "蓄電池との組み合わせで、高い電気を買わない仕組みを構築",
            "売電収入に依存しない、堅実なシミュレーションを作成します"],
    footer: [{ t: "売電ではなく、自分で電気を作って貯めて使う" }, { t: "「自家消費」", hl: true }, { t: "の時代です！" }],
  });

  /* ===== P37 目次(4) ===== */
  agendaSlide(3);

  /* ===== P38 太陽光のトラブル ===== */
  {
    const s = contentSlide({
      chapter: CH4, chip: "太陽光のトラブル",
      lead: "残念ながら、太陽光業界における相談・トラブルは後を絶たない。",
      source: "※相談件数の一次出典（国民生活センター等）は登壇前に確認してください",
      footer: "だからこそ、業者選びが非常に重要です",
    });
    const items = [
      ["訪問販売による契約トラブル", "その場で契約を迫られる。냉静に比較する時間を与えられない。"],
      ["説明と違う発電量", "有利な条件だけを並べたシミュレーション。実際は届かない。"],
      ["施工不良・雨漏り", "屋根に穴を開ける工事。施工の質が数年後に出てくる。"],
      ["連絡が取れない・倒産", "売ったあとは音信不通。保証書があっても使えない。"],
    ];
    items[0][1] = "その場で契約を迫られる。冷静に比較する時間を与えられない。";
    const cw = 5.36, ch = 2.14, gx = 0.53, gy = 0.32, x0 = 0.42, y0 = 1.86;
    items.forEach((it, i) => {
      const x = x0 + (i % 2) * (cw + gx), y = y0 + Math.floor(i / 2) * (ch + gy);
      K.card(s, x, y, cw, ch, { fill: WHITE, line: RED, lw: 2 });
      s.addShape(pres.ShapeType.rect, { x, y, w: cw, h: 0.62, fill: { color: RED }, line: { color: RED, width: 0 } });
      s.addText(it[0], { x: x + 0.1, y, w: cw - 0.2, h: 0.62, fontFace: F, fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      s.addText(it[1], { x: x + 0.24, y: y + 0.78, w: cw - 0.48, h: ch - 0.96, fontFace: F, fontSize: 16, bold: true, color: BLACK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.25 });
    });
    s.addText("設備の良し悪しよりも、誰から買うかで結果が変わります。", {
      x: ML, y: 6.32, w: CW, h: 0.44, fontFace: F, fontSize: 19, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
  }

  /* ===== P39 ①地域密着（表空） ===== */
  {
    const s = contentSlide({
      chapter: CH4, chip: "①地域密着",
      lead: "長岡の地に根を張り、雪国での実証を重ねてきた会社かどうかを見てほしい。",
      footer: "県外の会社ではなく、地元で実績のある会社を選ぶことが大切です",
      note: "差し替え：創業／事業年数／対応エリア／太陽光施工実績／長岡市との協業（名称・期間・成果）／講演会の登壇実績。太陽光の取り組み開始は記入済み。",
    });
    const rows = [
      [th("項目"), th("内容")],
      [lbl("創業"), blank()],
      [lbl("事業年数"), blank()],
      [lbl("太陽光の取り組み開始"), val("2014年（柏崎・長岡で自社発電所）", { bold: true })],
      [lbl("対応エリア"), blank()],
      [lbl("太陽光施工実績"), blank()],
      [lbl("長岡市との協業（名称）"), blank()],
      [lbl("協業の期間"), blank()],
      [lbl("実験・実証の成果"), blank()],
      [lbl("講演会の登壇実績"), blank()],
    ];
    table(s, rows, ML, 1.82, 6.9, [2.5, 4.4], { rowH: 0.47, fontSize: 14 });
    K.card(s, 7.28, 1.82, 4.25, 4.72, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("地元の会社を選ぶ、3つの理由", {
      x: 7.28, y: 2.0, w: 4.25, h: 0.5, fontFace: F, fontSize: 19, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    [["売り逃げができない", "同じ地域で商売を続ける以上、逃げるという選択肢がない"],
     ["雪国での実証がある", "他県の一般論ではなく、長岡の気候で確かめた知見で話せる"],
     ["何かあればすぐ行ける", "距離が近いことは、それだけで保証の一部になる"]]
      .forEach((it, i) => {
        const y = 2.66 + i * 1.28;
        bulletRow(s, 7.46, y, 3.9, 0.5, i + 1, it[0], NAVY, 17);
        s.addText(it[1], { x: 8.14, y: y + 0.5, w: 3.22, h: 0.68, fontFace: F, fontSize: 13, bold: true, color: "404040", align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
      });
    note(39, "表", "地域密着・経営基盤（創業／事業年数／対応エリア／施工実績／長岡市との協業／期間／成果／講演実績）");
  }

  /* ===== P40 ②適正価格 ===== */
  twoColSlide({
    chapter: CH4, chip: "②適正価格",
    lead: "価格は安いほど良い、とは限らない。安すぎる見積りには理由がある。",
    leftTitle: "【安すぎる見積りの危うさ】", rightTitle: "【ヤシロの考え方】",
    left: ["工事費を削るために、施工の質が落ちる。数年後に雨漏りで出てくる",
           "訪問販売の会社は、売ったあとの不具合に対応しないことがある",
           "「今日決めれば特別価格」は、比較させないための手法"],
    right: ["見積りの内訳を開示し、何にいくらかかるかを説明します",
            "相場から極端に外れた提案はしません。理由を添えて価格を出します",
            "他社と比較していただいて構いません。その時間をお取りください"],
    footer: "安さの理由を説明できる会社かどうかで判断してください",
  });

  /* ===== P41 ③国産メーカー（表空） ===== */
  {
    const s = contentSlide({
      chapter: CH4, chip: "③国産メーカーの取り扱い",
      lead: "弊社が国産メーカーの長州産業を扱う理由は、不具合時の対応にある。",
      footer: "設備そのものだけでなく、困ったときに動いてもらえるかで選びます",
      note: "差し替え：長州産業の保証内容（機器保証／出力保証／自然災害補償／施工保証／対応窓口の期間・内容）。",
    });
    K.card(s, ML, 1.94, 5.02, 4.6, { fill: PALE, line: NAVY, lw: 2 });
    s.addText("長州産業を取り扱う理由", {
      x: ML, y: 2.12, w: 5.02, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
    [["国内生産のメーカーである", "生産から供給まで国内で完結している"],
     ["不具合時の対応がしっかりしている", "連絡がつき、話が前に進む。これがいちばん大きい"],
     ["部材の供給が長期にわたって続く", "10年後、20年後に部品がある安心感"]]
      .forEach((it, i) => {
        const y = 2.8 + i * 1.24;
        bulletRow(s, 0.4, y, 4.6, 0.5, i + 1, it[0], NAVY, 17);
        s.addText(it[1], { x: 1.08, y: y + 0.5, w: 3.92, h: 0.64, fontFace: F, fontSize: 13, bold: true, color: "404040", align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
      });
    s.addText("保証内容", {
      x: 5.5, y: 1.94, w: 6.03, h: 0.46, fontFace: F, fontSize: 18, bold: true, color: WHITE,
      fill: { color: NAVY }, align: "center", valign: "middle", margin: 0,
    });
    const rows = [
      [th("保証項目"), th("期間・内容")],
      [lbl("機器保証"), blank()],
      [lbl("出力保証"), blank()],
      [lbl("自然災害補償"), blank()],
      [lbl("施工保証（弊社独自）"), blank()],
      [lbl("対応窓口・受付時間"), blank()],
    ];
    table(s, rows, 5.5, 2.5, 6.03, [2.3, 3.73], { rowH: 0.7, fontSize: 15 });
    note(41, "表", "長州産業の保証内容（機器保証／出力保証／自然災害補償／施工保証／対応窓口）");
  }

  /* ===== P42 ④現地調査 ===== */
  {
    const s = contentSlide({
      chapter: CH4, chip: "④現地調査の丁寧さ",
      lead: "工務店として住宅を扱ってきたからこそ、屋根の下まで見て判断できる。",
      footer: "屋根に載せる前に、屋根そのものを見ます",
    });
    s.addText("現地調査でヤシロが確認する項目", {
      x: ML, y: 1.86, w: CW, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: WHITE,
      fill: { color: NAVY }, align: "center", valign: "middle", margin: 0,
    });
    const checks = [
      ["屋根材の種類と状態", "瓦・スレート・金属。材質ごとに固定方法が変わる"],
      ["下地・野地板の健全性", "表面ではなく、留め付ける下地が持つかどうか"],
      ["雨仕舞い", "穴を開ける以上、水の処理をどうするか"],
      ["積雪荷重", "パネルと雪の重さに、構造が耐えられるか"],
      ["落雪の方向と着地点", "隣家・道路・カーポートに落とさない配置か"],
      ["電気容量・分電盤", "蓄電池・V2Hを入れられる余地があるか"],
    ];
    const cw = 3.55, ch = 1.42, gx = 0.32, gy = 0.26, x0 = (SW - (cw * 3 + gx * 2)) / 2, y0 = 2.6;
    checks.forEach((c, i) => {
      const x = x0 + (i % 3) * (cw + gx), y = y0 + Math.floor(i / 3) * (ch + gy);
      K.card(s, x, y, cw, ch, { fill: WHITE, line: NAVY, lw: 1.5 });
      s.addShape(pres.ShapeType.ellipse, { x: x + 0.18, y: y + 0.22, w: 0.42, h: 0.42, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
      s.addText(String(i + 1), { x: x + 0.18, y: y + 0.22, w: 0.42, h: 0.42, fontFace: F, fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      s.addText(c[0], { x: x + 0.7, y: y + 0.16, w: cw - 0.86, h: 0.5, fontFace: F, fontSize: 16, bold: true, color: NAVY, align: "left", valign: "middle", margin: 0 });
      s.addText(c[1], { x: x + 0.22, y: y + 0.72, w: cw - 0.44, h: 0.6, fontFace: F, fontSize: 13, bold: true, color: "404040", align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
    });
    s.addText("屋根の状態によっては「今はやめておきましょう」とお伝えすることもあります。", {
      x: ML, y: 5.98, w: CW, h: 0.5, fontFace: F, fontSize: 18, bold: true, color: RED,
      align: "center", valign: "middle", margin: 0,
    });
  }

  /* ===== P43 まとめ ===== */
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
      K.card(s, 0.6, y, SW - 1.2, h, { fill: i % 2 === 0 ? PALE : WHITE, line: NAVY, lw: 1.5 });
      s.addShape(pres.ShapeType.ellipse, { x: 0.86, y: y + (h - 0.58) / 2, w: 0.58, h: 0.58, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
      s.addText(String(i + 1), { x: 0.86, y: y + (h - 0.58) / 2, w: 0.58, h: 0.58, fontFace: F, fontSize: 22, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      s.addText(t, { x: 1.62, y, w: SW - 2.4, h, fontFace: F, fontSize: 20, bold: true, color: BLACK, align: "left", valign: "middle", margin: 0 });
    });
  }

  /* ===== P44 個別相談のご案内 ===== */
  {
    const s = contentSlide({
      chapter: "まとめ", chip: "個別相談のご案内",
      lead: "本講座は、まずは正しい知識を身につけていただくことを目的としている。",
      footer: "無理な売り込みは一切いたしません。お気軽にお声がけください",
    });
    K.card(s, 0.42, 1.86, 5.36, 4.68, { fill: WHITE, line: NAVY, lw: 2 });
    s.addShape(pres.ShapeType.rect, { x: 0.42, y: 1.86, w: 5.36, h: 0.62, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    s.addText("参加をご検討の方へ", { x: 0.42, y: 1.86, w: 5.36, h: 0.62, fontFace: F, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    s.addText(
      "当講座は、まずは正しい知識を身につけていただくことを目的としております。\n\n" +
      "無理な売り込みは一切いたしませんので、お気軽にご参加ください。\n\n" +
      "お見積り・現地調査をご希望の方は、恐れ入りますが講座終了後にスタッフへお声がけください。",
      { x: 0.72, y: 2.72, w: 4.76, h: 3.5, fontFace: F, fontSize: 17, bold: true, color: BLACK,
        align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.35 }
    );
    K.card(s, 5.92, 1.86, 5.36, 4.68, { fill: PALE, line: NAVY, lw: 2 });
    s.addShape(pres.ShapeType.rect, { x: 5.92, y: 1.86, w: 5.36, h: 0.62, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    s.addText("参加者特典", { x: 5.92, y: 1.86, w: 5.36, h: 0.62, fontFace: F, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    s.addText("『太陽光・蓄電池の\n　すべてが分かる小冊子』", {
      x: 6.22, y: 2.78, w: 4.76, h: 1.2, fontFace: F, fontSize: 24, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2,
    });
    s.addShape(pres.ShapeType.line, { x: 6.42, y: 4.06, w: 4.36, h: 0, line: { color: NAVY, width: 1 } });
    s.addText("太陽光の仕組み・よくある失敗から\nお得な使い方まで、すべて掲載しています。", {
      x: 6.22, y: 4.22, w: 4.76, h: 0.96, fontFace: F, fontSize: 15, bold: true, color: BLACK,
      align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.25,
    });
    K.card(s, 6.42, 5.3, 4.36, 0.96, { fill: WHITE, line: NAVY, lw: 1.5 });
    s.addText("株式会社ヤシロ　TEL 0120-17-5521\n営業時間 8:00〜17:00／定休日 日曜・祝日", {
      x: 6.42, y: 5.3, w: 4.36, h: 0.96, fontFace: F, fontSize: 14, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.25,
    });
  }

  /* ===== P45 御礼 ===== */
  {
    const s = pres.addSlide();
    s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: SW, h: SH, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    s.addShape(pres.ShapeType.rect, { x: 0, y: 1.9, w: SW, h: 4.4, fill: { color: "12386E" }, line: { color: "12386E", width: 0 } });
    s.addText("本日はお忙しいところ\nご来場いただき\n誠にありがとうございました！", {
      x: 0.6, y: 2.14, w: SW - 1.2, h: 2.36, fontFace: F, fontSize: 34, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.25,
    });
    s.addText("今後とも株式会社ヤシロを\n何卒よろしくお願いいたします。", {
      x: 0.6, y: 4.5, w: SW - 1.2, h: 1.6, fontFace: F, fontSize: 28, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.25,
    });
    s.addShape(pres.ShapeType.rect, { x: 0, y: 7.12, w: SW, h: SH - 7.12, fill: { color: "001845" }, line: { color: "001845", width: 0 } });
    s.addText([
      { text: "詳しくは、講座終了後に個別でご相談ください！", options: { fontFace: F, fontSize: 26, bold: true, color: YEL } },
    ], { x: 0.3, y: 7.12, w: SW - 0.6, h: SH - 7.12, align: "center", valign: "middle", margin: 0 });
    s.addNotes("差し替え：背景写真（自社施工事例など）を全面に敷き、現状のネイビー地と入れ替える。");
    note(45, "背景", "御礼ページ背景写真（現状はネイビー単色）");
  }
};
