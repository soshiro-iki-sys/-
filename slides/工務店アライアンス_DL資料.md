---
marp: true
theme: funai
size: 16:9
paginate: false
---

<style>
/* ===== 本資料の専用パーツ（テーマ本体は汚さない） ===== */

/* 3点カード */
.cards { display:flex; gap:20px; width:1236px; }
.cards > div { flex:1; border:2px solid #002060; padding:16px 18px 18px; }
.cards .no { display:inline-block; background:#002060; color:#fff; font-size:19px;
             padding:2px 14px; margin-bottom:10px; }
.cards .ttl { font-size:25px; color:#002060; line-height:1.35; margin-bottom:9px; }
.cards .txt { font-size:18px; font-weight:400; line-height:1.6; }

/* 分割カード（2列。.cols の中では1列に積む） */
.quad { display:flex; flex-wrap:wrap; gap:14px; width:1236px; }
.quad > div { flex:0 0 calc(50% - 9px); border-left:10px solid #002060;
              background:#F2F5FA; padding:11px 17px; }
.cols .quad { width:100%; }
.cols .quad > div { flex:0 0 100%; }
.quad .cat { font-size:17px; color:#5B6B85; }
.quad .ttl { font-size:24px; color:#002060; margin:2px 0 5px; }
.quad .txt { font-size:17.5px; font-weight:400; line-height:1.5; }
.quad .risk { font-size:16.5px; font-weight:400; color:#C0392B; margin-top:5px; }

/* 座組フロー図 */
.flow .cust { width:300px; margin:0 auto; background:#002060; color:#fff;
              font-size:21px; text-align:center; padding:8px 0; }
.flow .up { display:flex; justify-content:space-between; text-align:center;
            font-size:16px; font-weight:400; margin:11px 0; }
.flow .up > div { flex:1; line-height:1.45; }
.flow .up span { font-size:14px; color:#5B6B85; }
.flow .row { display:flex; gap:16px; }
.flow .node { flex:1; background:#002060; color:#fff; font-size:21px;
              text-align:center; padding:12px 0; line-height:1.35; }
.flow .node em { display:block; font-style:normal; font-size:14px; color:#C9D6EA; }
.flow .mid { margin-top:12px; font-size:16px; font-weight:400; }
.flow .mid .to   { background:#FDECEC; border-left:6px solid #FF0000; padding:8px 11px;
                   margin-bottom:7px; }
.flow .mid .from { background:#EDF1F8; border-left:6px solid #1F497D; padding:8px 11px; }

/* 棒グラフ */
.chart { display:flex; align-items:flex-end; gap:26px; height:270px;
         width:1000px; margin:10px 0 0 40px; border-bottom:2px solid #002060; }
.chart > div { flex:1; display:flex; flex-direction:column; justify-content:flex-end;
               align-items:center; height:100%; }
.chart .v { font-size:19px; color:#002060; margin-bottom:5px; }
.chart i { display:block; width:100%; background:#002060; }
.chart .last i { background:#FF0000; }
.chart .x { font-size:18px; color:#333; margin-top:7px; }

/* 方程式 */
.eq { width:1236px; text-align:center; border:3px solid #002060;
      padding:13px 0; font-size:30px; color:#002060; }
.eq b { color:#FF0000; }
.calc { display:flex; align-items:center; justify-content:center; gap:22px;
        width:1236px; margin-top:14px; }
.calc .box { border:2px solid #98A7BE; padding:11px 20px; text-align:center; font-size:19px;
             font-weight:400; line-height:1.5; }
.calc .box .hd { font-size:21px; font-weight:700; color:#002060; margin-bottom:5px; }
.calc .box .rs { font-size:32px; font-weight:700; color:#002060; margin-top:5px; }
.calc .box.after { border-color:#FF0000; }
.calc .box.after .rs { color:#FF0000; }
.calc .arw { font-size:24px; color:#002060; text-align:center; line-height:1.3; }

/* 表を詰める */
table.tight { font-size:18.5px; }
table.tight th, table.tight td { padding:6px 12px; }
table.tight td.q { color:#002060; }
table.dense { font-size:16.5px; }
table.dense th, table.dense td { padding:5px 10px; line-height:1.5; }
table.dense td.q { color:#002060; font-size:17.5px; }

/* チェックリスト */
.check { display:flex; gap:26px; width:1236px; font-weight:400; }
.check > div { flex:1; }
.check .grp { font-size:19px; font-weight:700; color:#fff; background:#002060;
              padding:3px 14px; margin:8px 0 4px; }
.check p { font-size:18px; line-height:1.5; margin:3px 0; }

/* 次の一歩 */
.step { display:flex; gap:22px; width:1236px; margin-top:12px; }
.step > div { flex:1; border:2px solid #002060; padding:14px 18px; }
.step .hd { background:#002060; color:#fff; font-size:20px; padding:3px 14px;
            display:inline-block; margin-bottom:9px; }
.step .txt { font-size:18px; font-weight:400; line-height:1.6; }
</style>

<!-- ============ P1 表紙 ============ -->
<!-- _class: title -->

<div class="cover">
<img src="../assets/表紙_サンプル.jpg">
<div class="panel"></div>
<div class="g-left"></div><div class="g-top"></div>
<div class="g-bottom"></div><div class="g-mask"></div><div class="g-right"></div>
<div class="logo"></div>
<div class="kicker">住宅用太陽光＆蓄電池</div>
</div>

# 工務店アライアンス<br>成功のポイントとは？

## 太陽光・蓄電池の受注を伸ばす5つの変数<br>受注件数の方程式と実践KPI

---

<!-- ============ P2 導入 ============ -->

# はじめに

## この資料でわかること

### 太陽光・蓄電池の販売店が、工務店との業務提携で受注を伸ばすための実務書です

<div class="cards">
<div>
<div class="no">01</div>
<div class="ttl">業界環境の<br>変化</div>
<div class="txt">新設住宅着工戸数は減少が続く一方、国のネットゼロ政策と2027年のZEH定義見直しにより、太陽光・蓄電池は「あれば良い設備」から「満たすべき要件」へ変わります。</div>
</div>
<div>
<div class="no">02</div>
<div class="ttl">工務店が<br>動けていない理由</div>
<div class="txt">必要性を理解していても対応できない工務店が数多く存在します。その理由を人材・知識・収益・運用の4つに分解し、提携の余地がどこにあるかを示します。</div>
</div>
<div>
<div class="no">03</div>
<div class="ttl">受注を分解する<br>方程式と打ち手</div>
<div class="txt">受注件数を5つの変数の掛け算として捉え、変数ごとのKPIと具体施策、そして自社の弱点を特定する診断チェックリストをご用意しました。</div>
</div>
</div>

> 提携の「やり方」ではなく、<b>どの変数を改善すべきか</b>が分かる資料です

<div class="pageno">2</div>

---

<!-- ============ P3 目次 ============ -->
<!-- _class: toc -->

# 目次

1. 本レポートで解説するアライアンスとは？
2. なぜ今なのか｜市場と制度の変化
3. 工務店が動けない4つの理由
4. 受注件数の方程式と打ち手

<div class="pageno">3</div>

---

<!-- ============ P4 定義 ============ -->

# 1. 本レポートで解説するアライアンスとは？

## アライアンスの定義

### 顧客名簿という資産を軸に、両社が中長期で収益を分け合う継続的な業務提携

工務店が持つ「顧客との接点と信頼」と、販売店が持つ「商品・提案・施工・アフターの実行力」を組み合わせ、双方の顧客に価値を提供する関係です。一度きりの案件紹介でも、元請けからの工事受注でもありません。

<table class="tight">
<tr><th style="width:200px"></th><th>下請け</th><th>単発の紹介</th><th style="background:#FF0000">アライアンス</th></tr>
<tr><td class="q">関係性</td><td>発注者と受注者の上下関係</td><td>都度のスポット取引</td><td>対等な継続的パートナー</td></tr>
<tr><td class="q">顧客との接点</td><td>元請け経由。自社は表に出ない</td><td>紹介された案件のみ</td><td>顧客名簿単位で継続的に接触</td></tr>
<tr><td class="q">価格の決定権</td><td>元請けが決定</td><td>案件ごとに交渉</td><td>提携時に条件を取り決め</td></tr>
<tr><td class="q">収益の形</td><td>工事単価のみ</td><td>単発の紹介フィー</td><td>継続的な受注＋紹介手数料</td></tr>
<tr><td class="q">再現性</td><td>元請けの受注量に依存</td><td>低い（属人的）</td><td>高い（仕組みで積み上がる）</td></tr>
</table>

> 目指すのは案件の獲得ではなく、<b>案件が生まれ続ける関係</b>をつくること

<div class="pageno">4</div>

---

<!-- ============ P5 座組 ============ -->

# 1. 本レポートで解説するアライアンスとは？

## 座組と役割分担

### 工務店は「紹介だけ」から「自社で営業し施工だけ依頼」まで関与度を選べる

<div class="cols">
<div>

<div class="flow">
<div class="cust">オーナー顧客</div>
<div class="up">
<div>② 訪問・関心の確認<br><span>（自社で営業する場合は<br>そのまま営業）</span></div>
<div>④ 営業代行＋施工<br><span>④' 施工のみ</span></div>
</div>
<div class="row">
<div class="node">工務店<em>顧客接点・地域での信頼</em></div>
<div class="node">販売店<em>提案・施工・アフター</em></div>
</div>
<div class="mid">
<div class="to">工務店 ▶ 販売店　③ 案件の紹介／③' 成約顧客リストの提供</div>
<div class="from">販売店 ▶ 工務店　① 商材の卸・営業ツール／⑤ バックマージン</div>
</div>
</div>

</div>
<div>

<table class="dense">
<tr><th style="width:100px"></th><th>工務店</th><th>販売店（自社）</th></tr>
<tr><td class="q">提供する<br>もの</td><td>顧客接点<br>地域での信頼<br>顧客名簿</td><td>商品調達／提案・見積<br>施工・アフター<br>補助金申請</td></tr>
<tr><td class="q">収益</td><td>紹介手数料<br><small>販売金額の5〜7％が目安</small></td><td>工事売上<br>継続的な提案機会</td></tr>
<tr><td class="q">リスク</td><td>顧客満足度の低下<br><small>施工品質は販売店に依存</small></td><td>提携先の紹介停滞<br>個人情報の管理責任</td></tr>
</table>

</div>
</div>

> 工務店の不安は「自社の顧客を任せること」。<b>そこを消す設計</b>が成否を決める

<div class="pageno">5</div>

---

<!-- ============ P6 市場環境 ============ -->

# 2. なぜ今なのか｜市場と制度の変化

## 新築着工棟数の推移

### 新設住宅着工戸数は減少が続き、1棟あたりの単価か既存顧客からの追加受注が問われる

<div class="chart">
<div><span class="v">90.5</span><i style="height:82%"></i><span class="x">2019年</span></div>
<div><span class="v">81.5</span><i style="height:74%"></i><span class="x">2020年</span></div>
<div><span class="v">85.6</span><i style="height:78%"></i><span class="x">2021年</span></div>
<div><span class="v">86.0</span><i style="height:78%"></i><span class="x">2022年</span></div>
<div><span class="v">82.0</span><i style="height:74%"></i><span class="x">2023年</span></div>
<div class="last"><span class="v">79.2</span><i style="height:72%"></i><span class="x">2024年</span></div>
</div>

<small class="src">新設住宅着工戸数（万戸・暦年）／出典：国土交通省「建築着工統計調査報告」　※数値は公表資料でご確認ください</small>

> 人口減少に伴う新築市場の縮小に対し、<b>従来の新築事業だけに依存した経営が岐路に</b>

<div class="pageno">6</div>

---

<!-- ============ P7 制度環境 ============ -->

# 2. なぜ今なのか｜市場と制度の変化

## ネットゼロと2027年ZEH定義の見直し

### 2027年のZEH定義見直しで、蓄電池が設備要件に加わる

<div class="cols">
<div>

<table class="dense">
<tr><th style="width:185px">ZEH（戸建）</th><th>現行定義</th><th style="background:#FF0000">2027年〜の新定義</th></tr>
<tr><td class="q">断熱性能</td><td>断熱等級5</td><td>断熱等級6</td></tr>
<tr><td class="q">一次エネ消費量削減率<br><small>（省エネのみ）</small></td><td>20％</td><td>35％</td></tr>
<tr><td class="q">設備要件</td><td>—</td><td>① 高度エネマネ<br>② 蓄電池<small>（PVありの場合）</small></td></tr>
<tr><td class="q">再エネ要件</td><td>『ZEH』100％</td><td>新ZEH＋：115％</td></tr>
</table>

</div>
<div>

<div class="quad">
<div><div class="ttl">補助金が取れない</div><div class="txt">GX志向型住宅は160万円／戸。要件を満たせない工務店は、施主に提示できる補助額で不利になります。</div></div>
<div><div class="ttl">比較検討で外れる</div><div class="txt">同一エリアの競合が創蓄を標準提案すれば、性能表で見劣りし検討段階で候補から落ちます。</div></div>
<div><div class="ttl">既存顧客が流出する</div><div class="txt">対応できない間に、オーナー顧客はリフォーム会社や訪販業者に囲い込まれます。</div></div>
</div>

</div>
</div>

<small class="src">出典：環境省「日本の新たな温室効果ガス削減目標（NDC）とGX推進政策について」／経済産業省 省エネルギー小委員会 資料　※2026年8月時点の公表情報</small>

> 「まだ先の話」では手遅れに。<b>2027年の定義変更は目前</b>

<div class="pageno">7</div>

---

<!-- ============ P8 章の締め ============ -->
<!-- _class: message -->

# 2. なぜ今なのか｜市場と制度の変化

## 第2章まとめ

必要性は感じている。
**しかし、動けていない。**

<div class="pageno">8</div>

---

<!-- ============ P9 課題の全体像 ============ -->

# 3. 工務店が動けない4つの理由

## 課題の全体像

### 動けない理由は意欲ではなく、人材・知識・収益・運用の4つの構造的な壁にある

<div class="quad">
<div>
<div class="cat">課題① 人　材</div>
<div class="ttl">営業リソースが足りない</div>
<div class="txt">新築の受注活動で手一杯。既存顧客に追加提案する担当を置けない。採用も難しく、有効求人倍率は2023年で1.31倍と高止まりしています。</div>
</div>
<div>
<div class="cat">課題② 知　識</div>
<div class="ttl">提案の仕方が分からない</div>
<div class="txt">容量設計、経済効果シミュレーション、見積の組み立てに知見がなく、聞かれても答えられないため提案自体を避けてしまう。</div>
</div>
<div>
<div class="cat">課題③ 収　益</div>
<div class="ttl">粗利を確保できない</div>
<div class="txt">仕入ルートがなく単価が高い。相見積で価格勝負になり、手間の割に利益が残らない。だから優先度が下がる。</div>
</div>
<div>
<div class="cat">課題④ 運　用</div>
<div class="ttl">補助金・法対応の負荷</div>
<div class="txt">補助金は年度ごとに要件が変わり、申請期限も短い。系統連系の申請も含め、事務負担が読めない。</div>
</div>
</div>

<small class="src">有効求人倍率の出典：厚生労働省「職業安定業務統計」</small>

> 動けない理由は「やる気」ではなく<b>構造</b>。だから外部との組み合わせで解ける

<div class="pageno">9</div>

---

<!-- ============ P10 課題の深掘り ============ -->

# 3. 工務店が動けない4つの理由

## 課題の深掘り

<div class="quad">
<div>
<div class="cat">課題① 営業リソース不足</div>
<div class="ttl">接点はあるのに提案されない</div>
<div class="txt">・オーナー顧客は増えているが、そのメリットを生かしきれていない<br>・点検やアフター訪問の機会があっても提案に踏み込めない<br>・担当者が現場管理と兼務で追客の時間が取れない</div>
<div class="risk">▶ 放置すると、顧客は他社の情報だけで判断してしまう</div>
</div>
<div>
<div class="cat">課題② 見積と提案の難しさ</div>
<div class="ttl">自信がないから切り出せない</div>
<div class="txt">・新規の売り方は分かるが、リフォーム商材の提案方法が分からない<br>・効果を語れないため会話が続かない<br>・切り出すタイミングが分からず点検で終わる</div>
<div class="risk">▶ 放置すると、受け身になり案件が生まれない</div>
</div>
<div>
<div class="cat">課題③ 粗利確保</div>
<div class="ttl">手間の割に利益が残らない</div>
<div class="txt">・仕入ルートがなく、単発発注のため仕入単価が高い<br>・訪販業者と相見積になり価格で負ける<br>・工事管理の手間に対して利益が見合わない</div>
<div class="risk">▶ 放置すると、社内で優先順位が下がり定着しない</div>
</div>
<div>
<div class="cat">課題④ 補助金・法対応</div>
<div class="ttl">事務負担が読めない</div>
<div class="txt">・補助金は年度ごとに要件・予算・締切が変わる<br>・予算到達で早期終了することがあり社内で追えない<br>・系統連系や各種申請の手順が確立していない</div>
<div class="risk">▶ 放置すると、申請漏れで施主の期待に応えられない</div>
</div>
</div>

> 4つの課題はすべて<b>「自社だけで抱えるから解けない」</b>という共通点を持つ

<div class="pageno">10</div>

---

<!-- ============ P11 だからこそ ============ -->

# 3. 工務店が動けない4つの理由

## だからこそのアライアンス

### 4つの課題は、販売店が既に持っている機能をそのまま当てれば解消できる

<table class="tight">
<tr><th style="width:140px"></th><th style="width:400px">工務店の課題</th><th>アライアンスによる解決</th></tr>
<tr>
<td class="q">① 人　材</td>
<td>提案する人がいない。採用もできない</td>
<td><b>営業代行</b>：販売店の営業が顧客訪問から提案・クロージングまで実施。工務店は紹介するだけでよい</td>
</tr>
<tr>
<td class="q">② 知　識</td>
<td>容量設計・経済効果・見積が分からない</td>
<td><b>ツールと勉強会</b>：アプローチブック・シミュレーション資料の提供と、営業マン向けの個社別勉強会を定期開催</td>
</tr>
<tr>
<td class="q">③ 収　益</td>
<td>仕入が高く粗利が残らない</td>
<td><b>商材卸＋紹介手数料</b>：まとめ仕入れによる原価低減と、販売金額の5〜7％を目安とした紹介フィー</td>
</tr>
<tr>
<td class="q">④ 運　用</td>
<td>補助金・申請の事務負担が読めない</td>
<td><b>申請代行と施工</b>：補助金申請・系統連系・施工・アフターまで販売店が一括対応</td>
</tr>
</table>

<blockquote class="tall">
<p>工務店は「手間なく満足度と売上が上がる」、販売店は「広告費ゼロで良質な案件が入る」<br>
<b>実績や規模は関係ない。営業マン1名の販売店でも提携は成立している</b></p>
</blockquote>

<div class="pageno">11</div>

---

<!-- ============ P12 方程式 ============ -->

# 4. 受注件数の方程式と打ち手

## 受注件数の方程式

### アライアンスの受注件数は、5つの変数の掛け算で決まる

<div class="eq">
受注件数 ＝ <b>①提携社数</b> × <b>②顧客名簿数</b> × <b>③紹介率</b> × <b>④商談率</b> × <b>⑤成約率</b>
</div>

<div class="calc">
<div class="box">
<div class="hd">改善前</div>
提携10社 × 名簿100件<br>× 紹介率5％ × 商談率60％<br>× 成約率30％
<div class="rs">年9件</div>
</div>
<div class="arw">各変数を<br><b>1.5倍</b>に<br>➡</div>
<div class="box after">
<div class="hd">改善後</div>
提携15社 × 名簿150件<br>× 紹介率7.5％ × 商談率90％<br>× 成約率45％
<div class="rs">年68件</div>
</div>
<div class="arw"><b>約7.6倍</b><br>（1.5⁵）</div>
</div>

<blockquote class="tall">
<p><b>1つでもゼロなら成果はゼロ</b>。提携社数を増やしても紹介率0％なら受注は生まれない<br>
各変数1.5倍で受注は約7.6倍。<b>一点突破より全変数の底上げ</b>が効く</p>
</blockquote>

<div class="pageno">12</div>

---

<!-- ============ P13 打ち手一覧 ============ -->

# 4. 受注件数の方程式と打ち手

## 5変数の打ち手一覧

<table class="dense">
<tr><th style="width:125px">変数</th><th style="width:245px">見るべきKPI</th><th>具体施策</th></tr>
<tr>
<td class="q">① 提携社数</td>
<td>アタックリスト数<br>面談率／提携率／継続率<br><small>目安：訪問120社→締結3社（2.5％）</small></td>
<td>地域密着の工務店を第一ターゲットに、SEO・MEO／SUUMO／工務店団体／商工会議所・銀行からリストアップ。訪問のゴールを「提携」ではなく<b>「勉強会への招待」</b>に下げる。3ヵ月に1回、メーカーをゲストに招いた勉強会を定例開催</td>
</tr>
<tr>
<td class="q">② 顧客名簿数</td>
<td>提携先のオーナー数<br>名簿の共有可否（0 or 1）<br>年間棟数</td>
<td><b>提案の幅×案件の透明化</b>。1年目PV・2年目蓄電池・3年目塗装・4年目床下・5年目内装と5年サイクルの提案設計を示す。顧客管理ツールで「誰が・いつ・何を提案したか」を提携先に開示し、預ける不安を消す</td>
</tr>
<tr>
<td class="q">③ 紹介率</td>
<td>提携先別の紹介率<br><small>（年間棟数に対する紹介件数）</small></td>
<td>紹介が止まる<b>阻害要因を特定する</b>。社長のトップダウン不足／営業の理解不足／提案資料がない／上下の熱量差。対策は個社別勉強会、営業マンへの直接説明、商談への同席、個別チラシの作成</td>
</tr>
<tr>
<td class="q">④ 商談率</td>
<td>紹介アポ→商談の到達率<br><small>目標：100％</small></td>
<td>紹介案件は見積提出のみで済むことが多く<b>100％を目指せる</b>。案件管理を紙からデジタルへ移行（スプレッドシート／kintone等）し、管理表を全社で共有して抜け漏れを防ぐ</td>
</tr>
<tr>
<td class="q">⑤ 成約率</td>
<td>商談→成約の到達率<br>営業担当者別のばらつき</td>
<td>エース営業の<b>「心構え・アポトーク・営業ツール・クロージングトーク」を言語化</b>しマニュアル化。営業同行・1on1・ロープレ・営業動画で組織に展開する</td>
</tr>
</table>

> 「背中を見て覚えろ」からの脱却。<b>売れる営業の育成は仕組み</b>

<div class="pageno">13</div>

---

<!-- ============ P14 診断チェックリスト ============ -->

# 4. 受注件数の方程式と打ち手

## 自社診断チェックリスト

### 「いいえ」が多い変数が、いま最優先で手を打つべき箇所

<div class="check">
<div>
<div class="grp">① 提携社数</div>
<p>□ 1　アタックリストが50社以上ある</p>
<p>□ 2　ターゲットを地域密着の工務店に絞れている</p>
<p>□ 3　訪問後に2回目以降の接触をしている</p>
<p>□ 4　勉強会など提携以外の出口を用意している</p>
<div class="grp">② 顧客名簿数</div>
<p>□ 5　提携先ごとのオーナー数を把握している</p>
<p>□ 6　太陽光以外の商材も提案できる</p>
<p>□ 7　5年先までの提案サイクルを設計している</p>
<p>□ 8　個人情報の管理体制を説明できる</p>
</div>
<div>
<div class="grp">③ 紹介率</div>
<p>□ 9　提携先ごとに紹介率を数値で把握している</p>
<p>□ 10　紹介が止まる理由を相手に確認している</p>
<p>□ 11　提携先の営業マンに直接説明する機会がある</p>
<div class="grp">④ 商談率</div>
<p>□ 12　案件管理をデジタルで一元化している</p>
<p>□ 13　追客・見積の抜け漏れが起きない仕組みがある</p>
<div class="grp">⑤ 成約率</div>
<p>□ 14　営業トークをマニュアル化している</p>
<p>□ 15　ロープレ・同行の機会を定例化している</p>
</div>
</div>

> <b>13個以上</b>「はい」なら仕組みは完成間近。<b>7個以下</b>なら伸びしろは大きい

<div class="pageno">14</div>

---

<!-- ============ P15 まとめ ============ -->

# まとめ

## 次の一歩

<div class="eq">
受注件数 ＝ <b>①提携社数</b> × <b>②顧客名簿数</b> × <b>③紹介率</b> × <b>④商談率</b> × <b>⑤成約率</b>
</div>

### 自社の「売りたい」ではなく、工務店の「困っている」から始める

まず方程式で自社の弱い変数を特定し、勉強会という形で価値提供から入る。そして仕組み化で属人性をなくす——この順序が成功の型です。

<div class="step">
<div>
<div class="hd">STEP 1</div>
<div class="txt"><b>自社の数値を方程式に入れる</b><br>P.14のチェックリストで弱い変数を特定してください。</div>
</div>
<div>
<div class="hd">STEP 2</div>
<div class="txt"><b>勉強会を1本企画する</b><br>メーカーをゲストに、地域の工務店を招く形から始めるのが取り組みやすい方法です。</div>
</div>
<div>
<div class="hd">STEP 3</div>
<div class="txt"><b>紹介率を計測する</b><br>提携先ごとに年間棟数と紹介件数を並べ、阻害要因を特定します。</div>
</div>
</div>

> 第二の集客軸の策定。そして<b>さらなる業績拡大へ</b>

<div class="pageno">15</div>

---

<!-- ============ P16 お問い合わせ ============ -->
<!-- _class: contact -->

<div class="contact-wrap">

<div class="c-lead">
経営に課題を感じている、コンサルティング内容や依頼方法が知りたい、専門家から事業に関するアドバイスが欲しい<br>
船井総合研究所では皆様の相談を <b>無料</b> で承ります。
</div>
<div class="c-note">※お電話、WEBをお選びいただけます</div>

<div class="c-row">
<div class="c-col">
<div class="c-hd">お電話でのお問い合わせ・ご相談</div>
<div class="c-sub">下記のフリーダイヤルからご連絡ください。</div>
<div class="c-tel">0120-958-270</div>
<div class="c-time">【受付時間】平日 9時45分〜17時30分（土日祝、年末年始を除く）</div>
</div>
<div class="c-col">
<div class="c-hd left">WEBでのお問い合わせ・ご相談</div>
<div class="c-sub left">下記ボタン、URLをクリックしてサイトへアクセスください。</div>
<div class="c-btn">無料　経営相談に申し込む</div>
<div class="c-url">https://www.funaisoken.co.jp/form/consulting</div>
</div>
</div>

<div class="c-search"><span class="c-box">船井総研　経営相談</span>　検索や船井総合研究所ウェブサイトURLからも上記ページへアクセスいただけます。</div>

<div class="c-foot">
<div class="c-company">株式会社船井総合研究所　会社概要<br><span>https://www.funaisoken.co.jp/info/company</span></div>
<div class="c-disc">
◆ 本資料に掲載のセミナーはすでに終了している場合があります。内容にご不明な点がある際は、別途お問い合わせください。◆ 本資料に掲載の情報は作成時点のものであり、将来的な内容を保証するものではありません。◆ 本資料に掲載の情報を目的とし、弊社承諾なく転載・改変することを禁じております。◆ 本資料に掲載の情報を利用したことによって発生する損害について、弊社は責任を負いかねます。◆ 本資料の内容は将来予告なく変更、または廃止されることがあります。<br>
◆ 本資料には生成AIにより生成した情報やPIXTAより提供された画像が含まれている場合があります。
</div>
<div class="c-logo"></div>
</div>

</div>
