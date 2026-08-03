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
.cards > div { flex:1; border:1px solid #000; padding:16px 18px 18px; }
.cards .no { display:inline-block; background:#404040; color:#fff; font-size:19px;
             padding:2px 14px; margin-bottom:10px; }
.cards .ttl { font-size:25px; color:#002060; line-height:1.35; margin-bottom:9px; }
.cards .txt { font-size:18px; line-height:1.6; }

/* 分割カード（2列。.cols の中では1列に積む） */
.quad { display:flex; flex-wrap:wrap; gap:14px; width:1236px; }
.quad > div { flex:0 0 calc(50% - 9px); border:1px solid #000; border-left:10px solid #002060;
              padding:11px 17px; }
.cols .quad { width:100%; }
.cols .quad > div { flex:0 0 100%; }
.quad .cat { font-size:17px; color:#404040; }
.quad .ttl { font-size:24px; color:#002060; margin:2px 0 5px; }
.quad .txt { font-size:17.5px; line-height:1.5; }
.quad .risk { font-size:16.5px; color:#FF0000; margin-top:5px; }

/* 座組フロー図 */
.flow .cust { width:300px; margin:0 auto; background:#002060; color:#fff;
              font-size:21px; text-align:center; padding:8px 0; }
.flow .up { display:flex; justify-content:space-between; text-align:center;
            font-size:16px; margin:11px 0; }
.flow .up > div { flex:1; line-height:1.45; }
.flow .up span { font-size:14px; color:#404040; }
.flow .row { display:flex; gap:16px; }
.flow .node { flex:1; background:#002060; color:#fff; font-size:21px;
              text-align:center; padding:12px 0; line-height:1.35; }
.flow .node em { display:block; font-style:normal; font-size:14px; color:#C9D6EA; }
.flow .mid { margin-top:12px; font-size:16px; }
.flow .mid .to   { border:1px solid #000; border-left:6px solid #FF0000; padding:8px 11px;
                   margin-bottom:7px; }
.flow .mid .from { border:1px solid #000; border-left:6px solid #002060; padding:8px 11px; }

/* 方程式 */
.eq { width:1236px; text-align:center; border:1px solid #000;
      padding:13px 0; font-size:30px; color:#002060; }
.eq b { color:#FF0000; }
.calc { display:flex; align-items:center; justify-content:center; gap:22px;
        width:1236px; margin-top:14px; }
.calc .box { border:1px solid #000; padding:11px 20px; text-align:center; font-size:19px;
             line-height:1.5; }
.calc .box .hd { font-size:21px; color:#fff; background:#404040; margin:-11px -20px 8px;
                 padding:4px 0; }
.calc .box .rs { font-size:32px; color:#002060; margin-top:5px; }
.calc .box.after { border:2px solid #FF0000; }
.calc .box.after .hd { background:#FF0000; }
.calc .box.after .rs { color:#FF0000; }
.calc .arw { font-size:24px; color:#002060; text-align:center; line-height:1.3; }

/* 表を詰める */
table.dense { font-size:17px; }
table tr.hl td { background:#FFF9D6; }
table.dense th, table.dense td { padding:6px 11px; line-height:1.5; }

/* 勉強会のKPIファネル */
.funnel { display:flex; align-items:center; justify-content:center; gap:10px;
          width:1236px; margin-top:6px; }
.funnel .st { flex:0 0 210px; height:112px; border:1px solid #000; text-align:center;
              font-size:22px; display:flex; flex-direction:column; justify-content:center; }
.funnel .st b { display:block; font-size:30px; color:#002060; }
.funnel .rt { flex:0 0 96px; text-align:center; font-size:22px; color:#FF0000; }

/* チェックリスト */
.check { display:flex; gap:26px; width:1236px; }
.check > div { flex:1; }
.check .grp { font-size:19px; color:#fff; background:#404040; padding:3px 14px; margin:8px 0 4px; }
.check p { font-size:18px; line-height:1.5; margin:3px 0; }

/* 次の一歩 */
.step { display:flex; gap:22px; width:1236px; margin-top:12px; }
.step > div { flex:1; border:1px solid #000; padding:14px 18px; }
.step .hd { background:#002060; color:#fff; font-size:20px; padding:3px 14px;
            display:inline-block; margin-bottom:9px; }
.step .txt { font-size:18px; line-height:1.6; }
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

## 受注件数の方程式と実践KPI<br>地域密着・営業5名規模の販売店向け

---

<!-- ============ P2 導入 ============ -->

# はじめに

## この資料でわかること

### 地域密着型・営業人員5名程度の販売店が、工務店との提携で受注を伸ばすための実務書です

<div class="cards">
<div>
<div class="no">01</div>
<div class="ttl">アライアンスの<br>座組と使う手法</div>
<div class="txt">工務店から新築施主の太陽光・蓄電池を発注いただく「元請け開拓」の座組を整理します。開拓手法は3つあり、自社の規模と商圏でどれを選ぶべきかを示します。</div>
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

1. アライアンスとは？
2. アライアンスをやるべき理由
3. アライアンス成功のポイント

<div class="pageno">3</div>

---

<!-- ============ P4 定義 ============ -->

# 1. アライアンスとは？

## 定義と本資料が扱う範囲

### 工務店から新築施主の太陽光・蓄電池導入を発注いただく「元請け開拓」を扱います

工務店が持つ「顧客との接点と信頼」と、販売店が持つ「商品・提案・施工・アフターの実行力」を組み合わせ、双方の顧客に価値を提供する継続的な業務提携です。一度きりの案件紹介でも、元請けからの工事受注でもありません。

<table class="dense">
<tr><th style="width:190px"></th><th>下請け</th><th>単発の紹介</th><th style="background:#FF0000">アライアンス</th></tr>
<tr><td>関係性</td><td>発注者と受注者の上下関係</td><td>都度のスポット取引</td><td>対等な継続的パートナー</td></tr>
<tr><td>顧客との接点</td><td>元請け経由。自社は表に出ない</td><td>紹介された案件のみ</td><td>顧客名簿単位で継続的に接触</td></tr>
<tr><td>価格の決定権</td><td>元請けが決定</td><td>案件ごとに交渉</td><td>提携時に条件を取り決め</td></tr>
<tr><td>収益の形</td><td>工事単価のみ</td><td>単発の紹介フィー</td><td>継続的な受注＋紹介手数料</td></tr>
<tr><td>再現性</td><td>元請けの受注量に依存</td><td>低い（属人的）</td><td>高い（仕組みで積み上がる）</td></tr>
</table>

> 目指すのは案件の獲得ではなく、<b>案件が生まれ続ける関係</b>をつくること

<div class="pageno">4</div>

---

<!-- ============ P5 座組 ============ -->

# 1. アライアンスとは？

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
<tr><td>提供する<br>もの</td><td>顧客接点<br>地域での信頼<br>顧客名簿</td><td>商品調達／提案・見積<br>施工・アフター<br>補助金申請</td></tr>
<tr><td>収益</td><td>紹介手数料<br><small>販売金額の5〜7％が目安</small></td><td>工事売上<br>継続的な提案機会</td></tr>
<tr><td>リスク</td><td>顧客満足度の低下<br><small>施工品質は販売店に依存</small></td><td>提携先の紹介停滞<br>個人情報の管理責任</td></tr>
</table>

</div>
</div>

> 工務店の不安は「自社の顧客を任せること」。<b>そこを消す設計</b>が成否を決める

<div class="pageno">5</div>

---

<!-- ============ P6 3手法 ============ -->

# 1. アライアンスとは？

## 工務店を開拓する3つの手法｜自社が使うのはどれか

### 地域密着・営業5名以下なら「セミナー（勉強会）」を選ぶ

<table class="dense">
<tr>
<th style="width:180px">施策</th><th>メリット</th><th>デメリット</th>
<th style="width:130px">商圏</th><th style="width:130px">対象とする販売店</th><th style="width:230px">出口</th>
</tr>
<tr>
<td>① 電話<br><small>TELマーケティング</small></td>
<td>自社のリソースで手間をかけずにアポが取れる／確実にアプローチできる</td>
<td>決裁者までたどり着くのに時間がかかる</td>
<td>関東圏／全国</td><td>営業10名以上</td>
<td>新築＋工務店リストの提供を受ける</td>
</tr>
<tr>
<td>② ネットワーク活用</td>
<td>上手くいけば一気に広げやすい／手っ取り早く決裁者に会いやすい</td>
<td>団体ごとの制約あり（有償、営業NGなど）</td>
<td>—</td><td>誰でも</td>
<td>新築＋工務店リストの提供を受ける</td>
</tr>
<tr class="hl">
<td><b>③ セミナー</b><br><small>勉強会</small></td>
<td>意欲の高い会社を集められる／外部講師などを含めてフックを使える</td>
<td>着手〜セミナー開催に時間を要する（3〜4か月）</td>
<td><b>地域密着型</b></td><td><b>5名以下</b></td>
<td><b>新築＋工務店リストの提供を受け、太陽光・蓄電池の営業を行う</b></td>
</tr>
</table>

> 本資料は以降、<b>セミナー（勉強会）による開拓</b>を前提に解説します

<div class="pageno">6</div>

---

<!-- ============ P7 4つの課題 ============ -->

# 2. アライアンスをやるべき理由

## 工務店が抱える4つの課題

### 新築着工棟数の減少と金利上昇で、工務店も施主からの追加受注を高めたい。しかし動けていない

<div class="quad">
<div>
<div class="cat">課題① 人　材</div>
<div class="ttl">営業リソースが足りない</div>
<div class="txt">新築の受注活動で手一杯。既存顧客に追加提案する担当を置けない。採用も難しく、有効求人倍率は高止まりしています。</div>
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

> 動けない理由は「やる気」ではなく<b>構造</b>。だから外部との組み合わせで解ける

<div class="pageno">7</div>

---

<!-- ============ P8 課題の深掘り ============ -->

# 2. アライアンスをやるべき理由

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

<div class="pageno">8</div>

---

<!-- ============ P9 だからこそ ============ -->

# 2. アライアンスをやるべき理由

## だからこそのアライアンス

### 4つの課題は、販売店が既に持っている機能をそのまま当てれば解消できる

<table class="dense">
<tr><th style="width:130px"></th><th style="width:380px">工務店の課題</th><th>アライアンスによる解決</th></tr>
<tr>
<td>① 人　材</td>
<td>提案する人がいない。採用もできない</td>
<td>営業代行：販売店の営業が顧客訪問から提案・クロージングまで実施。工務店は紹介するだけでよい</td>
</tr>
<tr>
<td>② 知　識</td>
<td>容量設計・経済効果・見積が分からない</td>
<td>ツールと勉強会：アプローチブック・シミュレーション資料の提供と、営業マン向けの個社別勉強会を定期開催</td>
</tr>
<tr>
<td>③ 収　益</td>
<td>仕入が高く粗利が残らない</td>
<td>商材卸＋紹介手数料：まとめ仕入れによる原価低減と、販売金額の5〜7％を目安とした紹介フィー</td>
</tr>
<tr>
<td>④ 運　用</td>
<td>補助金・申請の事務負担が読めない</td>
<td>申請代行と施工：補助金申請・系統連系・施工・アフターまで販売店が一括対応</td>
</tr>
</table>

<blockquote class="tall">
<p>工務店は「手間なく満足度と売上が上がる」、販売店は「広告費ゼロで良質な案件が入る」<br>
<b>実績や規模は関係ない。営業マン1名の販売店でも提携は成立している</b></p>
</blockquote>

<div class="pageno">9</div>

---

<!-- ============ P10 章の締め ============ -->
<!-- _class: message -->

# 2. アライアンスをやるべき理由

## 第2章まとめ

必要性は感じている。
**しかし、動けていない。**

<div class="pageno">10</div>

---

<!-- ============ P11 方程式 ============ -->

# 3. アライアンス成功のポイント

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

<div class="pageno">11</div>

---

<!-- ============ P12 5変数の打ち手 ============ -->

# 3. アライアンス成功のポイント

## 5変数の打ち手一覧

<table class="dense">
<tr><th style="width:125px">変数</th><th style="width:240px">見るべきKPI</th><th>具体施策</th></tr>
<tr>
<td>① 提携社数</td>
<td>アタックリスト数<br>面談率／提携率／継続率</td>
<td>地域密着の工務店を第一ターゲットに、SEO・MEO／SUUMO／工務店団体／商工会議所・銀行からリストアップ。訪問のゴールを「提携」ではなく「勉強会への招待」に下げる（詳細はP13）</td>
</tr>
<tr>
<td>② 顧客名簿数</td>
<td>提携先のオーナー数<br>名簿の共有可否（0 or 1）<br>年間棟数</td>
<td>提案の幅×案件の透明化。1年目PV・2年目蓄電池・3年目塗装・4年目床下・5年目内装と5年サイクルの提案設計を示す。顧客管理ツールで「誰が・いつ・何を提案したか」を開示し、預ける不安を消す</td>
</tr>
<tr>
<td>③ 紹介率</td>
<td>提携先別の紹介率<br><small>（年間棟数に対する紹介件数）</small></td>
<td>紹介が止まる阻害要因を特定する。社長のトップダウン不足／営業の理解不足／提案資料がない／上下の熱量差。対策は個社別勉強会、営業マンへの直接説明、商談への同席、個別チラシの作成</td>
</tr>
<tr>
<td>④ 商談率</td>
<td>紹介アポ→商談の到達率<br><small>目標：100％</small></td>
<td>紹介案件は見積提出のみで済むことが多く100％を目指せる。案件管理を紙からデジタルへ移行（スプレッドシート／kintone等）し、管理表を全社で共有して抜け漏れを防ぐ</td>
</tr>
<tr>
<td>⑤ 成約率</td>
<td>商談→成約の到達率<br>営業担当者別のばらつき</td>
<td>エース営業の「心構え・アポトーク・営業ツール・クロージングトーク」を言語化しマニュアル化。営業同行・1on1・ロープレ・営業動画で組織に展開する</td>
</tr>
</table>

> 「背中を見て覚えろ」からの脱却。<b>売れる営業の育成は仕組み</b>

<div class="pageno">12</div>

---

<!-- ============ P13 勉強会 ============ -->

# 3. アライアンス成功のポイント

## ①提携社数を増やす｜勉強会の設計とKPI

### 訪問120社から締結3社。数字で追うから改善できる

<div class="funnel">
<div class="st">訪問<b>120社</b></div>
<div class="rt">12.5％<br>▶</div>
<div class="st">勉強会 参加<b>15社</b><br><small>20名ほど</small></div>
<div class="rt">20％<br>▶</div>
<div class="st">提携 締結<b>3社</b></div>
</div>

<div class="cols">
<div>

<table class="dense">
<tr><th style="width:120px">企　画</th><th>やること</th></tr>
<tr><td>頻度</td><td>3ヵ月に1回。年間の講座構成を先に決める</td></tr>
<tr><td>ゲスト</td><td>メーカー・自治体を招き、集客のフックにする</td></tr>
<tr><td>会場</td><td>商工会議所など、地域で信頼される場所</td></tr>
<tr><td>集客</td><td>遅くとも開催の1か月半前から着手する</td></tr>
</table>

</div>
<div>

<table class="dense">
<tr><th style="width:120px">運　営</th><th>やること</th></tr>
<tr><td>当日オペ</td><td>分単位のタイムテーブルを組む</td></tr>
<tr><td>訪問</td><td>2週間に1回の訪問×電話で接点を保つ</td></tr>
<tr><td>面談率</td><td>役職者が訪問／雨の日を狙う／事務員と関係を作る</td></tr>
<tr><td>締結</td><td>社長同席のもと、先方の社長に条件を伝える</td></tr>
</table>

</div>
</div>

> 訪問のゴールは「提携」ではなく<b>「勉強会への招待」</b>。ハードルを下げるから続く

<div class="pageno">13</div>

---

<!-- ============ P14 チェックリスト ============ -->

# 3. アライアンス成功のポイント

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
<span class="mk">経営に課題を感じている</span>、<span class="mk">コンサルティング内容や依頼方法が知りたい</span>、<span class="mk">専門家から事業に関するアドバイスが欲しい</span><br>
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

<div class="c-search"><span class="c-box">船井総研　経営相談</span><span>検索や船井総合研究所ウェブサイトURLからも上記ページへアクセスいただけます。</span></div>

<div class="c-foot">
<div class="c-company">株式会社船井総合研究所　会社概要<br><span>https://www.funaisoken.co.jp/info/company</span></div>
<div class="c-disc">
◆ 本資料に掲載のセミナーはすでに終了している場合があります。内容にご不明な点がある際は、別途お問い合わせください。◆ 本資料に掲載の情報は作成時点のものであり、将来的な内容を保証するものではありません。◆ 本資料に掲載の情報を目的とし、弊社承諾なく転載・改変することを禁じております。◆ 本資料に掲載の情報を利用したことによって発生する損害について、弊社は責任を負いかねます。◆ 本資料の内容は将来予告なく変更、または廃止されることがあります。<br>
◆ 本資料には生成AIにより生成した情報やPIXTAより提供された画像が含まれている場合があります。
</div>
<div class="c-logo"></div>
</div>

</div>
