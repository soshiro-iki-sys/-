# -*- coding: utf-8 -*-
import html, io, json, os, sys

SIG = """――――――――――――――――
株式会社ワン・ミニット
【部署】【氏名】
〒【　　-　　　　】【住所】
TEL：【　　-　　　　-　　　　】
MAIL：【　　　　＠　　　　　】
URL：【　　　　　　　　　　】
――――――――――――――――"""

T1 = """【エリア】で新築を手がける工務店さまへ／太陽光の提案を、施工体制を持たずに始める方法

ご担当者さま

突然のご連絡を失礼いたします。
株式会社ワン・ミニットの【氏名】と申します。

施主さまから太陽光・蓄電池の相談を受けたときに、
「うちではできません」とお答えになったことはないでしょうか。

・機器の選定基準がわからない
・発電量のシミュレーションが出せない
・電気工事の職人がいない

この3つで提案を見送っている工務店さまが多くいらっしゃいます。
そこで、施工体制を持たないまま太陽光・蓄電池の提案を始める方法を
1冊にまとめた資料を公開しました。

■ 資料の内容（A4・全12ページ・無料）
・工務店が取りこぼしている売上はいくらか（棟数別の試算表）
・提携モデル3つの型（紹介型／販売代行型／共同施工型）
・型ごとの役割分担と、1棟あたりの収益
・年20棟の工務店での収益シミュレーション

▼ ダウンロードはこちら
【URL】

しつこいご連絡はいたしません。
資料をご覧になり、必要と感じられた場合のみご返信ください。

""" + SIG

T2 = """太陽光の提案を、施工体制を持たずに始める方法（無料資料）

突然のご連絡を失礼いたします。株式会社ワン・ミニットの【氏名】と申します。

施主さまから太陽光の相談を受けても、機器の選定・シミュレーション・電気工事の3つで
提案を見送っている工務店さまが多くいらっしゃいます。

施工体制を持たないまま提案を始める方法を、A4全12ページにまとめました。
提携モデル3つの型と、1棟あたりの収益まで載せています。無料です。

▼ ダウンロード
【URL】

株式会社ワン・ミニット／【氏名】
TEL【　　-　　　　-　　　　】／MAIL【　　　＠　　　　】"""

T3 = """賃貸・売買の管理物件に、太陽光という収益源を／オーナーさま向けのご提案資料

ご担当者さま

突然のご連絡を失礼いたします。
株式会社ワン・ミニットの【氏名】と申します。

管理されている物件のオーナーさまから、
電気代や太陽光についてご相談を受けることはございませんか。

管理手数料以外の収益源として、
オーナーさまへの太陽光・蓄電池のご案内に取り組む不動産会社さまが増えています。

御社がオーナーさまにご案内し、提案・施工・保証は当社が引き受ける形です。
物件の資産価値と入居率にも効きます。

■ 資料の内容（A4・全12ページ・無料）
・オーナー紹介モデルの仕組みと、役割分担
・1件あたりの収益と、管理戸数別の試算表
・物件価値・入居率への影響
・導入事例と、個別相談のご案内

▼ ダウンロードはこちら
【URL】

しつこいご連絡はいたしません。
資料をご覧になり、必要と感じられた場合のみご返信ください。

""" + SIG

T4 = """【勉強会】◯月◯日オンライン／太陽光・蓄電池で収益化する工務店の進め方

ご担当者さま

突然のご連絡を失礼いたします。
株式会社ワン・ミニットの【氏名】と申します。

工務店さま向けに、太陽光・蓄電池の収益化をテーマにした
オンライン勉強会を開催いたします。

■ 開催概要
日　時：◯月◯日（◯）◯◯:◯◯〜◯◯:◯◯（90分）
形　式：Zoomウェビナー
対　象：新築を手がける工務店・ビルダーさま
定　員：30社（1社2名さままで）
参加費：無料

■ 当日の内容
第1部　市場と補助金の最新
第2部　提携モデルの解説（工事を持たずに売上をつくる3つの型）
第3部　収益試算のデモ（御社の棟数で試算します）
第4部　質疑と、個別相談のご案内

▼ お申し込みはこちら
【URL】

売り込みの場ではありません。
まず情報だけ持ち帰っていただく前提で構いません。

""" + SIG

T5 = """先日の資料のご案内／太陽光・蓄電池 提携モデル解説（再送）

ご担当者さま

株式会社ワン・ミニットの【氏名】です。
以前にも資料のご案内をお送りしております。重ねてのご連絡を失礼いたします。

前回から【　　】ヶ月が経ち、補助金の要件と機器の価格が変わりました。
資料も【　　　　　　　】の内容を追加して更新しております。

■ 今回の追加分
・【　　　　　　　　　　　　　　　　　】
・【　　　　　　　　　　　　　　　　　】

▼ 更新版のダウンロードはこちら
【URL】

もし今後のご案内が不要でしたら、その旨ご返信ください。
以後お送りいたしません。

""" + SIG

T6 = """資料をダウンロードいただきありがとうございました

【会社名】
【ご担当者名】さま

株式会社ワン・ミニットの【氏名】です。
このたびは資料をダウンロードいただき、ありがとうございました。

資料の3つの型のうち、どこから始めるのが御社に合うかは、
年間棟数と、社内に電気工事士がいらっしゃるかで変わります。

もしよろしければ、オンラインで60分ほどお時間をいただき、
御社の棟数に合わせた試算をお出しします。無料です。

▼ ご都合のよい日程をお選びください
【URL】

もちろん、資料だけで十分という場合はご放念ください。

""" + SIG

TEMPLATES = [
    ("01", "基本形｜工務店向け・DL訴求", "第1回・第4回・第6回",
     "文字数制限が1,000字以上のフォーム。年8回のうち、DL訴求の主力に使う。", T1),
    ("02", "短縮形｜工務店向け・DL訴求", "文字数制限のあるフォーム",
     "500字以下しか入らないフォーム用。01と同じ内容を、冒頭2行と資料の中身だけに削った。", T2),
    ("03", "不動産会社向け・DL訴求", "第5回・第7回",
     "案B（不動産会社）に送るとき。オーナー名簿が母数になる相手なので、訴求を「管理手数料以外の収益源」に置き換えている。", T3),
    ("04", "勉強会の告知回", "第3回・第7回",
     "セミナー開催の1.5ヶ月前に送る。DLではなく申込フォームへ直接誘導する。", T4),
    ("05", "再送用｜2回目以降", "第2回以降の同一宛先",
     "同じ会社への2通目以降。前回からの変化を必ず1つ書く。停止依頼の受け口を明記する。", T5),
    ("06", "DL後のフォロー", "DL獲得の当日〜翌営業日",
     "フォームではなくメールで送る。DL者を個別相談に載せる導線。", T6),
]

SUBJECTS = [
    ("第1回", "提携モデル解説", "太陽光の提案を、施工体制を持たずに始める方法（無料資料）"),
    ("第2回", "制度・補助金", "【◯年度】太陽光・蓄電池の補助金、締切と要件をまとめました"),
    ("第3回", "勉強会告知", "【勉強会】◯月◯日オンライン／太陽光・蓄電池で収益化する工務店の進め方"),
    ("第4回", "成功事例", "年20棟の工務店が、棟数を増やさずに粗利を◯◯万円上げた話"),
    ("第5回", "不動産向け", "管理物件のオーナーさまに、太陽光という収益源を（無料資料）"),
    ("第6回", "市場データ", "太陽光の設置動向と価格の推移／◯年◯月時点のまとめ"),
    ("第7回", "勉強会告知", "【勉強会】◯月◯日オンライン／オーナー提案で収益をつくる不動産会社の進め方"),
    ("第8回", "総括レポート", "◯年の太陽光・蓄電池まとめ／来期に効く3つの変化"),
]

FIELDS = [
    ("会社名", "株式会社ワン・ミニット", "略さない。「(株)」は使わない"),
    ("部署名", "【部署】", "空欄不可のフォームが多い。統一しておく"),
    ("氏名", "【氏名】", "毎回同じ人にする。認知されるまで変えない"),
    ("フリガナ", "【フリガナ】", "全角カタカナ。半角だと弾かれるフォームがある"),
    ("メール", "【　　　＠　　　】", "返信を受ける実在のアドレス。noreply は使わない"),
    ("電話番号", "【　　-　　　-　　　】", "つながる番号。ここで信用が決まる"),
    ("住所", "〒【　　-　　　】【住所】", "本社住所"),
    ("問い合わせ種別", "その他 ／ 資料のご案内", "選択肢がある場合。「営業」があれば正直に選ぶ"),
    ("件名", "下の件名案から", "件名欄がないフォームは、本文1行目を件名にする"),
    ("本文", "下の文面から", "改行が消えるフォームがある。送信前にプレビューで確認"),
]

def esc(t):
    return html.escape(t)

def cnt(t):
    return len(t.replace('\n', '').replace(' ', ''))

blocks = []
for no, name, use, note, body in TEMPLATES:
    blocks.append(f'''<article class="tpl">
  <header class="tpl-h">
    <span class="tpl-no">{no}</span>
    <div class="tpl-id">
      <h3>{esc(name)}</h3>
      <p class="tpl-use">{esc(use)}</p>
    </div>
    <span class="tpl-len">{cnt(body)}<span class="unit">字</span></span>
  </header>
  <p class="tpl-note">{esc(note)}</p>
  <div class="tpl-body">
    <button class="copy" type="button" data-i="{no}">コピー</button>
    <pre id="t{no}">{esc(body)}</pre>
  </div>
</article>''')

subj = '\n'.join(
    f'<tr><td class="mono">{esc(a)}</td><td class="tag">{esc(b)}</td><td>{esc(c)}</td></tr>'
    for a, b, c in SUBJECTS)
fields = '\n'.join(
    f'<tr><th scope="row">{esc(a)}</th><td class="fill">{esc(b)}</td><td class="hint">{esc(c)}</td></tr>'
    for a, b, c in FIELDS)

HTML = '''<title>フォーム営業の送信文</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Zen+Kaku+Gothic+New:wght@500;700&family=Noto+Sans+JP:wght@400;500;700&family=IBM+Plex+Mono:wght@500;600&display=swap">
<style>
:root{
  --ground:#FBF9F6; --surface:#FFFFFF; --sunk:#F4EFE8;
  --ink:#2B2A28; --muted:#7A736A; --rule:#E5DED4; --rule-soft:#EFE9E1;
  --accent:#E7752F; --amber:#FAA413; --ok:#3F7A52;
}
@media (prefers-color-scheme:dark){
  :root:not([data-theme="light"]){
    --ground:#191715; --surface:#221F1C; --sunk:#14120F;
    --ink:#EFE9E1; --muted:#A29A8F; --rule:#37312B; --rule-soft:#2A2521;
    --accent:#F0894A; --amber:#F2A83F; --ok:#6FB183;
  }
}
:root[data-theme="dark"]{
  --ground:#191715; --surface:#221F1C; --sunk:#14120F;
  --ink:#EFE9E1; --muted:#A29A8F; --rule:#37312B; --rule-soft:#2A2521;
  --accent:#F0894A; --amber:#F2A83F; --ok:#6FB183;
}
*{box-sizing:border-box;}
body{
  margin:0; background:var(--ground); color:var(--ink);
  font-family:"Noto Sans JP","Hiragino Kaku Gothic ProN","Yu Gothic",Meiryo,sans-serif;
  font-size:15px; line-height:1.85; -webkit-font-smoothing:antialiased;
}
.wrap{max-width:940px; margin:0 auto; padding:52px 22px 90px;}
h1,h2,h3{font-family:"Zen Kaku Gothic New","Noto Sans JP",sans-serif; margin:0; text-wrap:balance;}
.mono{font-family:"IBM Plex Mono",ui-monospace,monospace; font-variant-numeric:tabular-nums;}

.eyebrow{font-size:11px; letter-spacing:.18em; color:var(--accent); font-weight:700;
  font-family:"IBM Plex Mono",monospace; margin:0 0 12px;}
h1{font-size:clamp(27px,4vw,37px); font-weight:700; line-height:1.35;}
.stand{color:var(--muted); max-width:64ch; margin:18px 0 0; font-size:15px;}

.facts{display:grid; grid-template-columns:repeat(4,1fr); gap:1px; margin:34px 0 0;
  background:var(--rule); border:1px solid var(--rule); border-radius:5px; overflow:hidden;}
.fact{background:var(--surface); padding:15px 17px;}
.fact dt{font-size:10.5px; letter-spacing:.11em; color:var(--muted); font-weight:700; margin:0;}
.fact dd{margin:3px 0 0; font-size:21px; font-weight:600;
  font-family:"IBM Plex Mono",monospace; font-variant-numeric:tabular-nums;}
.fact dd span{font-size:12px; color:var(--muted); margin-left:3px;}

section{margin-top:56px;}
.sec-h{border-top:2px solid var(--ink); padding-top:13px; margin-bottom:22px;}
.sec-h .n{font-family:"IBM Plex Mono",monospace; font-size:11px; font-weight:600;
  color:var(--accent); letter-spacing:.1em;}
.sec-h h2{font-size:21px; font-weight:700; margin-top:3px;}
.sec-h p{color:var(--muted); font-size:14px; margin:9px 0 0; max-width:68ch;}

.tpls{display:flex; flex-direction:column; gap:26px;}
.tpl{background:var(--surface); border:1px solid var(--rule); border-radius:5px; overflow:hidden;}
.tpl-h{display:flex; align-items:center; gap:14px; padding:15px 18px; border-bottom:1px solid var(--rule-soft);}
.tpl-no{font-family:"IBM Plex Mono",monospace; font-size:13px; font-weight:600;
  color:var(--surface); background:var(--accent); padding:3px 9px; border-radius:3px; flex:none;}
.tpl-id{flex:1; min-width:0;}
.tpl-id h3{font-size:16px; font-weight:700;}
.tpl-use{margin:1px 0 0; font-size:11.5px; color:var(--muted);
  font-family:"IBM Plex Mono",monospace;}
.tpl-len{font-family:"IBM Plex Mono",monospace; font-size:19px; font-weight:600;
  color:var(--muted); flex:none; font-variant-numeric:tabular-nums;}
.tpl-len .unit{font-size:11px; margin-left:2px;}
.tpl-note{margin:0; padding:12px 18px; font-size:13.5px; color:var(--muted);
  background:var(--sunk); border-bottom:1px solid var(--rule-soft);}
.tpl-body{position:relative;}
pre{margin:0; padding:20px 18px; white-space:pre-wrap; word-break:break-word;
  font-family:"Noto Sans JP",sans-serif; font-size:13.5px; line-height:2.0;}
.copy{position:absolute; top:12px; right:12px; font:600 11.5px/1 "IBM Plex Mono",monospace;
  color:var(--ink); background:var(--surface); border:1px solid var(--rule);
  border-radius:3px; padding:7px 11px; cursor:pointer;}
.copy:hover{border-color:var(--accent); color:var(--accent);}
.copy:focus-visible{outline:2px solid var(--accent); outline-offset:2px;}
.copy.done{color:var(--ok); border-color:var(--ok);}

.tbl{overflow-x:auto; background:var(--surface); border:1px solid var(--rule); border-radius:5px;}
table{border-collapse:collapse; width:100%; min-width:560px; font-size:14px;}
th{text-align:left; font-size:10.5px; letter-spacing:.1em; color:var(--muted); font-weight:700;
  padding:12px 15px; border-bottom:1px solid var(--rule); white-space:nowrap;}
td{padding:11px 15px; border-bottom:1px solid var(--rule-soft); vertical-align:top;}
tr:last-child td{border-bottom:none;}
th[scope=row]{font-size:14px; color:var(--ink); letter-spacing:0; white-space:nowrap;
  border-bottom:1px solid var(--rule-soft); vertical-align:top;}
td.mono{font-family:"IBM Plex Mono",monospace; color:var(--muted); white-space:nowrap;}
td.tag{white-space:nowrap;}
td.tag::before{content:""; display:inline-block; width:6px; height:6px; border-radius:50%;
  background:var(--amber); margin-right:7px; vertical-align:middle;}
td.fill{font-family:"IBM Plex Mono",monospace; font-size:13px;}
td.hint{color:var(--muted); font-size:13px;}

.rules{display:grid; grid-template-columns:1fr 1fr; gap:18px;}
.rule{background:var(--surface); border:1px solid var(--rule); border-radius:5px; padding:19px 21px;}
.rule h3{font-size:15px; font-weight:700; margin-bottom:11px;}
.rule ul{margin:0; padding:0; list-style:none; display:flex; flex-direction:column; gap:9px;}
.rule li{font-size:13.5px; padding-left:19px; position:relative; line-height:1.75;}
.rule li::before{content:""; position:absolute; left:2px; top:9px; width:6px; height:6px;
  border-radius:50%; background:var(--accent);}
.rule li b{font-weight:700;}

.warn{margin-top:22px; padding:19px 22px; background:var(--surface);
  border:1px solid var(--rule); border-left:3px solid var(--accent); border-radius:0 5px 5px 0;}
.warn h3{font-size:15px; font-weight:700; margin-bottom:8px;}
.warn p{margin:0; color:var(--muted); font-size:13.5px;}
.warn p + p{margin-top:10px;}
.warn strong{color:var(--ink);}

@media (max-width:760px){
  .facts{grid-template-columns:1fr 1fr;}
  .rules{grid-template-columns:1fr;}
}
@media (prefers-reduced-motion:reduce){*{transition:none!important; animation:none!important;}}
</style>

<div class="wrap">
  <p class="eyebrow">ONE MINUTE ／ FORM OUTREACH</p>
  <h1>フォーム営業の送信文</h1>
  <p class="stand">
    商圏3,000社の問い合わせフォームに、1.5ヶ月に1本のペースで送るための文面集です。
    目的はDL資料のダウンロード。企画書の「フォーム・メルマガ・DL」の数字に合わせています。
  </p>

  <dl class="facts">
    <div class="fact"><dt>送信先</dt><dd>3,000<span>社</span></dd></div>
    <div class="fact"><dt>頻度</dt><dd>年8<span>回</span></dd></div>
    <div class="fact"><dt>月あたり</dt><dd>1,920<span>通</span></dd></div>
    <div class="fact"><dt>年間DL目標</dt><dd>276<span>件</span></dd></div>
  </dl>

  <section>
    <div class="sec-h">
      <span class="n">TEMPLATES</span>
      <h2>6つの文面</h2>
      <p>そのまま貼れる形にしています。【　】だけ埋めてください。文字数は空白と改行を除いた本文の実数です。</p>
    </div>
    <div class="tpls">
''' + '\n'.join(blocks) + '''
    </div>
  </section>

  <section>
    <div class="sec-h">
      <span class="n">SUBJECTS</span>
      <h2>年8回の件名案</h2>
      <p>件名欄のあるフォーム用。件名欄がない場合は、本文の1行目をそのまま件名にしてください。</p>
    </div>
    <div class="tbl">
      <table>
        <thead><tr><th>回</th><th>ネタ</th><th>件名</th></tr></thead>
        <tbody>
''' + subj + '''
        </tbody>
      </table>
    </div>
  </section>

  <section>
    <div class="sec-h">
      <span class="n">FIELDS</span>
      <h2>フォームの記入欄</h2>
      <p>どのフォームもだいたい同じ項目を聞かれます。毎回同じ内容を入れてください。</p>
    </div>
    <div class="tbl">
      <table>
        <thead><tr><th>項目</th><th>入れる内容</th><th>注意</th></tr></thead>
        <tbody>
''' + fields + '''
        </tbody>
      </table>
    </div>
  </section>

  <section>
    <div class="sec-h">
      <span class="n">RULES</span>
      <h2>運用のルール</h2>
    </div>
    <div class="rules">
      <div class="rule">
        <h3>送り方</h3>
        <ul>
          <li><b>同じ会社には年8回まで。</b>8回送って反応がなければ、その先はリストから外す</li>
          <li><b>差出人と署名は毎回同じにする。</b>担当者名を変えない。認知されるまで続ける</li>
          <li><b>変えるのは冒頭2行だけ。</b>資料の中身と署名は共通のまま回す</li>
          <li><b>1回の配信を1.5ヶ月に分けて流す。</b>月1,920通でツール枠（月3,000社）の64％</li>
          <li><b>送信前にプレビューで改行を確認する。</b>改行が消えるフォームがある</li>
        </ul>
      </div>
      <div class="rule">
        <h3>数字の見方</h3>
        <ul>
          <li><b>1回あたりのDL率は1.2％。</b>3,000社に送って36件が目安</li>
          <li><b>年8回の累積で9.2％＝276件。</b>回を重ねるほど未反応の母数が減る</li>
          <li><b>DL者はその日のうちにメルマガへ載せる。</b>ここが次の母数になる</li>
          <li><b>返信が来たら24時間以内に返す。</b>フォーム営業は反応の鮮度で決まる</li>
          <li><b>件名だけを毎回記録する。</b>どの件名でDLが増えたかが唯一の学習材料</li>
        </ul>
      </div>
    </div>

    <div class="warn">
      <h3>停止依頼の受け口を必ず入れる</h3>
      <p>
        問い合わせフォームを営業目的で使うことを、利用規約で禁じている企業があります。
        文面には必ず<strong>「今後のご案内が不要でしたらご返信ください。以後お送りいたしません」</strong>
        の一文を入れ、返信があった先は<strong>即日リストから外してください</strong>。
      </p>
      <p>
        これはクレームを抑えるためだけではありません。外した先を記録していくと、
        残ったリストの精度が回を追うごとに上がります。3,000社を8回回すあいだに、
        送るべき相手だけが残る状態をつくるのが狙いです。
      </p>
    </div>
  </section>
</div>

<script>
document.querySelectorAll('.copy').forEach(function (b) {
  b.addEventListener('click', function () {
    var pre = document.getElementById('t' + b.dataset.i);
    var text = pre.textContent;
    var done = function () {
      b.textContent = 'コピーしました'; b.classList.add('done');
      setTimeout(function () { b.textContent = 'コピー'; b.classList.remove('done'); }, 1800);
    };
    var fallback = function () {
      var ta = document.createElement('textarea');
      ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta); ta.select();
      var ok = false;
      try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
      document.body.removeChild(ta);
      if (ok) { done(); } else { b.textContent = '選択してコピーしてください'; }
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(done, fallback);
    } else { fallback(); }
  });
});
</script>
'''
# HTML（アーティファクト用）の出力先は第1引数で指定できる
html_out = sys.argv[1] if len(sys.argv) > 1 else 'output/form-outreach.html'
os.makedirs(os.path.dirname(html_out) or '.', exist_ok=True)
io.open(html_out, 'w', encoding='utf-8').write(HTML)

md = ['# フォーム営業の送信文（株式会社ワン・ミニット）', '',
      '商圏3,000社へ1.5ヶ月に1本・年8回。目的はDL資料のダウンロード。', '']
for no, name, use, note, body in TEMPLATES:
    md += ['## %s %s' % (no, name), '', '- 使いどころ：%s' % use, '- %s' % note,
           '- 文字数：%d字（空白・改行を除く）' % cnt(body), '', '```', body, '```', '']
md += ['## 年8回の件名案', '', '| 回 | ネタ | 件名 |', '|---|---|---|']
md += ['| %s | %s | %s |' % s for s in SUBJECTS]
md += ['', '## フォームの記入欄', '', '| 項目 | 入れる内容 | 注意 |', '|---|---|---|']
md += ['| %s | %s | %s |' % f for f in FIELDS]
os.makedirs('output', exist_ok=True)
io.open('output/フォーム営業_送信文.md', 'w', encoding='utf-8').write('\n'.join(md) + '\n')
print('ok / 文字数:', [(t[0], cnt(t[4])) for t in TEMPLATES])
