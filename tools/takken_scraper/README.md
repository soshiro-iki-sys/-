# 宅建業者リスト取得スクリプト

国交省「[建設業者・宅建業者等企業情報検索システム](https://etsuran2.mlit.go.jp/TAKKEN/takkenKensaku.do)」から、
指定した都道府県の宅建業者を取得して CSV に出力します。

出力列: `都道府県, 会社名, 住所, 電話番号, 社長電話番号, URL, 根拠`

| 列 | 取得元 |
|---|---|
| 会社名・住所・電話番号 | 各業者の詳細画面（商号又は名称／主たる事務所の所在地／電話番号） |
| 社長電話番号 | 検索システムに掲載がないため常に空欄 |
| URL | 詳細画面にホームページ欄があれば取得。なければ空欄 |
| 根拠 | 免許行政庁・免許証番号・取得日・取得元 URL |

## セットアップ

```bash
pip install -r requirements.txt
playwright install chromium   # Chromium が未導入の場合のみ
```

## 使い方

```bash
# まず数件だけ取得して、画面の解析がうまくいくか確認する
python takken_scraper.py --pref 東京都 --limit 10 --dump-dir debug_html

# 本番（知事免許の業者を全件取得）
python takken_scraper.py --pref 東京都 --out output/takken_東京都.csv

# 国土交通大臣免許のうち、主たる事務所が指定県にある業者も含める
python takken_scraper.py --pref 大阪府 --include-minister

# 複数県をまとめて取得
python takken_scraper.py --pref 埼玉県 千葉県 神奈川県
```

主なオプション:

- `--delay 秒`: 画面操作の間隔（既定 2 秒）。サーバに負荷をかけないよう短くしすぎないでください
- `--no-detail`: 詳細画面を開かず一覧だけで出力する（速いが電話番号が取れない場合あり）
- `--headful`: ブラウザ画面を表示して動きを確認する
- `--dump-dir`: 取得した画面の HTML を保存する（うまく動かないときの調査用）
- `--pref-select` / `--disp-select`: 免許行政庁・表示件数プルダウンの `name` 属性を直接指定する

途中で止まっても、同じ `--out` を指定して再実行すれば取得済みの免許証番号を飛ばして続きから再開します
（検索結果一覧に免許証番号の列がある場合）。

## 仕組みと注意点

- 検索システムはフォームの POST と JavaScript で画面を切り替えるため、Playwright でブラウザを操作します。
  詳細画面は別ウィンドウで開き、一覧ページを残したままページ送りします。
- 画面の `name` 属性は変わることがあるので、「免許行政庁」の選択肢・「商号」「電話番号」などの見出しといった
  **画面上の文言**を手がかりに要素を探しています。画面構成が大きく変わった場合は `--dump-dir` で HTML を確認し、
  `takken_scraper.py` 冒頭の `*_LABELS` を調整してください。
- 都道府県知事免許の業者は「免許行政庁＝その都道府県」で検索します。複数の都道府県に事務所がある業者は
  国土交通大臣免許なので、`--include-minister` を付けたときだけ住所で絞り込んで含めます（全国分の大臣免許業者を
  たどるため時間がかかります）。
- 検索システムの利用規約・注意事項を確認のうえ、業務に必要な範囲で、間隔を空けて利用してください。

## 動作確認用の模擬サイト

`tests/mock_server.py` は検索システムの画面遷移（POST での検索・ページ送り・詳細表示、「戻る」で一覧に戻れない挙動）を
模した簡易サーバです。

```bash
python tests/mock_server.py 8765 &
python takken_scraper.py --pref 東京都 --include-minister --delay 0 \
  --url http://127.0.0.1:8765/TAKKEN/takkenKensaku.do --out /tmp/test.csv
```
