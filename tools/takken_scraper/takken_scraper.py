#!/usr/bin/env python3
"""国交省「宅建業者等企業情報検索システム」から指定都道府県の宅建業者一覧を取得し CSV に出力する。

出力列: 都道府県, 会社名, 住所, 電話番号, 社長電話番号, URL, 根拠

検索画面はフォームの POST と JavaScript で画面遷移するため Playwright でブラウザを操作し、
各ページの HTML を BeautifulSoup で解析する。画面の項目名（name 属性など）は変わることがあるので、
「商号」「電話番号」といった画面上の文言を手がかりに要素を探す作りにしている。
うまく動かない場合は --dump-dir で HTML を保存して中身を確認し、--pref-select などで要素を指定する。

使い方:
    python takken_scraper.py --pref 東京都 --out output/takken_東京都.csv
    python takken_scraper.py --pref 大阪府 --include-minister --limit 20 --headful
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import logging
import re
import sys
import time
import unicodedata
from pathlib import Path

from bs4 import BeautifulSoup
from playwright.sync_api import Page, sync_playwright

DEFAULT_URL = "https://etsuran2.mlit.go.jp/TAKKEN/takkenKensaku.do"
COLUMNS = ["都道府県", "会社名", "住所", "電話番号", "社長電話番号", "URL", "根拠"]
MINISTER = "国土交通大臣"
PREFECTURES = [
    "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
    "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
    "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
    "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県",
    "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県",
    "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県",
    "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県",
]

# 一覧・詳細画面の項目名の候補（部分一致）
NAME_LABELS = ["商号又は名称", "商号", "名称"]
NAME_EXCLUDE = ["フリガナ", "カナ", "ｶﾅ", "事務所"]
ADDRESS_LABELS = ["主たる事務所の所在地", "主たる事務所所在地", "所在地", "住所"]
PHONE_LABELS = ["電話番号", "電話"]
PHONE_EXCLUDE = ["FAX", "ＦＡＸ", "ファクス"]
URL_LABELS = ["ホームページ", "URL", "ＵＲＬ", "Webサイト", "ウェブサイト"]
LICENSE_LABELS = ["免許証番号", "免許番号"]

ROW_ATTR = "data-takken-row"
log = logging.getLogger("takken")


# ---------------------------------------------------------------- HTML 解析

def clean(text: str | None) -> str:
    return re.sub(r"\s+", " ", (text or "").replace("　", " ")).strip()


def nfkc(text: str) -> str:
    return clean(unicodedata.normalize("NFKC", text))


def label_matches(label: str, include: list[str], exclude: list[str] = ()) -> bool:
    label = nfkc(label).replace(" ", "")
    return any(nfkc(k) in label for k in include) and not any(nfkc(k) in label for k in exclude)


def label_value_pairs(soup: BeautifulSoup) -> list[tuple[str, str, object]]:
    """表（th/td）と定義リスト（dt/dd）から (項目名, 値, 値の要素) を順に取り出す。"""
    pairs = []
    for tr in soup.find_all("tr"):
        cells = tr.find_all(["th", "td"], recursive=False)
        if any(c.name == "th" for c in cells):
            for i, cell in enumerate(cells[:-1]):
                if cell.name == "th" and cells[i + 1].name == "td":
                    pairs.append((clean(cell.get_text(" ")), clean(cells[i + 1].get_text(" ")), cells[i + 1]))
        elif len(cells) == 2:
            pairs.append((clean(cells[0].get_text(" ")), clean(cells[1].get_text(" ")), cells[1]))
    for dt_tag in soup.find_all("dt"):
        dd = dt_tag.find_next_sibling("dd")
        if dd:
            pairs.append((clean(dt_tag.get_text(" ")), clean(dd.get_text(" ")), dd))
    return pairs


def pick(pairs, include, exclude=()) -> tuple[str, object]:
    # 候補語の優先順に探す（「商号又は名称」を「名称」より優先するため）
    for key in include:
        for label, value, elem in pairs:
            if value and label_matches(label, [key], list(exclude)):
                return value, elem
    return "", None


def normalize_phone(text: str) -> str:
    m = re.search(r"0\d{1,4}[-(（]?\d{1,4}[-)）]?\d{3,4}", nfkc(text))
    return m.group(0).replace("(", "-").replace(")", "-").replace("（", "-").replace("）", "-") if m else nfkc(text)


def parse_detail(html: str) -> dict:
    soup = BeautifulSoup(html, "html.parser")
    pairs = label_value_pairs(soup)
    name, _ = pick(pairs, NAME_LABELS, NAME_EXCLUDE)
    address, _ = pick(pairs, ADDRESS_LABELS, ["フリガナ", "カナ"])
    phone, _ = pick(pairs, PHONE_LABELS, PHONE_EXCLUDE)
    license_no, _ = pick(pairs, LICENSE_LABELS)
    url, url_elem = pick(pairs, URL_LABELS)
    if url_elem is not None and url_elem.find("a", href=True):
        url = url_elem.find("a", href=True)["href"]
    m = re.search(r"https?://[^\s　]+", url)
    return {
        "会社名": name,
        "住所": address,
        "電話番号": normalize_phone(phone) if phone else "",
        "URL": m.group(0) if m else "",
        "免許証番号": nfkc(license_no),
    }


def parse_result_list(html: str) -> list[dict]:
    """MARK_ROWS_JS で番号を振った一覧の各行を {列名: 値} として返す。"""
    soup = BeautifulSoup(html, "html.parser")
    first = soup.find("tr", attrs={ROW_ATTR: True})
    if first is None:
        return []
    table = first.find_parent("table")
    heads = []
    for tr in table.find_all("tr"):
        cells = [clean(c.get_text(" ")) for c in tr.find_all(["th", "td"], recursive=False)]
        if len(cells) > 2 and any(label_matches(c, NAME_LABELS, NAME_EXCLUDE) for c in cells):
            heads = cells
            break
    rows = []
    for tr in table.find_all("tr", attrs={ROW_ATTR: True}):
        cells = [clean(c.get_text(" ")) for c in tr.find_all(["th", "td"], recursive=False)]
        row = {h: (cells[i] if i < len(cells) else "") for i, h in enumerate(heads)}
        row["_index"] = int(tr[ROW_ATTR])
        row["_has_link"] = tr.find(["a", "input", "button"]) is not None
        rows.append(row)
    return rows


def column(row: dict, include, exclude=()) -> str:
    for key in include:
        for label, value in row.items():
            if not label.startswith("_") and label_matches(label, [key], list(exclude)):
                return value
    return ""


# ---------------------------------------------------------------- ブラウザ操作

MARK_ROWS_JS = """
(attr) => {
  const nameLabels = ['商号', '名称'];
  const exclude = ['フリガナ', 'カナ', '事務所'];
  const norm = s => (s || '').replace(/\\s+/g, '');
  for (const table of document.querySelectorAll('table')) {
    const trs = [...table.querySelectorAll('tr')].filter(tr => tr.closest('table') === table);
    const headIdx = trs.findIndex(tr => {
      const cells = [...tr.children].map(c => norm(c.textContent));
      return cells.length > 2 && cells.some(c => nameLabels.some(k => c.includes(k)) && !exclude.some(k => c.includes(k)));
    });
    if (headIdx < 0) continue;
    let n = 0;
    trs.slice(headIdx + 1).forEach(tr => {
      if (tr.querySelector('td')) tr.setAttribute(attr, String(n++));
    });
    return n;
  }
  return 0;
}
"""

# name を渡すとフォームとリンクの target を差し替え、null を渡すと元に戻す
SET_TARGET_JS = """
(name) => {
  const els = [...document.querySelectorAll('form')].concat([...document.querySelectorAll('a[href]')].filter(a => {
    const h = (a.getAttribute('href') || '').toLowerCase();
    return !h.startsWith('#') && !h.startsWith('javascript:');
  }));
  for (const el of els) {
    if (name) {
      if (!el.hasAttribute('data-orig-target')) el.setAttribute('data-orig-target', el.getAttribute('target') || '');
      el.target = name;
    } else if (el.hasAttribute('data-orig-target')) {
      const orig = el.getAttribute('data-orig-target');
      if (orig) el.setAttribute('target', orig); else el.removeAttribute('target');
      el.removeAttribute('data-orig-target');
    }
  }
}
"""


class Scraper:
    def __init__(self, page: Page, args):
        self.page = page
        self.args = args
        self.dump_count = 0

    def pause(self):
        time.sleep(self.args.delay)

    def dump(self, tag: str, html: str | None = None):
        if not self.args.dump_dir:
            return
        self.dump_count += 1
        path = Path(self.args.dump_dir) / f"{self.dump_count:04d}_{tag}.html"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(html if html is not None else self.page.content(), encoding="utf-8")

    def wait(self):
        self.page.wait_for_load_state("domcontentloaded")
        try:
            self.page.wait_for_load_state("networkidle", timeout=10_000)
        except Exception:
            pass

    # --- 検索条件の入力

    def select_option_by_label(self, label: str, override: str | None, prefer: str | None = None):
        if override:
            self.page.select_option(f"select[name='{override}']", label=label)
            return
        candidates = []
        for sel in self.page.locator("select").all():
            texts = [clean(t) for t in sel.locator("option").all_inner_texts()]
            if label in texts:
                candidates.append((prefer is not None and prefer in texts, sel))
        if not candidates:
            raise RuntimeError(f"「{label}」を選べるプルダウンが見つかりません。--dump-dir で画面を確認し --pref-select で指定してください。")
        candidates.sort(key=lambda c: not c[0])
        candidates[0][1].select_option(label=label)

    def set_display_count(self):
        if self.args.disp_select:
            sel = self.page.locator(f"select[name='{self.args.disp_select}']")
        else:
            sel = None
            for cand in self.page.locator("select").all():
                texts = [nfkc(t) for t in cand.locator("option").all_inner_texts()]
                nums = [re.sub(r"[件\s]", "", t) for t in texts]
                if len(nums) >= 2 and all(n.isdigit() for n in nums):
                    sel = cand
                    break
            if sel is None:
                log.info("表示件数のプルダウンが見つからないため既定の件数で検索します")
                return
        texts = sel.locator("option").all_inner_texts()
        best = max(texts, key=lambda t: int(re.sub(r"\D", "", nfkc(t)) or 0))
        sel.select_option(label=best)

    def click_search(self):
        loc = self.page.locator(
            "input[type=submit][value*='検索'], input[type=button][value*='検索'], "
            "input[type=image][alt*='検索'], button:has-text('検索'), a:has-text('検索する'), "
            "img[alt*='検索']"
        ).filter(has_not_text="クリア")
        if loc.count() == 0:
            raise RuntimeError("検索ボタンが見つかりません。--dump-dir で画面を確認してください。")
        loc.first.click()
        self.wait()

    def search(self, authority: str):
        log.info("検索: 免許行政庁=%s", authority)
        for attempt in range(3):
            try:
                self.page.goto(self.args.url)
                break
            except Exception:
                if attempt == 2:
                    raise
                self.page.wait_for_timeout(1000)
        self.wait()
        self.dump("search_form")
        self.select_option_by_label(authority, self.args.pref_select, prefer=MINISTER)
        self.set_display_count()
        self.pause()
        self.click_search()

    # --- 一覧ページ

    def list_rows(self) -> list[dict]:
        if self.page.evaluate(MARK_ROWS_JS, ROW_ATTR) == 0:
            return []
        return parse_result_list(self.page.content())

    def next_page(self) -> bool:
        loc = self.page.locator(
            "a:text-is('次へ'), a:has-text('次へ'), a:text-is('次ページ'), a:text-is('>'), a:text-is('＞'), "
            "input[value='次へ'], input[value*='次ページ'], input[value='>'], button:has-text('次へ')"
        )
        for i in range(loc.count()):
            el = loc.nth(i)
            if el.is_visible() and el.is_enabled():
                el.click()
                self.wait()
                return True
        return False

    def go_to_page(self, authority: str, page_no: int):
        """戻るボタンで一覧に戻れない場合に、検索し直して page_no ページ目まで進める。"""
        self.search(authority)
        for _ in range(page_no - 1):
            self.pause()
            if not self.next_page():
                raise RuntimeError(f"{page_no} ページ目まで戻れませんでした")

    # --- 詳細ページ

    def open_detail(self, index: int) -> tuple[str, Page | None]:
        """行のリンクを押して詳細画面の HTML を返す。

        検索システムは POST で画面遷移するため、同じタブで詳細を開くと「戻る」で一覧に戻れない。
        そこでフォームとリンクの target を別ウィンドウにして、一覧ページを残したまま詳細を開く。
        """
        self.page.evaluate(SET_TARGET_JS, "_takken_detail")
        row = self.page.locator(f"tr[{ROW_ATTR}='{index}']")
        link = row.locator("a, input[type=button], input[type=submit], button").first
        try:
            with self.page.context.expect_page(timeout=10_000) as info:
                link.click()
            popup = info.value
            popup.wait_for_load_state("domcontentloaded")
            return popup.content(), popup
        except Exception as e:  # noqa: BLE001  別ウィンドウにならない遷移（location.href 等）
            log.debug("別ウィンドウで開けなかったため同じタブで表示: %s", e)
            self.wait()
            return self.page.content(), None

    def back_to_list(self, authority: str, page_no: int, popup: Page | None) -> list[dict]:
        if popup is not None:
            popup.close()
            self.page.evaluate(SET_TARGET_JS, None)  # ページ送り等が別ウィンドウに飛ばないよう元に戻す
            rows = self.list_rows()
            if rows:
                return rows
        try:
            self.page.go_back()
            self.wait()
            rows = self.list_rows()
            if rows:
                return rows
        except Exception as e:  # noqa: BLE001
            log.debug("go_back 失敗: %s", e)
        log.info("一覧に戻れないため再検索して %d ページ目へ移動します", page_no)
        self.go_to_page(authority, page_no)
        return self.list_rows()


# ---------------------------------------------------------------- メイン処理

def load_done(out: Path) -> set[str]:
    """既存 CSV の「根拠」列から取得済みの免許証番号を集める（途中再開用）。"""
    if not out.exists():
        return set()
    done = set()
    with out.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            m = re.search(r"免許証番号:([^\s;]+)", row.get("根拠", ""))
            if m:
                done.add(m.group(1))
    return done


def license_key(text: str) -> str:
    return nfkc(text).replace(" ", "")


def scrape_authority(scraper: Scraper, authority: str, pref: str, writer, done: set[str], counter: list[int]):
    args = scraper.args
    today = dt.date.today().isoformat()
    scraper.search(authority)
    page_no = 1
    seen_first = set()
    while True:
        rows = scraper.list_rows()
        scraper.dump(f"list_{page_no}")
        if not rows:
            if page_no == 1:
                log.warning("検索結果の一覧が見つかりません（%s）。--dump-dir の HTML を確認してください。", authority)
            break
        first_key = tuple(sorted((k, v) for k, v in rows[0].items() if not k.startswith("_")))
        if first_key in seen_first:  # 次へを押しても同じページなら終了
            break
        seen_first.add(first_key)
        log.info("%s %d ページ目: %d 件", authority, page_no, len(rows))

        for i in range(len(rows)):
            if args.limit and counter[0] >= args.limit:
                return
            row = rows[i]
            list_license = license_key(column(row, LICENSE_LABELS))
            list_address = column(row, ADDRESS_LABELS, ["フリガナ", "カナ"])
            if list_license and list_license in done:
                continue
            # 大臣免許は全国分なので、一覧に所在地があれば対象県以外を詳細を開かずに除外
            if authority == MINISTER and list_address and not nfkc(list_address).startswith(pref):
                continue

            detail = {}
            if row["_has_link"] and not args.no_detail:
                scraper.pause()
                html, popup = scraper.open_detail(row["_index"])
                detail = parse_detail(html)
                scraper.dump(f"detail_{page_no}_{i}", html)
                rows = scraper.back_to_list(authority, page_no, popup)

            address = detail.get("住所") or list_address
            if authority == MINISTER and not nfkc(address).startswith(pref):
                continue
            license_no = detail.get("免許証番号") or nfkc(column(row, LICENSE_LABELS))
            name = detail.get("会社名") or column(row, NAME_LABELS, NAME_EXCLUDE)
            phone = detail.get("電話番号") or normalize_phone(column(row, PHONE_LABELS, PHONE_EXCLUDE))
            writer.writerow({
                "都道府県": pref,
                "会社名": name,
                "住所": address,
                "電話番号": phone,
                "社長電話番号": "",  # 国交省の検索システムには掲載されない
                "URL": detail.get("URL", ""),
                "根拠": f"国交省 宅建業者等企業情報検索システム 免許行政庁:{authority} "
                        f"免許証番号:{license_key(license_no) or '不明'} 取得日:{today} {args.url}",
            })
            counter[0] += 1
            if license_no:
                done.add(license_key(license_no))

        scraper.pause()
        if not scraper.next_page():
            break
        page_no += 1


def main(argv=None):
    p = argparse.ArgumentParser(description="国交省の宅建業者検索から指定都道府県の業者一覧を CSV に出力する")
    p.add_argument("--pref", required=True, nargs="+", help="都道府県名（例: 東京都 大阪府）")
    p.add_argument("--out", help="出力 CSV（既定: output/takken_<都道府県>.csv）")
    p.add_argument("--include-minister", action="store_true",
                   help="国土交通大臣免許の業者のうち、主たる事務所が指定都道府県にあるものも含める（時間がかかる）")
    p.add_argument("--no-detail", action="store_true", help="詳細画面を開かず一覧の情報だけで出力する（電話番号が取れない場合あり）")
    p.add_argument("--limit", type=int, default=0, help="取得件数の上限（動作確認用）")
    p.add_argument("--delay", type=float, default=2.0, help="画面操作の間隔（秒）。サーバに負荷をかけないため 1 秒以上を推奨")
    p.add_argument("--url", default=DEFAULT_URL, help="検索画面の URL")
    p.add_argument("--pref-select", help="免許行政庁プルダウンの name 属性（自動判定できない場合）")
    p.add_argument("--disp-select", help="表示件数プルダウンの name 属性（自動判定できない場合）")
    p.add_argument("--dump-dir", help="取得した画面の HTML を保存するディレクトリ（調査用）")
    p.add_argument("--headful", action="store_true", help="ブラウザ画面を表示して実行する")
    p.add_argument("--browser-path", help="Chromium 実行ファイルのパス")
    p.add_argument("-v", "--verbose", action="store_true")
    args = p.parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(asctime)s %(message)s")

    unknown = [x for x in args.pref if x not in PREFECTURES]
    if unknown:
        p.error(f"都道府県名が不正です: {', '.join(unknown)}（「東京都」「北海道」のように指定）")

    out = Path(args.out or f"output/takken_{'_'.join(args.pref)}.csv")
    out.parent.mkdir(parents=True, exist_ok=True)
    done = load_done(out)
    if done:
        log.info("既存の %s から %d 件を取得済みとして続きから再開します", out, len(done))

    counter = [0]
    is_new = not out.exists() or out.stat().st_size == 0
    # Excel で文字化けしないよう新規作成時のみ BOM 付き UTF-8 にする（追記時に BOM を重ねない）
    with sync_playwright() as pw, out.open("a", encoding="utf-8-sig" if is_new else "utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        if is_new:
            writer.writeheader()
        launch = {"headless": not args.headful}
        if args.browser_path:
            launch["executable_path"] = args.browser_path
        browser = pw.chromium.launch(**launch)
        page = browser.new_context(locale="ja-JP").new_page()
        scraper = Scraper(page, args)
        try:
            for pref in args.pref:
                authorities = [pref] + ([MINISTER] if args.include_minister else [])
                for authority in authorities:
                    scrape_authority(scraper, authority, pref, writer, done, counter)
                    f.flush()
        finally:
            browser.close()
    log.info("完了: %d 件を %s に出力しました", counter[0], out)


if __name__ == "__main__":
    sys.exit(main())
