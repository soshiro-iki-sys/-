"""宅建業者検索システムを模した簡易サーバ（スクレイパーの動作確認用）。

検索・ページ送り・詳細表示をすべてフォーム POST と JavaScript で行い、
ブラウザの「戻る」で一覧に戻れない（再送信が必要になる）挙動も再現する。
    python tests/mock_server.py 8765
"""
import sys
import urllib.parse
from http.server import BaseHTTPRequestHandler, HTTPServer

PREFS = ["北海道", "東京都", "大阪府"]
COMPANIES = (
    [("東京都", f"東京都知事(1)第{10000 + i}号", f"株式会社テスト不動産{i}", f"東京都千代田区丸の内{i}-1", f"03-1234-{1000 + i}") for i in range(23)]
    + [("国土交通大臣", "国土交通大臣(3)第5001号", "大臣免許東京株式会社", "東京都港区芝1-1", "03-9999-0001"),
       ("国土交通大臣", "国土交通大臣(2)第5002号", "大臣免許大阪株式会社", "大阪府大阪市北区梅田1-1", "06-9999-0002")]
)


def page(body):
    return f"<html><head><meta charset='utf-8'></head><body>{body}</body></html>".encode()


class H(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def send(self, html):
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(html)

    def do_GET(self):
        opts = "".join(f"<option value='{i}'>{p}</option>" for i, p in enumerate(["国土交通大臣"] + PREFS))
        self.send(page(f"""
<form method=post action=/TAKKEN/takkenKensaku.do>
<table><tr><th>免許行政庁</th><td><select name=licenseNoKbn><option value=''>選択してください</option>{opts}</select></td></tr>
<tr><th>商号又は名称</th><td><input name=comNameKj></td></tr></table>
表示件数 <select name=dispCount><option>10</option><option>20</option></select>
<input type=hidden name=CMD value=search><input type=hidden name=pageNo value=1>
<input type=submit value='検索する'> <input type=reset value='条件クリア'>
</form>"""))

    def do_POST(self):
        n = int(self.headers["Content-Length"])
        q = {k: v[0] for k, v in urllib.parse.parse_qs(self.rfile.read(n).decode()).items()}
        auth = (["国土交通大臣"] + PREFS)[int(q["licenseNoKbn"])]
        rows = [c for c in COMPANIES if c[0] == auth]
        if q.get("CMD") == "detail":
            c = rows[int(q["sv"])]
            hp = "<a href='https://example.com/'>https://example.com/</a>" if int(q["sv"]) % 2 == 0 else ""
            self.send(page(f"""<h1>宅地建物取引業者 詳細</h1><table>
<tr><th>免許証番号</th><td>{c[1]}</td></tr><tr><th>商号又は名称</th><td>{c[2]}</td></tr>
<tr><th>フリガナ</th><td>テスト</td></tr><tr><th>代表者名</th><td>山田 太郎</td></tr>
<tr><th>主たる事務所の所在地</th><td>{c[3]}</td></tr><tr><th>電話番号</th><td>{c[4]}</td></tr>
<tr><th>ホームページ</th><td>{hp}</td></tr></table>"""))
            return
        per = int(q["dispCount"]); pno = int(q.get("pageNo", 1))
        chunk = rows[(pno - 1) * per: pno * per]
        trs = "".join(
            f"<tr><td>{(pno - 1) * per + i + 1}</td><td>{c[0]}</td><td>{c[1]}</td>"
            f"<td><a href='#' onclick=\"js_Detail({(pno - 1) * per + i});return false;\">{c[2]}</a></td><td>山田 太郎</td></tr>"
            for i, c in enumerate(chunk))
        nxt = f"<a href='#' onclick=\"js_Page({pno + 1});return false;\">次へ</a>" if pno * per < len(rows) else ""
        hidden = "".join(f"<input type=hidden name={k} value='{v}'>" for k, v in q.items() if k not in ("CMD", "pageNo", "sv"))
        self.send(page(f"""<p>{len(rows)}件</p>
<form id=f method=post action=/TAKKEN/takkenKensaku.do>{hidden}<input type=hidden name=CMD><input type=hidden name=pageNo value={pno}><input type=hidden name=sv></form>
<script>function js_Detail(i){{f.CMD.value='detail';f.sv.value=i;f.submit();}}
function js_Page(p){{f.CMD.value='search';f.pageNo.value=p;f.submit();}}</script>
<table border=1><tr><th>番号</th><th>免許行政庁</th><th>免許証番号</th><th>商号又は名称</th><th>代表者名</th></tr>{trs}</table>{nxt}"""))


if __name__ == "__main__":
    HTTPServer(("127.0.0.1", int(sys.argv[1]) if len(sys.argv) > 1 else 8765), H).serve_forever()
