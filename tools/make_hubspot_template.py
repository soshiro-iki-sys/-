# -*- coding: utf-8 -*-
"""分解版LPから、HubSpot のコード化テンプレート（HubL）を生成する。

    python3 tools/make_hubspot_template.py site/lp-image/index.html lp/イエオモイ_画像LP_HubSpot.html

デザインマネージャーにそのまま貼れる形にする。素のHTMLとの違いは3点。
  1. head/body は HubSpot 側の雛形に合わせ、standard_header_includes /
     standard_footer_includes を残す（トラッキングとテーマCSSの読み込みに必要）
  2. 画像は相対パスが効かないので、先頭の IMG_BASE 1か所だけ書き換えれば
     全部差し替わるようにする
  3. 送信先を HubSpot のフォーム送信API（JSON形式）に組み替える

LP本体の CSS と markup は分解版から機械的に取り込むので、LP を直したら
これを再実行すれば追従できる。
"""
import re
import sys

HEAD = '''<!--
    templateType: page
    isAvailableForNewContent: true
-->
{# 画像の置き場所。HubSpot のファイルマネージャーにアップロードして、
   末尾スラッシュなしのフォルダURLをここに入れる。ここ1か所で全画像に効く。 #}
{% set IMG_BASE = "https://xxxxxxxx.fs1.hubspotusercontent-na1.net/hubfs/xxxxxxxx/ieomoi" %}
<!doctype html>
<html lang="ja">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    {% if content.html_title %}<title>{{ content.html_title }}</title>{% endif %}
    <meta name="description" content="{{ content.meta_description }}">
    {% if brand_settings.primaryFavicon.src %}
      <link rel="shortcut icon" href="{{ brand_settings.primaryFavicon.src }}" />
    {% endif %}
    {{ standard_header_includes }}
    <style>
'''

# テーマCSSに負けないための最低限の打ち消し。standard_header_includes より
# 後ろに置いているので、同じ詳細度ならこちらが勝つ。
GUARD = '''
  /* --- HubSpot のテーマCSSがフォーム部品を上書きしてくる場合の打ち消し --- */
  .lp img { max-width: 100%; }
  .ov-form .fld, .ov-form .fbtn {
    margin: 0; min-height: 0; text-transform: none; box-shadow: none;
  }
  .ov-form .fbtn { text-decoration: none; }
  .lp a, .lp a:hover { text-decoration: none; }
'''

SCRIPT = '''    <script>
    (function () {
      // ▼▼▼ HubSpot の設定 ▼▼▼
      // フォーム編集画面の「共有 → 埋め込みコード」に出てくる
      // portalId と formId をそのまま入れる。
      var PORTAL_ID = '00000000';
      var FORM_GUID = '00000000-0000-0000-0000-000000000000';
      // このフォームの name → HubSpot のプロパティ内部名
      var FIELD_MAP = {
        name:    'lastname',
        tel:     'phone',
        email:   'email',
        address: 'city',
        plan:    'desired_plan'   // 事前にカスタムプロパティを作成しておく
      };
      // ▲▲▲

      var ENDPOINT = 'https://api.hsforms.com/submissions/v3/integration/submit/'
                     + PORTAL_ID + '/' + FORM_GUID;
      var form = document.getElementById('lp-form');
      if (!form) { return; }
      var msg = form.querySelector('.fmsg');
      var done = form.querySelector('.fdone');
      var btn = form.querySelector('.fbtn');
      var plan = form.querySelector('[name="plan"]');

      // プラン別ボタンから来たら、そのプランを選択済みにする
      Array.prototype.forEach.call(document.querySelectorAll('[data-plan]'), function (el) {
        el.addEventListener('click', function () { plan.value = el.getAttribute('data-plan'); });
      });
      var q = (location.search.match(/[?&]plan=([^&]+)/) || [])[1];
      if (q) { plan.value = decodeURIComponent(q); }

      form.addEventListener('submit', function (e) {
        e.preventDefault();
        msg.textContent = '';
        msg.classList.remove('is-error');
        if (!form.checkValidity()) { form.reportValidity(); return; }
        if (PORTAL_ID === '00000000') {
          msg.textContent = 'PORTAL_ID と FORM_GUID が未設定です';
          msg.classList.add('is-error');
          return;
        }
        var fields = [];
        new FormData(form).forEach(function (v, k) {
          if (FIELD_MAP[k]) {
            fields.push({ objectTypeId: '0-1', name: FIELD_MAP[k], value: v });
          }
        });
        var label = btn.textContent;
        btn.disabled = true;
        btn.textContent = '送信中…';
        fetch(ENDPOINT, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            fields: fields,
            context: { pageUri: location.href, pageName: document.title }
          })
        }).then(function (r) {
          if (!r.ok) { throw new Error(r.status); }
          done.hidden = false;
          Array.prototype.forEach.call(form.querySelectorAll('.fl, .fbtn, .fmsg'), function (el) {
            el.style.display = 'none';
          });
        }).catch(function () {
          msg.textContent = '送信に失敗しました。時間をおいてお試しください。';
          msg.classList.add('is-error');
          btn.disabled = false;
          btn.textContent = label;
        });
      });
    })();
    </script>
'''


def main(src, dst):
    html = open(src, encoding='utf-8').read()

    css = re.search(r'<style>(.*?)</style>', html, re.S).group(1)
    body = re.search(r'<body>(.*?)</body>', html, re.S).group(1)
    body = re.sub(r'<script>.*?</script>\s*', '', body, flags=re.S)
    # 相対パスの画像を HubL 変数に差し替える
    body, n = re.subn(r'src="images/([^"]+)"', r'src="{{ IMG_BASE }}/\1"', body)

    out = (HEAD + css.rstrip() + '\n' + GUARD + '    </style>\n'
           + '  </head>\n  <body>\n' + body.rstrip() + '\n'
           + SCRIPT + '    {{ standard_footer_includes }}\n  </body>\n</html>\n')
    open(dst, 'w', encoding='utf-8').write(out)
    print('%s  %.1f KB  画像参照 %d 箇所を IMG_BASE に置換' % (dst, len(out.encode()) / 1024, n))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
