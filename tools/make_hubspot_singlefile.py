# -*- coding: utf-8 -*-
"""HubSpot に「1回貼るだけ」で完結するテンプレートを作る。

    python3 tools/make_hubspot_singlefile.py site/lp-image/index.html lp/イエオモイ_画像LP_HubSpot_1ファイル版.html

画像を外部参照ではなく data URI で埋め込むので、ファイルマネージャーへの
アップロードも URL の書き換えも要らない。そのままでは 6MB あって貼れないため、
写真の3枚は WebP に変換して圧縮する（実測 4.36MB → 約0.5MB。文字の劣化は
等倍では判別できない）。透過が要るロゴ・アイコンは PNG のまま。

送信先の PORTAL_ID / FORM_GUID だけは HubSpot 側の値なので、
探さなくて済むようファイル先頭にまとめて置く。
"""
import base64
import io
import re
import sys

from PIL import Image

WEBP_QUALITY = 85

HEAD = '''<!--
    templateType: page
    isAvailableForNewContent: true
-->
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
    <script>
    /* ===== ここだけ書き換えてください（申し込みの送信先） =====
       HubSpot のフォーム編集画面「共有 → 埋め込みコード」に出てくる
       portalId と formId を入れます。空のままでも表示はされますが、
       申し込みは送信されず、ボタンの上に赤字で警告が出ます。          */
    var IEOMOI_PORTAL_ID = '';
    var IEOMOI_FORM_GUID = '';
    /* ===== 書き換えるのはここまで ===== */
    </script>
    <style>
'''

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
      // フォームの name → HubSpot のプロパティ内部名。名前が違う場合はここを直す。
      var FIELD_MAP = {
        name:    'lastname',
        tel:     'phone',
        email:   'email',
        address: 'city',
        plan:    'desired_plan'
      };

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
        if (!IEOMOI_PORTAL_ID || !IEOMOI_FORM_GUID) {
          msg.textContent = '送信先が未設定です（ファイル先頭の PORTAL_ID / FORM_GUID）';
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
        fetch('https://api.hsforms.com/submissions/v3/integration/submit/'
              + IEOMOI_PORTAL_ID + '/' + IEOMOI_FORM_GUID, {
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


def to_data_uri(path):
    """写真は WebP に変換して圧縮、透過が要る小さい素材は PNG のまま。"""
    im = Image.open(path)
    if im.mode == 'RGBA' or im.size[0] < 200:
        return 'data:image/png;base64,' + base64.b64encode(open(path, 'rb').read()).decode()
    buf = io.BytesIO()
    im.convert('RGB').save(buf, 'WEBP', quality=WEBP_QUALITY, method=6)
    return 'data:image/webp;base64,' + base64.b64encode(buf.getvalue()).decode()


def main(src, dst):
    html = open(src, encoding='utf-8').read()
    css = re.search(r'<style>(.*?)</style>', html, re.S).group(1)
    body = re.search(r'<body>(.*?)</body>', html, re.S).group(1)
    body = re.sub(r'<script>.*?</script>\s*', '', body, flags=re.S)

    base = src.rsplit('/', 1)[0]
    seen = {}

    def swap(m):
        name = m.group(1)
        if name not in seen:
            seen[name] = to_data_uri(base + '/images/' + name)
        return 'src="' + seen[name] + '"'

    body = re.sub(r'src="images/([^"]+)"', swap, body)

    out = (HEAD + css.rstrip() + '\n' + GUARD + '    </style>\n'
           + '  </head>\n  <body>\n' + body.rstrip() + '\n'
           + SCRIPT + '    {{ standard_footer_includes }}\n  </body>\n</html>\n')
    open(dst, 'w', encoding='utf-8').write(out)
    print('%s  %.0f KB  画像%d枚を埋め込み' % (dst, len(out.encode()) / 1024, len(seen)))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
