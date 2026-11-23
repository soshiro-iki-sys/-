# -*- coding: utf-8 -*-
"""撮影指示書の「撮影完成イメージ」を ChatGPT などで生成するためのプロンプトを書き出す。

photobook_cuts.py の仕様7項目からプロンプトを組み立てるので、カットを直せば
プロンプトも自動で追従する。

実行： python3.13 slides/build/make_prompts.py
出力： docs/画像生成プロンプト_撮影指示書.md
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from photobook_cuts import CUTS  # noqa: E402

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   '..', '..', 'docs', '画像生成プロンプト_撮影指示書.md')

# 全カット共通の指定（毎回つける）
COMMON = (
    '写真。日本、現代。自然光中心、落ち着いた色調、ドキュメンタリー調。'
    '横位置・アスペクト比3:2。'
    '文字・ロゴ・看板・書類の文字・車のナンバー・表札・部屋番号は一切写さない。'
    '実在の人物に似せない。過度に演出された笑顔にしない。'
)

HEAD = """# 画像生成プロンプト（撮影指示書の「撮影完成イメージ」用）

撮影指示書 `2026撮影指示書_マルユウハウジー_お片付けサイト.pptx` の写真枠に入れる
**参考画像**を ChatGPT（画像生成）で作るためのプロンプト集です。全52カット。

> 実際の撮影写真ではなく、**カメラマンに構図を伝えるための完成イメージ**です。
> 見本の指示書も同じ使い方をしていて、写真の下に「AI生成・撮影参考」と明記されています。

---

## 使い方（3ステップ）

### 1. ChatGPT で画像を作る

下のプロンプトをそのまま貼り付けてください。**1カットにつき1枚**で十分です。
うまくいかないときは、末尾の共通指定（文字を写さない等）はそのまま残して、
前半の場面描写だけ言い換えてみてください。

### 2. ファイル名を「カット番号」にして保存する

```
A-1.png   A-2.png   A-3.png   A-4.png
B-1.png   B-2.png   ...
C-1.png   C-2.png   C-4.png   ...
```

- **ファイル名＝カット番号**（半角ハイフン）。拡張子は `.png` でも `.jpg` でも可
- 全部そろっていなくて構いません。**ある分だけ差し込み、無いカットは枠のまま**にします
- サイズは**横長 3:2**（例 1536×1024）が理想。多少違っても自動で中央トリミングします

### 3. 私に渡す

まとめて送っていただければ、`assets/photobook/` に置いてビルドし直し、
**全ページに自動で差し込みます。**

---

## 共通指定（全プロンプトの末尾についています）

```
{common}
```

---

## プロンプト一覧（全{n}カット）

"""


LABELS = ['場面', '人物', '服装', '小物', '動作', '構図', '表情']

# 群ごとの補足（プロンプトだけでは伝わらない作り方）
GROUP_NOTE = {
    'B': 'B-5 は3枚を1組で使うので、1枚作ったあと「同じ画角・同じ明るさで、'
         '次は◯◯の場面」と続けて指示すると色味がそろいます。',
    'F': 'F-1 と F-2 は**同じ部屋**でなければ意味がありません。'
         'F-1 を作ってから、その画像に対して「この部屋から家財をすべて取り除いて、'
         'カメラ位置はそのまま」と指示して F-2 を作ってください。別々に作ると別の部屋になります。',
    'J': '6点とも背景・光・影の向きをそろえます。J-1 を作ったあと、'
         '「同じ背景・同じ光で、次は洗濯機」と続けて指示するのが確実です。',
}


def compose(c):
    """指示書の仕様7項目を、そのままラベル付きプロンプトにする。

    値が「—」で始まるもの（指定なし）は行ごと落とす。
    人物が「なし」で始まるものは「人物を入れない」に言い換える。
    """
    lines = []
    for label, val in zip(LABELS, c['spec']):
        if val.startswith('—'):
            continue
        if label == '人物' and val.startswith('なし'):
            val = '人物を入れない（無人）'
        if label == '小物' and val.startswith('なし'):
            continue
        lines.append('%s：%s' % (label, val))
    lines.append('仕上がり：' + COMMON)
    return '\n'.join(lines)


def main():
    out = os.path.normpath(OUT)
    with open(out, 'w', encoding='utf-8') as f:
        f.write(HEAD.format(common=COMMON, n=len(CUTS)))
        group = None
        for c in CUTS:
            g = c['no'].split('-')[0]
            if g != group:
                group = g
                f.write('\n### %s群\n\n' % g)
                if g in GROUP_NOTE:
                    f.write('> %s\n\n' % GROUP_NOTE[g])
            f.write('#### %s　%s\n\n' % (c['no'], c['name']))
            f.write('保存名：`%s.png`\n\n' % c['no'])
            f.write('```\n%s\n```\n\n' % compose(c))
    print(out, len(CUTS), 'prompts')


if __name__ == '__main__':
    main()
