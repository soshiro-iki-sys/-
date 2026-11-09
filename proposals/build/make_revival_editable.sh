#!/bin/bash
# REVIVAL LINK様ご提案書を「全ページ文字まで編集可能」な形で出力する手順
# 使い方: make_revival_editable.sh <企画書PDF> <build_revival.jsの出力をPDF化したもの> <作業フォルダ> <出力pptx>
# 1) 企画書PDFと貴社向けページのPDFを、提案書の順番に並べて1つのPDFにする
# 2) LibreOfficeのPDF取り込み（impress_pdf_import）でpptxに変換（文字はテキストボックスとして復元される）
# 3) フォント名・言語設定を整える
set -e
K=$1; R=$2; W=$3; OUT=$4
mkdir -p "$W/parts"; rm -f "$W"/parts/*.pdf
i=0; add(){ i=$((i+1)); pdfseparate -f $2 -l $2 "$1" "$W/parts/$(printf %03d $i).pdf"; }
add "$R" 1
for p in 2 3 4 5 $(seq 7 19) $(seq 45 76); do add "$K" $p; done
add "$R" 51; add "$R" 52
for p in 24 25 26; do add "$K" $p; done
for p in $(seq 56 64); do add "$R" $p; done
for p in $(seq 77 96); do add "$K" $p; done
pdfunite "$W"/parts/*.pdf "$W/all.pdf"
python3 "$PPTX_SKILL/scripts/office/soffice.py" --headless --infilter="impress_pdf_import" --convert-to pptx --outdir "$W" "$W/all.pdf"
rm -rf "$W/x"; mkdir "$W/x"; (cd "$W/x" && unzip -q ../all.pptx)
python3 - "$W/x" <<'PY'
import glob,sys
M={'WenQuanYi Zen Hei':'Meiryo','Inter-Regular':'Meiryo','Inter-Bold':'Meiryo','MeiryoUI':'Meiryo UI','YuGothic':'Yu Gothic','BIZ-UDPGothic':'BIZ UDPGothic','BIZUDPGothic':'BIZ UDPGothic','MS-Gothic':'MS Gothic','Times New Roman':'Meiryo'}
for f in glob.glob(sys.argv[1]+'/ppt/slides/slide*.xml'):
    x=open(f,encoding='utf-8').read()
    if f.endswith('/slide1.xml'): x=x.replace('typeface="WenQuanYi Zen Hei"','typeface="Yu Mincho"')
    for a,b in M.items(): x=x.replace(f'typeface="{a}"',f'typeface="{b}"')
    x=x.replace('lang="zh-CN"','lang="ja-JP"')
    open(f,'w',encoding='utf-8').write(x)
PY
rm -f "$OUT"; (cd "$W/x" && zip -qXrD "$OUT" .)
# 4) 1行ずつのボックスを段落単位にまとめる／貴社向けページは元のスライドに差し替え／表はグループ化
D="$(dirname "$0")"
python3 "$D/merge_textboxes.py" "$W/x"
# NATIVE: build_revival.js で出力したpptxの展開フォルダ（環境変数で指定）
python3 "$D/transplant_native_slides.py" "$W/x" "$NATIVE" "$D/assets/kikaku"
python3 "$D/group_tables.py" "$W/x"
rm -f "$OUT"; (cd "$W/x" && zip -qXrD "$OUT" .)
echo "written: $OUT"
