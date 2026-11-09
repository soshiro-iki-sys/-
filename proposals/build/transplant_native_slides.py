"""PDFから変換した提案書（展開済みフォルダ）の指定ページを、pptxgenjsで作った元のスライド（ネイティブ）に置き換える。

使い方: python3 transplant_native_slides.py <変換後の展開フォルダ> <ネイティブpptxの展開フォルダ> <背景画像フォルダ>
  置き換えるページ：1（表紙）, 51（中扉）, 52, 56〜64（貴社向けページ）。番号は両方で同じ。
"""
import os
import re
import shutil
import sys

DST, SRC, ASSET = sys.argv[1:4]
PAGES = [1, 51, 52] + list(range(56, 65))
IMG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"


def rels_of(root, n):
    return os.path.join(root, "ppt", "slides", "_rels", f"slide{n}.xml.rels")


for n in PAGES:
    srels = open(rels_of(SRC, n), encoding="utf-8").read()
    drels = open(rels_of(DST, n), encoding="utf-8").read()
    layout_target = re.search(r'Type="%s" Target="([^"]+)"' % re.escape(LAYOUT), drels) or re.search(r'Target="([^"]+)"[^>]*Type="%s"' % re.escape(LAYOUT), drels)
    layout_target = layout_target.group(1)
    out = []
    for rid, typ, target in re.findall(r'<Relationship Id="([^"]+)" Type="([^"]+)" Target="([^"]+)"/>', srels):
        if typ == IMG:
            src = os.path.normpath(os.path.join(SRC, "ppt", "slides", target))
            name = f"native-{n}-{os.path.basename(src)}"
            shutil.copy(src, os.path.join(DST, "ppt", "media", name))
            out.append(f'<Relationship Id="{rid}" Type="{IMG}" Target="../media/{name}"/>')
        elif typ == LAYOUT:
            out.append(f'<Relationship Id="{rid}" Type="{LAYOUT}" Target="{layout_target}"/>')
        # ノートは持ち込まない
    xml = open(os.path.join(SRC, "ppt", "slides", f"slide{n}.xml"), encoding="utf-8").read()
    # 背景（ネイティブ版はマスターの背景を使っているので、スライドに直接入れる）
    if n == 51:
        bg = '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="1F2D50"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
    else:
        bgfile = "bg_cover.jpg" if n == 1 else "bg_content.jpg"
        name = f"native-bg-{bgfile}"
        shutil.copy(os.path.join(ASSET, bgfile), os.path.join(DST, "ppt", "media", name))
        out.append(f'<Relationship Id="rIdBg" Type="{IMG}" Target="../media/{name}"/>')
        bg = '<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rIdBg"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>'
    xml = re.sub(r"<p:bg>.*?</p:bg>", "", xml, flags=re.S)
    xml = re.sub(r"(<p:cSld[^>]*>)", r"\1" + bg, xml, count=1)
    # ページ番号（他のページと同じ位置・書式）
    if n != 1 and n != 51:
        num = (f'<p:sp><p:nvSpPr><p:cNvPr id="9999" name="PageNumber"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
               f'<p:spPr><a:xfrm><a:off x="9580000" y="6680000"/><a:ext cx="300000" cy="180000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
               f'<p:txBody><a:bodyPr wrap="none" lIns="0" rIns="0" tIns="0" bIns="0"/><a:p><a:pPr algn="r"/><a:r><a:rPr lang="en-US" sz="1110"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
               f'<a:latin typeface="Calibri"/></a:rPr><a:t>{n}</a:t></a:r></a:p></p:txBody></p:sp>')
        xml = xml.replace("</p:spTree>", num + "</p:spTree>", 1)
    open(os.path.join(DST, "ppt", "slides", f"slide{n}.xml"), "w", encoding="utf-8").write(xml)
    head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    open(rels_of(DST, n), "w", encoding="utf-8").write(head + "".join(out) + "</Relationships>")
print("transplanted:", PAGES)
