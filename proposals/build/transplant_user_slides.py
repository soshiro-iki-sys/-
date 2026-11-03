"""ユーザー作成の表紙（1枚目）と会社紹介（2枚目）を、生成したデッキに差し替える。
usage: python3 transplant_user_slides.py <deck.pptx>
"""
import sys, os, re, shutil, zipfile, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
A = os.path.join(HERE, "assets", "user_slides")
SLIDES = {
    1: ("cover.xml", {"rId3": "u_cover_bg.jpg", "rId4": "u_cover_2.png", "rId5": "u_cover_3.png"}),
    2: ("company.xml", {"rId3": "u_company_4.png", "rId4": "u_company_5.jpeg", "rId5": "u_company_6.jpeg", "rId6": "u_company_7.jpeg"}),
}
IMG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

deck = sys.argv[1]
tmp = tempfile.mkdtemp()
with zipfile.ZipFile(deck) as z:
    z.extractall(tmp)
os.makedirs(os.path.join(tmp, "ppt", "media"), exist_ok=True)
orphans = set()
for n, (xml, imgs) in SLIDES.items():
    shutil.copy(os.path.join(A, xml), os.path.join(tmp, "ppt", "slides", f"slide{n}.xml"))
    rp = os.path.join(tmp, "ppt", "slides", "_rels", f"slide{n}.xml.rels")
    rels = open(rp, encoding="utf-8").read()
    # レイアウトとノートだけ残し、リレーションIDを振り直して画像を追加する
    keep = re.findall(r'<Relationship [^>]*/>', rels)
    for r in keep:
        if IMG in r:
            old_media = os.path.normpath(os.path.join(tmp, "ppt", "slides", re.search(r'Target="([^"]+)"', r).group(1)))
            orphans.add(old_media)
    lay = next(r for r in keep if "slideLayout" in r)
    notes = [r for r in keep if "notesSlide" in r]
    lay = re.sub(r'Id="rId\d+"', 'Id="rId1"', lay)
    notes = [re.sub(r'Id="rId\d+"', 'Id="rId2"', r) for r in notes]
    add = lay + "".join(notes)
    for rid, f in imgs.items():
        shutil.copy(os.path.join(A, f), os.path.join(tmp, "ppt", "media", f))
        add += f'<Relationship Id="{rid}" Type="{IMG}" Target="../media/{f}"/>'
    rels = re.sub(r"<Relationships([^>]*)>.*</Relationships>", lambda m: f"<Relationships{m.group(1)}>" + add + "</Relationships>", rels, flags=re.S)
    open(rp, "w", encoding="utf-8").write(rels)
# 差し替えで参照されなくなった画像を削除（他スライドから参照されていれば残す）
allrels = ""
for root, _, files in os.walk(os.path.join(tmp, "ppt")):
    for f in files:
        if f.endswith(".rels"):
            allrels += open(os.path.join(root, f), encoding="utf-8").read()
for m in orphans:
    if os.path.exists(m) and ("media/" + os.path.basename(m) + '"') not in allrels:
        os.remove(m)
ct = os.path.join(tmp, "[Content_Types].xml")
c = open(ct, encoding="utf-8").read()
for ext, mt in [("jpg", "image/jpeg"), ("jpeg", "image/jpeg"), ("png", "image/png")]:
    if f'Extension="{ext}"' not in c:
        c = c.replace("<Default ", f'<Default Extension="{ext}" ContentType="{mt}"/><Default ', 1)
open(ct, "w", encoding="utf-8").write(c)
os.remove(deck)
with zipfile.ZipFile(deck, "w", zipfile.ZIP_DEFLATED) as z:
    ct_path = os.path.join(tmp, "[Content_Types].xml")
    z.write(ct_path, "[Content_Types].xml")
    for root, _, files in os.walk(tmp):
        for f in files:
            p = os.path.join(root, f)
            arc = os.path.relpath(p, tmp)
            if arc != "[Content_Types].xml":
                z.write(p, arc)
shutil.rmtree(tmp)
print("transplanted slides 1-2 into", deck)
