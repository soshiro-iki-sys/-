"""PDFから変換したスライドで、表（行×列に並んだテキストと罫線）をグループ化する。

PowerPointの「表」には戻せないため、表全体を1つのグループにまとめ、まとめて移動・拡大縮小できるようにする。
セルの文字は、グループの中をクリックすればそのまま編集できる。

使い方: python3 group_tables.py <展開したpptxのフォルダ>
"""
import glob
import re
import sys
from statistics import median

from lxml import etree

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main", "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
P = "{%s}" % NS["p"]
A = "{%s}" % NS["a"]
IN = 914400


def geom(el):
    off = el.find("./p:spPr/a:xfrm/a:off", NS)
    if off is None:
        off = el.find("./p:grpSpPr/a:xfrm/a:off", NS)
    ext = el.find("./p:spPr/a:xfrm/a:ext", NS)
    if ext is None:
        ext = el.find("./p:grpSpPr/a:xfrm/a:ext", NS)
    if off is None or ext is None:
        return None
    x, y = int(off.get("x")), int(off.get("y"))
    return x, y, x + int(ext.get("cx")), y + int(ext.get("cy"))


def text_of(sp):
    return "".join(t.text or "" for t in sp.iter(A + "t")).strip()


def find_tables(texts):
    """texts: [(sp, x0, y0, x1, y1)] -> 表の範囲のリスト"""
    items = [t for t in texts if (t[4] - t[2]) < 0.45 * IN]
    if not items:
        return []
    items.sort(key=lambda t: (t[2] + t[4]) / 2)
    hs = median([t[4] - t[2] for t in items])
    rows = []
    for t in items:
        cy = (t[2] + t[4]) / 2
        if rows and abs(cy - rows[-1]["cy"]) < 0.45 * hs:
            rows[-1]["items"].append(t)
        else:
            rows.append({"cy": cy, "items": [t]})
    for r in rows:
        r["items"].sort(key=lambda t: t[1])

    def cols(r):
        return [(t[1], (t[1] + t[3]) / 2, t[3]) for t in r["items"]]

    def match(r1, r2):
        tol = 0.12 * IN
        n = 0
        for a in cols(r1):
            if any(abs(a[0] - b[0]) < tol or abs(a[1] - b[1]) < tol or abs(a[2] - b[2]) < tol for b in cols(r2)):
                n += 1
        return n

    tables = []
    run = []
    gap_max = max(3.2 * hs, 1.2 * IN)
    for r in rows:
        if len(r["items"]) < 3:
            continue  # 1〜2個だけの行（記号や別の図）は表の切れ目にしない
        if run and (r["cy"] - run[-1]["cy"] > gap_max or match(run[-1], r) < 3):
            if len(run) >= 3:
                tables.append(run)
            run = []
        run.append(r)
    if len(run) >= 3:
        tables.append(run)
    boxes = []
    tol = 0.12 * IN
    for run in tables:
        its = [t for r in run for t in r["items"]]
        # 3行以上で同じ列位置に並ぶものだけを表の一部とみなす
        keep = []
        for t in its:
            c = (t[1], (t[1] + t[3]) / 2, t[3])
            rows_hit = sum(1 for r in run if any(abs(c[0] - u[1]) < tol or abs(c[1] - (u[1] + u[3]) / 2) < tol or abs(c[2] - u[3]) < tol for u in r["items"]))
            if rows_hit >= 3:
                keep.append(t)
        if len(keep) >= 9:
            boxes.append((min(t[1] for t in keep), min(t[2] for t in keep), max(t[3] for t in keep), max(t[4] for t in keep)))
    return boxes


def make_group(tree_el, members, gid):
    xs = [geom(m) for m in members]
    x0, y0 = min(g[0] for g in xs), min(g[1] for g in xs)
    x1, y1 = max(g[2] for g in xs), max(g[3] for g in xs)
    grp = etree.Element(P + "grpSp")
    nv = etree.SubElement(grp, P + "nvGrpSpPr")
    c = etree.SubElement(nv, P + "cNvPr")
    c.set("id", str(gid))
    c.set("name", f"表グループ {gid}")
    etree.SubElement(nv, P + "cNvGrpSpPr")
    etree.SubElement(nv, P + "nvPr")
    gp = etree.SubElement(grp, P + "grpSpPr")
    xf = etree.SubElement(gp, A + "xfrm")
    for tag, a, b in (("off", x0, y0), ("ext", x1 - x0, y1 - y0), ("chOff", x0, y0), ("chExt", x1 - x0, y1 - y0)):
        e = etree.SubElement(xf, A + tag)
        if tag in ("off", "chOff"):
            e.set("x", str(a))
            e.set("y", str(b))
        else:
            e.set("cx", str(a))
            e.set("cy", str(b))
    last = members[-1]
    last.addnext(grp)
    for m in members:
        grp.append(m)


def process(path):
    tree = etree.parse(path)
    sptree = tree.getroot().find(".//p:cSld/p:spTree", NS)
    kids = [k for k in sptree if k.tag in (P + "sp", P + "cxnSp", P + "pic")]
    texts = []
    for k in kids:
        if k.tag == P + "sp" and text_of(k):
            g = geom(k)
            if g:
                texts.append((k, *g))
    tables = find_tables(texts)
    n = 0
    gid = 90000
    for (x0, y0, x1, y1) in tables:
        pad = 0.18 * IN
        X0, Y0, X1, Y1 = x0 - pad, y0 - pad, x1 + pad, y1 + pad
        members = []
        for k in kids:
            g = geom(k)
            # 文字のボックスだけをまとめる（背景の図形は重なり順が崩れないよう、そのまま残す）
            if g and g[0] >= X0 and g[1] >= Y0 and g[2] <= X1 and g[3] <= Y1 and k.getparent() is sptree and k.tag == P + "sp" and text_of(k):
                members.append(k)
        if len(members) >= 9:
            gid += 1
            make_group(sptree, members, gid)
            n += 1
    tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
    return n


if __name__ == "__main__":
    total = 0
    hit = []
    for f in sorted(glob.glob(sys.argv[1] + "/ppt/slides/slide*.xml"), key=lambda s: int(re.findall(r"(\d+)\.xml", s)[0])):
        n = process(f)
        if n:
            hit.append((int(re.findall(r"(\d+)\.xml", f)[0]), n))
        total += n
    print("table groups:", total, hit)
