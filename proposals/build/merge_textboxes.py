"""PDFから変換したpptxの「1行ごとに分かれたテキストボックス」を、段落単位のテキストボックスにまとめる。

使い方: python3 merge_textboxes.py <展開したpptxのフォルダ>
  - 同じ行で文字装飾ごとに分かれたボックスを横方向に結合
  - 上下に続く行を1つのボックスにまとめ、折り返し行は1つの段落につなぐ（行間は元の位置を保つ固定値）
  - 箇条書き（・■①など）や、句点で終わる行のあとは新しい段落にする
"""
import glob
import re
import sys
from copy import deepcopy
from statistics import median

from lxml import etree

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
A = "{%s}" % NS["a"]
EMU_IN = 914400
EMU_PT = 12700
MARKERS = tuple("・■◼□●○◆◇▶►➢★☆※✕×✓①②③④⑤⑥⑦⑧⑨⑩➊➋➌➍➎「【(（") + ("STEP", "POINT")
ENDERS = tuple("。．！？!?」』）)】")
NUM_START = re.compile(r"^(\d+[\.．、)）]|[０-９]+[．、])")


class Line:
    def __init__(self, sp):
        self.sp = sp
        off = sp.find(".//a:xfrm/a:off", NS)
        ext = sp.find(".//a:xfrm/a:ext", NS)
        self.x, self.y = int(off.get("x")), int(off.get("y"))
        self.w, self.h = int(ext.get("cx")), int(ext.get("cy"))
        self.runs = sp.findall(".//a:p/a:r", NS)
        szs = [int(r.find("a:rPr", NS).get("sz", "1800")) for r in self.runs if r.find("a:rPr", NS) is not None]
        self.sz = max(szs) if szs else 1800
        self.text = "".join((r.findtext("a:t", namespaces=NS) or "") for r in self.runs)

    @property
    def r(self):
        return self.x + self.w

    @property
    def b(self):
        return self.y + self.h

    @property
    def cx(self):
        return self.x + self.w / 2

    @property
    def cy(self):
        return self.y + self.h / 2

    @property
    def fh(self):  # font height in EMU
        return self.sz / 100 * EMU_PT


def is_textbox(sp):
    c = sp.find("p:nvSpPr/p:cNvSpPr", NS)
    if c is None or c.get("txBox") != "1":
        return False
    ps = sp.findall(".//p:txBody/a:p", NS)
    return len(ps) == 1 and sp.find(".//a:p/a:r", NS) is not None


def h_merge(lines):
    changed = True
    while changed:
        changed = False
        lines.sort(key=lambda l: (l.y, l.x))
        for i, a in enumerate(lines):
            for b in lines[i + 1:]:
                if b.y > a.b:
                    break
                same_row = abs(a.cy - b.cy) < 0.35 * max(a.h, b.h) or (b.y >= a.y - 0.1 * a.h and b.b <= a.b + 0.1 * a.h)
                gap = b.x - a.r
                fh = max(a.fh, b.fh)
                if same_row and -0.15 * fh <= gap < 0.9 * fh:
                    p = a.sp.find(".//a:p", NS)
                    end = p.find("a:endParaRPr", NS)
                    if gap > 0.4 * fh:
                        sp_run = deepcopy(a.runs[-1])
                        sp_run.find("a:t", NS).text = " "
                        (end.addprevious(sp_run) if end is not None else p.append(sp_run))
                    for r in b.runs:
                        (end.addprevious(r) if end is not None else p.append(r))
                    nb = max(a.b, b.b)
                    a.y = min(a.y, b.y)
                    a.h = nb - a.y
                    a.w = max(a.r, b.r) - a.x
                    a.runs = a.sp.findall(".//a:p/a:r", NS)
                    a.sz = max(a.sz, b.sz)
                    a.text += b.text
                    b.sp.getparent().remove(b.sp)
                    lines.remove(b)
                    write_geom(a)
                    changed = True
                    break
            if changed:
                break
    return lines


def write_geom(l):
    l.sp.find(".//a:xfrm/a:off", NS).set("x", str(int(l.x)))
    l.sp.find(".//a:xfrm/a:off", NS).set("y", str(int(l.y)))
    l.sp.find(".//a:xfrm/a:ext", NS).set("cx", str(int(l.w)))
    l.sp.find(".//a:xfrm/a:ext", NS).set("cy", str(int(l.h)))


def aligned(p, l):
    tol = 0.08 * EMU_IN
    if abs(l.x - p.x) < 0.3 * EMU_IN:
        return True
    if abs(l.cx - p.cx) < tol:
        return True
    if abs(l.r - p.r) < tol:
        return True
    return False


def v_blocks(lines):
    lines.sort(key=lambda l: (l.y, l.x))
    blocks = []
    for l in lines:
        best = None
        for bk in blocks:
            p = bk[-1]
            if not (0.96 <= l.sz / p.sz <= 1.04):
                continue
            gap = l.y - p.b
            if not (-0.15 * p.h <= gap <= 0.8 * p.h):
                continue
            if min(l.r, p.r) - max(l.x, p.x) <= 0 and abs(l.cx - p.cx) > 0.08 * EMU_IN:
                continue
            if not aligned(p, l):
                continue
            # 間に別のボックスが挟まっていないか
            if any(o is not l and o not in bk and p.b <= o.cy <= l.y + 0.1 * l.h and min(o.r, l.r) > max(o.x, l.x) for o in lines):
                continue
            if best is None or gap < best[1]:
                best = (bk, gap)
        if best:
            best[0].append(l)
        else:
            blocks.append([l])
    return blocks


def starts_new(prev, l, maxr):
    t = l.text.strip()
    if not t:
        return True
    if t.startswith(MARKERS) or NUM_START.match(t):
        return True
    pt = prev.text.strip()
    if pt.endswith(ENDERS):
        return True
    # 前の行が右端まで届いていなければ、そこで段落が終わっている
    if prev.r < maxr - 1.6 * prev.fh:
        return True
    return False


def build_block(bk):
    first = bk[0]
    minx = min(l.x for l in bk)
    maxr = max(l.r for l in bk)
    miny = first.y
    maxb = max(l.b for l in bk)
    tol = 0.06 * EMU_IN
    if all(abs(l.x - first.x) < tol for l in bk):
        algn = "l"
    elif all(abs(l.cx - first.cx) < tol for l in bk):
        algn = "ctr"
    elif all(abs(l.r - first.r) < tol for l in bk):
        algn = "r"
    else:
        algn = "l"
    pitches = [bk[i + 1].y - bk[i].y for i in range(len(bk) - 1)]
    pitch_pt = median(pitches) / EMU_PT if pitches else None

    # 段落ごとに行をまとめる
    paras = [[first]]
    for prev, l in zip(bk, bk[1:]):
        if algn == "l" and starts_new(prev, l, maxr):
            paras.append([l])
        elif algn != "l":
            # 中央・右揃えは折り返し位置の判定が難しいので、行ごとに段落を分ける
            paras.append([l])
        else:
            paras[-1].append(l)

    txBody = first.sp.find("p:txBody", NS)
    body = txBody.find("a:bodyPr", NS)
    body.set("wrap", "square")
    for p in txBody.findall("a:p", NS):
        txBody.remove(p)
    for pl in paras:
        p = etree.SubElement(txBody, A + "p")
        pPr = etree.SubElement(p, A + "pPr")
        pPr.set("algn", algn)
        # ぶら下げインデント（2行目以降が字下げされている場合）
        if len(pl) > 1 and algn == "l":
            d = median([l.x - pl[0].x for l in pl[1:]])
            if d > 0.04 * EMU_IN:
                pPr.set("marL", str(int(d + pl[0].x - minx)))
                pPr.set("indent", str(int(-d)))
            elif pl[0].x - minx > 0.04 * EMU_IN:
                pPr.set("marL", str(int(pl[0].x - minx)))
        elif pl[0].x - minx > 0.04 * EMU_IN and algn == "l":
            pPr.set("marL", str(int(pl[0].x - minx)))
        if pitch_pt:
            ln = etree.SubElement(pPr, A + "lnSpc")
            sp = etree.SubElement(ln, A + "spcPts")
            sp.set("val", str(int(round(pitch_pt * 100))))
        last_rpr = None
        for l in pl:
            for r in l.runs:
                p.append(r)
                last_rpr = r.find("a:rPr", NS)
        if last_rpr is not None:
            end = deepcopy(last_rpr)
            end.tag = A + "endParaRPr"
            p.append(end)
    # 位置・大きさ（折り返しの余裕を少し持たせる）
    width = maxr - minx
    slack = int(width * 0.04 + first.fh * 0.5)
    x = minx - (slack // 2 if algn == "ctr" else (slack if algn == "r" else 0))
    first.x, first.y, first.w, first.h = x, miny, width + slack, maxb - miny
    write_geom(first)
    for l in bk[1:]:
        l.sp.getparent().remove(l.sp)


def process(path):
    tree = etree.parse(path)
    root = tree.getroot()
    sps = [sp for sp in root.iter("{%s}sp" % NS["p"]) if is_textbox(sp)]
    before = len(sps)
    lines = [Line(sp) for sp in sps]
    lines = h_merge(lines)
    blocks = v_blocks(lines)
    for bk in blocks:
        if len(bk) > 1:
            build_block(bk)
    after = len(blocks)
    tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
    return before, after


if __name__ == "__main__":
    tb = ta = 0
    for f in sorted(glob.glob(sys.argv[1] + "/ppt/slides/slide*.xml"), key=lambda s: int(re.findall(r"(\d+)\.xml", s)[0])):
        b, a = process(f)
        tb += b
        ta += a
    print(f"text boxes: {tb} -> {ta}")
