"""DOCX -> HWPX converter tuned for the uploaded report (paragraphs, tables, inline pictures, page numbers)."""
import re
import sys
import zipfile
from xml.sax.saxutils import escape
from lxml import etree
from PIL import Image
import io

SRC, SKEL, OUT = sys.argv[1], sys.argv[2], sys.argv[3]

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS = {"w": W, "a": A, "r": R, "wp": WP}
FONT = "맑은 고딕"


def w(tag):
    return f"{{{W}}}{tag}"


def wval(el, path, attr="val", default=None):
    if el is None:
        return default
    x = el.find(path, NS)
    if x is None:
        return default
    return x.get(w(attr), default)


def twip(v):  # twips -> HWPUNIT
    return int(round(int(v) * 5))


src = zipfile.ZipFile(SRC)
doc = etree.fromstring(src.read("word/document.xml"))
rels = etree.fromstring(src.read("word/_rels/document.xml.rels"))
relmap = {r.get("Id"): r.get("Target") for r in rels}

# ---------------------------------------------------------------- registries
char_props, para_props, border_fills = {}, {}, {}
CHAR_BASE, PARA_BASE, BF_BASE = 7, 20, 3


def char_id(size, bold, color):
    key = (size, bold, color)
    if key not in char_props:
        char_props[key] = CHAR_BASE + len(char_props)
    return char_props[key]


def para_id(**kw):
    key = tuple(sorted(kw.items()))
    if key not in para_props:
        para_props[key] = PARA_BASE + len(para_props)
    return para_props[key]


def mm_width(sz_eighths):
    pt = int(sz_eighths) / 8.0
    mm = pt * 0.3528
    choices = [0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0]
    return min(choices, key=lambda c: abs(c - mm))


def side(el):
    """Return (type, width_mm, color) for a w:top/w:left... border element."""
    if el is None:
        return ("NONE", 0.12, "#000000")
    val = el.get(w("val"), "none")
    if val in ("none", "nil"):
        return ("NONE", 0.12, "#000000")
    color = el.get(w("color"), "000000")
    color = "#000000" if color == "auto" else "#" + color.upper()
    return ("SOLID", mm_width(el.get(w("sz"), "4")), color)


def bf_id(top, left, bottom, right, fill=None):
    key = (top, left, bottom, right, fill)
    if key not in border_fills:
        border_fills[key] = BF_BASE + len(border_fills)
    return border_fills[key]


NO_BORDER = ("NONE", 0.12, "#000000")

# ---------------------------------------------------------------- pictures
bin_items = []  # (id, name, bytes, ext)
pic_counter = [1000]


def picture_xml(drawing):
    ext = drawing.find(".//wp:extent", NS)
    cx, cy = int(ext.get("cx")), int(ext.get("cy"))
    width, height = round(cx / 127), round(cy / 127)
    blip = drawing.find(".//a:blip", NS)
    target = relmap[blip.get(f"{{{R}}}embed")]
    data = src.read("word/" + target)
    fmt = target.rsplit(".", 1)[-1].lower()
    fmt = "jpg" if fmt == "jpeg" else fmt
    idx = len(bin_items) + 1
    bid = f"BIN{idx:04d}"
    bin_items.append((bid, f"{bid}.{fmt}", data, fmt))
    px_w, px_h = Image.open(io.BytesIO(data)).size
    dw, dh = px_w * 75, px_h * 75
    pic_counter[0] += 1
    pid = pic_counter[0]
    return (
        f'<hp:pic id="{pid}" zOrder="{pid - 1000}" numberingType="PICTURE" textWrap="TOP_AND_BOTTOM" '
        f'textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="{pid}" reverse="0">'
        f'<hp:offset x="0" y="0"/><hp:orgSz width="{width}" height="{height}"/>'
        f'<hp:curSz width="{width}" height="{height}"/><hp:flip horizontal="0" vertical="0"/>'
        f'<hp:rotationInfo angle="0" centerX="{width // 2}" centerY="{height // 2}" rotateimage="1"/>'
        f'<hp:renderingInfo><hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
        f'<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
        f'<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo>'
        f'<hp:imgRect><hc:pt0 x="0" y="0"/><hc:pt1 x="{width}" y="0"/><hc:pt2 x="{width}" y="{height}"/>'
        f'<hc:pt3 x="0" y="{height}"/></hp:imgRect>'
        f'<hp:imgClip left="0" right="{dw}" top="0" bottom="{dh}"/>'
        f'<hp:inMargin left="0" right="0" top="0" bottom="0"/>'
        f'<hp:imgDim dimwidth="{dw}" dimheight="{dh}"/>'
        f'<hc:img binaryItemIDRef="{bid}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>'
        f'<hp:effects/>'
        f'<hp:sz width="{width}" widthRelTo="ABSOLUTE" height="{height}" heightRelTo="ABSOLUTE" protect="0"/>'
        f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" '
        f'vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
        f'<hp:outMargin left="0" right="0" top="0" bottom="0"/><hp:shapeComment/></hp:pic>'
    ), height


# ---------------------------------------------------------------- paragraphs
pid_counter = [0]
pending_break = [False]


def next_pid():
    pid_counter[0] += 1
    return pid_counter[0]


def paragraph_xml(p, sec_pr_xml=""):
    ppr = p.find("w:pPr", NS)
    align = {"center": "CENTER", "right": "RIGHT", "both": "JUSTIFY"}.get(wval(ppr, "w:jc"), "LEFT")
    sp = ppr.find("w:spacing", NS) if ppr is not None else None
    before = int(sp.get(w("before"), "0")) if sp is not None else 0
    after = int(sp.get(w("after"), "0")) if sp is not None else 0
    line = int(sp.get(w("line"), "240")) if sp is not None else 240
    line_pct = int(round(line / 240 * 133))
    ind = ppr.find("w:ind", NS) if ppr is not None else None
    left = int(ind.get(w("left"), "0")) if ind is not None else 0
    hanging = int(ind.get(w("hanging"), "0")) if ind is not None else 0
    first = int(ind.get(w("firstLine"), "0")) if ind is not None else 0
    keep = 1 if (ppr is not None and ppr.find("w:keepNext", NS) is not None) else 0

    border = None
    pbdr = ppr.find("w:pBdr", NS) if ppr is not None else None
    b_off = (0, 0)
    if pbdr is not None:
        t, b = pbdr.find("w:top", NS), pbdr.find("w:bottom", NS)
        border = bf_id(side(t), NO_BORDER, side(b), NO_BORDER)
        spc = (t if t is not None else b).get(w("space"), "0")
        off = int(spc) * 100
        b_off = (off if t is not None else 0, off if b is not None else 0)

    # Word hanging indent: first line at left-hanging, rest at left.
    if hanging:
        h_left, h_intent = twip(left - hanging), -twip(hanging)
    else:
        h_left, h_intent = twip(left), twip(first)

    runs_xml = []
    first_size = None
    has_page_break = False
    for r in p.iter(w("r")):
        rpr = r.find("w:rPr", NS)
        size = int(wval(rpr, "w:sz", default="21"))
        bold = rpr is not None and rpr.find("w:b", NS) is not None and wval(rpr, "w:b", default="true") not in ("0", "false")
        color = wval(rpr, "w:color", default="000000")
        color = "000000" if color == "auto" else color.upper()
        cid = char_id(size, bold, color)
        if first_size is None:
            first_size = cid
        parts = []
        for child in r:
            tag = etree.QName(child).localname
            if tag == "t":
                parts.append(f"<hp:t>{escape(child.text or '')}</hp:t>")
            elif tag == "tab":
                parts.append("<hp:t><hp:tab/></hp:t>")
            elif tag == "br":
                if child.get(w("type")) == "page":
                    has_page_break = True
                else:
                    parts.append("<hp:t><hp:lineBreak/></hp:t>")
            elif tag == "drawing":
                xml, _ = picture_xml(child)
                parts.append(xml)
        runs_xml.append(f'<hp:run charPrIDRef="{cid}">{"".join(parts)}</hp:run>')
    if first_size is None:
        first_size = char_id(21, False, "000000")
        runs_xml.append(f'<hp:run charPrIDRef="{first_size}"/>')

    ppid = para_id(align=align, left=h_left, intent=h_intent, prev=twip(before), next=twip(after),
                   line=line_pct, border=border or 0, boff=b_off, keep=keep)
    text_only = "".join(t.text or "" for t in p.iter(w("t")))
    if has_page_break and not text_only and p.find(".//w:drawing", NS) is None and not sec_pr_xml:
        pending_break[0] = True  # break-only paragraph: move the break onto the next block
        return ""
    page_break = 1 if (pending_break[0] or has_page_break) else 0
    pending_break[0] = False
    head = ""
    if sec_pr_xml:
        head = f'<hp:run charPrIDRef="{first_size}">{sec_pr_xml}</hp:run>'
    return (f'<hp:p id="{next_pid()}" paraPrIDRef="{ppid}" styleIDRef="0" pageBreak="{page_break}" '
            f'columnBreak="0" merged="0">{head}{"".join(runs_xml)}</hp:p>')


# ---------------------------------------------------------------- tables
def table_xml(tbl):
    grid = [twip(g.get(w("w"))) for g in tbl.findall("w:tblGrid/w:gridCol", NS)]
    rows = tbl.findall("w:tr", NS)
    # first pass: layout matrix
    layout = []  # per row: list of (tc, col, span, vmerge)
    for tr in rows:
        col, cells = 0, []
        for tc in tr.findall("w:tc", NS):
            tcpr = tc.find("w:tcPr", NS)
            span = int(wval(tcpr, "w:gridSpan", default="1"))
            vm = tcpr.find("w:vMerge", NS) if tcpr is not None else None
            vmerge = None if vm is None else (vm.get(w("val")) or "continue")
            cells.append((tc, col, span, vmerge))
            col += span
        layout.append(cells)

    def rowspan(ri, col):
        n = 1
        for rj in range(ri + 1, len(layout)):
            hit = [c for c in layout[rj] if c[1] == col and c[3] == "continue"]
            if not hit:
                break
            n += 1
        return n

    header_row = rows[0].find("w:trPr/w:tblHeader", NS) is not None if rows else False
    rows_xml = []
    for ri, cells in enumerate(layout):
        tcs = []
        for tc, col, span, vmerge in cells:
            if vmerge == "continue":
                continue
            tcpr = tc.find("w:tcPr", NS)
            bd = tcpr.find("w:tcBorders", NS)
            sides = [side(bd.find(f"w:{s}", NS)) if bd is not None else ("SOLID", 0.12, "#000000")
                     for s in ("top", "left", "bottom", "right")]
            shd = tcpr.find("w:shd", NS)
            fill = None
            if shd is not None and shd.get(w("fill")) not in (None, "auto", "FFFFFF"):
                fill = "#" + shd.get(w("fill")).upper()
            bid = bf_id(*sides, fill=fill)
            mar = tcpr.find("w:tcMar", NS)
            m = {s: twip(mar.find(f"w:{s}", NS).get(w("w"))) if mar is not None and mar.find(f"w:{s}", NS) is not None else 0
                 for s in ("top", "left", "bottom", "right")}
            valign = {"center": "CENTER", "bottom": "BOTTOM"}.get(wval(tcpr, "w:vAlign"), "TOP")
            rs = rowspan(ri, col) if vmerge == "restart" else 1
            width = sum(grid[col:col + span])
            paras = "".join(paragraph_xml(p) for p in tc.findall("w:p", NS))
            tcs.append(
                f'<hp:tc name="" header="{1 if (header_row and ri == 0) else 0}" hasMargin="1" protect="0" '
                f'editable="0" dirty="0" borderFillIDRef="{bid}">'
                f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="{valign}" '
                f'linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
                f'{paras}</hp:subList>'
                f'<hp:cellAddr colAddr="{col}" rowAddr="{ri}"/><hp:cellSpan colSpan="{span}" rowSpan="{rs}"/>'
                f'<hp:cellSz width="{width}" height="282"/>'
                f'<hp:cellMargin left="{m["left"]}" right="{m["right"]}" top="{m["top"]}" bottom="{m["bottom"]}"/></hp:tc>')
        rows_xml.append("<hp:tr>" + "".join(tcs) + "</hp:tr>")
    total = sum(grid)
    tid = next_pid() + 500000
    return (f'<hp:tbl id="{tid}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" '
            f'lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="{1 if header_row else 0}" '
            f'rowCnt="{len(rows)}" colCnt="{len(grid)}" cellSpacing="0" borderFillIDRef="1" noAdjust="0">'
            f'<hp:sz width="{total}" widthRelTo="ABSOLUTE" height="{282 * len(rows)}" heightRelTo="ABSOLUTE" protect="0"/>'
            f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" '
            f'vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
            f'<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
            f'<hp:inMargin left="300" right="300" top="200" bottom="200"/>'
            + "".join(rows_xml) + "</hp:tbl>")


# ---------------------------------------------------------------- section properties
body = doc.find("w:body", NS)
sect = body.find("w:sectPr", NS)
pg = sect.find("w:pgSz", NS)
mar = sect.find("w:pgMar", NS)
page_w, page_h = twip(pg.get(w("w"))), twip(pg.get(w("h")))
m_top, m_bot = twip(mar.get(w("top"))), twip(mar.get(w("bottom")))
m_l, m_r = twip(mar.get(w("left"))), twip(mar.get(w("right")))
m_foot = twip(mar.get(w("footer")))
title_pg = sect.find("w:titlePg", NS) is not None
# Word measures body from page edge; Hancom stacks header/footer between margin and body.
foot_area = max(m_bot - m_foot, 0)

skel_sec = open(f"{SKEL}/Contents/section0.xml", encoding="utf-8").read()
sec_pr = re.search(r"<hp:secPr.*?</hp:secPr>", skel_sec, re.S).group(0)
sec_pr = re.sub(r'<hp:pagePr [^>]*>', f'<hp:pagePr landscape="WIDELY" width="{page_w}" height="{page_h}" gutterType="LEFT_ONLY">', sec_pr)
sec_pr = re.sub(r'<hp:margin [^>]*/>',
                f'<hp:margin header="0" footer="{foot_area}" gutter="0" left="{m_l}" right="{m_r}" top="{m_top}" bottom="{m_foot}"/>',
                sec_pr)
sec_pr = sec_pr.replace('hideFirstPageNum="0"', f'hideFirstPageNum="{1 if title_pg else 0}"')
sec_pr = sec_pr.replace('<hp:startNum pageStartsOn="BOTH" page="0"', '<hp:startNum pageStartsOn="BOTH" page="1"')
col_ctrl = '<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>'
page_num = '<hp:ctrl><hp:pageNum pos="BOTTOM_CENTER" formatType="DIGIT" sideChar="-"/></hp:ctrl>'
first_run_ctrl = sec_pr + col_ctrl + page_num

# ---------------------------------------------------------------- body
out = []
first = True
for el in body:
    tag = etree.QName(el).localname
    if tag == "p":
        out.append(paragraph_xml(el, first_run_ctrl if first else ""))
        first = False
    elif tag == "tbl":
        if first:
            out.append(paragraph_xml(etree.Element(w("p")), first_run_ctrl))
            first = False
        tpid = para_id(align="LEFT", left=0, intent=0, prev=0, next=0, line=133, border=0, boff=(0, 0), keep=0)
        pb = 1 if pending_break[0] else 0
        pending_break[0] = False
        out.append(f'<hp:p id="{next_pid()}" paraPrIDRef="{tpid}" styleIDRef="0" pageBreak="{pb}" columnBreak="0" merged="0">'
                   f'<hp:run charPrIDRef="{char_id(21, False, "000000")}">{table_xml(el)}<hp:t/></hp:run></hp:p>')

sec_head = skel_sec[: skel_sec.index("<hp:p ")]
section_xml = sec_head + "".join(out) + "</hs:sec>"

# ---------------------------------------------------------------- header.xml
header = open(f"{SKEL}/Contents/header.xml", encoding="utf-8").read()
header = header.replace('face="함초롬돋움"', f'face="{FONT}"').replace('face="함초롬바탕"', f'face="{FONT}"')
# page number style (char 1): 9pt
header = re.sub(r'(<hh:charPr id="1" height=")\d+', r"\g<1>900", header)

CHAR_T = ('<hh:charPr id="{id}" height="{h}" textColor="#{c}" shadeColor="none" useFontSpace="0" useKerning="0" '
          'symMark="NONE" borderFillIDRef="2"><hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>'
          '<hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>'
          '<hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>'
          '<hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>'
          '<hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>{b}'
          '<hh:underline type="NONE" shape="SOLID" color="#000000"/><hh:strikeout shape="NONE" color="#000000"/>'
          '<hh:outline type="NONE"/><hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/></hh:charPr>')
chars = "".join(CHAR_T.format(id=i, h=size * 50, c=color, b="<hh:bold/>" if bold else "")
                for (size, bold, color), i in sorted(char_props.items(), key=lambda kv: kv[1]))


def margin(intent, left, prev, nxt, line, k):
    return (f'<hh:margin><hc:intent value="{intent * k}" unit="HWPUNIT"/><hc:left value="{left * k}" unit="HWPUNIT"/>'
            f'<hc:right value="0" unit="HWPUNIT"/><hc:prev value="{prev * k}" unit="HWPUNIT"/>'
            f'<hc:next value="{nxt * k}" unit="HWPUNIT"/></hh:margin>'
            f'<hh:lineSpacing type="PERCENT" value="{line}" unit="HWPUNIT"/>')


paras = []
for key, i in sorted(para_props.items(), key=lambda kv: kv[1]):
    d = dict(key)
    bfr = d["border"] or 2
    bt, bb = d["boff"]
    paras.append(
        f'<hh:paraPr id="{i}" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="0" suppressLineNumbers="0" '
        f'checked="0" textDir="LTR"><hh:align horizontal="{d["align"]}" vertical="BASELINE"/>'
        f'<hh:heading type="NONE" idRef="0" level="0"/>'
        f'<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="KEEP_WORD" widowOrphan="0" '
        f'keepWithNext="{d["keep"]}" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>'
        f'<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
        f'<hp:switch><hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">'
        f'{margin(d["intent"], d["left"], d["prev"], d["next"], d["line"], 1)}</hp:case>'
        f'<hp:default>{margin(d["intent"], d["left"], d["prev"], d["next"], d["line"], 2)}</hp:default></hp:switch>'
        f'<hh:border borderFillIDRef="{bfr}" offsetLeft="0" offsetRight="0" offsetTop="{bt}" offsetBottom="{bb}" '
        f'connect="0" ignoreMargin="1"/></hh:paraPr>')


def bline(name, s):
    return f'<hh:{name} type="{s[0]}" width="{s[1]} mm" color="{s[2]}"/>'


bfs = []
for (t, l, b, r, fill), i in sorted(border_fills.items(), key=lambda kv: kv[1]):
    fb = (f'<hc:fillBrush><hc:winBrush faceColor="{fill}" hatchColor="#999999" alpha="0"/></hc:fillBrush>'
          if fill else "")
    bfs.append(
        f'<hh:borderFill id="{i}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
        f'<hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
        + bline("leftBorder", l) + bline("rightBorder", r) + bline("topBorder", t) + bline("bottomBorder", b)
        + '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>' + fb + "</hh:borderFill>")


def append_list(xml, tag, items):
    m = re.search(rf'<hh:{tag} itemCnt="(\d+)">', xml)
    cnt = int(m.group(1)) + len(items)
    xml = xml[:m.start()] + f'<hh:{tag} itemCnt="{cnt}">' + xml[m.end():]
    end = xml.index(f"</hh:{tag}>")
    return xml[:end] + "".join(items) + xml[end:]


header = append_list(header, "borderFills", bfs)
header = append_list(header, "charProperties", [chars])
m = re.search(r'<hh:charProperties itemCnt="(\d+)">', header)
header = header.replace(m.group(0), f'<hh:charProperties itemCnt="{CHAR_BASE + len(char_props)}">')
header = append_list(header, "paraProperties", paras)

# ---------------------------------------------------------------- package
hpf = open(f"{SKEL}/Contents/content.hpf", encoding="utf-8").read()
items = "".join(
    f'<opf:item id="{bid}" href="BinData/{name}" media-type="{"image/png" if fmt == "png" else "image/jpeg"}" isEmbeded="1"/>'
    for bid, name, _, fmt in bin_items)
hpf = hpf.replace('<opf:item id="settings"', items + '<opf:item id="settings"')
hpf = hpf.replace("<opf:title/>", "<opf:title>전국우수시장박람회 전통시장 안전관리 홍보부스 운영</opf:title>")

with zipfile.ZipFile(OUT, "w") as z:
    z.writestr(zipfile.ZipInfo("mimetype"), open(f"{SKEL}/mimetype", "rb").read(), compress_type=zipfile.ZIP_STORED)
    for name in ["version.xml", "META-INF/container.xml", "META-INF/container.rdf", "META-INF/manifest.xml",
                 "settings.xml", "Preview/PrvImage.png"]:
        z.write(f"{SKEL}/{name}", name, compress_type=zipfile.ZIP_DEFLATED)
    z.writestr("Contents/header.xml", header, compress_type=zipfile.ZIP_DEFLATED)
    z.writestr("Contents/section0.xml", section_xml, compress_type=zipfile.ZIP_DEFLATED)
    z.writestr("Contents/content.hpf", hpf, compress_type=zipfile.ZIP_DEFLATED)
    text = " ".join(t.text or "" for t in doc.iter(w("t")))[:1000]
    z.writestr("Preview/PrvText.txt", text, compress_type=zipfile.ZIP_DEFLATED)
    for bid, name, data, _ in bin_items:
        z.writestr(f"BinData/{name}", data, compress_type=zipfile.ZIP_STORED)

print("chars", len(char_props), "paras", len(para_props), "borderFills", len(border_fills), "pics", len(bin_items))
