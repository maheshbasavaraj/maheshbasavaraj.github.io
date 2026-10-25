"""Build Mahesh_Basavaraj_Resume.docx from data.json (two-column, shaded sidebar, photo).

Layout: one 2-column table, one row per page. Left cell = shaded sidebar, right = main.
Fonts: Georgia (headings) + Calibri (body) -- present on every Office install.
"""
import json, pathlib
from docx import Document
from docx.shared import Pt, Mm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = pathlib.Path(__file__).parent
D = json.loads((ROOT / "data.json").read_text(encoding="utf-8"))

INK = RGBColor(0x16, 0x20, 0x2A); MUT = RGBColor(0x5B, 0x66, 0x70); BODY = RGBColor(0x33, 0x3F, 0x4A)
ACC = RGBColor(0x0E, 0x5A, 0x5C); ACC2 = RGBColor(0xB8, 0x74, 0x2A); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SIDE = "F3F1EC"; LINE = "DCDFE2"
HEAD, SANS = "Georgia", "Calibri"

PAGE_W, PAGE_H = Mm(210), Mm(297)
SIDE_W = Mm(66); MAIN_W = PAGE_W - SIDE_W


# ---------------------------------------------------------------- helpers ---
def set_cell_bg(cell, hex_):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_)
    tcPr.append(shd)

def set_cell_margins(cell, top=0, left=0, bottom=0, right=0):
    tcPr = cell._tc.get_or_add_tcPr(); mar = OxmlElement("w:tcMar")
    for k, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{k}"); el.set(qn("w:w"), str(int(v))); el.set(qn("w:type"), "dxa"); mar.append(el)
    tcPr.append(mar)

def set_cell_border_right(cell, hex_):
    tcPr = cell._tc.get_or_add_tcPr(); b = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        if side == "right":
            el.set(qn("w:val"), "single"); el.set(qn("w:sz"), "4"); el.set(qn("w:color"), hex_)
        else:
            el.set(qn("w:val"), "nil")
        b.append(el)
    tcPr.append(b)

def set_cell_width(cell, w):
    cell.width = w
    tcPr = cell._tc.get_or_add_tcPr(); tcW = tcPr.find(qn("w:tcW"))
    if tcW is None: tcW = OxmlElement("w:tcW"); tcPr.append(tcW)
    tcW.set(qn("w:w"), str(int(w / 635))); tcW.set(qn("w:type"), "dxa")

def para_border_bottom(p, hex_=LINE, sz=4, space=1):
    pPr = p._p.get_or_add_pPr(); pbdr = OxmlElement("w:pBdr"); bt = OxmlElement("w:bottom")
    bt.set(qn("w:val"), "single"); bt.set(qn("w:sz"), str(sz)); bt.set(qn("w:space"), str(space)); bt.set(qn("w:color"), hex_)
    pbdr.append(bt); pPr.append(pbdr)

def char_spacing(run, twips):
    rPr = run._r.get_or_add_rPr(); sp = OxmlElement("w:spacing"); sp.set(qn("w:val"), str(twips)); rPr.append(sp)

def shade_run(run, hex_):
    rPr = run._r.get_or_add_rPr(); shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_); rPr.append(shd)

def hyperlink(p, url, text, color=ACC, size=9, bold=True, font=SANS):
    part = p.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    h = OxmlElement("w:hyperlink"); h.set(qn("r:id"), r_id)
    r = OxmlElement("w:r"); rPr = OxmlElement("w:rPr")
    f = OxmlElement("w:rFonts"); f.set(qn("w:ascii"), font); f.set(qn("w:hAnsi"), font); rPr.append(f)
    c = OxmlElement("w:color"); c.set(qn("w:val"), str(color)); rPr.append(c)
    if bold: rPr.append(OxmlElement("w:b"))
    s = OxmlElement("w:sz"); s.set(qn("w:val"), str(int(size * 2))); rPr.append(s)
    r.append(rPr); t = OxmlElement("w:t"); t.text = text; t.set(qn("xml:space"), "preserve"); r.append(t)
    h.append(r); p._p.append(h)

def para(cell_or_doc, first=False, before=0, after=2, line=1.15, align=None, keep=False):
    p = cell_or_doc.paragraphs[0] if first else cell_or_doc.add_paragraph()
    pf = p.paragraph_format; pf.space_before = Pt(before); pf.space_after = Pt(after); pf.line_spacing = line
    if align: p.alignment = align
    if keep: pf.keep_with_next = True
    return p

def run(p, text, size=9, color=BODY, bold=False, italic=False, font=SANS, caps=False, spacing=None):
    r = p.add_run(text); r.font.size = Pt(size); r.font.color.rgb = color; r.font.bold = bold
    r.font.italic = italic; r.font.name = font; r.font.all_caps = caps
    r._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), font)
    if spacing: char_spacing(r, spacing)
    return r

def heading(cell, text, first=False, before=10):
    p = para(cell, first=first, before=0 if first else before, after=4, keep=True)
    run(p, text, size=8, color=ACC, bold=True, caps=True, spacing=40)
    para_border_bottom(p)
    return p

def bullet(cell, text, size=9):
    p = para(cell, after=1.5, line=1.12)
    pf = p.paragraph_format; pf.left_indent = Mm(3.6); pf.first_line_indent = Mm(-3.6)
    run(p, "▪\u2002", size=size - 2, color=ACC2); run(p, text, size=size)
    return p


# ------------------------------------------------------------- sidebar 1 ---
def sidebar_p1(c):
    p = para(c, first=True, after=6, align=WD_ALIGN_PARAGRAPH.LEFT)
    photo = ROOT / "assets" / "photo_circle.png"
    if photo.exists(): p.add_run().add_picture(str(photo), width=Mm(32))
    p = para(c, after=0, line=0.95); run(p, "Mahesh", size=24, color=INK, bold=True, font=HEAD)
    p = para(c, after=4, line=0.95); run(p, "Basavaraj", size=24, color=ACC, bold=True, font=HEAD)
    p = para(c, after=0); run(p, "Asst. Professor · AI Researcher", size=8, color=ACC2, bold=True, caps=True, spacing=10)

    heading(c, "Contact")
    for lab, val in (("Location", D["location"]), ("Phone", D["phone"])):
        p = para(c, after=0); run(p, lab, size=6.5, color=MUT, caps=True, spacing=20)
        p = para(c, after=3); run(p, val, size=8.6, color=INK)
    for m in D["emails"]:
        p = para(c, after=0); run(p, "Email", size=6.5, color=MUT, caps=True, spacing=20)
        p = para(c, after=3); hyperlink(p, "mailto:" + m, m, color=INK, size=8.6, bold=False)

    heading(c, "Online")
    for l in D["links"]:
        p = para(c, after=2.5)
        p.paragraph_format.tab_stops.add_tab_stop(Mm(15))
        p.paragraph_format.left_indent = Mm(15); p.paragraph_format.first_line_indent = Mm(-15)
        run(p, l["label"] + "\t", size=6.5, color=MUT, caps=True, spacing=20)
        hyperlink(p, l["url"], l["handle"], size=8.4)

    heading(c, "At a glance")
    t = c.add_table(rows=2, cols=2); t.autofit = False
    for i, s in enumerate(D["stats"]):
        cell = t.cell(i // 2, i % 2); set_cell_width(cell, Mm(25)); set_cell_margins(cell, 0, 0, 0, 0)
        p = para(cell, first=True, after=0, line=1.0); run(p, s["n"], size=16, color=ACC, bold=True, font=HEAD)
        p = para(cell, after=4, line=1.0); run(p, s["l"], size=7, color=MUT)

    heading(c, "Skills")
    for g in D["skills"]:
        p = para(c, after=0, keep=True); run(p, g["group"], size=8.2, color=INK, bold=True)
        p = para(c, after=4, line=1.1); run(p, " · ".join(g["items"]), size=8.2)


# ---------------------------------------------------------------- main 1 ---
def job_block(c, j):
    p = para(c, before=2, after=0, keep=True)
    p.paragraph_format.tab_stops.add_tab_stop(MAIN_W - Mm(20), alignment=2)  # right tab
    run(p, j["role"], size=10, color=INK, bold=True); run(p, "\t" + j["dates"], size=8, color=MUT, bold=True)
    p = para(c, after=2, keep=True); run(p, j["org"], size=9, color=ACC, bold=True)
    for pt in j["points"]: bullet(c, pt)

def main_p1(c):
    p = para(c, first=True, after=2, line=1.25); run(p, D["tagline"], size=12, color=BODY, font=HEAD, italic=True)
    heading(c, "Profile", before=8)
    p = para(c, after=0, line=1.2); run(p, D["summary"], size=9.3, color=BODY)
    heading(c, "Academic Experience")
    for j in D["academic"]: job_block(c, j)
    heading(c, "Industry Experience")
    for j in D["industry"]: job_block(c, j)


# ------------------------------------------------------------- sidebar 2 ---
def sidebar_p2(c):
    heading(c, "Education", first=True)
    for x in D["education"]:
        p = para(c, after=0, keep=True); run(p, x["degree"], size=8.8, color=INK, bold=True)
        p = para(c, after=0, keep=True); run(p, x["school"], size=8.2)
        p = para(c, after=7); run(p, x["when"], size=8, color=ACC2, bold=True)
    heading(c, "Certifications & FDPs")
    for x in D["certifications"]:
        p = para(c, after=0, keep=True); run(p, x["name"], size=8.6, color=INK, bold=True)
        p = para(c, after=6); run(p, f'{x["by"]} · {x["when"]}', size=8, color=MUT)
    heading(c, "Writing & Media")
    for w in D["writing"]:
        p = para(c, after=3, line=1.1); p.paragraph_format.left_indent = Mm(3); p.paragraph_format.first_line_indent = Mm(-3)
        run(p, "▪\u2002", size=7, color=ACC2); hyperlink(p, w["url"], w["title"], size=8.4)
    p = para(c, before=4, after=0, line=1.15)
    run(p, "More essays on Substack and AI explainers on the ", size=7.8, color=MUT)
    run(p, "Echoes of AI", size=7.8, color=MUT, bold=True); run(p, " YouTube channel.", size=7.8, color=MUT)


# ---------------------------------------------------------------- main 2 ---
def main_p2(c):
    heading(c, "Patents", first=True)
    for i, pt in enumerate(D["patents"], 1):
        p = para(c, after=4, line=1.15); p.paragraph_format.left_indent = Mm(8); p.paragraph_format.first_line_indent = Mm(-8)
        p.paragraph_format.tab_stops.add_tab_stop(Mm(8))
        run(p, f"0{i}\t", size=11, color=ACC2, bold=True); run(p, pt, size=9.4, color=INK)
        run(p, "").add_break(); run(p, "Utility patent · 2026", size=7.6, color=MUT)
    heading(c, "Research Publications")
    for x in D["publications"]:
        p = para(c, after=4, line=1.15)
        v = run(p, f'\u00a0{x["venue"]}\u00a0', size=7, color=WHITE, bold=True); shade_run(v, "0E5A5C")
        run(p, "  " + x["title"], size=9, color=INK, bold=True)
        run(p, "  " + x["where"], size=8, color=MUT)
    heading(c, "Open-Source & Teaching Projects")
    p = para(c, after=3); run(p, "github.com/" + D["github_user"] + " · selected repositories", size=8, color=MUT)
    for x in D["projects"][:5]:
        p = para(c, after=0, keep=True); hyperlink(p, x["url"], x["name"], size=9.2)
        run(p, "   " + x["tech"], size=7.4, color=ACC2, bold=True)
        p = para(c, after=4, line=1.12); run(p, x["desc"], size=8.6)


# ------------------------------------------------------------------ build ---
def build(out):
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = SANS; st.font.size = Pt(9)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), SANS)
    s = doc.sections[0]
    s.page_width, s.page_height = PAGE_W, PAGE_H
    s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Mm(0)
    s.header_distance = s.footer_distance = Mm(0)
    # empty header/footer paragraphs still reserve ~12pt at page edges; shrink them to ~0
    for hf in (s.header, s.footer):
        for hp in hf.paragraphs:
            hp.paragraph_format.space_before = hp.paragraph_format.space_after = Pt(0)
            hp.paragraph_format.line_spacing = Pt(1)
            pPr = hp._p.get_or_add_pPr(); rpr = OxmlElement("w:rPr")
            sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "2"); rpr.append(sz); pPr.append(rpr)
    settings = doc.settings.element
    compat = settings.find(qn("w:compat"))
    if compat is None: compat = OxmlElement("w:compat"); settings.append(compat)
    for cs in compat.findall(qn("w:compatSetting")): compat.remove(cs)
    cs = OxmlElement("w:compatSetting"); cs.set(qn("w:name"), "compatibilityMode")
    cs.set(qn("w:uri"), "http://schemas.microsoft.com/office/word"); cs.set(qn("w:val"), "15"); compat.append(cs)
    ind = OxmlElement("w:tblInd"); ind.set(qn("w:w"), "0"); ind.set(qn("w:type"), "dxa")
    # full-height sidebar strip drawn in the header (behind text, every page) so shading never stops short
    hp = s.header.paragraphs[0]
    strip = f'''<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:drawing>
<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="1" locked="1" layoutInCell="0" allowOverlap="1">
<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>
<wp:extent cx="{int(SIDE_W)}" cy="{int(PAGE_H)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>
<wp:docPr id="99" name="SidebarStrip"/><wp:cNvGraphicFramePr/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/>
<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(SIDE_W)}" cy="{int(PAGE_H)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{SIDE}"/></a:solidFill><a:ln w="6350"><a:solidFill><a:srgbClr val="{LINE}"/></a:solidFill></a:ln></wps:spPr>
<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'''
    from docx.oxml import parse_xml
    hp._p.append(parse_xml(strip))
    props = doc.core_properties; props.title = "Mahesh Basavaraj — Resume"; props.author = "Mahesh Basavaraj"

    t = doc.add_table(rows=2, cols=2); t.alignment = WD_TABLE_ALIGNMENT.LEFT; t.autofit = False
    tblPr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay); tblPr.append(ind)
    # first row 14mm top padding etc. (dxa = 1/20 pt; 1 mm ≈ 56.7 dxa)
    mm = 56.7
    for r, row in enumerate(t.rows):
        row.height = PAGE_H - (Mm(0.6) if r == 0 else Mm(10)); row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        trPr = row._tr.get_or_add_trPr(); cs = OxmlElement("w:cantSplit"); trPr.append(cs)
        l, m = row.cells
        set_cell_width(l, SIDE_W); set_cell_width(m, MAIN_W)
        set_cell_bg(l, SIDE); set_cell_border_right(l, LINE)
        set_cell_margins(l, top=13 * mm, left=10 * mm, bottom=8 * mm, right=6 * mm)
        set_cell_margins(m, top=13 * mm, left=9 * mm, bottom=8 * mm, right=11 * mm)
        tcb = m._tc.get_or_add_tcPr().find(qn("w:tcBorders"))
        if tcb is None:
            b = OxmlElement("w:tcBorders")
            for side in ("top", "left", "bottom", "right"):
                el = OxmlElement(f"w:{side}"); el.set(qn("w:val"), "nil"); b.append(el)
            m._tc.get_or_add_tcPr().append(b)
    grid = t._tbl.tblGrid.findall(qn("w:gridCol"))
    grid[0].set(qn("w:w"), str(int(SIDE_W / 635))); grid[1].set(qn("w:w"), str(int(MAIN_W / 635)))

    sidebar_p1(t.cell(0, 0)); main_p1(t.cell(0, 1))
    sidebar_p2(t.cell(1, 0)); main_p2(t.cell(1, 1))

    # Word requires a paragraph after a table; make it 1pt so it doesn't create a 3rd page
    last = doc.paragraphs[-1] if doc.paragraphs else doc.add_paragraph()
    tail = doc.element.body[-2] if doc.element.body[-1].tag == qn("w:sectPr") else doc.element.body[-1]
    for p in doc.paragraphs:
        p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = Pt(1)
        for r_ in p.runs: r_.font.size = Pt(1)
        # shrink + hide the paragraph mark itself so Word doesn't spill it onto page 3
        pPr = p._p.get_or_add_pPr(); rpr = OxmlElement("w:rPr")
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "2"); rpr.append(sz)
        rpr.append(OxmlElement("w:vanish")); pPr.append(rpr)
    doc.save(out)

if __name__ == "__main__":
    out = ROOT / "Mahesh_Basavaraj_Resume.docx"
    build(out); print("wrote", out)
