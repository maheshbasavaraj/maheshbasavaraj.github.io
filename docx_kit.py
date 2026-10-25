"""Shared DOCX design kit - the same typography, colours and two-column layout
used by the resume, so every document in the application packet looks like one set.

Fonts: Georgia (headings) + Calibri (body) - present on every Office install.
Colours: deep teal accent, amber secondary, warm-grey sidebar.
"""
import pathlib
from docx.shared import Pt, Mm, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
from docx.text.paragraph import Paragraph

ROOT = pathlib.Path(__file__).parent

INK = RGBColor(0x16, 0x20, 0x2A); MUT = RGBColor(0x5B, 0x66, 0x70); BODY = RGBColor(0x33, 0x3F, 0x4A)
ACC = RGBColor(0x0E, 0x5A, 0x5C); ACC2 = RGBColor(0xB8, 0x74, 0x2A); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SIDE = "F3F1EC"; LINE = "DCDFE2"; TEAL = "0E5A5C"; AMBER = "B8742A"
HEAD, SANS = "Georgia", "Calibri"

PAGE_W, PAGE_H = Mm(210), Mm(297)
SIDE_W = Mm(66); MAIN_W = PAGE_W - SIDE_W
MM = 56.7  # dxa per millimetre


# ------------------------------------------------------------------ atoms ---
def set_cell_bg(cell, hex_):
    tcPr = cell._tc.get_or_add_tcPr(); shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_)
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


def no_borders(cell):
    tcPr = cell._tc.get_or_add_tcPr(); b = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}"); el.set(qn("w:val"), "nil"); b.append(el)
    tcPr.append(b)


def set_cell_width(cell, w):
    cell.width = w
    tcPr = cell._tc.get_or_add_tcPr(); tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW"); tcPr.append(tcW)
    tcW.set(qn("w:w"), str(int(w / 635))); tcW.set(qn("w:type"), "dxa")


def para_border_bottom(p, hex_=LINE, sz=4, space=1):
    pPr = p._p.get_or_add_pPr(); pbdr = OxmlElement("w:pBdr"); bt = OxmlElement("w:bottom")
    bt.set(qn("w:val"), "single"); bt.set(qn("w:sz"), str(sz)); bt.set(qn("w:space"), str(space))
    bt.set(qn("w:color"), hex_); pbdr.append(bt); pPr.append(pbdr)


def char_spacing(run_, twips):
    rPr = run_._r.get_or_add_rPr(); sp = OxmlElement("w:spacing"); sp.set(qn("w:val"), str(twips)); rPr.append(sp)


def shade_run(run_, hex_):
    rPr = run_._r.get_or_add_rPr(); shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_); rPr.append(shd)


def hyperlink(p, url, text, color=ACC, size=9, bold=True, font=SANS):
    part = p.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                          is_external=True)
    h = OxmlElement("w:hyperlink"); h.set(qn("r:id"), r_id)
    r = OxmlElement("w:r"); rPr = OxmlElement("w:rPr")
    f = OxmlElement("w:rFonts"); f.set(qn("w:ascii"), font); f.set(qn("w:hAnsi"), font); rPr.append(f)
    c = OxmlElement("w:color"); c.set(qn("w:val"), str(color)); rPr.append(c)
    if bold:
        rPr.append(OxmlElement("w:b"))
    s = OxmlElement("w:sz"); s.set(qn("w:val"), str(int(size * 2))); rPr.append(s)
    r.append(rPr); t = OxmlElement("w:t"); t.text = text; t.set(qn("xml:space"), "preserve"); r.append(t)
    h.append(r); p._p.append(h)


def para(cell_or_doc, first=False, before=0, after=2, line=1.15, align=None, keep=False):
    p = cell_or_doc.paragraphs[0] if (first and cell_or_doc.paragraphs) else cell_or_doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after); pf.line_spacing = line
    if align:
        p.alignment = align
    if keep:
        pf.keep_with_next = True
    return p


def run(p, text, size=9, color=BODY, bold=False, italic=False, font=SANS, caps=False, spacing=None):
    r = p.add_run(text)
    r.font.size = Pt(size); r.font.color.rgb = color; r.font.bold = bold
    r.font.italic = italic; r.font.name = font; r.font.all_caps = caps
    r._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), font)
    if spacing:
        char_spacing(r, spacing)
    return r


def heading(cell, text, first=False, before=10, color=ACC, size=8):
    p = para(cell, first=first, before=0 if first else before, after=4, keep=True)
    run(p, text, size=size, color=color, bold=True, caps=True, spacing=40)
    para_border_bottom(p)
    return p


def bullet(cell, text, size=9, color=BODY, after=1.5):
    p = para(cell, after=after, line=1.12)
    pf = p.paragraph_format; pf.left_indent = Mm(3.6); pf.first_line_indent = Mm(-3.6)
    run(p, "\u25aa\u2002", size=size - 2, color=ACC2); run(p, text, size=size, color=color)
    return p


def numbered(cell, n, text, size=9):
    p = para(cell, after=3, line=1.14)
    pf = p.paragraph_format; pf.left_indent = Mm(6); pf.first_line_indent = Mm(-6)
    pf.tab_stops.add_tab_stop(Mm(6))
    run(p, f"{n}\t", size=size + 0.6, color=ACC2, bold=True); run(p, text, size=size)
    return p


def badge(p, text, fill=TEAL, size=7):
    v = run(p, f"\u00a0{text}\u00a0", size=size, color=WHITE, bold=True)
    shade_run(v, fill)
    return v


# ------------------------------------------------------------ page set-up ---
def base_styles(doc):
    st = doc.styles["Normal"]; st.font.name = SANS; st.font.size = Pt(9)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), SANS)


def _quiet_headers(s):
    for hf in (s.header, s.footer):
        for hp in hf.paragraphs:
            hp.paragraph_format.space_before = hp.paragraph_format.space_after = Pt(0)
            hp.paragraph_format.line_spacing = Pt(1)
            pPr = hp._p.get_or_add_pPr(); rpr = OxmlElement("w:rPr")
            sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "2"); rpr.append(sz); pPr.append(rpr)


def _compat15(doc):
    settings = doc.settings.element
    compat = settings.find(qn("w:compat"))
    if compat is None:
        compat = OxmlElement("w:compat"); settings.append(compat)
    for cs in compat.findall(qn("w:compatSetting")):
        compat.remove(cs)
    cs = OxmlElement("w:compatSetting"); cs.set(qn("w:name"), "compatibilityMode")
    cs.set(qn("w:uri"), "http://schemas.microsoft.com/office/word"); cs.set(qn("w:val"), "15"); compat.append(cs)


def _sidebar_strip(s):
    """Full-height sidebar rectangle anchored in the header, behind the text."""
    hp = s.header.paragraphs[0]
    strip = f'''<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:drawing>
<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="1" locked="1" layoutInCell="0" allowOverlap="1">
<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>
<wp:extent cx="{int(SIDE_W)}" cy="{int(PAGE_H)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>
<wp:docPr id="99" name="SidebarStrip"/><wp:cNvGraphicFramePr/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/>
<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(SIDE_W)}" cy="{int(PAGE_H)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{SIDE}"/></a:solidFill><a:ln w="6350"><a:solidFill><a:srgbClr val="{LINE}"/></a:solidFill></a:ln></wps:spPr>
<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'''
    hp._p.append(parse_xml(strip))


def single_page(doc, top=20, bottom=18, left=20, right=20):
    """A plain one-column page with real margins (for letters)."""
    base_styles(doc)
    s = doc.sections[0]
    s.page_width, s.page_height = PAGE_W, PAGE_H
    s.top_margin, s.bottom_margin = Mm(top), Mm(bottom)
    s.left_margin, s.right_margin = Mm(left), Mm(right)
    s.header_distance = s.footer_distance = Mm(0)
    _quiet_headers(s); _compat15(doc)
    return s


def two_col(doc, pages, top_pad=13, left_pad=10, right_pad=11, main_left=9, main_right=11,
            row1_trim=0.6, row2_trim=10):
    """Build the shaded-sidebar layout: one table row per page, each row exactly
    one page tall so the sidebar shading and column split never drift.

    `pages` is a list of (sidebar_fn, main_fn) taking a table cell.
    """
    base_styles(doc)
    s = doc.sections[0]
    s.page_width, s.page_height = PAGE_W, PAGE_H
    s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Mm(0)
    s.header_distance = s.footer_distance = Mm(0)
    _quiet_headers(s); _compat15(doc); _sidebar_strip(s)

    t = doc.add_table(rows=len(pages), cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT; t.autofit = False
    tblPr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    ind = OxmlElement("w:tblInd"); ind.set(qn("w:w"), "0"); ind.set(qn("w:type"), "dxa"); tblPr.append(ind)

    for i, row in enumerate(t.rows):
        row.height = PAGE_H - Mm(row1_trim if i == 0 else row2_trim)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        l, m = row.cells
        set_cell_width(l, SIDE_W); set_cell_width(m, MAIN_W)
        set_cell_bg(l, SIDE); set_cell_border_right(l, LINE)
        set_cell_margins(l, top=top_pad * MM, left=left_pad * MM, bottom=8 * MM, right=6 * MM)
        set_cell_margins(m, top=top_pad * MM, left=main_left * MM, bottom=8 * MM, right=main_right * MM)
        no_borders(m)

    grid = t._tbl.tblGrid.findall(qn("w:gridCol"))
    grid[0].set(qn("w:w"), str(int(SIDE_W / 635))); grid[1].set(qn("w:w"), str(int(MAIN_W / 635)))

    for i, (side_fn, main_fn) in enumerate(pages):
        side_fn(t.cell(i, 0)); main_fn(t.cell(i, 1))
    return t


def finish(doc, out, title=None, author="Mahesh Basavaraj", shrink=True):
    """Shrink Word's mandatory post-table paragraph to ~0 so it cannot spill a
    blank page, then save. Set shrink=False for one-column documents (letters),
    where every paragraph is real content."""
    if shrink:
        # Word needs exactly one paragraph after a table, but it inherits Normal
        # spacing (~22pt) and spills a blank page. Keep exactly one, pinned to 1pt.
        body = doc.element.body
        ps = body.findall(qn("w:p"))
        if not ps:
            new_p = OxmlElement("w:p")
            sect = body.find(qn("w:sectPr"))
            if sect is not None:
                sect.addprevious(new_p)
            else:
                body.append(new_p)
            ps = [new_p]
        keep = ps[-1]
        for extra in ps[:-1]:
            body.remove(extra)
        p = Paragraph(keep, doc)
        pf = p.paragraph_format
        pf.space_before = Pt(0); pf.space_after = Pt(0); pf.line_spacing = Pt(1)
        pPr = keep.get_or_add_pPr()
        rpr = OxmlElement("w:rPr")
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "2"); rpr.append(sz)
        szcs = OxmlElement("w:szCs"); szcs.set(qn("w:val"), "2"); rpr.append(szcs)
        rpr.append(OxmlElement("w:vanish")); pPr.append(rpr)
    props = doc.core_properties
    if title:
        props.title = title
    props.author = author
    doc.save(out)
    return out


# ------------------------------------------------------- shared fragments ---
def name_block(c, photo=None, tagline="AI Researcher \u00b7 Educator", name_size=24):
    if photo and pathlib.Path(photo).exists():
        p = para(c, first=True, after=6)
        p.add_run().add_picture(str(photo), width=Mm(28))
        first = False
    else:
        first = True
    p = para(c, first=first, after=0, line=0.95); run(p, "Mahesh", size=name_size, color=INK, bold=True, font=HEAD)
    p = para(c, after=4, line=0.95); run(p, "Basavaraj", size=name_size, color=ACC, bold=True, font=HEAD)
    p = para(c, after=0); run(p, tagline, size=8, color=ACC2, bold=True, caps=True, spacing=10)


def contact_block(c, D, ids=()):
    heading(c, "Contact")
    for lab, val in (("Location", D["location"]), ("Phone", D["phone"])):
        p = para(c, after=0); run(p, lab, size=6.5, color=MUT, caps=True, spacing=20)
        p = para(c, after=3); run(p, val, size=8.6, color=INK)
    for m in D["emails"]:
        p = para(c, after=0); run(p, "Email", size=6.5, color=MUT, caps=True, spacing=20)
        p = para(c, after=3); hyperlink(p, "mailto:" + m, m, color=INK, size=8.6, bold=False)
    if ids:
        for lab, val, url in ids:
            p = para(c, after=0); run(p, lab, size=6.5, color=MUT, caps=True, spacing=20)
            p = para(c, after=3)
            if url:
                hyperlink(p, url, val, color=INK, size=8.4, bold=False)
            else:
                run(p, val, size=8.4, color=INK)
