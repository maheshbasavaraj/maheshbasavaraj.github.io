"""Render each DOCX through Microsoft Word -> PDF, then prove nothing was clipped:
every word that exists in the DOCX must also appear in the rendered PDF text layer.
A fixed-height table row hides overflow silently, so this is the real test.
"""
import pathlib, re, sys, win32com.client, pymupdf
from docx import Document

R = pathlib.Path(__file__).parent
DOCS = ["Mahesh_Basavaraj_Research_Publications.docx",
        "Mahesh_Basavaraj_Cover_Letter_UNSW.docx",
        "Mahesh_Basavaraj_Teaching_Methodology.docx"]


def docx_words(path):
    d = Document(str(path))
    out = []
    for p in d.paragraphs:
        out.append(p.text)
    for t in d.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    out.append(p.text)
                for t2 in cell.tables:
                    for r2 in t2.rows:
                        for c2 in r2.cells:
                            out.extend(p.text for p in c2.paragraphs)
    return " ".join(out)


def norm(s):
    s = s.replace("\u00ad", "").replace("\u2019", "'").replace("\u201c", '"').replace("\u201d", '"')
    s = s.replace("\u2013", "-").replace("\u2014", "-").replace("\u00b7", " ").replace("\u00a0", " ")
    return re.findall(r"[A-Za-z0-9']+", s.lower())


word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
ok = True
try:
    for name in DOCS:
        src = R / name
        pdf = R / (src.stem + "_render.pdf")
        d = word.Documents.Open(str(src), ReadOnly=False)
        d.SaveAs(str(pdf), FileFormat=17)
        pages = d.ComputeStatistics(2)
        # how much room is left at the bottom of each page
        d.Close(False)
        doc = pymupdf.open(pdf)
        ptext = "".join(pg.get_text() for pg in doc)
        # collapse ALL whitespace: letter-spaced headings (C O N T A C T) and
        # line-broken words both normalise to the same string this way
        flat = "".join(norm(ptext))
        missing = []
        for w in norm(docx_words(src)):
            if len(w) > 2 and w not in flat:
                missing.append(w)
        missing = sorted(set(missing))
        bottom = max((max((b[3] for b in pg.get_text("blocks")), default=0) for pg in doc), default=0)
        low = [round(max((b[3] for b in pg.get_text("blocks")), default=0)) for pg in doc]
        status = "OK" if not missing and pages == doc.page_count else "PROBLEM"
        if status != "OK":
            ok = False
        print(f"{name}\n  word pages={pages} pdf pages={doc.page_count}  {status}")
        print(f"  lowest text y per page: {low} (page height {round(doc[0].rect.height)})")
        if missing:
            print(f"  MISSING FROM PDF ({len(missing)}): {' '.join(missing[:40])}")
        doc.close()
finally:
    word.Quit()
print("\nALL CLEAN" if ok else "\nFIX NEEDED")
