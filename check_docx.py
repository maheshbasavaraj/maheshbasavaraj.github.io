"""Render the DOCX through real Microsoft Word -> PDF + PNGs, report page count."""
import pathlib, sys, win32com.client, pymupdf
R = pathlib.Path(__file__).parent
src = (R / "Mahesh_Basavaraj_Resume.docx").resolve(); out = R / "_docx_render.pdf"
w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
try:
    d = w.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False)
    print("word pages:", d.ComputeStatistics(2))
    d.SaveAs2(str(out.resolve()), FileFormat=17); d.Close(False)
finally:
    w.Quit()
pdf = pymupdf.open(out); print("pdf pages:", len(pdf))
for i, p in enumerate(pdf):
    p.get_pixmap(dpi=100).save(R / f"_docx_p{i+1}.png")
    blocks = [b for b in p.get_text("blocks") if b[4].strip()]
    print(f"p{i+1} lowest text y={max(b[3] for b in blocks):.0f} of {p.rect.height:.0f}" if blocks else f"p{i+1} empty")
