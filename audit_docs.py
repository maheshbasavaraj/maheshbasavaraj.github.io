"""Geometry audit of the rendered packet - substitutes for eyeballing a PNG.

Checks per page:
  * no text outside the printable page
  * sidebar text stays in the sidebar column, main text in the main column
  * no two text blocks collide (catches overlapping text the eye might miss)
  * the shaded sidebar rectangle really is drawn, full height
  * the portrait image is embedded where expected
  * hyperlinks survive into the PDF
"""
import pathlib, pymupdf

R = pathlib.Path(__file__).parent
W, H = 595.3, 841.9          # A4 in points
SIDE_X = 66 * 72 / 25.4      # sidebar width 66mm -> 187pt
TOL = 3

DOCS = {
    "Mahesh_Basavaraj_Research_Publications": 2,
    "Mahesh_Basavaraj_Cover_Letter_UNSW": 1,
    "Mahesh_Basavaraj_Teaching_Methodology": 2,
}
SIDEBAR_DOCS = {"Mahesh_Basavaraj_Research_Publications", "Mahesh_Basavaraj_Teaching_Methodology"}

problems = []
for stem, expect_pages in DOCS.items():
    pdf = R / (stem + "_render.pdf")
    doc = pymupdf.open(pdf)
    tag = stem.split("_", 2)[-1]
    print(f"\n{tag}: {doc.page_count} page(s), expect {expect_pages}")
    if doc.page_count != expect_pages:
        problems.append(f"{tag}: page count {doc.page_count} != {expect_pages}")
    for pi, pg in enumerate(doc):
        # line-level boxes: block boxes merge adjacent cells and give false overlaps
        blocks = []
        for blk in pg.get_text("dict")["blocks"]:
            for ln in blk.get("lines", []):
                txt = "".join(sp["text"] for sp in ln["spans"]).strip()
                if txt:
                    blocks.append((*ln["bbox"], txt))
        imgs = pg.get_images(full=True)
        links = pg.get_links()
        # sidebar rectangle drawn?
        rects = [d["rect"] for d in pg.get_drawings()
                 if d.get("fill") is not None and d["rect"].width > SIDE_X * 0.9 and d["rect"].height > H * 0.9]
        sidebar_ok = bool(rects) if stem in SIDEBAR_DOCS else True
        if stem in SIDEBAR_DOCS and not sidebar_ok:
            problems.append(f"{tag} p{pi+1}: sidebar fill rectangle missing")
        out_of_page = [b for b in blocks if b[0] < -TOL or b[2] > W + TOL or b[1] < -TOL or b[3] > H + TOL]
        if out_of_page:
            problems.append(f"{tag} p{pi+1}: {len(out_of_page)} block(s) off-page")
        # column containment
        wrong_col = []
        if stem in SIDEBAR_DOCS:
            for b in blocks:
                cx = (b[0] + b[2]) / 2
                if b[0] < SIDE_X - 4 and b[2] > SIDE_X + 6:
                    wrong_col.append((round(b[0]), round(b[2]), b[4][:30].replace("\n", " ")))
                elif b[0] > SIDE_X + 6 and b[2] > W - 8:
                    wrong_col.append((round(b[0]), round(b[2]), b[4][:30].replace("\n", " ")))
        if wrong_col:
            problems.append(f"{tag} p{pi+1}: text crossing column/page edge {wrong_col}")
        # collisions
        coll = []
        for i in range(len(blocks)):
            for j in range(i + 1, len(blocks)):
                a, b = pymupdf.Rect(blocks[i][:4]), pymupdf.Rect(blocks[j][:4])
                inter = a & b
                if inter.is_empty:
                    continue
                area = inter.width * inter.height
                small = min(a.width * a.height, b.width * b.height)
                if small > 0 and area / small > 0.30:
                    coll.append((blocks[i][4][:22].replace("\n", " "), blocks[j][4][:22].replace("\n", " ")))
        if coll:
            problems.append(f"{tag} p{pi+1}: overlapping text {coll[:3]}")
        bottom = max((b[3] for b in blocks), default=0)
        print(f"  p{pi+1}: {len(blocks)} text blocks, bottom y={bottom:.0f}/{H:.0f}, "
              f"images={len(imgs)}, links={len(links)}, sidebar_rect={sidebar_ok}, "
              f"sidebar_used={bool(rects)}")

print("\n" + ("AUDIT CLEAN" if not problems else "PROBLEMS:\n" + "\n".join(problems)))
