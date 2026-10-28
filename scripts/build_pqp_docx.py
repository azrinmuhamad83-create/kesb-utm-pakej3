"""Tukar output/CP-QC-draf.md jadi output/Construction-Plan-QC.docx (Word).

Guna:  py -3.12 scripts/build_pqp_docx.py
       py -3.12 scripts/build_pqp_docx.py <fail.md> [<fail.docx>]   (fail lain, cth method statement)
Sokong: # ## ### tajuk, perenggan, senarai (- / 1.), jadual markdown, **tebal**.
"""
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

SRC = Path("output/CP-QC-draf.md")
OUT = Path("output/Construction-Plan-QC.docx")


def add_inline(par, text):
    """Tambah teks dengan sokongan **tebal**; tanda [PERLU SAHKAN] dalam merah."""
    for part in re.split(r"(\*\*[^*]+\*\*|\[PERLU SAHKAN[^\]]*\])", text):
        if not part:
            continue
        if part.startswith("**"):
            par.add_run(part[2:-2]).bold = True
        elif part.startswith("[PERLU SAHKAN"):
            run = par.add_run(part)
            run.bold = True
            run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
        else:
            par.add_run(part)


def add_page_number(par):
    run = par.add_run()
    for kind, text in (("begin", None), (None, "PAGE"), ("end", None)):
        if kind:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = text
        run._r.append(el)


def add_table(doc, lines, lebar_mm=None):
    rows = [[c.strip() for c in l.strip().strip("|").split("|")] for l in lines
            if not re.match(r"^\s*\|?\s*:?-{2,}", l)]
    table = doc.add_table(rows=len(rows), cols=max(len(r) for r in rows))
    table.style = "Table Grid"
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.text = ""
            add_inline(cell.paragraphs[0], val)
            if i == 0:
                for r in cell.paragraphs[0].runs:
                    r.bold = True
                shade = OxmlElement("w:shd")
                shade.set(qn("w:fill"), "DDEBF7")
                cell._tc.get_or_add_tcPr().append(shade)
    # Lebar lajur ikut panjang teks (lajur huraian lebih lebar), dalam lebar teks A4 159 mm
    ncol = len(table.columns)
    if lebar_mm and len(lebar_mm) == ncol:
        table.autofit = False
        for j, w in enumerate(lebar_mm):
            for cell in table.columns[j].cells:
                cell.width = Mm(w)
    elif ncol > 1:
        berat = [max(1.0, sum(len(r[j]) for r in rows if j < len(r)) ** 0.5) for j in range(ncol)]
        jumlah = sum(berat)
        lebar = [max(24.0, 159.0 * b / jumlah) for b in berat]
        skala = 159.0 / sum(lebar)
        table.autofit = False
        for j, w in enumerate(lebar):
            for cell in table.columns[j].cells:
                cell.width = Mm(w * skala)
    doc.add_paragraph()


def main(src=SRC, out=OUT):
    lines = src.read_text(encoding="utf-8").splitlines()
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(10.5)

    title = next((l[2:].strip() for l in lines if l.startswith("# ")), "Construction Plan (Quality Control)")
    section = doc.sections[0]
    section.page_width, section.page_height = Mm(210), Mm(297)  # A4
    section.header.paragraphs[0].text = title
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run("Muka surat ")
    add_page_number(footer)

    i = 0
    lebar = None  # lebar lajur (mm) untuk jadual seterusnya: baris "<!-- lebar: 8, 12, 40 -->"
    while i < len(lines):
        line = lines[i].rstrip()
        if m := re.match(r"^<!--\s*lebar:\s*([\d.,\s]+)-->", line):
            lebar = [float(x) for x in m.group(1).split(",") if x.strip()]
            i += 1
            continue
        if line.startswith("|"):
            block = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                block.append(lines[i])
                i += 1
            add_table(doc, block, lebar)
            lebar = None
            continue
        if m := re.match(r"^(#{1,3})\s+(.*)", line):
            level = len(m.group(1))
            if level == 1:
                p = doc.add_heading(m.group(2), level=0)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                doc.add_heading(m.group(2), level=level - 1)
        elif m := re.match(r"^\s*[-*]\s+(.*)", line):
            add_inline(doc.add_paragraph(style="List Bullet"), m.group(1))
        elif m := re.match(r"^\s*\d+[.)]\s+(.*)", line):
            add_inline(doc.add_paragraph(style="List Number"), m.group(1))
        elif line.strip() == "---":
            doc.add_page_break()
        elif line.strip():
            add_inline(doc.add_paragraph(), line.strip())
        i += 1

    doc.save(out)
    print(f"Siap: {out}")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        src = Path(sys.argv[1])
        main(src, Path(sys.argv[2]) if len(sys.argv) > 2 else src.with_suffix(".docx"))
    else:
        main()
