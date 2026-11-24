"""Jana dokumen Word ikut Sistem Reka Bentuk Dokumen KESB (sistem-reka-bentuk/index.html).

Guna:  py -3.12 scripts/build_reka_bentuk_docx.py            (jana panduan gaya + templat PQP & SOM, dan
                                                              isi semula blok AUTO dalam index.html)
       py -3.12 scripts/build_reka_bentuk_docx.py <fail.md> [<fail.docx>]

Fail .md bermula dengan blok meta (antara dua baris ---):
    jenis: PROJECT QUALITY PLAN
    tajuk: Construction Plan (Quality Control)
    no: KESB/UTM-P3/CP-QC/001
    rev: 0
    tarikh: __/__/2026
    disediakan / disemak / diluluskan / diterima: nama, jawatan
Kulit dan blok kawalan dijana daripada meta. Badan menyokong: # ## ### tajuk, perenggan, senarai (- / 1.),
jadual markdown, **tebal**, `kod`, [PERLU SAHKAN ...] (merah), "> Nota: ..." (kotak nota), "> ! ..." (amaran),
"<!-- lebar: 10, 30, ... -->" (lebar lajur mm), "<!-- sertakan: fail.md -->" (sisip fail bersama),
"---" (muka surat baharu). Jadual dengan baris pengepala kosong = jadual label (lajur 1 kelabu);
baris yang sel pertamanya **tebal** dan sel lain kosong = jalur kategori penuh lebar.
"""
import html
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

# ===== Token reka bentuk (sama dengan sistem-reka-bentuk/index.html) =====
FONT = "Arial"
NAVY = "1F4E78"
BIRU_PUCAT = "DDEBF7"
BIRU_SEDERHANA = "BDD7EE"
KELABU_LABEL = "F2F2F2"
KELABU_TEKS = "595959"
MERAH = "C00000"
LEBAR_TEKS = 159.0  # mm, A4 dengan margin kiri/kanan 25 mm (+1 mm toleransi)
SYARIKAT = "KARIM ENGINEERING SDN BHD"
PROJEK = ("UTM – Baikpulih Makmal P&P di Bangunan Akademik (RP1-RMKe-13)",
          "Pakej 3: Fakulti Komputeran (FKOM) & Fakulti Kejuruteraan Mekanikal (FKM)",
          "No. kontrak TN/2026-21/D-1/SP/BPA")

ROOT = Path(__file__).resolve().parent.parent
DIR = ROOT / "sistem-reka-bentuk"
LALAI = [
    (DIR / "sumber/panduan-gaya.md", DIR / "Panduan-Gaya-Dokumen-KESB.docx"),
    (DIR / "sumber/templat-pqp.md", DIR / "Templat-PQP.docx"),
    (DIR / "sumber/templat-som.md", DIR / "Templat-SOM.docx"),
]

# Unsur pPr yang mesti selepas w:pBdr / w:shd (susunan skema OOXML)
PPR_SELEPAS_SHD = ("w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
                   "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
                   "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
                   "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
                   "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")


def rgb(hexa):
    return RGBColor.from_string(hexa)


def shade(cell, color):
    el = OxmlElement("w:shd")
    el.set(qn("w:val"), "clear")
    el.set(qn("w:fill"), color)
    cell._tc.get_or_add_tcPr().append(el)


def para_border(par_or_style, side, color, sz=6, space=4):
    """Garis sempadan perenggan (cth garis bawah pengepala, bingkai kiri nota)."""
    ppr = par_or_style.element.get_or_add_pPr() if hasattr(par_or_style, "element") else par_or_style._p.get_or_add_pPr()
    bdr = ppr.find(qn("w:pBdr"))
    if bdr is None:
        bdr = OxmlElement("w:pBdr")
        ppr.insert_element_before(bdr, *PPR_SELEPAS_SHD, "w:shd")
    el = OxmlElement(f"w:{side}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), str(space))
    el.set(qn("w:color"), color)
    bdr.append(el)


def para_shade(style, color):
    ppr = style.element.get_or_add_pPr()
    el = OxmlElement("w:shd")
    el.set(qn("w:val"), "clear")
    el.set(qn("w:fill"), color)
    ppr.insert_element_before(el, *PPR_SELEPAS_SHD)


def set_font(style, size=None, bold=None, italic=None, color=None):
    style.font.name = FONT
    rpr = style.element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.append(fonts)
    for k in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        fonts.attrib.pop(qn(k), None)  # fon tema (Aptos/Calibri) mengatasi Arial
    for k in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(k), FONT)
    if size:
        style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if italic is not None:
        style.font.italic = italic
    if color:
        style.font.color.rgb = rgb(color)


def setup_styles(doc):
    st = doc.styles
    normal = st["Normal"]
    set_font(normal, 10.5)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.15

    h1 = st["Heading 1"]
    set_font(h1, 13, True, False, NAVY)
    h1.font.all_caps = True
    h1.paragraph_format.space_before = Pt(18)
    h1.paragraph_format.space_after = Pt(8)
    h1.paragraph_format.page_break_before = True
    h1.paragraph_format.keep_with_next = True
    para_border(h1, "bottom", NAVY, sz=8, space=2)

    h2 = st["Heading 2"]
    set_font(h2, 11, True, False, "000000")
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after = Pt(4)
    h2.paragraph_format.keep_with_next = True

    h3 = st["Heading 3"]
    set_font(h3, 10.5, True, True, "000000")
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(3)
    h3.paragraph_format.keep_with_next = True

    for name in ("List Bullet", "List Number"):
        set_font(st[name], 10.5)
        st[name].paragraph_format.space_after = Pt(3)

    for name, warna, isi in (("Nota KESB", NAVY, "F4F8FC"), ("Amaran KESB", MERAH, "FDF3F3")):
        s = st.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        s.base_style = normal
        set_font(s, 10)
        s.paragraph_format.left_indent = Mm(3)
        s.paragraph_format.space_before = Pt(4)
        s.paragraph_format.space_after = Pt(8)
        para_border(s, "left", warna, sz=24, space=6)
        para_shade(s, isi)

    s = st.add_style("Jadual KESB", WD_STYLE_TYPE.PARAGRAPH)
    s.base_style = normal
    set_font(s, 9.5)
    s.paragraph_format.space_after = Pt(0)
    s.paragraph_format.line_spacing = 1.0

    s = st.add_style("PERLU SAHKAN", WD_STYLE_TYPE.CHARACTER)
    s.font.bold = True
    s.font.color.rgb = rgb(MERAH)


def add_inline(par, text):
    for part in re.split(r"(\*\*[^*]+\*\*|\[PERLU SAHKAN[^\]]*\]|`[^`]+`)", text):
        if not part:
            continue
        if part.startswith("**"):
            par.add_run(part[2:-2]).bold = True
        elif part.startswith("[PERLU SAHKAN"):
            par.add_run(part, style="PERLU SAHKAN")
        elif part.startswith("`"):
            r = par.add_run(part[1:-1])
            r.font.name = "Consolas"
            r.font.size = Pt(9.5)
        else:
            par.add_run(part)


def field(run, code):
    for kind, text in (("begin", None), (None, code), ("end", None)):
        if kind:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = text
        run._r.append(el)


def repeat_header(row):
    trpr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trpr.append(el)


def table_cell_text(cell, text, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.style = "Jadual KESB"
    add_inline(p, text)
    if bold:
        for r in p.runs:
            r.bold = True


def add_table(doc, lines, lebar_mm=None):
    rows = [[c.strip() for c in l.strip().strip("|").split("|")] for l in lines
            if not re.match(r"^\s*\|?\s*:?-{2,}", l)]
    ncol = max(len(r) for r in rows)
    label = rows[0] and all(not c for c in rows[0])  # pengepala kosong -> jadual label (lajur 1 kelabu)
    if label:
        rows = rows[1:]
    t = doc.add_table(rows=len(rows), cols=ncol)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j in range(ncol):
            val = row[j] if j < len(row) else ""
            cell = t.cell(i, j)
            if label:
                table_cell_text(cell, val, bold=(j == 0))
                if j == 0:
                    shade(cell, KELABU_LABEL)
            else:
                table_cell_text(cell, val, bold=(i == 0))
                if i == 0:
                    shade(cell, BIRU_PUCAT)
    if not label:
        repeat_header(t.rows[0])
    if not (lebar_mm and len(lebar_mm) == ncol):
        if label and ncol == 2:
            lebar_mm = [50, LEBAR_TEKS - 50]
        else:
            berat = [max(1.0, sum(len(r[j]) for r in rows if j < len(r)) ** 0.5) for j in range(ncol)]
            lebar_mm = [max(14.0, LEBAR_TEKS * b / sum(berat)) for b in berat]
    skala = LEBAR_TEKS / sum(lebar_mm)
    t.autofit = False
    for j, w in enumerate(lebar_mm):
        for cell in t.columns[j].cells:
            cell.width = Mm(w * skala)
    if not label:
        # baris kategori: sel pertama **tebal** dan sel lain kosong -> jalur penuh lebar biru sederhana
        for i, row in enumerate(rows[1:], start=1):
            if row[0].startswith("**") and all(not c for c in row[2:]):
                teks = " ".join(c.strip("*") for c in row if c)
                merged = t.cell(i, 0).merge(t.cell(i, ncol - 1))
                table_cell_text(merged, teks, bold=True)
                shade(merged, BIRU_SEDERHANA)
    gap = doc.add_paragraph()
    gap.paragraph_format.space_after = Pt(0)
    gap.paragraph_format.line_spacing = Pt(6)


def read_meta(lines):
    meta = {}
    if lines and lines[0].strip() == "---":
        end = lines.index("---", 1)
        for l in lines[1:end]:
            if ":" in l:
                k, v = l.split(":", 1)
                meta[k.strip().lower()] = v.strip()
        lines = lines[end + 1:]
    return meta, lines


def expand_includes(lines, base):
    """Baris "<!-- sertakan: fail.md -->" diganti dengan isi fail itu (cth daftar perundangan bersama)."""
    out = []
    for l in lines:
        if m := re.match(r"^<!--\s*sertakan:\s*(.+?)\s*-->", l):
            out += (base / m.group(1)).read_text(encoding="utf-8").splitlines()
        else:
            out.append(l)
    return out


def header_footer(section, meta):
    section.different_first_page_header_footer = True  # kulit tanpa pengepala/pengaki
    hp = section.header.paragraphs[0]
    hp.text = ""
    hp.paragraph_format.tab_stops.add_tab_stop(Mm(LEBAR_TEKS), alignment=2)  # kanan
    r = hp.add_run(SYARIKAT)
    r.bold = True
    r.font.size = Pt(8)
    r.font.color.rgb = rgb(NAVY)
    r = hp.add_run(f"\t{meta.get('no', '')}  ·  Rev {meta.get('rev', '0')}")
    r.font.size = Pt(8)
    r.font.color.rgb = rgb(KELABU_TEKS)
    para_border(hp, "bottom", NAVY, sz=6, space=1)

    fp = section.footer.paragraphs[0]
    fp.text = ""
    fp.paragraph_format.tab_stops.add_tab_stop(Mm(LEBAR_TEKS), alignment=2)
    para_border(fp, "top", "999999", sz=4, space=1)
    parts = [("Salinan cetak tanpa cop SALINAN TERKAWAL adalah salinan tidak terkawal\tMuka surat ", None),
             (None, "PAGE"), (" daripada ", None), (None, "NUMPAGES")]
    for text, code in parts:
        r = fp.add_run(text or "")
        r.font.size = Pt(8)
        r.font.color.rgb = rgb(KELABU_TEKS)
        if code:
            field(r, code)


def cover(doc, meta):
    def line(text, size, bold=False, color=None, before=0, after=0, align=WD_ALIGN_PARAGRAPH.CENTER):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        r = p.add_run(text)
        r.font.size = Pt(size)
        r.bold = bold
        if color:
            r.font.color.rgb = rgb(color)
        return p

    line(SYARIKAT, 12, True, before=24)
    # Jalur navy: jenis dokumen + tajuk
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = t.cell(0, 0)
    c.width = Mm(LEBAR_TEKS)
    shade(c, NAVY)
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    r = p.add_run(meta.get("jenis", "").upper())
    r.font.size = Pt(9)
    r.font.color.rgb = rgb("FFFFFF")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:val"), "40")
    r._r.get_or_add_rPr().append(spacing)
    p = c.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(14)
    r = p.add_run(meta.get("tajuk", ""))
    r.font.size = Pt(16)
    r.bold = True
    r.font.color.rgb = rgb("FFFFFF")

    line(PROJEK[0], 11, before=36, after=2)
    line(PROJEK[1], 11, True, after=2)
    line(PROJEK[2], 10, after=110)

    rows = [("No. dokumen", meta.get("no", "")),
            ("Revisi / tarikh", f"{meta.get('rev', '0')} / {meta.get('tarikh', '')}"),
            ("Disediakan oleh", meta.get("disediakan", "")),
            ("Disemak oleh", meta.get("disemak", "")),
            ("Diluluskan oleh", meta.get("diluluskan", "")),
            ("Diterima oleh", meta.get("diterima", ""))]
    rows = [r for r in rows if r[1]]
    add_table(doc, ["|  |  |", "|---|---|"] + [f"| {k} | {v} |" for k, v in rows], [50, LEBAR_TEKS - 50])
    p = line("No. salinan terkawal: ______", 9, color=KELABU_TEKS, align=WD_ALIGN_PARAGRAPH.RIGHT, before=6)
    p.runs[0].add_break(WD_BREAK.PAGE)


def build(src, out):
    lines = src.read_text(encoding="utf-8").splitlines()
    meta, lines = read_meta(lines)
    lines = expand_includes(lines, src.parent)
    doc = Document()
    setup_styles(doc)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.top_margin, sec.bottom_margin = Mm(25), Mm(20)
    sec.left_margin, sec.right_margin = Mm(25), Mm(25)
    sec.header_distance = sec.footer_distance = Mm(10)
    header_footer(sec, meta)
    if meta.get("tajuk"):
        cover(doc, meta)

    first_h1 = True
    lebar = None
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        if m := re.match(r"^<!--\s*lebar:\s*([\d.,\s]+)-->", line):
            lebar = [float(x) for x in m.group(1).split(",") if x.strip()]
        elif line.startswith("|"):
            block = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                block.append(lines[i])
                i += 1
            add_table(doc, block, lebar)
            lebar = None
            continue
        elif m := re.match(r"^(#{1,3})\s+(.*)", line):
            level = len(m.group(1))
            p = doc.add_heading(m.group(2), level=level)
            if level == 1 and first_h1:
                p.paragraph_format.page_break_before = False  # kulit sudah berakhir dengan pemisah muka surat
                first_h1 = False
        elif m := re.match(r"^>\s*!\s*(.*)", line):
            add_inline(doc.add_paragraph(style="Amaran KESB"), m.group(1))
        elif m := re.match(r"^>\s*(.*)", line):
            add_inline(doc.add_paragraph(style="Nota KESB"), m.group(1))
        elif m := re.match(r"^\s*[-*]\s+(.*)", line):
            add_inline(doc.add_paragraph(style="List Bullet"), m.group(1))
        elif m := re.match(r"^\s*\d+[.)]\s+(.*)", line):
            add_inline(doc.add_paragraph(style="List Number"), m.group(1))
        elif line.strip() == "---":
            doc.add_page_break()
        elif line.strip():
            add_inline(doc.add_paragraph(), line.strip())
        i += 1

    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    print(f"Siap: {out}")


def md_ke_html(text):
    """Tukar serpihan .md bersama (jadual, perenggan, nota) kepada HTML untuk index.html."""
    def fmt(t):
        t = html.escape(t)
        t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
        t = re.sub(r"`([^`]+)`", r"<code>\1</code>", t)
        return re.sub(r"(\[PERLU SAHKAN[^\]]*\])", r'<span class="psah">\1</span>', t)

    out, lines, i = [], text.splitlines(), 0
    while i < len(lines):
        l = lines[i].strip()
        if l.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                if not re.match(r"^\|?\s*:?-{2,}", lines[i].strip()):
                    rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            n = max(len(r) for r in rows)
            t = ['<table class="doc legal">', "<tr>" + "".join(f"<th>{fmt(c)}</th>" for c in rows[0]) + "</tr>"]
            for r in rows[1:]:
                if r[0].startswith("**") and not any(r[2:]):
                    t.append(f'<tr class="kat"><td colspan="{n}">{fmt(" ".join(c for c in r if c))}</td></tr>')
                else:
                    t.append("<tr>" + "".join(f"<td>{fmt(c)}</td>" for c in r) + "</tr>")
            out.append("\n".join(t + ["</table>"]))
            continue
        if m := re.match(r"^>\s*!\s*(.*)", l):
            out.append(f'<div class="callout r">{fmt(m.group(1))}</div>')
        elif m := re.match(r"^>\s*(.*)", l):
            out.append(f'<div class="callout">{fmt(m.group(1))}</div>')
        elif l and not l.startswith("<!--"):
            out.append(f"<p>{fmt(l)}</p>")
        i += 1
    return "\n".join(out)


def kemas_kini_html():
    """Isi semula blok <!-- AUTO:fail.md --> ... <!-- /AUTO --> dalam index.html daripada sumber .md."""
    page = DIR / "index.html"
    s = page.read_text(encoding="utf-8")

    def isi(m):
        body = md_ke_html((DIR / "sumber" / m.group(2)).read_text(encoding="utf-8"))
        return f"{m.group(1)}\n{body}\n{m.group(3)}"

    s = re.sub(r"(<!-- AUTO:(\S+) -->).*?(<!-- /AUTO -->)", isi, s, flags=re.S)
    page.write_text(s, encoding="utf-8")
    print(f"Siap: {page} (blok AUTO dikemas kini)")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        src = Path(sys.argv[1])
        build(src, Path(sys.argv[2]) if len(sys.argv) > 2 else src.with_suffix(".docx"))
    else:
        for src, out in LALAI:
            build(src, out)
        kemas_kini_html()
