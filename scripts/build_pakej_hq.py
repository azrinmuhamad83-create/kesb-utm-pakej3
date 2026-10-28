"""Himpun pakej serahan draf CP-QC kepada HQ KESB.

Guna:  py -3.12 scripts/build_pakej_hq.py
Perlu: Microsoft Word & Excel (tukar ke PDF), PyMuPDF, PyYAML, python-docx, openpyxl.
Hasil: output/Serahan-HQ-CP-QC-Rev0/ – semua dokumen dalam PDF + satu PDF gabungan berbookmark;
      fail Word & Excel asal dalam subfolder "Fail Asal (Word & Excel)"; dan fail .zip yang sama nama.
Jana dokumen sumber dahulu (build_itp_xlsx, build_item_terbuka, build_pqp_docx, build_borang, build_carta_alir).
"""
import shutil
import sys
from pathlib import Path

import fitz
import yaml
import win32com.client
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

sys.path.insert(0, str(Path(__file__).parent))
import build_pqp_docx  # noqa: E402

ROOT = Path.cwd()
PKG = ROOT / "output" / "Serahan-HQ-CP-QC-Rev0"
MS_DIR = "06 Lampiran H - Method Statement"
QF_DIR = "07 Lampiran I - Borang"
ASAL_DIR = "Fail Asal (Word & Excel)"

HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DISIPLIN = {"G": "Umum", "S": "Seni bina", "P": "Perabot", "M": "Mekanikal", "E": "Elektrik"}


def daftar_lukisan(out):
    data = yaml.safe_load((ROOT / "docs/drawings/sheets/sheet_index.yaml").read_text(encoding="utf-8"))
    wb = Workbook()
    ws = wb.active
    ws.title = "Daftar Lukisan"
    ws.append(["Lampiran C: Daftar Lukisan – UTM Baikpulih Makmal P&P, Pakej 3 (lukisan tender)"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append(["S-FK21 dan E-FKM32 disenaraikan dalam senarai lukisan tetapi tiada dalam set (ISS-2026-0001/0002). "
               "Tiada ruangan revisi pada blok tajuk."])
    ws.append(["No", "No. helaian", "Tajuk", "Disiplin", "No. lukisan", "Set PDF", "Muka surat"])
    for col, w in zip("ABCDEFG", [6, 16, 60, 12, 40, 42, 10]):
        ws.column_dimensions[col].width = w
    for c in ws[3]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.border = BORDER
    for i, p in enumerate(data["pages"], start=1):
        ws.append([i, p.get("sheet_number", ""), p.get("title", ""), DISIPLIN.get(p.get("discipline"), p.get("discipline", "")),
                   p.get("drawing_no", ""), Path(p.get("source_pdf", "")).name, p.get("page_index", 0) + 1])
        for c in ws[ws.max_row]:
            c.border = BORDER
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:G{ws.max_row}"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.print_title_rows = "3:3"
    wb.save(out)
    return len(data["pages"])


def to_pdf(word, docx_files):
    pdfs = []
    for f in docx_files:
        doc = word.Documents.Open(str(f), ReadOnly=True)
        pdf = f.with_suffix(".pdf")
        doc.SaveAs2(str(pdf), FileFormat=17)
        doc.Close(False)
        pdfs.append(pdf)
    return pdfs


def gabung(pdfs, out):
    merged = fitz.open()
    for p in pdfs:
        with fitz.open(p) as d:
            merged.insert_pdf(d)
    merged.save(out)
    n = merged.page_count
    merged.close()
    return n


def xlsx_to_pdf(excel, xlsx_files, out_dir):
    """Setiap sheet: landskap A4, muat lebar 1 halaman, panjang ikut kandungan."""
    pdfs = []
    for f in xlsx_files:
        wb = excel.Workbooks.Open(str(f), ReadOnly=True)
        try:
            for ws in wb.Worksheets:
                ps = ws.PageSetup
                ps.Orientation = 2
                ps.PaperSize = 9
                ps.Zoom = False
                ps.FitToPagesWide = 1
                ps.FitToPagesTall = False
                ps.CenterFooter = "Muka surat &P"
            pdf = out_dir / (f.stem + ".pdf")
            wb.ExportAsFixedFormat(0, str(pdf))
            pdfs.append(pdf)
        finally:
            wb.Close(False)
    return pdfs


def a4_landskap(pdf):
    """Excel mengeksport saiz halaman ikut kandungan; letak setiap halaman pada A4 landskap (skala dikekalkan)."""
    src = fitz.open(pdf)
    out = fitz.open()
    for i in range(src.page_count):
        pg = out.new_page(width=842, height=595)
        pg.show_pdf_page(fitz.Rect(20, 20, 822, 575), src, i)
    src.close()
    out.save(pdf.with_suffix(".a4.pdf"))
    out.close()
    pdf.with_suffix(".a4.pdf").replace(pdf)


def main():
    # Kosongkan fail lama sahaja; folder dikekalkan kerana mungkin sedang dibuka dalam Explorer / OneDrive
    if PKG.exists():
        for f in PKG.rglob("*"):
            if f.is_file():
                f.unlink()
        for d in sorted((p for p in PKG.rglob("*") if p.is_dir()), reverse=True):
            try:
                d.rmdir()
            except OSError:
                pass
    asal = PKG / ASAL_DIR
    (asal / MS_DIR).mkdir(parents=True, exist_ok=True)
    (asal / QF_DIR).mkdir(parents=True, exist_ok=True)

    # Fail asal (boleh diedit) dalam subfolder; PDF di aras atas pakej
    build_pqp_docx.main(ROOT / "output/memo-hq.md", asal / "00 Memo Iringan kepada HQ.docx")
    shutil.copy(ROOT / "output/Construction-Plan-QC.docx", asal / "01 Construction Plan (QC) Rev0.docx")
    shutil.copy(ROOT / "output/ITP.xlsx", asal / "02 Lampiran A - ITP.xlsx")
    n_dwg = daftar_lukisan(asal / "04 Lampiran C - Daftar Lukisan.xlsx")
    shutil.copy(ROOT / "output/Lampiran-E-Jadual-Loji.xlsx", asal / "05 Lampiran E - Jadual Loji.xlsx")
    shutil.copy(ROOT / "output/Daftar-Item-Terbuka.xlsx", asal / "08 Lampiran J - Daftar Item Terbuka.xlsx")
    shutil.copy(ROOT / "docs/Carta_Organisasi_UTM_Pakej3_LATEST_6.pdf", PKG / "03 Lampiran B - Carta Organisasi.pdf")
    shutil.copy(ROOT / "output/Lampiran-K-Carta-Alir.pdf", PKG / "09 Lampiran K - Carta Alir.pdf")
    ms = sorted((ROOT / "output/method-statement").glob("MS-*.docx"))
    qf = sorted((ROOT / "output/borang").glob("KESB-QF-*.docx"))
    for f in ms:
        shutil.copy(f, asal / MS_DIR / f.name)
    for f in qf:
        shutil.copy(f, asal / QF_DIR / f.name)

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        utama = to_pdf(word, [asal / "00 Memo Iringan kepada HQ.docx", asal / "01 Construction Plan (QC) Rev0.docx"])
        ms_pdf = to_pdf(word, sorted((asal / MS_DIR).glob("*.docx")))
        qf_pdf = to_pdf(word, sorted((asal / QF_DIR).glob("*.docx")))
    finally:
        word.Quit()
    for p in utama:
        shutil.move(p, PKG / p.name)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        xl_pdf = xlsx_to_pdf(excel, sorted(asal.glob("*.xlsx")), PKG)
    finally:
        excel.Quit()
    for p in xl_pdf:
        a4_landskap(p)

    n_ms = gabung(ms_pdf, PKG / "06 Lampiran H - Method Statement MS-01 hingga MS-14.pdf")
    n_qf = gabung(qf_pdf, PKG / "07 Lampiran I - Borang QF-01 hingga QF-29.pdf")

    # PDF gabungan lengkap dengan bookmark (MS & borang sebagai sub-bookmark)
    sub = {"06": [(p.stem, p) for p in ms_pdf], "07": [(p.stem, p) for p in qf_pdf]}
    lengkap, toc = fitz.open(), []
    for p in sorted(PKG.glob("0*.pdf")):
        toc.append([1, p.stem, lengkap.page_count + 1])
        if p.name[:2] in sub:
            for title, sp in sub[p.name[:2]]:
                toc.append([2, title, lengkap.page_count + 1])
                with fitz.open(sp) as d:
                    lengkap.insert_pdf(d)
        else:
            with fitz.open(p) as d:
                mula = lengkap.page_count
                toc += [[lvl + 1, t, mula + pg] for lvl, t, pg in d.get_toc()]
                lengkap.insert_pdf(d)
    lengkap.set_toc(toc)
    out_all = PKG / "99 Pakej Lengkap CP-QC Rev0 (gabungan).pdf"
    lengkap.save(out_all)
    n_all = lengkap.page_count
    lengkap.close()
    for p in ms_pdf + qf_pdf:
        p.unlink()

    zip_path = shutil.make_archive(str(PKG), "zip", PKG.parent, PKG.name)
    for p in sorted(PKG.glob("0*.pdf")):
        with fitz.open(p) as d:
            print(f"{p.name}: {d.page_count} hlm")
    print(f"Method statement: {len(ms)} fail, {n_ms} hlm; borang: {len(qf)} fail, {n_qf} hlm; daftar lukisan: {n_dwg} helaian")
    print(f"Gabungan: {n_all} hlm; Excel ditukar: {len(xl_pdf)}")
    print(f"Siap: {PKG}\n      {zip_path}")


if __name__ == "__main__":
    main()
