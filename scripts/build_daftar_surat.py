"""Jana Daftar Surat-Menyurat UTM Pakej 3 (Excel): Surat Keluar dan Surat Masuk.

Guna:  py -3.12 scripts/build_daftar_surat.py           (cipta fail jika belum ada)
       py -3.12 scripts/build_daftar_surat.py --tambah  (masukkan surat baharu dalam output/surat/ ke baris kosong
                                                          Surat Keluar; baris sedia ada tidak diusik)
       py -3.12 scripts/build_daftar_surat.py --paksa   (tulis ganti – SEMUA isian Kerani hilang)
Hasil: output/Daftar-Surat-UTM.xlsx

Fail ini rekod hidup. Maklumat surat (ruj., tarikh, perkara, lampiran, fail) diambil daripada fail .md dalam output/surat/
yang ada baris "Ruj. Kami". Tarikh hantar, cara hantar dan akuan terima UTM diisi oleh Kerani Tapak / PM dalam Excel.
Surat bertanda DRAF belum dikeluarkan – ia hanya dianggap keluar selepas diluluskan PM.
"""
import re
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path("output/Daftar-Surat-UTM.xlsx")
SURAT = Path("output/surat")
HEAD = PatternFill("solid", fgColor="1F4E78")
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
TARIKH = "DD/MM/YYYY"
BARIS_KOSONG = 30

CARA = ["Serahan tangan", "Pos berdaftar", "Kurier", "E-mel", "Portal UTM"]
STATUS_KELUAR = ["Draf", "Disemak PM", "Diluluskan PM", "Dihantar", "Akuan diterima", "Dibatalkan"]
STATUS_MASUK = ["Baharu", "Perlu tindakan", "Dalam tindakan", "Dibalas", "Untuk rekod"]

KELUAR = [  # (tajuk, lebar)
    ("Bil", 5), ("Ruj. Kami", 26), ("Tarikh surat", 12), ("Perkara", 52), ("Lampiran", 28), ("Status", 15),
    ("Tarikh hantar", 12), ("Cara hantar", 15), ("Akuan terima UTM (Ya / Tidak)", 14), ("Tarikh akuan", 12),
    ("Nama penerima / cop UTM", 22), ("Balasan UTM (Ruj. Tuan)", 24), ("Fail (output/surat/)", 34), ("Catatan", 30),
]
MASUK = [
    ("Bil", 5), ("Ruj. Tuan (UTM)", 28), ("Tarikh surat", 12), ("Tarikh terima", 12), ("Daripada", 24), ("Perkara", 52),
    ("Tindakan diperlukan", 34), ("Tarikh akhir jawapan", 13), ("Ruj. balasan KESB", 24), ("Status", 15),
    ("Fail / imbasan", 30), ("Catatan", 30),
]

# Surat masuk sedia ada dalam rekod projek (CLAUDE.md – SST).
MASUK_AWAL = [[
    1, "UTM.J.17.03/16.11/667(4)", date(2026, 8, 10), date(2026, 8, 14), "Universiti Teknologi Malaysia",
    "Surat Setuju Terima (SST) – Baikpulih Makmal P & P di Bangunan Akademik (RP1-RMKe-13), Pakej 3, No. Kontrak TN/2026-21/D-1/SP/BPA",
    "Syarat SST: mula kerja ≤ 2 minggu dari milik tapak; polisi insurans ≤ 30 hari; senarai subkon & Program Kerja sebelum mula kerja",
    None, "", "Untuk rekod", "docs/SST UTM JOHOR.pdf (hlm 1–6; Lampiran A tiada)", "Diakui KESB 14/08/2026",
]]


def baca_surat(fail):
    """Ambil maklumat daripada satu fail .md surat. Pulangkan None jika bukan surat berkepala rujukan."""
    teks = fail.read_text(encoding="utf-8")
    ruj = re.search(r"^Ruj\.\s*Kami\s*:\s*(.+)$", teks, re.M)
    if not ruj:
        return None
    tarikh = re.search(r"^Tarikh\s*:\s*(.+)$", teks, re.M)
    tajuk = re.search(r"^#\s+(.+)$", teks, re.M)
    perkara = tajuk.group(1).strip() if tajuk else fail.stem
    draf = bool(re.match(r"draf\b", perkara, re.I))
    perkara = re.sub(r"^draf\s+surat\s*[–-]\s*", "", perkara, flags=re.I)
    # Lampiran: baris "Lampiran X" / "LAMPIRAN X" (tajuk atau perenggan)
    lamp = []
    for m in re.finditer(r"^\W*(lampiran\s+[A-Z0-9]+\b[^\n]*)$", teks, re.M | re.I):
        s = m.group(1).strip(" *#")
        if s not in lamp:
            lamp.append(s)
    ruj = ruj.group(1).strip()
    tarikh = tarikh.group(1).strip() if tarikh else ""
    return {
        "ruj": "" if "[" in ruj else ruj,
        "tarikh": "" if "[" in tarikh else tarikh,
        "perkara": perkara,
        "lampiran": "; ".join(lamp),
        "status": "Draf" if draf else "",
        "fail": fail.name,
    }


def senarai_surat():
    if not SURAT.exists():
        return []
    out = [baca_surat(f) for f in sorted(SURAT.glob("*.md"))]
    return [s for s in out if s]


def tarikh_sel(nilai):
    """Tukar 'dd/mm/yyyy' kepada tarikh sebenar; jika tidak boleh, biar sebagai teks."""
    m = re.fullmatch(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})", nilai or "")
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    return nilai or None


def baris_keluar(n, s):
    return [n, s["ruj"], tarikh_sel(s["tarikh"]), s["perkara"], s["lampiran"], s["status"], None, "", "", None,
            "", "", s["fail"], ""]


def gaya_tajuk(ws, baris, lajur):
    for cell, (_, w) in zip(ws[baris], lajur):
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEAD
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        cell.border = BORDER
        ws.column_dimensions[cell.column_letter].width = w


def gaya_baris(ws, tarikh_lajur):
    for c in ws[ws.max_row]:
        c.alignment = WRAP
        c.border = BORDER
    for col in tarikh_lajur:
        ws[f"{col}{ws.max_row}"].number_format = TARIKH


def susun_cetak(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.print_title_rows = "3:3"


def sheet_keluar(wb, surat):
    ws = wb.active
    ws.title = "Surat Keluar"
    ws.append(["Daftar Surat Keluar – UTM Baikpulih Makmal P&P, Pakej 3 (KESB/UTM-P3) – No. Kontrak TN/2026-21/D-1/SP/BPA"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append(["Maklumat surat diambil daripada output/surat/. Isi tarikh hantar, cara hantar dan akuan terima selepas surat dikeluarkan. "
               "Kuning = dihantar lebih 7 hari belum ada akuan UTM; kelabu = draf."])
    ws.append([t for t, _ in KELUAR])
    gaya_tajuk(ws, 3, KELUAR)
    for n, s in enumerate(surat, start=1):
        ws.append(baris_keluar(n, s))
        gaya_baris(ws, "CGJ")
    for n in range(len(surat) + 1, len(surat) + BARIS_KOSONG + 1):
        ws.append([n] + [""] * (len(KELUAR) - 1))
        gaya_baris(ws, "CGJ")
    last = ws.max_row
    dv = DataValidation(type="list", formula1='"' + ",".join(STATUS_KELUAR) + '"', allow_blank=True)
    dv2 = DataValidation(type="list", formula1='"' + ",".join(CARA) + '"', allow_blank=True)
    dv3 = DataValidation(type="list", formula1='"Ya,Tidak"', allow_blank=True)
    for d, col in ((dv, "F"), (dv2, "H"), (dv3, "I")):
        ws.add_data_validation(d)
        d.add(f"{col}4:{col}{last}")
    rng = f"A4:N{last}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=['AND($G4<>"",$I4<>"Ya",TODAY()-$G4>7)'], fill=PatternFill("solid", fgColor="FFF2CC")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['$F4="Draf"'], fill=PatternFill("solid", fgColor="EDEDED")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['$F4="Akuan diterima"'], fill=PatternFill("solid", fgColor="E2EFDA")))
    ws.freeze_panes = "C4"
    ws.auto_filter.ref = f"A3:N{last}"
    susun_cetak(ws)
    return ws


def sheet_masuk(wb):
    ws = wb.create_sheet("Surat Masuk")
    ws.append(["Daftar Surat Masuk – UTM Baikpulih Makmal P&P, Pakej 3 (KESB/UTM-P3) – No. Kontrak TN/2026-21/D-1/SP/BPA"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append(["Catat setiap surat, arahan S.O. dan jawapan RFI daripada UTM pada hari ia diterima. Merah = tarikh akhir jawapan telah lepas "
               "dan belum dibalas; hijau = sudah dibalas."])
    ws.append([t for t, _ in MASUK])
    gaya_tajuk(ws, 3, MASUK)
    for r in MASUK_AWAL:
        ws.append(r)
        gaya_baris(ws, "CDH")
    for n in range(len(MASUK_AWAL) + 1, len(MASUK_AWAL) + BARIS_KOSONG + 1):
        ws.append([n] + [""] * (len(MASUK) - 1))
        gaya_baris(ws, "CDH")
    last = ws.max_row
    dv = DataValidation(type="list", formula1='"' + ",".join(STATUS_MASUK) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"J4:J{last}")
    rng = f"A4:L{last}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=['AND($H4<>"",$H4<TODAY(),$J4<>"Dibalas",$J4<>"Untuk rekod")'], fill=PatternFill("solid", fgColor="F8CBAD")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['$J4="Dibalas"'], fill=PatternFill("solid", fgColor="E2EFDA")))
    ws.freeze_panes = "C4"
    ws.auto_filter.ref = f"A3:L{last}"
    susun_cetak(ws)


def sheet_arahan(wb):
    ar = wb.create_sheet("Arahan")
    for line in [
        "Arahan kemas kini (Kerani Tapak)",
        "",
        "1. Semua surat rasmi kepada UTM / S.O. dihantar melalui Pengurus Projek dan direkod di sini pada hari ia dikeluarkan.",
        "2. Surat Keluar: isi Tarikh hantar, Cara hantar dan tukar Status kepada 'Dihantar'. Untuk surat yang diserah tangan, minta cop / tandatangan penerima.",
        "3. Bila UTM mengakui terima (cop, e-mel atau WhatsApp): isi Akuan terima = Ya, Tarikh akuan dan nama penerima; Status 'Akuan diterima'. "
        "Simpan salinan bercop dalam folder surat.",
        "4. Baris kuning = surat sudah dihantar lebih 7 hari tetapi belum ada akuan terima – hubungi pejabat S.O.",
        "5. Surat Masuk: isi Ruj. Tuan, Tarikh surat dan Tarikh terima. Jika surat meminta jawapan, catat tarikh akhir; baris merah bermakna sudah lewat.",
        "6. Bila KESB membalas, isi Ruj. balasan KESB dan tukar Status kepada 'Dibalas'. Catat juga Ruj. Tuan balasan pada baris Surat Keluar yang berkaitan.",
        "7. Surat Draf belum dikeluarkan sehingga diluluskan PM. Jangan padam baris; tandakan 'Dibatalkan' jika surat tidak jadi dihantar.",
        "8. Surat baharu dalam output/surat/ dimasukkan ke baris kosong dengan: py -3.12 scripts/build_daftar_surat.py --tambah",
        "9. Jangan jalankan skrip dengan --paksa kecuali mahu mula semula; semua isian akan hilang.",
    ]:
        ar.append([line])
    ar["A1"].font = Font(bold=True, size=13)
    ar.column_dimensions["A"].width = 140


def tambah():
    """Masukkan surat baharu (belum ada dalam lajur Fail) ke baris kosong pertama Surat Keluar."""
    wb = load_workbook(OUT)
    ws = wb["Surat Keluar"]
    ada = {str(r[12].value) for r in ws.iter_rows(min_row=4) if r[12].value}
    baharu = [s for s in senarai_surat() if s["fail"] not in ada]
    if not baharu:
        print("Tiada surat baharu.")
        return
    kosong = [r[0].row for r in ws.iter_rows(min_row=4) if not any(c.value for c in r[1:])]
    if len(kosong) < len(baharu):
        print(f"Baris kosong tidak cukup ({len(kosong)} baris, {len(baharu)} surat) – tambah baris secara manual.")
        return
    for row, s in zip(kosong, baharu):
        for i, v in enumerate(baris_keluar(ws.cell(row, 1).value, s)[1:], start=2):
            ws.cell(row, i).value = v
        print(f"Ditambah: {s['fail']}")
    wb.save(OUT)


def main():
    if "--tambah" in sys.argv and OUT.exists():
        tambah()
        return
    if OUT.exists() and "--paksa" not in sys.argv:
        print(f"{OUT} sudah ada – tidak ditulis ganti (isian Kerani dikekalkan). Guna --tambah untuk surat baharu, atau --paksa untuk cipta semula.")
        return
    OUT.parent.mkdir(parents=True, exist_ok=True)
    surat = senarai_surat()
    wb = Workbook()
    sheet_keluar(wb, surat)
    sheet_masuk(wb)
    sheet_arahan(wb)
    wb.save(OUT)
    print(f"Siap: {OUT} ({len(surat)} surat keluar daripada {SURAT}/)")


if __name__ == "__main__":
    main()
