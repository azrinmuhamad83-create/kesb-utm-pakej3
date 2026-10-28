"""Jana Lampiran E: Jadual Penyelenggaraan & Kalibrasi serta Jadual Pemeriksaan & Ujian Loji (Prelim 3.0 (iv)).

Guna:  py -3.12 scripts/build_lampiran_e.py
Hasil: output/Lampiran-E-Jadual-Loji.xlsx  (KESB/UTM-P3/PLT/001)

Senarai loji & peralatan dikumpul daripada Seksyen 5.1 method statement MS-01 – MS-14.
Bilangan unit, no. siri dan kekerapan yang tiada dalam dokumen kontrak ditanda [PERLU SAHKAN].
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

PROJEK = "UTM – Baikpulih Makmal P&P di Bangunan Akademik (RP1-RMKe-13), Pakej 3 – No. kontrak TN/2026-21/D-1/SP/BPA"
HEAD = PatternFill("solid", fgColor="1F4E78")
CAT = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
PS = "[PERLU SAHKAN]"

# (loji / peralatan, MS, pemilik, kalibrasi?, pemeriksaan & ujian)
LOJI = [
    ("Perancah & platform kerja", "MS-01, 02, 05, 06, 09, 14", "SSS", False,
     "Pemeriksaan oleh orang kompeten & tag sebelum digunakan; mingguan " + PS),
    ("Platform beroda / tangga", "MS-01–07, 10, 12", "Site Supervisor", False, "Pemeriksaan visual sebelum digunakan"),
    ("Jentera angkat (AHU Blok E04)", "MS-02", "SSS", False,
     "Sijil pemeriksaan sah & insurans; pengendali bertauliah; pemeriksaan sebelum setiap angkatan"),
    ("Janakuasa (T&C / bekalan sementara / mesin sembur)", "MS-02, 08", "Site Supervisor", False,
     "Pemeriksaan sebelum digunakan; ELCB & pembumian; bulanan " + PS),
    ("Mesin coring", "MS-02, 12", "Site Supervisor", False, "Pemeriksaan sebelum digunakan; kabel & ELCB"),
    ("Alat kuasa (pemecah, pemotong, gerudi, router)", "MS-01, 05, 10, 12", "Site Supervisor", False,
     "Pemeriksaan sebelum digunakan; kabel & penebat"),
    ("Pengekstrak habuk / vakum industri", "MS-01, 07", "Site Supervisor", False, "Pemeriksaan sebelum digunakan; penapis"),
    ("Kipas ekzos / pengudaraan sementara", "MS-01, 07", "Site Supervisor", False, "Pemeriksaan sebelum digunakan"),
    ("Mesin kimpalan", "MS-05, 10, 14", "SSS", False, "Pemeriksaan sebelum digunakan; kabel & pemegang; permit kerja panas"),
    ("Set brazing (oxy-acetylene / nitrogen) & regulator", "MS-02", "Ketua pasukan mekanikal", False,
     "Pemeriksaan hos, regulator & flashback arrestor sebelum digunakan; permit kerja panas"),
    ("Silinder nitrogen & regulator", "MS-02", "Ketua pasukan mekanikal", False, "Silinder diikat tegak; regulator diperiksa"),
    ("Pam vakum", "MS-02", "Ketua pasukan mekanikal", False, "Ujian fungsi & minyak pam sebelum digunakan"),
    ("Mesin pengisar lantai / shot blast", "MS-07", "Site Supervisor", False, "Pemeriksaan sebelum digunakan"),
    ("Pengadun mortar / konkrit", "MS-05, 07, 14", "Site Supervisor", False, "Pemeriksaan sebelum digunakan"),
    ("Mesin sembur polyurea (plural component)", "MS-08", "Pemasang polyurea", False,
     "Pemeriksaan tekanan & pemanas ikut pengilang sebelum digunakan"),
    ("Mesin sembur cat tanpa udara", "MS-09", "Site Supervisor", False, "Pemeriksaan sebelum digunakan"),
    ("Mesin gilap jubin / terrazzo", "MS-07", "Site Supervisor", False, "Pemeriksaan sebelum digunakan"),
    ("Penggetar & pemadat konkrit", "MS-14", "Site Supervisor", False, "Pemeriksaan sebelum digunakan"),
    ("Lori, troli & dolly penghantaran", "MS-11, 13", "Site Supervisor", False, "Pemeriksaan sebelum digunakan; cukai & insurans kenderaan"),
]
PERALATAN = [
    ("Tolok manifold & tolok mikron", "MS-02", "Ketua pasukan mekanikal"),
    ("Pengesan kebocoran refrigeran", "MS-02", "Ketua pasukan mekanikal"),
    ("Penguji penebatan (IR tester)", "MS-03, 04, 12", "Orang Kompeten PW4"),
    ("Penguji ELR / RCD", "MS-03, 12", "Orang Kompeten PW4"),
    ("Alat ujian relay EF/OC (IDMT)", "MS-03", "Orang Kompeten PW4"),
    ("Multimeter & clamp meter", "MS-02, 03, 12", "Orang Kompeten PW4"),
    ("Pengesan voltan", "MS-03", "Orang Kompeten PW4"),
    ("Penguji detector haba / asap", "MS-04", "Ketua pasukan mekanikal"),
    ("Penguji kabel data (certifier)", "MS-12", "H&Z " + PS),
    ("OTDR / power meter fiber", "MS-12", "H&Z " + PS),
    ("Aras laser", "MS-05, 06, 07", "Site Supervisor"),
    ("Meter kelembapan substrat", "MS-07, 08", "Site Supervisor"),
    ("Tolok ketebalan basah", "MS-07, 08", "Site Supervisor"),
    ("Tolok ketebalan filem cat", "MS-09", "Site Supervisor"),
    ("Termometer & higrometer", "MS-02, 08", "Site Supervisor"),
    ("Penimbang bahan", "MS-08", "Site Supervisor"),
    ("Pengesan gas (ujian udara ruang terkurung)", "CP-QC 8.3; KESB-QF-18", "SSS"),
    ("Pengimbas / pengesan kabel & servis tersembunyi", "MS-01, 02", "Penyelaras M&E"),
    ("Kon slump & acuan kiub " + PS, "MS-14", "Site Supervisor"),
]


def sheet(wb, title, headers, widths):
    ws = wb.create_sheet(title)
    ws.append([f"Lampiran E (KESB/UTM-P3/PLT/001 Rev 0) – {title}"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append([PROJEK])
    ws.append(["Prelim 3.0 (iv); Prelim 4.0. Rekod pemeriksaan loji: KESB-QF-25; rekod kalibrasi: KESB-QF-10."])
    ws.append(headers)
    for c in ws[4]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD
        c.alignment = WRAP
        c.border = BORDER
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w
    ws.freeze_panes = "B5"
    return ws


def row(ws, values, fill=None):
    ws.append(values)
    for c in ws[ws.max_row]:
        c.alignment = WRAP
        c.border = BORDER
        if fill:
            c.fill = fill


def main():
    wb = Workbook()
    wb.remove(wb.active)

    ws = sheet(wb, "Penyelenggaraan & Kalibrasi",
               ["Bil", "Loji / peralatan", "Digunakan dalam", "Bil. unit", "No. siri / pendaftaran",
                "Penyelenggaraan (jenis & kekerapan)", "Kalibrasi (kekerapan & makmal)", "Sijil / insurans",
                "Pemilik", "Rekod"],
               [5, 38, 20, 9, 18, 30, 28, 18, 22, 14])
    row(ws, ["", "A. LOJI"], CAT)
    for i, (n, ms, own, _, _) in enumerate(LOJI, start=1):
        row(ws, [i, n, ms, PS, "", "Ikut manual pengilang; " + PS, "–", "Sah sepanjang penggunaan " + PS, own, "KESB-QF-25"])
    row(ws, ["", "B. PERALATAN UKUR & UJIAN"], CAT)
    for i, (n, ms, own) in enumerate(PERALATAN, start=len(LOJI) + 1):
        row(ws, [i, n, ms, PS, "", "Ikut manual pengilang", "Sijil kalibrasi sah; " + PS + " (cadangan 12 bulan)", "–", own, "KESB-QF-10"])

    ws = sheet(wb, "Pemeriksaan & Ujian",
               ["Bil", "Loji / peralatan", "Jenis pemeriksaan / ujian", "Kekerapan", "Kriteria",
                "Pemeriksa", "Rekod", "Tindakan jika gagal"],
               [5, 38, 45, 22, 26, 22, 14, 28])
    row(ws, ["", "A. LOJI"], CAT)
    for i, (n, _, own, _, insp) in enumerate(LOJI, start=1):
        row(ws, [i, n, insp, "Sebelum digunakan; berkala " + PS, "Berfungsi selamat; sijil/insurans sah",
                 own, "KESB-QF-25", "Tag \"JANGAN GUNA\"; baiki / ganti"])
    row(ws, ["", "B. PERALATAN UKUR & UJIAN"], CAT)
    for i, (n, _, own) in enumerate(PERALATAN, start=len(LOJI) + 1):
        row(ws, [i, n, "Semakan sijil kalibrasi & semakan fungsi sebelum digunakan", "Sebelum digunakan",
                 "Kalibrasi belum luput", own, "KESB-QF-10",
                 "Label \"JANGAN GUNA\"; nilai semula keputusan ujian sejak kalibrasi terakhir (CP-QC 10.1)"])

    out = Path("output/Lampiran-E-Jadual-Loji.xlsx")
    wb.save(out)
    print(f"Siap: {out}  ({len(LOJI)} loji, {len(PERALATAN)} peralatan ukur & ujian)")


if __name__ == "__main__":
    main()
