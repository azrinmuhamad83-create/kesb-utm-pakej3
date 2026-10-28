"""Jana Lampiran K: Carta Alir Sistem Kualiti (PDF, A4 landskap) untuk CP-QC.

Guna:  py -3.12 scripts/build_carta_alir.py
Hasil: output/Lampiran-K-Carta-Alir.pdf
Kandungan carta diambil daripada CP-QC (output/CP-QC-draf.md); kemas kini di sini jika proses dalam CP-QC berubah.

Jenis nod: s = mula, p = proses, d = keputusan, e = tamat.
Cabang keputusan: {"label": "Tidak", "loop": i} (kembali ke nod i) atau {"label": "Ya", "side": "teks"} (tindakan sisi).
"""
from pathlib import Path

import fitz

OUT = Path("output/Lampiran-K-Carta-Alir.pdf")
W, H = 842, 595
FONT = "C:/Windows/Fonts/arial.ttf"
FONT_B = "C:/Windows/Fonts/arialbd.ttf"
NAVY = (0.12, 0.31, 0.47)
LIGHT = (0.87, 0.92, 0.97)
YELLOW = (1.0, 0.95, 0.75)
SIDE = (0.99, 0.89, 0.84)
GREY = (0.35, 0.35, 0.35)
WHITE = (1, 1, 1)
BLACK = (0, 0, 0)

# Kolum: peranan | kotak aliran | saluran gelung | tindakan sisi | rekod / borang
X_ROLE = (30, 200)
X_BOX = (215, 455)
X_LOOP = 468
X_SIDE = (482, 625)
X_REC = (640, 812)
Y_TOP, Y_BOT = 108, 560

DOC = "KESB/UTM-P3/CP-QC/001 Rev 0 – Lampiran K: Carta Alir Sistem Kualiti"

CARTA = [
    ("K1", "Kelulusan bahan & submittal", "CP-QC Seksyen 6.1", [
        ("s", "Keperluan bahan, shop drawing atau method statement dikenal pasti (programme, BQ, lukisan)", "Pengurus Projek / Site Supervisor", "Programme; Daftar Submittal KESB-QF-21"),
        ("p", "Sediakan submittal: katalog, sampel, sijil IKRAM/SIRIM, laporan ujian, shop drawing", "Pembekal / pasukan KESB", "–"),
        ("p", "Semak pematuhan BQ, lukisan dan Preamble 16.0 (bahan tempatan)", "Wakil Kualiti", "–"),
        ("p", "Kemukakan kepada S.O. (shop drawing sekurang-kurangnya 1 bulan sebelum kerja – Prelim 11.0)", "Pengurus Projek", "KESB-QF-01 (MAR) / QF-07 / QF-15"),
        ("d", "Diluluskan S.O.?", "S.O.", "", {"label": "Tidak – pinda & kemuka semula", "loop": 1}),
        ("p", "Sampel / mock-up disimpan di pejabat tapak; pemasangan contoh diperiksa (H)", "Site Supervisor", "KESB-QF-02"),
        ("e", "Bahan boleh dipesan (Carta K2)", "Pengurus Projek", "KESB-QF-21 dikemas kini"),
    ]),
    ("K2", "Permohonan, pesanan & pengeluaran bahan", "CP-QC Seksyen 9 (subkon upah pasang; bahan dibeli KESB)", [
        ("s", "Subkon memaklumkan keperluan bahan mengikut ruang", "Subkon (upah pasang)", "–"),
        ("p", "Mohon bahan dengan no. MAR yang telah diluluskan", "Site Supervisor", "KESB-QF-27 (MR)"),
        ("d", "MAR lulus, kuantiti dalam had BQ, tarikh ikut programme?", "Pengurus Projek", "", {"label": "Tidak – pinda MR", "loop": 1}),
        ("p", "Luluskan MR; pejabat keluarkan PO (catat no. MAR & no. MR)", "Pengurus Projek / pejabat KESB", "PO KESB"),
        ("p", "Terima bahan di tapak: semak kuantiti, jenama, model, sijil & keadaan", "Site Supervisor", "KESB-QF-03"),
        ("d", "Sepadan dengan bahan yang diluluskan?", "Site Supervisor", "", {"label": "Tidak – tanda DITOLAK, keluarkan; NCR pembekal (K4)", "side": True}),
        ("p", "Simpan di stor KESB, terlindung cuaca", "Site Supervisor", "–"),
        ("p", "Keluarkan kepada subkon ikut MR/ruang; wakil subkon tandatangan terima; baki dipulangkan", "Site Supervisor / wakil subkon", "KESB-QF-28"),
        ("e", "Bahan dipasang (Carta K3)", "Subkon", "–"),
    ]),
    ("K3", "Pemeriksaan kerja & WIR", "CP-QC Seksyen 8.1; ITP (Lampiran A)", [
        ("s", "Aktiviti ITP siap untuk diperiksa", "Subkon / pasukan", "ITP"),
        ("p", "Pemeriksaan dalaman ikut senarai semak trade", "Site Supervisor", "Senarai semak; KESB-QF-08 (pra-tutup)"),
        ("d", "Akur?", "Site Supervisor", "", {"label": "Tidak – baiki", "loop": 0}),
        ("p", "Sahkan & kemukakan WIR kepada S.O. [notis minimum: PERLU SAHKAN]", "Wakil Kualiti / Pengurus Projek", "KESB-QF-02 Bahagian A"),
        ("p", "Pemeriksaan: H – kelulusan bertulis S.O.; W – S.O. boleh hadir; R – semakan rekod / permit", "S.O. / Site Supervisor", "KESB-QF-02 Bahagian B"),
        ("d", "Lulus?", "S.O.", "", {"label": "Tidak – NCR (K4), periksa semula", "loop": 3}),
        ("e", "Kerja diteruskan ke peringkat berikutnya", "Site Supervisor", "KESB-QF-02 difailkan; Site Diary UTM"),
    ]),
    ("K4", "Ketidakakuran (NCR) & tindakan pembetulan", "CP-QC Seksyen 11 (tempoh: sasaran dalaman KESB)", [
        ("s", "Ketidakakuran dikenal pasti: pemeriksaan dalaman, WIR gagal, arahan S.O., audit, penerimaan bahan", "Sesiapa", "–"),
        ("p", "Keluarkan NCR; hentikan atau tanda kerja terlibat (1 hari bekerja)", "Wakil Kualiti / Site Supervisor", "KESB-QF-04; Daftar NCR"),
        ("p", "Analisis punca; cadang tindakan: baiki, kerja semula, ganti atau konsesi (3 hari bekerja)", "Subkon / pasukan + Site Supervisor", "KESB-QF-04"),
        ("d", "Tindakan diluluskan? (konsesi hanya dengan kelulusan bertulis S.O.)", "Pengurus Projek / S.O.", "", {"label": "Tidak – cadang semula", "loop": 2}),
        ("p", "Laksana tindakan & periksa semula", "Site Supervisor", "KESB-QF-02 baharu"),
        ("d", "Berkesan?", "Wakil Kualiti", "", {"label": "Tidak – laksana semula", "loop": 4}),
        ("e", "Tutup NCR (sasaran ≤ 14 hari); trend dilapor dalam semakan pengurusan", "Pengurus Projek / S.O. (NCR S.O.)", "Daftar NCR"),
    ]),
    ("K5", "Aduan pengguna & makluman WhatsApp", "CP-QC Seksyen 8.4", [
        ("s", "Wakil pengguna menghantar aduan dalam group WhatsApp blok", "Wakil pengguna", "Mesej WhatsApp"),
        ("d", "Segera? (henti kerja, keselamatan, kerosakan)", "PM / Kerani Tapak", "", {"label": "Ya – telefon Site Supervisor untuk tindakan serta-merta; henti kerja ikut Carta K6", "side": True}),
        ("p", "Rekod dalam Log Isu Pengguna: no. ISU-xxx, saluran, tangkap layar mesej asal", "PM / Site Supervisor / Kerani Tapak", "Log Isu Pengguna"),
        ("d", "Isu kualiti?", "Site Supervisor", "", {"label": "Ya – NCR (Carta K4)", "side": True}),
        ("p", "Ambil tindakan; tetapkan tanggungjawab & tarikh sasaran", "Site Supervisor", "Log Isu Pengguna"),
        ("p", "Admin balas dalam group (no. ISU & tindakan); blast ke group blok terlibat jika menjejaskan pengguna lain", "Admin group (PM; Kerani Tapak)", "Tangkap layar makluman"),
        ("p", "Isu terbuka disemak dalam mesyuarat koordinasi mingguan", "Pengurus Projek", "Minit mesyuarat"),
        ("e", "Isu ditutup; makluman susulan dalam group", "Admin group", "Log Isu: Selesai"),
    ]),
    ("K6", "Henti kerja oleh pengguna / UTM", "CP-QC Seksyen 8.5; Preamble 34.1–34.3", [
        ("s", "Permintaan henti kerja oleh pengguna atau UTM", "Wakil pengguna / UTM", "–"),
        ("p", "Henti kerja dengan selamat: matikan peralatan, lindungi kerja separuh siap; pulihkan zon kebakaran", "Site Supervisor / subkon", "Foto keadaan"),
        ("p", "Rekod peminta, masa, lokasi, sebab, kerja & pekerja terjejas", "Site Supervisor", "Log Isu Pengguna"),
        ("p", "Maklumkan PM; admin blast ke group blok terlibat", "Site Supervisor / admin group", "Makluman WhatsApp"),
        ("p", "Maklum S.O. secara bertulis; minta pengesahan arahan rasmi, tempoh & tarikh sambung (1 hari bekerja)", "Pengurus Projek", "Surat / e-mel; KESB-QF-15"),
        ("p", "Atur semula pekerja & subkon ke lokasi lain; kemas kini programme", "Pengurus Projek", "Programme; minit koordinasi"),
        ("d", "Kualiti kerja separuh siap terjejas?", "Site Supervisor / Wakil Kualiti", "", {"label": "Ya – pemeriksaan semula / NCR (K4)", "side": True}),
        ("p", "Sambung kerja selepas disahkan pengguna / S.O.", "Site Supervisor", "KESB-QF-02"),
        ("e", "Jika menjejaskan masa / kos: kemukakan rekod & mohon kelulusan bertulis UTM (Preamble 34.3)", "Pengurus Projek", "Surat kepada S.O. / UTM"),
    ]),
    ("K7", "Pengurusan sisa (termasuk sisa roboh)", "CP-QC Seksyen 8.6; Prelim 22.0, 24.0, 25.0", [
        ("s", "Kerja roboh, tanggal atau pemasangan menghasilkan sisa", "Subkon / pasukan", "–"),
        ("d", "Aset UTM? (A/C, pencegah kebakaran, exhaust fan, WiFi AP, AV)", "Site Supervisor", "", {"label": "Ya – label & lindungi; pump down refrigeran; pindah / serah dengan Gate Pass (QF-14)", "side": True}),
        ("d", "Buangan terjadual? (tiub pendarfluor, cat, pelarut, kimia, sisa makmal)", "Site Supervisor / SSS", "", {"label": "Ya – asing dalam bekas berlabel di stor berbumbung; kontraktor berlesen JAS (QF-22)", "side": True}),
        ("p", "Kumpul sisa roboh & binaan di ruang; logam & kabel diasingkan", "Subkon", "–"),
        ("p", "Bersihkan ruang & laluan pengguna setiap hari kerja; sisa ke tong RoRo bernama KESB", "Subkon / Site Supervisor", "Site Diary UTM"),
        ("p", "Bawa keluar ke tapak dibenarkan PBT, sekurang-kurangnya 2 kali seminggu", "Pengangkut lantikan PM", "KESB-QF-22; tiket pelupusan"),
        ("e", "Ruang bersih sebelum penyerahan", "Site Supervisor", "KESB-QF-06"),
    ]),
    ("K8", "Penyerahan & pengambilan semula ruang", "CP-QC Seksyen 14.1; Lampiran Preamble/1", [
        ("s", "UTM menyerahkan ruang mengikut fasa", "UTM / S.O.", "–"),
        ("p", "Joint survey: rekod keadaan, aset & kecacatan sedia ada dengan foto", "Pengurus Projek / wakil pengguna", "KESB-QF-06"),
        ("p", "Kesan servis, tinjauan asbestos, pasang penghadang; notis kerja kepada pengguna", "Site Supervisor / SSS", "KESB-QF-02; QF-29"),
        ("p", "Kerja ikut ITP (Carta K1–K4); sisa harian (Carta K7)", "Pasukan & subkon", "ITP; Site Diary UTM"),
        ("p", "T&C mekanikal, elektrik, ICT & AV", "Penyelaras M&E / Orang Kompeten PW4", "KESB-QF-11"),
        ("p", "Snag dalaman; semua NCR ruang ditutup", "Site Supervisor", "KESB-QF-09"),
        ("d", "Lulus pemeriksaan bersama S.O. & pengguna?", "S.O. / wakil pengguna", "", {"label": "Tidak – baiki kecacatan", "loop": 5}),
        ("p", "Serah aset UTM yang ditanggal", "Site Supervisor", "KESB-QF-14"),
        ("e", "Ruang diserah semula; Pengesahan Pelanggan & maklum balas pelanggan", "Pengurus Projek", "KESB-QF-06; QF-20"),
    ]),
    ("K9", "RFI & arahan S.O.", "CP-QC Seksyen 5.1; Preamble 15.3", [
        ("s", "Percanggahan atau maklumat tidak lengkap dalam lukisan, BQ atau spesifikasi", "Sesiapa", "–"),
        ("p", "Daftar isu; sediakan RFI dengan rujukan lukisan / BQ", "Pengurus Projek", "KESB-QF-16"),
        ("p", "Kemukakan RFI kepada S.O.; kerja berkaitan tidak dimulakan", "Pengurus Projek", "KESB-QF-15"),
        ("p", "Jawapan S.O., arahan (AI/VO) atau revisi lukisan diterima & didaftarkan", "Kerani Tapak", "Daftar Dokumen"),
        ("p", "Nilai impak ke atas ITP, method statement, programme & bahan", "Pengurus Projek", "–"),
        ("p", "Kemas kini dokumen terkawal; edar semula; cop BATAL pada revisi lama", "Kerani Tapak", "Daftar Dokumen; KESB-QF-15"),
        ("e", "Kerja diteruskan ikut maklumat terkini", "Site Supervisor", "–"),
    ]),
]


def fit(page, rect, text, size=9, bold=False, color=BLACK, align=1):
    """Masukkan teks dalam kotak; kecilkan fon sehingga muat."""
    while size >= 5.5:
        rc = page.insert_textbox(rect, text, fontname="fb" if bold else "f", fontsize=size, color=color, align=align)
        if rc >= 0:
            return
        size -= 0.5
    page.insert_textbox(rect, text, fontname="fb" if bold else "f", fontsize=5.5, color=color, align=align)


def boxed(page, rect, text, fill, size=9, bold=False, color=BLACK, round_=False):
    page.draw_rect(rect, color=NAVY, fill=fill, width=0.8, radius=0.35 if round_ else None)
    inner = fitz.Rect(rect.x0 + 4, rect.y0 + 2, rect.x1 - 4, rect.y1 - 2)
    # Pusat menegak: anggar tinggi teks
    lines = max(1, len(text) * size * 0.5 // max(1, inner.width) + 1)
    off = max(0, (inner.height - lines * size * 1.2) / 2)
    fit(page, fitz.Rect(inner.x0, inner.y0 + off, inner.x1, inner.y1), text, size, bold, color)


def arrow(page, p1, p2, color=NAVY):
    page.draw_line(p1, p2, color=color, width=0.9)
    v = fitz.Point(p2) - fitz.Point(p1)
    ln = abs(v)
    if ln == 0:
        return
    u = fitz.Point(v.x / ln, v.y / ln)
    n = fitz.Point(-u.y, u.x)
    a = fitz.Point(p2) - u * 6
    page.draw_polyline([a + n * 3, fitz.Point(p2), a - n * 3, a + n * 3], color=color, fill=color, width=0.5)


def page_frame(doc, title, sub, no):
    page = doc.new_page(width=W, height=H)
    page.insert_font(fontname="f", fontfile=FONT)
    page.insert_font(fontname="fb", fontfile=FONT_B)
    page.draw_rect(fitz.Rect(0, 0, W, 58), color=None, fill=NAVY)
    page.insert_text((30, 26), title, fontname="fb", fontsize=15, color=WHITE)
    page.insert_text((30, 46), sub, fontname="f", fontsize=9.5, color=WHITE)
    page.insert_text((30, H - 14), DOC, fontname="f", fontsize=7.5, color=GREY)
    page.insert_text((W - 90, H - 14), f"Muka surat {no}", fontname="f", fontsize=7.5, color=GREY)
    return page


def carta(doc, kod, tajuk, sumber, nodes, no):
    page = page_frame(doc, f"Carta {kod}: {tajuk}", f"UTM Baikpulih Makmal P&P, Pakej 3 – {sumber}", no)
    for (x0, x1), lbl in [(X_ROLE, "Oleh"), (X_BOX, "Aliran"), (X_SIDE, "Tindakan sisi"), (X_REC, "Rekod / borang")]:
        page.draw_rect(fitz.Rect(x0, 70, x1, 90), color=None, fill=LIGHT)
        fit(page, fitz.Rect(x0, 74, x1, 90), lbl, 9, True)
    n = len(nodes)
    step = (Y_BOT - Y_TOP) / n
    bh = min(44, step - 14)
    cx = (X_BOX[0] + X_BOX[1]) / 2
    rects = []
    for i, nd in enumerate(nodes):
        yc = Y_TOP + step * i + step / 2
        rects.append(fitz.Rect(X_BOX[0], yc - bh / 2, X_BOX[1], yc + bh / 2))
    for i, nd in enumerate(nodes):
        kind, text, oleh, rekod = nd[:4]
        r = rects[i]
        yc = (r.y0 + r.y1) / 2
        if kind == "d":
            pts = [fitz.Point(cx, r.y0), fitz.Point(X_BOX[1] - 20, yc), fitz.Point(cx, r.y1), fitz.Point(X_BOX[0] + 20, yc), fitz.Point(cx, r.y0)]
            page.draw_polyline(pts, color=NAVY, fill=YELLOW, width=0.8)
            fit(page, fitz.Rect(X_BOX[0] + 62, yc - 10, X_BOX[1] - 62, yc + 16), text, 8)
        elif kind in "se":
            boxed(page, r, text, NAVY, 8.5, True, WHITE, round_=True)
        else:
            boxed(page, r, text, LIGHT, 8.5)
        fit(page, fitz.Rect(X_ROLE[0], yc - 10, X_ROLE[1], yc + 14), oleh, 8, False, GREY, align=0)
        if rekod:
            fit(page, fitz.Rect(X_REC[0], yc - 10, X_REC[1], yc + 14), rekod, 8, False, GREY, align=0)
        if i < n - 1:
            arrow(page, (cx, r.y1), (cx, rects[i + 1].y0))
            if kind == "d":
                lbl = "Tidak" if nd[4]["label"].startswith("Ya") else "Ya"
                page.insert_text((cx + 5, r.y1 + 9), lbl, fontname="fb", fontsize=7.5, color=NAVY)
        if kind == "d":
            br = nd[4]
            right = fitz.Point(X_BOX[1] - 20, yc)
            if "loop" in br:
                t = rects[br["loop"]]
                ty = (t.y0 + t.y1) / 2
                page.draw_polyline([right, fitz.Point(X_LOOP, yc), fitz.Point(X_LOOP, ty)], color=(0.75, 0.2, 0.15), width=0.9)
                arrow(page, (X_LOOP, ty), (X_BOX[1], ty), color=(0.75, 0.2, 0.15))
                fit(page, fitz.Rect(X_SIDE[0], yc - 12, X_SIDE[1], yc + 14), br["label"], 7.5, True, (0.75, 0.2, 0.15), align=0)
            else:
                sr = fitz.Rect(X_SIDE[0], yc - bh / 2 - 4, X_SIDE[1], yc + bh / 2 + 4)
                arrow(page, right, (X_SIDE[0], yc), color=(0.75, 0.2, 0.15))
                boxed(page, sr, br["label"], SIDE, 7.5)


def peta(doc):
    page = page_frame(doc, "Carta K0: Peta Sistem Kualiti Projek",
                      "Kitar hayat setiap ruang (atas) dan proses sokongan berterusan (bawah) – CP-QC KESB/UTM-P3/CP-QC/001", 1)
    fit(page, fitz.Rect(30, 68, 812, 84), "KITAR HAYAT SETIAP RUANG / FASA", 9.5, True, NAVY, align=0)
    kitar = [
        ("1. Penyerahan ruang daripada UTM", "Joint survey – K8\nQF-06"),
        ("2. Persediaan kerja", "Servis, asbestos, penghadang, notis pengguna\nQF-02, QF-29"),
        ("3. Bahan", "Kelulusan – K1\nMAR → MR → PO → terima → keluar – K2\nQF-01, 27, 03, 28"),
        ("4. Kerja & pemeriksaan", "ITP (H/W/R), WIR – K3\nQF-02, QF-08"),
        ("5. Ujian & pentauliahan", "M&E, elektrik, ICT, AV\nQF-11"),
        ("6. Snag & penyerahan semula", "K8\nQF-09, QF-14, QF-06, QF-20"),
    ]
    bw, gap, y0 = 118, 15, 92
    for i, (t, d) in enumerate(kitar):
        x = 30 + i * (bw + gap)
        top = fitz.Rect(x, y0, x + bw, y0 + 30)
        boxed(page, top, t, NAVY, 8.5, True, WHITE, round_=True)
        body = fitz.Rect(x, y0 + 30, x + bw, y0 + 100)
        page.draw_rect(body, color=NAVY, fill=LIGHT, width=0.8)
        fit(page, fitz.Rect(x + 4, y0 + 35, x + bw - 4, y0 + 100), d, 8)
        if i < len(kitar) - 1:
            arrow(page, (x + bw, y0 + 15), (x + bw + gap, y0 + 15))
    # Gelung NCR daripada bahan & pemeriksaan
    ncr = fitz.Rect(30 + 2 * (bw + gap) + 40, y0 + 125, 30 + 4 * (bw + gap) - 50, y0 + 160)
    boxed(page, ncr, "Ketidakakuran → NCR & tindakan pembetulan – K4 (QF-04)", SIDE, 8.5, True)
    red = (0.75, 0.2, 0.15)
    for i in (2, 3):
        x = 30 + i * (bw + gap) + bw / 2
        arrow(page, (x, y0 + 100), (x, ncr.y0), color=red)
    fit(page, fitz.Rect(30, 290, 812, 306), "PROSES SOKONGAN (BERTERUSAN SEPANJANG PROJEK)", 9.5, True, NAVY, align=0)
    sokong = [
        ("Aduan pengguna & makluman WhatsApp", "Group blok; admin blast isu – K5\nLog Isu Pengguna"),
        ("Henti kerja oleh pengguna / UTM", "Rekod, maklum S.O. bertulis – K6\nLog Isu; surat S.O."),
        ("Pengurusan sisa", "Asing di punca; RoRo; buangan terjadual – K7\nQF-22; Site Diary UTM"),
        ("RFI & arahan S.O.", "Percanggahan dokumen – K9\nQF-16, QF-15, Daftar Dokumen"),
        ("Rekod harian", "Site Diary UTM (ITP 1.10)\nPembersihan harian (ITP 1.11)"),
        ("Koordinasi & semakan", "Mesyuarat koordinasi mingguan; audit dalaman & semakan pengurusan setiap 3 bulan"),
    ]
    for i, (t, d) in enumerate(sokong):
        x = 30 + i * (bw + gap)
        boxed(page, fitz.Rect(x, 310, x + bw, 340), t, (0.30, 0.45, 0.60), 8, True, WHITE, round_=True)
        page.draw_rect(fitz.Rect(x, 340, x + bw, 410), color=NAVY, fill=WHITE, width=0.8)
        fit(page, fitz.Rect(x + 4, 345, x + bw - 4, 410), d, 8)
    # Jadual peranan
    fit(page, fitz.Rect(30, 425, 812, 441), "PERANAN UTAMA", 9.5, True, NAVY, align=0)
    peranan = [
        ("Pengurus Projek cum Site Agent (Wakil Kualiti interim)", "Saluran rasmi kepada S.O.; luluskan MR, WIR keluar, NCR & surat; admin group WhatsApp"),
        ("Site Supervisor", "Pemeriksaan dalaman, penerimaan & pengeluaran bahan, Site Diary UTM, pembersihan harian"),
        ("Kerani Tapak", "Kawalan dokumen, Daftar Dokumen & Log Isu; admin kedua group WhatsApp"),
        ("SSS / Penyelaras M&E", "Permit kerja panas & ruang terkurung, stor buangan terjadual / penyelarasan M&E & T&C"),
        ("Subkon (upah pasang)", "Mutu kerja; akui terima bahan (QF-28); bersihkan sisa kerja sendiri setiap hari"),
        ("Ibu Pejabat KESB", "Penyelaras Projek menyemak, Pengarah Projek meluluskan CP-QC & method statement"),
    ]
    y = 445
    for a, b in peranan:
        page.draw_rect(fitz.Rect(30, y, 812, y + 18), color=(0.8, 0.8, 0.8), width=0.5)
        fit(page, fitz.Rect(34, y + 4, 290, y + 18), a, 8, True, BLACK, align=0)
        fit(page, fitz.Rect(295, y + 4, 810, y + 18), b, 8, False, BLACK, align=0)
        y += 18


def main():
    doc = fitz.open()
    peta(doc)
    for no, (kod, tajuk, sumber, nodes) in enumerate(CARTA, start=2):
        carta(doc, kod, tajuk, sumber, nodes, no)
    doc.set_toc([[1, "Carta K0: Peta Sistem Kualiti Projek", 1]] +
                [[1, f"Carta {k}: {t}", i] for i, (k, t, _, _) in enumerate(CARTA, start=2)])
    doc.save(OUT)
    print(f"Siap: {OUT}  ({doc.page_count} hlm)")


if __name__ == "__main__":
    main()
