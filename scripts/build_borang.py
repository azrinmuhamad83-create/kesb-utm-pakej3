"""Jana borang kualiti KESB-QF-01 hingga QF-29 (Word) untuk CP-QC.

Guna:  py -3.12 scripts/build_borang.py            (semua borang)
       py -3.12 scripts/build_borang.py 02 08      (borang tertentu)
Hasil: output/borang/KESB-QF-xx <tajuk>.docx

Setiap borang: kepala projek, maklumat (label | ruang isi), jadual rekod (lajur + baris kosong),
senarai semak (Ya / Tidak / T.B.), dan blok tandatangan.
"""
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt

OUT = Path("output/borang")
PROJEK = [
    ("Projek", "UTM – Baikpulih Makmal P & P di Bangunan Akademik (RP1-RMKe-13), Pakej 3: FKOM & FKM"),
    ("No. kontrak", "TN/2026-21/D-1/SP/BPA"),
    ("Kontraktor", "Karim Engineering Sdn Bhd (KESB)"),
]

KESB = ["Disediakan (KESB)", "Disemak (KESB)"]
SO = ["Pegawai Penguasa (S.O.) / Wakil"]

# Setiap borang: (no, tajuk, [blok...], tandatangan)
# Blok: ("info", [label...]) | ("grid", tajuk, [lajur...], bil_baris) | ("semak", tajuk, [item...])
#       | ("pilih", tajuk, [pilihan...])  (satu baris kotak ☐) | ("teks", tajuk, bil_baris)
FORMS = [
    ("01", "Permohonan Kelulusan Bahan (MAR)", [
        ("info", ["No. MAR", "Tarikh", "Trade / no. ITP", "Bahan", "Jenama / model", "Pengilang / pembekal",
                  "Rujukan BQ / lukisan / Jadual Kemasan", "Kuantiti anggaran", "Lokasi penggunaan"]),
        ("pilih", "Status bahan tempatan (Preamble 16.0)", [
            "Tersenarai IKRAM QA / SIRIM QA", "Diuji & disahkan IKRAM / SIRIM", "Bahan import – kelulusan melalui MTO",
            "Bahan Green / alternatif (Preamble 22.1)"]),
        ("semak", "Lampiran", ["Katalog / data teknikal", "Sampel", "Sijil IKRAM / SIRIM / BOMBA",
                               "Laporan ujian makmal bertauliah", "Arahan pemasangan pengilang", "Waranti"]),
        ("pilih", "Keputusan S.O.", ["Lulus", "Lulus bersyarat", "Ditolak", "Kemuka semula"]),
        ("teks", "Ulasan S.O.", 2),
    ], KESB + SO),
    ("02", "Request for Inspection (WIR)", [
        ("bahagian", "BAHAGIAN A – PERMOHONAN PEMERIKSAAN KERJA (KESB)"),
        ("info", ["No. WIR / tarikh mohon", "Tarikh / masa periksa (cadangan)",
                  "Blok / ruang / trade / subkon", "ITP no. & titik (H / W)", "Aktiviti / elemen",
                  "Lukisan / MS / MAR"]),
        ("pilih", "Semakan dalaman KESB sebelum permohonan", [
            "Diperiksa Site Supervisor & subkon", "Ikut lukisan & jadual kemasan", "Bahan diluluskan",
            "Rekod ujian dilampirkan", "Senarai semak (QF-08) dilampirkan", "Foto dilampirkan"]),
        ("tt", ["Diperiksa (Site Supervisor)", "Dimohon (Pengurus Projek)"]),
        ("bahagian", "BAHAGIAN B – PENERIMAAN & PEMERIKSAAN (S.O.)"),
        ("info", ["Diterima oleh / tarikh", "Diperiksa pada (tarikh / masa)", "Hadir (S.O. / KESB / subkon)"]),
        ("pilih", "Keputusan pemeriksaan", ["Lulus – kerja boleh diteruskan", "Lulus bersyarat",
                                            "Gagal – NCR dikeluarkan (no.: ........)",
                                            "Dijadual semula ke: ............"]),
        ("teks", "Ulasan / syarat S.O.", 2),
        ("tt", ["Pegawai Penguasa (S.O.) / Wakil"]),
        ("bahagian", "BAHAGIAN C – PEMERIKSAAN SEMULA (jika bersyarat / gagal)"),
        ("info", ["Pembetulan siap / NCR no."]),
        ("pilih", "Keputusan", ["Diterima – kerja boleh diteruskan", "Tidak diterima"]),
        ("tt", ["Pengurus Projek (KESB)", "Pegawai Penguasa (S.O.) / Wakil"]),
    ], []),
    ("03", "Penerimaan Bahan di Tapak", [
        ("info", ["Tarikh", "Pembekal", "No. DO / invois", "No. MAR / No. MR / No. PO", "Diterima di (blok / stor)"]),
        ("grid", "Butiran bahan", ["Bil", "Bahan", "Jenama / model", "Kuantiti", "Keadaan", "Sepadan MAR (Y/T)", "Catatan"], 8),
        ("pilih", "Keputusan", ["Diterima", "Diterima sebahagian", "Ditolak – NCR pembekal (no.: ........)"]),
    ], ["Diterima oleh (Site Supervisor)", "Disahkan (Wakil Kualiti)"]),
    ("04", "Laporan Ketidakakuran (NCR)", [
        ("info", ["No. NCR", "Tarikh", "Jenis (dalaman / S.O. / pembekal)", "Blok / ruang", "Trade / no. ITP",
                  "Subkon / pembekal terlibat", "Dikeluarkan oleh"]),
        ("teks", "Butiran ketidakakuran (dengan rujukan lukisan / spesifikasi / BQ)", 3),
        ("pilih", "Tindakan segera", ["Kerja dihentikan", "Kawasan / bahan ditanda", "S.O. dimaklumkan"]),
        ("teks", "Punca", 2),
        ("teks", "Tindakan pembetulan dicadang (baiki / kerja semula / ganti / konsesi S.O.)", 2),
        ("info", ["Tarikh sasaran", "Diluluskan oleh / tarikh", "Konsesi S.O. (jika ada) – rujukan",
                  "WIR pemeriksaan semula (no.)", "Tindakan disahkan berkesan (Y / T)"]),
        ("teks", "Tindakan pencegahan", 2),
    ], ["Pengesah tindakan (Wakil Kualiti)", "Ditutup oleh / tarikh", "Penutupan S.O. (NCR S.O. / konsesi)"]),
    ("05", "Rekod Ujian", [
        ("info", ["No. rekod", "Tarikh", "Jenis ujian", "Blok / ruang / sistem", "No. aktiviti ITP",
                  "Kriteria / standard", "Peralatan ujian & no. sijil kalibrasi"]),
        ("grid", "Keputusan", ["Bil", "Titik / litar / unit", "Nilai diperlukan", "Nilai sebenar", "Lulus / Gagal", "Catatan"], 10),
    ], ["Diuji oleh", "Disaksikan (KESB)", "Disaksikan (S.O.)"]),
    ("06", "Penyerahan & Pengambilan Semula Ruang", [
        ("info", ["Tarikh", "Blok / ruang", "Fasa", "Jenis (UTM → KESB / KESB → UTM)", "Rujukan Lampiran Preamble/1"]),
        ("grid", "Keadaan ruang", ["Bil", "Elemen", "Keadaan", "Rujukan foto", "Catatan"], 6),
        ("grid", "Aset UTM dalam ruang", ["Bil", "Aset", "No. tag / siri", "Kuantiti", "Keadaan", "Tindakan"], 5),
        ("semak", "Penyerahan kepada UTM (bila berkaitan)", [
            "T&C selesai (KESB-QF-11)", "Snag ditutup (KESB-QF-09)", "Ruang bersih", "Aset dikembalikan",
            "Pengesahan Pelanggan pada pelan ditandatangani"]),
    ], ["Wakil KESB", "Wakil UTM / pengguna fakulti"] + SO),
    ("07", "Transmittal Shop Drawing", [
        ("info", ["No. transmittal", "Tarikh", "Trade", "Tarikh kerja dijadual (≥ 1 bulan – Prelim 11.0)"]),
        ("grid", "Lukisan", ["Bil", "No. lukisan", "Tajuk", "Rev", "Bil. salinan"], 7),
        ("pilih", "Status S.O.", ["A – Diluluskan", "B – Diluluskan dengan ulasan", "C – Tidak diluluskan, kemuka semula"]),
        ("teks", "Ulasan", 2),
    ], KESB + SO),
    ("08", "Senarai Semak Pra-tutup (Siling / Dinding)", [
        ("info", ["Tarikh", "Blok / ruang", "Elemen (siling / dinding)", "No. aktiviti ITP (7.4 / 3.5 / 17.6)"]),
        ("semak", "Servis dalam siling / dinding", [
            "Paip refrigeran – ujian tekanan lulus (KESB-QF-05 no.: ......)", "Paip refrigeran – vakum lulus",
            "Penebat paip lengkap", "Paip kondensat – kecerunan & ujian aliran", "Ducting / diffuser / grille",
            "Konduit & kabel elektrik – terbenam, diuji", "Kotak suis / soket pada kedudukan disahkan",
            "Kabel data / AV", "Pengesan kebakaran & pendawaian signal", "Lampu & lampu kecemasan",
            "Lubang skrin projektor / bukaan", "Sokongan & penggantung servis", "Manhole siling"]),
        ("grid", "Pengesahan trade", ["Trade", "Nama", "Tandatangan", "Tarikh"], 5),
    ], ["Penyelaras M&E", "Site Supervisor"] + SO),
    ("09", "Senarai Kecacatan (Snag)", [
        ("info", ["Tarikh pemeriksaan", "Blok / ruang", "Diperiksa oleh"]),
        ("grid", "Kecacatan", ["Bil", "Lokasi", "Elemen", "Kecacatan", "Trade", "Tarikh sasaran", "Disahkan siap"], 15),
    ], KESB),
    ("10", "Rekod Kalibrasi Peralatan", [
        ("grid", "Peralatan ukur & ujian", ["Bil", "Peralatan", "No. siri", "Makmal kalibrasi", "No. sijil",
                                            "Tarikh kalibrasi", "Tarikh luput", "Status"], 15),
    ], KESB),
    ("11", "Laporan Ujian & Pentauliahan (T&C)", [
        ("info", ["No. laporan", "Tarikh", "Sistem (hawa dingin / kebakaran / elektrik / ICT / AV)", "Blok / ruang",
                  "No. aktiviti ITP", "Kriteria penerimaan"]),
        ("grid", "Rekod setiap unit / peranti", ["Bil", "Unit / peranti", "Model", "Kapasiti", "Voltan / arus",
                                                 "Bacaan (suhu / tekanan / fungsi)", "Lulus / Gagal"], 8),
        ("semak", "Ujian fungsi", ["Operasi & kawalan", "Penggera / isyarat ke panel", "Tiada bunyi / getaran luar biasa",
                                   "Label & penandaan", "Inventori dikemas kini"]),
        ("teks", "Ulasan", 2),
    ], ["Diuji oleh", "Pengurus Projek (KESB)", "Ujiterima S.O."]),
    ("12", "Permit Kerja Panas", [
        ("info", ["No. permit", "Tarikh", "Masa mula – tamat", "Blok / ruang", "Jenis kerja (kimpalan / brazing / potong)",
                  "Pelaksana"]),
        ("semak", "Langkah kawalan sebelum kerja", [
            "Pemadam api di lokasi", "Bahan mudah terbakar dialih / dilindungi",
            "Zon pengesan diasingkan dengan permit (KESB-QF-19)", "Pengudaraan mencukupi",
            "Pemerhati api ditetapkan", "PPE (pelindung muka, sarung tangan)"]),
        ("info", ["Pemeriksaan selepas kerja – masa / oleh", "Zon pengesan dipulihkan – masa"]),
    ], ["Pemohon", "Diluluskan (SSS)", "Permit ditutup (SSS)"]),
    ("13", "Rekod Latihan & Taklimat", [
        ("info", ["Tajuk", "Tarikh / masa", "Tempat", "Disampaikan oleh", "Jenis (taklimat CP-QC / toolbox / OSH / lain)"]),
        ("grid", "Peserta", ["Bil", "Nama", "Jawatan", "Syarikat", "Tandatangan"], 15),
    ], ["Disediakan oleh"]),
    ("14", "Serahan Aset UTM / Gate Pass", [
        ("info", ["No. rekod", "Tarikh", "Blok / ruang asal", "No. Gate Pass UTM"]),
        ("grid", "Aset", ["Bil", "Aset", "No. tag / siri", "Kuantiti", "Keadaan sebelum", "Destinasi / penerima", "Keadaan selepas"], 10),
    ], ["Diserah oleh (KESB)", "Diterima oleh (UTM)"]),
    ("15", "Transmittal Dokumen", [
        ("info", ["No. transmittal", "Tarikh", "Kepada", "Tujuan (kelulusan / makluman / rekod)"]),
        ("grid", "Dokumen", ["Bil", "No. dokumen", "Tajuk", "Rev", "Bil. salinan"], 10),
    ], ["Dihantar oleh (KESB)", "Diterima oleh"]),
    ("16", "Permintaan Maklumat (RFI)", [
        ("info", ["No. RFI", "Tarikh", "Perkara", "Rujukan (lukisan / BQ / spesifikasi)", "Tarikh jawapan diperlukan",
                  "No. isu daftar (jika ada)"]),
        ("teks", "Soalan / percanggahan", 4),
        ("teks", "Cadangan kontraktor", 3),
        ("teks", "Jawapan S.O.", 4),
    ], ["Pengurus Projek (KESB)"] + SO),
    ("17", "Permit Kerja Elektrik", [
        ("info", ["No. permit", "Tarikh", "Masa mula – tamat", "DB / litar / suis", "Lokasi", "Jenis kerja",
                  "Kebenaran pemutusan UTM (oleh / masa)"]),
        ("semak", "Pengasingan", ["Laluan kabel dikenal pasti (BQ E nota V)", "Suis dibuka & dikunci",
                                  "Tag dipasang (no. ......)", "Ketiadaan voltan diuji", "Pengguna dimaklumkan"]),
        ("info", ["Kerja siap & kunci dibuka – masa", "Bekalan dipulihkan & diuji – masa"]),
    ], ["Orang Kompeten PW4", "Pelaksana", "Permit ditutup"]),
    ("18", "Permit Masuk Ruang Terkurung & Ujian Udara", [
        ("info", ["No. permit", "Tarikh", "Lokasi (ruang siling / lain)", "Masa masuk – keluar", "Pekerja masuk", "Pemerhati"]),
        ("grid", "Ujian udara", ["Parameter", "Had selamat [PERLU SAHKAN]", "Bacaan", "Masa"], 4),
        ("semak", "Kawalan", ["Pengudaraan", "Komunikasi dengan pemerhati", "Pelan menyelamat", "PPE"]),
    ], ["Pemohon", "Diluluskan (SSS)", "Permit ditutup (SSS)"]),
    ("19", "Log Pengasingan Zon Pengesan Kebakaran", [
        ("grid", "Log", ["Tarikh", "Blok / zon", "Panel", "Masa diasingkan", "Kebenaran UTM (nama)",
                         "Masa dipulihkan", "Disahkan oleh"], 15),
    ], ["Penyelaras M&E"]),
    ("20", "Maklum Balas Pelanggan", [
        ("info", ["Tarikh", "Blok / ruang", "Nama & jawatan wakil fakulti"]),
        ("grid", "Penilaian (1 = sangat lemah, 5 = sangat baik)", ["Perkara", "1", "2", "3", "4", "5"], 0),
        ("teks", "Komen / cadangan", 3),
    ], ["Wakil fakulti"]),
    ("21", "Daftar Submittal & Pesanan Bahan", [
        ("grid", "Daftar", ["No. MAR / SD", "Trade", "Bahan / lukisan", "Tarikh hantar", "Tarikh perlu",
                            "Status S.O.", "Tarikh lulus", "Tarikh pesanan", "Tarikh tiba"], 15),
    ], ["Wakil Kualiti"]),
    ("22", "Rekod Pelupusan Sisa", [
        ("grid", "Rekod", ["Tarikh", "Jenis sisa (binaan / terjadual / domestik)", "Kuantiti", "Pengangkut",
                           "Destinasi diluluskan", "No. tiket / manifest"], 15),
    ], ["Site Supervisor"]),
    ("23", "Rekod Ukuran Bersama", [
        ("info", ["Tarikh", "Blok / ruang", "Bahagian BQ (D / E – kuantiti sementara)"]),
        ("grid", "Ukuran", ["Item BQ", "Keterangan", "Unit", "Kuantiti BQ", "Kuantiti diukur", "Catatan"], 12),
    ], ["Wakil KESB", "Wakil S.O."]),
    ("24", "Penilaian Pembekal / Subkontraktor", [
        ("info", ["Nama syarikat", "Skop", "Pendaftaran CIDB / ST / lain & tarikh tamat", "Tempoh penilaian",
                  "Jenis (pemilihan / suku tahunan)"]),
        ("grid", "Kriteria (1 = lemah, 5 = cemerlang)", ["Kriteria", "Skor", "Catatan"], 0),
        ("pilih", "Keputusan", ["Diterima / diteruskan", "Diteruskan dengan pemantauan", "Tidak diterima"]),
    ], ["Dinilai oleh", "Diluluskan (Pengurus Projek)"]),
    ("25", "Rekod Pemeriksaan Loji", [
        ("grid", "Pemeriksaan", ["Tarikh", "Loji / peralatan", "No. / siri", "Sijil / insurans sah (Y/T)",
                                 "Keadaan", "Tindakan", "Diperiksa oleh"], 15),
    ], ["Site Supervisor", "SSS"]),
    ("26", "Daftar Kecacatan DLP", [
        ("grid", "Daftar", ["No.", "Tarikh aduan", "Blok / ruang", "Kecacatan", "Dilapor oleh", "Tarikh periksa",
                            "Tarikh dibaiki", "Disahkan UTM"], 15),
    ], ["Pengurus Projek (KESB)"]),
    ("27", "Permohonan Bahan (MR)", [
        ("info", ["No. MR / tarikh", "Blok / ruang (lokasi penggunaan)", "Trade / subkon",
                  "Tarikh diperlukan di tapak", "Rujukan item BQ / lukisan"]),
        ("grid", "Bahan dimohon", ["Bil", "Bahan / spesifikasi", "No. MAR diluluskan", "Unit",
                                   "Kuantiti dimohon", "Kuantiti terdahulu / baki BQ", "Catatan"], 10),
        ("pilih", "Semakan Pengurus Projek sebelum lulus", [
            "Bahan diluluskan S.O. (MAR)", "Kuantiti dalam had BQ / ukuran tapak", "Tarikh perlu ikut programme",
            "Stor & laluan penghantaran sedia"]),
        ("pilih", "Keputusan", ["Diluluskan", "Diluluskan sebahagian", "Ditolak"]),
        ("info", ["No. PO dikeluarkan / tarikh", "Pembekal"]),
    ], ["Dimohon (Site Supervisor)", "Disemak & diluluskan (Pengurus Projek)", "Diproses (pejabat / pembelian)"]),
    ("28", "Pengeluaran & Penerimaan Bahan (Subkon)", [
        ("info", ["No. rekod / tarikh", "Subkon (upah pasang)", "Blok / ruang", "Trade / aktiviti",
                  "No. MR / no. rekod penerimaan (QF-03)"]),
        ("grid", "Bahan dikeluarkan dari stor KESB", ["Bil", "Bahan / spesifikasi", "No. MAR", "Unit",
                                                     "Kuantiti keluar", "Keadaan semasa diterima",
                                                     "Baki pulang", "Catatan"], 10),
        ("pilih", "Pengakuan penerimaan oleh subkon", [
            "Bahan diterima dalam kuantiti & keadaan seperti di atas",
            "Bahan sepadan dengan bahan diluluskan (MAR)",
            "Subkon bertanggungjawab menjaga bahan sehingga dipasang & memulangkan baki ke stor KESB"]),
        ("info", ["Baki dipulangkan – tarikh", "Kerosakan / kehilangan (jika ada) – no. NCR"]),
    ], ["Dikeluarkan oleh (stor / Site Supervisor KESB)", "Diterima oleh (wakil subkon)",
        "Baki diterima semula (KESB)"]),
    ("29", "Notis Kerja kepada Pengguna", [
        ("info", ["No. notis / tarikh dikeluarkan", "Kepada (wakil pengguna / fakulti / pengurusan bangunan)",
                  "Blok / ruang terlibat", "Tarikh & masa kerja (mula – tamat)"]),
        ("pilih", "Jenis kerja / gangguan", [
            "Kerja bising / bergetar", "Pemutusan bekalan elektrik", "Pengasingan zon pengesan kebakaran",
            "Penutupan laluan / ruang", "Penghantaran bahan / perabot", "Kerja panas", "Kerja luar waktu",
            "Lain-lain: ............"]),
        ("teks", "Butiran kerja & kawasan terjejas", 3),
        ("teks", "Langkah kawalan (penghadang, laluan alternatif, bekalan sementara, kawalan habuk & bunyi)", 2),
        ("info", ["Pegawai untuk dihubungi (KESB) / telefon"]),
        ("pilih", "Maklum balas pengguna", ["Dipersetujui", "Dipersetujui dengan syarat", "Minta dijadual semula ke: ............"]),
        ("teks", "Syarat / ulasan pengguna", 2),
    ], ["Dikeluarkan oleh (Pengurus Projek KESB)", "Diterima / dipersetujui (wakil pengguna)"]),
]

# Lebar lajur jadual rekod (cm) bagi borang yang lajur pertamanya perlu ruang tulis lebih
WIDTHS = {
    "27": [1.0, 5.6, 2.6, 1.4, 2.0, 2.4, 2.4],
    "28": [0.9, 4.8, 2.0, 1.2, 1.9, 2.4, 1.9, 2.3],
}

RATINGS = {
    "20": ["Kualiti kemasan", "Kefungsian peralatan & servis", "Kebersihan semasa penyerahan",
           "Gangguan kepada P&P semasa kerja", "Komunikasi & maklum balas KESB", "Ketepatan masa penyerahan"],
    "24": ["Kualiti kerja / bahan", "Ketepatan masa", "Pematuhan K&K", "Dokumentasi & rekod",
           "Respons kepada NCR", "Kompetensi personel"],
}


def spacer(doc, pt=4):
    """Perenggan kosong kecil sebagai jarak antara jadual."""
    r = doc.add_paragraph().add_run(" ")
    r.font.size = Pt(pt)


def heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    r = p.add_run(text)
    r.bold = True


def part(doc, text):
    """Tajuk bahagian berlorek (cth BAHAGIAN A) untuk borang berbilang bahagian."""
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    c = t.cell(0, 0)
    c.text = text
    c.paragraphs[0].runs[0].bold = True
    shade(c, "BDD7EE")
    spacer(doc, 3)


def sign_row(doc, roles):
    """Blok tandatangan ringkas dalam bahagian: peranan, kemudian tandatangan / nama / tarikh."""
    t = doc.add_table(rows=2, cols=len(roles))
    t.style = "Table Grid"
    for j, r in enumerate(roles):
        c = t.cell(0, j)
        c.text = r
        c.paragraphs[0].runs[0].bold = True
        shade(c)
        t.cell(1, j).text = "Tandatangan:\nNama:                              Tarikh:"
    t.rows[1].height = Cm(1.1)
    spacer(doc, 2)


def choices(doc, items):
    """Pilihan dalam satu baris dengan kotak ☐."""
    p = doc.add_paragraph("     ".join(f"☐ {it}" for it in items))
    p.paragraph_format.space_after = Pt(2)


def shade(cell, color="DDEBF7"):
    el = OxmlElement("w:shd")
    el.set(qn("w:fill"), color)
    cell._tc.get_or_add_tcPr().append(el)


def label_table(doc, labels):
    t = doc.add_table(rows=len(labels), cols=2)
    t.style = "Table Grid"
    for i, lab in enumerate(labels):
        c = t.cell(i, 0)
        c.text = lab
        c.paragraphs[0].runs[0].bold = True
        shade(c)
        c.width = Cm(6)
        t.cell(i, 1).width = Cm(11)
        t.rows[i].height = Cm(0.6)
    spacer(doc)


def grid(doc, cols, rows, first_col=None, widths=None):
    n = len(first_col) if first_col else rows
    t = doc.add_table(rows=n + 1, cols=len(cols))
    t.style = "Table Grid"
    if widths:
        t.autofit = False
        for row in t.rows:
            for c, w in zip(row.cells, widths):
                c.width = Cm(w)
    for j, h in enumerate(cols):
        c = t.cell(0, j)
        c.text = h
        c.paragraphs[0].runs[0].bold = True
        shade(c)
    for i in range(1, n + 1):
        if first_col:
            t.cell(i, 0).text = first_col[i - 1]
        t.rows[i].height = Cm(0.6)
    spacer(doc)


def checklist(doc, items):
    t = doc.add_table(rows=len(items) + 1, cols=4)
    t.style = "Table Grid"
    for j, h in enumerate(["Perkara", "Ya", "Tidak", "T.B."]):
        c = t.cell(0, j)
        c.text = h
        c.paragraphs[0].runs[0].bold = True
        shade(c)
    for i, it in enumerate(items, start=1):
        t.cell(i, 0).text = it
    for row in t.rows:
        row.cells[0].width = Cm(12)
        for c in row.cells[1:]:
            c.width = Cm(1.6)
    spacer(doc)


def signatures(doc, roles):
    t = doc.add_table(rows=4, cols=len(roles))
    t.style = "Table Grid"
    for j, r in enumerate(roles):
        c = t.cell(0, j)
        c.text = r
        c.paragraphs[0].runs[0].bold = True
        shade(c)
        t.cell(1, j).text = "Tandatangan:\n"
        t.cell(2, j).text = "Nama:"
        t.cell(3, j).text = "Tarikh:"


def build(no, title, blocks, sign):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(9.5)
    st.paragraph_format.space_after = Pt(0)
    st.paragraph_format.space_before = Pt(0)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)  # A4
    sec.left_margin = sec.right_margin = Cm(1.8)
    sec.top_margin = sec.bottom_margin = Cm(1.2)
    sec.header_distance = Cm(0.6)
    sec.header.paragraphs[0].text = f"KESB-QF-{no}  |  Rev 0  |  CP-QC KESB/UTM-P3/CP-QC/001"
    sec.header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("KARIM ENGINEERING SDN BHD")
    r.bold = True
    r.font.size = Pt(12)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"KESB-QF-{no}: {title.upper()}")
    r.bold = True
    r.font.size = Pt(13)

    t = doc.add_table(rows=len(PROJEK), cols=2)
    t.style = "Table Grid"
    for i, (k, v) in enumerate(PROJEK):
        t.cell(i, 0).text = k
        t.cell(i, 0).paragraphs[0].runs[0].bold = True
        shade(t.cell(i, 0), "F2F2F2")
        t.cell(i, 1).text = v
    spacer(doc)

    for b in blocks:
        kind = b[0]
        if kind == "info":
            label_table(doc, b[1])
            continue
        if kind == "tt":
            sign_row(doc, b[1])
            continue
        if kind == "bahagian":
            part(doc, b[1])
            continue
        heading(doc, b[1])
        if kind == "pilih":
            choices(doc, b[2])
        elif kind == "grid":
            grid(doc, b[2], b[3], RATINGS.get(no) if b[3] == 0 else None, WIDTHS.get(no))
        elif kind == "semak":
            checklist(doc, b[2])
        elif kind == "teks":
            tt = doc.add_table(rows=1, cols=1)
            tt.style = "Table Grid"
            tt.cell(0, 0).text = "\n" * b[2]
            spacer(doc)

    if sign:
        signatures(doc, sign)
    else:
        # perenggan terakhir selepas jadual dikecilkan supaya tidak menolak ke muka surat baharu
        last = doc.paragraphs[-1]
        for r in last.runs:
            r.font.size = Pt(1)
        last.paragraph_format.line_spacing = Pt(1)
    OUT.mkdir(parents=True, exist_ok=True)
    safe = title.replace("/", "-")
    out = OUT / f"KESB-QF-{no} {safe}.docx"
    doc.save(out)
    return out


if __name__ == "__main__":
    want = set(sys.argv[1:])
    n = 0
    for no, title, blocks, sign in FORMS:
        if want and no not in want:
            continue
        build(no, title, blocks, sign)
        n += 1
    print(f"Siap: {n} borang dalam {OUT}")
