"""Jana draf surat kepada S.O. (senarai wakil pengguna & group WhatsApp koordinasi) dan panduan group WhatsApp.

Guna:  py -3.12 scripts/build_surat_whatsapp.py
Hasil: output/surat/Surat-SO-Wakil-Pengguna-WhatsApp.md/.docx/.pdf
       output/surat/Panduan-Group-WhatsApp.md/.docx/.pdf
Senarai ruang Lampiran A daripada output/lokasi.json. Surat ialah DRAF – dikeluarkan hanya selepas diluluskan PM.
"""
import json
import sys
from pathlib import Path

import win32com.client

sys.path.insert(0, str(Path(__file__).parent))
import build_pqp_docx  # noqa: E402

OUT = Path("output/surat")
PROJEK = ("UTM – Baikpulih Makmal P & P di Bangunan Akademik (RP1-RMKe-13) Kampus Universiti Teknologi Malaysia, "
          "Johor Bahru – Pakej 3: Fakulti Komputeran (FKOM) & Fakulti Kejuruteraan Mekanikal (FKM)")


def lampiran_a():
    lok = json.loads(Path("output/lokasi.json").read_text(encoding="utf-8"))["lokasi"]
    rows = ["<!-- lebar: 10, 13, 32, 30, 22, 19, 16, 17 -->",
            "| No | Blok | Ruang (ikut BQ B) | Nama wakil pengguna | Jawatan / jabatan | No. telefon | Setuju (Ya / Tidak) | Tandatangan |",
            "|---|---|---|---|---|---|---|---|"]
    for i, r in enumerate(lok, start=1):
        rows.append(f"| {i} | {r['blok']} | {r['ruang']} |  |  |  |  |  |")
    return "\n".join(rows)


BLOK = "N28 (FKOM); C23, C24, C25, E01, E03, E04, E05, E06, P22 dan P23 (FKM)"

PERATURAN = """1. Kumpulan ini ialah saluran **makluman dan koordinasi kerja** antara wakil pengguna blok dan pasukan tapak KESB bagi projek UTM Pakej 3.
2. Kandungan: notis kerja (kerja bising, pemutusan bekalan, pengasingan zon kebakaran, penutupan laluan, penghantaran, kerja panas, kerja luar waktu), jadual kerja, aduan pengguna dan makluman isu.
3. **Arahan rasmi** kepada KESB tetap melalui Pegawai Penguasa (S.O.) secara bertulis. Mesej dalam kumpulan ini bukan arahan kontrak atau rekod kontrak.
4. Aduan dijawab oleh admin KESB dengan no. rujukan ISU-xxx; aduan segera (keselamatan, kerosakan, henti kerja) juga boleh dibuat melalui telefon kepada Pengurus Projek atau Site Supervisor.
5. Admin KESB akan membuat makluman (blast) apabila timbul isu yang menjejaskan pengguna, dan makluman susulan apabila isu selesai.
6. Tiada dokumen kontrak, harga atau maklumat sulit dikongsi dalam kumpulan ini.
7. Ahli ditambah hanya oleh admin KESB dengan persetujuan ahli; no. telefon digunakan untuk tujuan projek sahaja (Akta Perlindungan Data Peribadi 2010).
8. Kumpulan dibubarkan selepas projek siap."""


def surat():
    return f"""# Surat – Permohonan Senarai Wakil Pengguna & Penubuhan Kumpulan WhatsApp Koordinasi Kerja

**[KEPALA SURAT KARIM ENGINEERING SDN BHD]**

Ruj. Kami : KESB/UTM-P3/SURAT/______
Ruj. Tuan : UTM.J.17.03/16.11/667(4)
Tarikh : 28 Oktober 2026

Pengarah
Bahagian Pengurusan Aset
Universiti Teknologi Malaysia
81310 UTM Johor Bahru
Johor

u.p.: En. Nazri (wakil S.O.)

Tuan,

**{PROJEK.upper()}**

**NO. KONTRAK: TN/2026-21/D-1/SP/BPA**

**PERMOHONAN SENARAI WAKIL PENGGUNA BAGI SETIAP LOKASI DAN MAKLUMAN PENUBUHAN KUMPULAN WHATSAPP KOORDINASI KERJA**

Dengan hormatnya perkara di atas adalah dirujuk.

**2.** Kerja baikpulih projek ini dilaksanakan di bangunan yang masih digunakan bagi pengajaran dan pembelajaran, iaitu di Blok {BLOK}. Selaras dengan kehendak kontrak untuk mengurangkan gangguan kepada bangunan yang masih digunakan dan mematuhi arahan S.O. (BQ A Preamble 17.0), kawalan bunyi (Preamble 25.1–25.2) dan kerja luar waktu (Preamble 29.0), KESB memerlukan saluran koordinasi terus dengan wakil pengguna di setiap lokasi.

**3.** Sehubungan itu, KESB bercadang menubuhkan **kumpulan WhatsApp koordinasi kerja mengikut blok** (di bawah satu WhatsApp Community projek) bagi tujuan:
   a) makluman awal notis kerja – kerja bising, pemutusan bekalan elektrik atau air, pengasingan zon pengesan kebakaran, penutupan laluan, penghantaran bahan dan kerja luar waktu;
   b) penerimaan aduan pengguna dan makluman tindakan KESB; dan
   c) makluman segera apabila timbul isu yang menjejaskan pengguna, serta makluman susulan apabila isu selesai.

**4.** Untuk tujuan tersebut, KESB dengan hormatnya memohon jasa baik pihak tuan untuk:
   a) mengemukakan **nama, jawatan dan no. telefon wakil pengguna** bagi setiap blok dan ruang seperti dalam **Lampiran A**;
   b) mendapatkan **persetujuan wakil pengguna** untuk dimasukkan ke dalam kumpulan WhatsApp berkenaan (Lampiran A, lajur persetujuan), selaras dengan Akta Perlindungan Data Peribadi 2010;
   c) menamakan wakil S.O. yang akan menyertai kumpulan tersebut; dan
   d) memaklumkan sekiranya UTM mempunyai saluran koordinasi rasmi lain yang perlu digunakan.

**5.** Peraturan penggunaan kumpulan dilampirkan sebagai **Lampiran B**. Ditegaskan bahawa kumpulan ini adalah saluran makluman dan koordinasi sahaja; segala arahan rasmi kepada KESB kekal melalui S.O. secara bertulis.

**6.** Kerjasama pihak tuan untuk mengemukakan maklumat di atas sebelum **11 November 2026**, atau sekurang-kurangnya sebelum kerja dimulakan di setiap lokasi, amatlah dihargai.

Sekian, terima kasih.

Yang benar,

.......................................................
**MUHAMAD AZRIN BIN RAHMAT**
Pengurus Projek cum Site Agent
Karim Engineering Sdn Bhd

s.k.
- Pengarah Projek, Karim Engineering Sdn Bhd
- Fail projek (KESB/UTM-P3)

---

## Lampiran A: Maklumat & Persetujuan Wakil Pengguna Mengikut Lokasi

Senarai ruang mengikut tajuk item BQ B. Seorang wakil boleh mewakili lebih daripada satu ruang dalam blok yang sama.

{lampiran_a()}

Dengan menandatangani borang ini, wakil pengguna bersetuju no. telefonnya digunakan oleh KESB bagi tujuan koordinasi kerja projek ini sahaja dan dimasukkan ke dalam kumpulan WhatsApp blok berkenaan. Persetujuan boleh ditarik balik dengan memaklumkan admin kumpulan.

| | Disahkan oleh (S.O. / wakil S.O.) |
|---|---|
| Nama | |
| Jawatan | |
| Tandatangan & cop | |
| Tarikh | |

## Lampiran B: Peraturan Kumpulan WhatsApp Koordinasi Kerja

{PERATURAN}
"""


def panduan():
    return f"""# Panduan Penubuhan & Pengurusan Kumpulan WhatsApp Koordinasi Lokasi – UTM Pakej 3

Rujukan: CP-QC KESB/UTM-P3/CP-QC/001 Seksyen 8.4 dan 8.5; Lampiran K Carta K5 dan K6. Admin: Pengurus Projek; admin kedua: Kerani Tapak.

## 1. Sebelum group dibuka

| Langkah | Tindakan | Oleh | Rekod |
|---|---|---|---|
| 1 | Luluskan dan hantar surat permohonan wakil pengguna kepada S.O. (fail "Surat-SO-Wakil-Pengguna-WhatsApp") bersama Lampiran A & B | Pengurus Projek | Surat; KESB-QF-15 Transmittal |
| 2 | Terima Lampiran A yang diisi dan disahkan S.O. – nama, jawatan, no. telefon & persetujuan setiap wakil | Kerani Tapak | Fail surat |
| 3 | Masukkan maklumat wakil ke Daftar Lokasi (ITP.xlsx sheet "Daftar Lokasi") | Kerani Tapak | Daftar Lokasi |
| 4 | Wakil yang tidak bersetuju dimasukkan ke group: dihubungi melalui telefon / e-mel sahaja; catat dalam Daftar Lokasi ("Dalam group WhatsApp" = T) | Kerani Tapak | Daftar Lokasi |

## 2. Buka WhatsApp Community & group blok

1. WhatsApp → tab **Communities** → **New community** → nama: "KESB – UTM Pakej 3 (Koordinasi Kerja)"; penerangan: tujuan & peraturan ringkas (Lampiran B surat).
2. Dalam Community, cipta satu group bagi setiap blok (11 group): "UTM P3 – Blok N28 (FKOM)", "UTM P3 – Blok C23", "C24", "C25", "E01", "E03", "E04", "E05", "E06", "P22", "P23". Blok yang mempunyai ruang sedikit boleh digabung jika dipersetujui S.O.
3. Group **Announcements** Community digunakan untuk makluman yang melibatkan lebih daripada satu blok (contohnya pemutusan bekalan utama atau perubahan jadual fasa).
4. Tetapan setiap group (Group info → **Group settings**):
   - "Edit group settings": **Only admins**
   - "Add other members": **Only admins**
   - "Approve new members": **On**
   - "Disappearing messages": **Off** (mesej perlu kekal untuk rekod)
5. Jadikan **Kerani Tapak** admin kedua dalam setiap group dan dalam Community.
6. Tambah ahli: Pengurus Projek, Site Supervisor, Kerani Tapak, wakil pengguna blok tersebut (yang bersetuju) dan wakil S.O. (jika dinamakan). Subkon tidak dimasukkan; arahan kepada subkon melalui Site Supervisor.
7. Tampal **peraturan group** (Lampiran B surat) dan **pin** mesej tersebut; hantar mesej alu-aluan (templat 1).

## 3. Penggunaan harian

| Situasi | Tindakan admin | Masa | Rekod |
|---|---|---|---|
| Kerja yang mengganggu pengguna dirancang | Hantar notis kerja (templat 2) | Sebelum kerja, sekurang-kurangnya [PERLU SAHKAN] hari bekerja | KESB-QF-29 |
| Aduan diterima | Akui penerimaan; rekod dalam Log Isu (ISU-xxx); balas dengan no. ISU & tindakan | Hari yang sama | Log Isu Pengguna; tangkap layar |
| Isu yang menjejaskan pengguna (henti kerja, pemutusan bekalan, insiden, kelewatan) | **Blast** ke group blok terlibat (templat 3); Announcements jika lebih daripada satu blok | Segera, sekurang-kurangnya hari yang sama | Tangkap layar dalam rekod ISU |
| Isu selesai | Makluman susulan (templat 4) | Sebaik selesai | Log Isu: Selesai |
| Keputusan / persetujuan penting daripada wakil S.O. dalam group | **Tangkap layar serta-merta**; minta pengesahan bertulis jika melibatkan kelulusan atau arahan | Serta-merta | Fail rekod; surat / e-mel pengesahan |

Mesej boleh dipadam oleh pengirim – sebab itu tangkap layar dibuat serta-merta bagi perkara penting.

## 4. Rekod & penutupan
- **Setiap bulan**: eksport chat setiap group (Group info → **Export chat** → "Include media") dan simpan dalam fail rekod projek (CP-QC Seksyen 8.4).
- Kemas kini ahli apabila wakil pengguna bertukar (maklum S.O.).
- Selepas projek siap dan penyerahan akhir: makluman penutupan, kemudian group dan Community dibubarkan; eksport terakhir disimpan.

## 5. Templat mesej

**Templat 1 – Alu-aluan**
Assalamualaikum dan salam sejahtera. Kumpulan ini ditubuhkan oleh Karim Engineering Sdn Bhd (KESB) bagi koordinasi kerja projek baikpulih makmal UTM Pakej 3 di Blok [blok]. Makluman notis kerja, jadual dan aduan akan dibuat di sini. Sila rujuk peraturan kumpulan yang dipin. Untuk perkara segera, hubungi Pengurus Projek / Site Supervisor KESB. Terima kasih.

**Templat 2 – Notis kerja**
Makluman KESB (UTM Pakej 3) – NOTIS KERJA
Lokasi: Blok [ ], [ruang]
Kerja: [jenis kerja – cth pecah lantai / pemutusan bekalan / kerja panas]
Tarikh & masa: [ ] hingga [ ]
Kesan kepada pengguna: [bunyi / laluan ditutup / bekalan terputus]
Langkah KESB: [penghadang, kawalan habuk, laluan alternatif]
Mohon maklum balas sekiranya terdapat kuliah / peperiksaan pada waktu tersebut. Terima kasih.

**Templat 3 – Makluman isu (blast)**
Makluman KESB (UTM Pakej 3) – ISU-[xxx]
Lokasi: Blok [ ], [ruang]
Isu: [ ]
Kesan: [ ]
Tindakan: [ ]
Jangkaan selesai: [tarikh / masa]
Mohon maaf atas kesulitan. Sebarang pertanyaan sila hubungi Pengurus Projek KESB.

**Templat 4 – Isu selesai**
Makluman KESB (UTM Pakej 3) – ISU-[xxx] SELESAI
[ringkasan tindakan]. Terima kasih atas kerjasama pihak tuan/puan.
"""


def to_pdf(docx_files):
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        for f in docx_files:
            d = word.Documents.Open(str(f.resolve()), ReadOnly=True)
            d.SaveAs2(str(f.with_suffix(".pdf").resolve()), FileFormat=17)
            d.Close(False)
    finally:
        word.Quit()


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    docs = []
    for name, text in [("Surat-SO-Wakil-Pengguna-WhatsApp", surat()), ("Panduan-Group-WhatsApp", panduan())]:
        md = OUT / f"{name}.md"
        md.write_text(text, encoding="utf-8")
        build_pqp_docx.main(md, md.with_suffix(".docx"))
        docs.append(md.with_suffix(".docx"))
    to_pdf(docs)
    print("Siap:", ", ".join(str(d.with_suffix('.pdf')) for d in docs))


if __name__ == "__main__":
    main()
