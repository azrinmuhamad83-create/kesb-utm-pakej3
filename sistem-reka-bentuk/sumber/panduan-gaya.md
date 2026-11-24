---
jenis: Panduan Gaya Dokumen
tajuk: Sistem Reka Bentuk Dokumen SOM & PQP
no: KESB/UTM-P3/RB/001
rev: 0
tarikh: 03/11/2026
disediakan: Muhamad Azrin bin Rahmat, Pengurus Projek cum Site Agent
disemak: Muhammad Syahidzul Ikhwan, Penyelaras Projek (Ibu Pejabat KESB)
diluluskan: Zaillan Md Ismail, Pengarah Projek KESB
---
## Tentang panduan ini
Panduan ini menetapkan rupa dan struktur semua dokumen kualiti KESB bagi projek UTM Pakej 3: Project Quality Plan (PQP / CP-QC), Standard Operation Manual (SOM), prosedur operasi standard (SOP), method statement, borang KESB-QF, ITP dan daftar. Token reka bentuk diekstrak daripada skrip penjana sedia ada (`scripts/build_*.py`) dan CP-QC Rev 5. Versi interaktif: `sistem-reka-bentuk/index.html`.

# 1. Prinsip
- **Kontrak dahulu.** Setiap prosedur merujuk klausa kontrak; amalan tambahan ditanda "cadangan KESB" atau "sasaran dalaman KESB".
- **Satu warna utama.** Navy KESB untuk struktur; warna lain hanya membawa makna status.
- **Boleh dicetak dan difotokopi.** A4, Arial, garis hitam nipis. Makna tidak bergantung pada warna sahaja.
- **Yang belum sah ditunjuk jelas.** Maklumat yang belum ada dalam dokumen kontrak ditanda [PERLU SAHKAN] – jangan diteka.

# 2. Warna

<!-- lebar: 38, 22, 99 -->
| Token | Heks | Kegunaan |
|---|---|---|
| Navy KESB | `1F4E78` | Pengepala jadual Excel, jalur kulit, Tajuk 1, carta alir |
| Biru pucat | `DDEBF7` | Pengepala jadual Word, baris trade Excel, kotak proses carta alir |
| Biru sederhana | `BDD7EE` | Jalur "BAHAGIAN A/B" dalam borang; jalur kategori dalam daftar |
| Kelabu label | `F2F2F2` | Sel label (maklumat projek, blok kawalan) |
| Kelabu teks | `595959` | Pengepala, pengaki, keterangan |
| Garis Excel | `999999` | Sempadan nipis lembaran kerja |
| Merah sahkan | `C00000` | [PERLU SAHKAN], cop BATAL dan SALINAN TERKAWAL |
| Lulus / Selesai | `E2EFDA` | Status Diluluskan; titik ITP R |
| Dikemukakan | `FFF2CC` | Status Dikemukakan / menunggu; titik ITP W |
| Hold / Ditolak | `F8CBAD` | Titik ITP H; status Ditolak / Lewat |
| Keputusan | `FFF2BF` | Rombus keputusan carta alir |
| Cabang sisi | `FCE3D6` | Kotak pengecualian carta alir |
| Gelung | `BF3326` | Anak panah "Tidak" kembali ke langkah terdahulu |

# 3. Tipografi
Satu keluarga fon sahaja: **Arial**. Gaya Word dalam templat sudah ditetapkan – pilih gaya, jangan format manual.

<!-- lebar: 38, 30, 91 -->
| Gaya Word | Saiz | Ciri |
|---|---|---|
| Tajuk dokumen (kulit) | 16 pt | Tebal, putih atas jalur navy |
| Heading 1 | 13 pt | Tebal, navy, HURUF BESAR, garis bawah navy, mula muka surat baharu |
| Heading 2 | 11 pt | Tebal, hitam, bernombor (5.3) |
| Heading 3 | 10.5 pt | Tebal condong |
| Normal | 10.5 pt | Jarak baris 1.15; 6 pt selepas perenggan |
| Jadual KESB | 9.5 pt | Teks dalam sel; pengepala tebal |
| Nota KESB | 10 pt | Bingkai kiri navy, latar biru sangat pucat |
| Amaran KESB | 10 pt | Bingkai kiri merah – K&K, bangunan beroperasi dan larangan undang-undang sahaja |
| PERLU SAHKAN (aksara) | – | Merah C00000 tebal |
| Pengepala / pengaki | 8 pt | Kelabu; nama syarikat navy tebal |

# 4. Halaman & Grid

| Jenis | Saiz | Margin (atas × bawah × kiri × kanan) | Lebar teks | Fon badan |
|---|---|---|---|---|
| PQP / SOM / SOP / MS | A4 menegak | 25 × 20 × 25 × 25 mm | 159 mm | 10.5 pt |
| Borang KESB-QF | A4 menegak | 12 × 12 × 18 × 18 mm | 174 mm | 9.5 pt |
| ITP / daftar / carta alir | A4 melintang | 15 mm semua | 267 mm | 9 pt |

- **Pengepala** (semua muka surat kecuali kulit): kiri KARIM ENGINEERING SDN BHD; kanan no. dokumen dan revisi; garis bawah navy.
- **Pengaki:** kiri nota salinan terkawal; kanan "Muka surat X daripada Y".
- **Borang:** pengepala kanan sahaja – KESB-QF-xx | Rev n | KESB/UTM-P3/CP-QC/001.
- **Jarak:** Heading 1 mula muka surat baharu; satu perenggan kecil selepas setiap jadual.
- **Lebar lajur:** jumlah = lebar teks. Lajur huraian paling lebar; lajur Bil / Rev 10–14 mm. Dalam fail .md tetapkan dengan baris `<!-- lebar: 10, 70, 30, 22, 27 -->` sebelum jadual.

# 5. Penomboran Dokumen

<!-- lebar: 40, 50, 69 -->
| Dokumen | No. dokumen | Catatan |
|---|---|---|
| Project Quality Plan | KESB/UTM-P3/CP-QC/001 | Nama kontrak: Construction Plan (Quality Control), Prelim 3.0. Nombor sedia ada dikekalkan |
| Standard Operation Manual | KESB/UTM-P3/SOM/001 | Baharu – cadangan KESB |
| Prosedur operasi standard | KESB/UTM-P3/SOP-01 … SOP-nn | Satu prosedur satu nombor |
| Method statement | KESB/UTM-P3/MS-01 … MS-14 | Sedia ada; Prelim 3.0 (ii) |
| Borang | KESB-QF-01 … KESB-QF-29 | Sedia ada; PQP Seksyen 15.3 |
| Panduan gaya ini | KESB/UTM-P3/RB/001 | Baharu – cadangan KESB |
| Rekod berturutan | NCR-001, RFA-01, ISU-001 | Ikut daftar masing-masing |

Revisi ialah integer bermula Rev 0 (draf pertama untuk semakan HQ). Setiap revisi direkod dalam Sejarah Revisi bersama seksyen yang dipinda. Versi lama diarkibkan dan dicop **BATAL**.

# 6. Kulit & Blok Kawalan
Setiap PQP, SOM dan SOP (jika diterbitkan berasingan) bermula dengan:
1. **Kulit** – nama syarikat; jalur navy dengan jenis dokumen dan tajuk; nama projek, pakej dan no. kontrak; blok kawalan (no. dokumen, revisi/tarikh, disediakan, disemak, diluluskan, diterima); ruang no. salinan terkawal.
2. **Sejarah revisi** – Rev · Tarikh · Perubahan (seksyen dipinda) · Disediakan · Diluluskan.
3. **Senarai edaran salinan terkawal** – No. salinan · Pemegang.

Rantaian kelulusan (ikut CP-QC):

| Peranan | Nama | Jawatan |
|---|---|---|
| Disediakan | Muhamad Azrin bin Rahmat | Pengurus Projek cum Site Agent |
| Disemak | Muhammad Syahidzul Ikhwan | Penyelaras Projek (Ibu Pejabat KESB) |
| Diluluskan | Zaillan Md Ismail | Pengarah Projek KESB |
| Diterima | [PERLU SAHKAN: nama S.O.] | Pegawai Penguasa (S.O.), UTM |

> ! Penyedia, penyemak dan pelulus mestilah orang berbeza. Tiada orang meluluskan kerja sendiri.

Nilai kontrak, harga BQ dan isi SST tidak dicetak dalam SOM / SOP kerana manual itu diedarkan kepada subkon. Logo UTM hanya digunakan dengan kebenaran S.O. [PERLU SAHKAN]

# 7. Cop & Status
- **Cop fizikal:** SALINAN TERKAWAL (merah, bernombor), BATAL (merah), SALINAN TIDAK TERKAWAL dan DRAF – UNTUK SEMAKAN (navy).
- **Titik ITP:** H – Hold (kerja berhenti hingga kelulusan bertulis S.O.); W – Witness (S.O. dimaklumkan, boleh hadir); R – Review (semakan rekod / permit).
- **Status daftar:** Dikemukakan, Diluluskan, Selesai, Ditolak, Lewat, Batal. Status sentiasa ditulis dalam teks; warna hanya membantu imbasan.

# 8. Jadual
**Word ("Table Grid"):** garis hitam nipis; baris pengepala tebal berlatar `DDEBF7` dan diulang pada setiap muka surat; teks 9.5 pt. Jadual label (dua lajur, tanpa pengepala) – lajur kiri tebal berlatar `F2F2F2`.

**Excel (ITP, daftar):** tajuk lembaran 14 pt tebal di A1; pengepala navy `1F4E78` teks putih tebal; garis `999999`; baris trade `DDEBF7` tebal; baris pengepala dibekukan; tapis automatik aktif; format bersyarat untuk status.

**Borang:** jalur bahagian – satu sel penuh lebar, `BDD7EE`, tebal (cth BAHAGIAN B: PEMERIKSAAN OLEH S.O.).

# 9. Carta Alir Prosedur
Format Lampiran K: lajur Peranan · Langkah · Cabang sisi · Rekod. Setiap SOP mempunyai satu carta.

<!-- lebar: 52, 18, 89 -->
| Simbol | Kod | Kegunaan |
|---|---|---|
| Kapsul navy, teks putih | s / e | Mula / tamat |
| Segi empat biru pucat | p | Langkah proses – mula dengan kata kerja |
| Rombus kuning | d | Keputusan – soalan Ya / Tidak |
| Segi empat salmon | side | Cabang pengecualian |
| Anak panah merah | loop | Gelung "Tidak" ke langkah terdahulu |

Jalur tajuk navy dengan tajuk putih 15 pt; pengaki: no. dokumen di kiri, muka surat di kanan. Penjana: `scripts/build_carta_alir.py`.

# 10. Borang & Tandatangan
Komponen tetap borang (ikut `scripts/build_borang.py`): nama syarikat dan tajuk "KESB-QF-xx: TAJUK" di tengah; jadual maklumat projek (label kelabu); jadual label; pilihan "☐"; senarai semak Perkara · Ya · Tidak · T.B.; ruang teks bebas; blok tandatangan (Tandatangan · Nama · Tarikh) bagi setiap peranan.

| Disediakan | Disemak | Diluluskan | Diterima (S.O.) |
|---|---|---|---|
| Tandatangan: | Tandatangan: | Tandatangan: | Tandatangan & cop: |
| Nama: | Nama: | Nama: | Nama: |
| Tarikh: | Tarikh: | Tarikh: | Tarikh: |

# 11. Nota & Penanda
- [PERLU SAHKAN: perkara] – merah tebal dalam kurungan siku, menyatakan apa yang perlu disahkan.
- Rujukan kontrak dalam kurungan selepas ayat: (BQ A Prelim 11.0), (SST para 9(c)).
- Frasa tetap: "cadangan KESB", "sasaran dalaman KESB".

> Nota: Kotak nota (gaya Nota KESB) untuk penjelasan, pengecualian atau rujukan silang. Dalam .md: baris bermula dengan "> ".

> ! Kotak amaran (gaya Amaran KESB) hanya untuk risiko K&K, bangunan beroperasi, permit kerja dan larangan undang-undang. Dalam .md: baris bermula dengan "> !".

# 12. Pematuhan Perundangan
Setiap dokumen KESB mesti menunjukkan bukan sahaja kewajipan kontrak, tetapi juga kewajipan undang-undang KESB sebagai **majikan** dan **kontraktor binaan**. Komponen berikut digunakan:

**a) Daftar Perundangan** – satu jadual induk (PQP Seksyen 4.4 dan SOM Seksyen 6; sumber bersama `sumber/_daftar-perundangan.md`). Lajur tetap: Bil · Akta / peraturan · Keperluan utama · Bukti / rekod · SOP · Pemilik. Kategori dipisahkan dengan jalur biru sederhana: A JKKP, B CIDB, C Pekerjaan, D Kewajipan majikan, E Alam sekitar & utiliti.

**b) Gaya sitasi** – nama penuh dan nombor akta pada penggunaan pertama, kemudian singkatan: "Akta Keselamatan dan Kesihatan Pekerjaan 1994 (Akta 514)", seterusnya "OSHA 1994". Peraturan: nama penuh + tahun. Nombor seksyen hanya ditulis selepas disemak dengan teks rasmi; jika belum, tanda [PERLU SAHKAN].

**c) Baris Perundangan dalam blok tajuk SOP** – menyenaraikan no. daftar (cth B1, B2, D1) supaya setiap prosedur boleh dikesan ke akta.

**d) Kotak amaran** – untuk larangan undang-undang (cth pekerja tanpa Kad CIDB yang sah tidak dibenarkan bekerja).

**e) Semakan berkala** – daftar disemak setiap 6 bulan atau apabila undang-undang dipinda; perubahan direkod dalam Sejarah Revisi.

> Nota: "Akta Majikan" bukan nama satu akta. Kewajipan majikan diliputi oleh Akta Kerja 1955, Akta Keselamatan Sosial Pekerja 1969 (PERKESO), Akta Sistem Insurans Pekerjaan 2017, Akta KWSP 1991, Akta 446 dan Perintah Gaji Minimum – kategori C dan D dalam daftar.

<!-- sertakan: _daftar-perundangan.md -->

> ! Daftar ini ialah ringkasan untuk kawalan dokumen, bukan nasihat undang-undang. Sahkan teks dan pindaan terkini dengan JKKP, CIDB, JTK dan HQ KESB sebelum dokumen dikemukakan kepada S.O.

# 13. Saluran Komunikasi & Notifikasi
Setiap notifikasi dan maklumat dihantar melalui saluran yang ditetapkan. Saluran diberi kod (R rasmi kontrak, P pengguna bangunan, D dalaman tapak, K kecemasan & statutori) supaya prosedur SOP boleh merujuk kod, cth "R1 + R3". Jadual ini dimuatkan dalam PQP Seksyen 5.5 dan SOM Seksyen 7.

<!-- sertakan: _saluran-komunikasi.md -->

# 14. Struktur PQP (CP-QC)
Ikut CP-QC Rev 5 – 15 seksyen, selari MS ISO 9001:2015 dan BQ A Prelim 3.0. Templat: `Templat-PQP.docx`.
1. Pengenalan
2. Polisi & Objektif Kualiti
3. Organisasi Projek
4. Dokumen Rujukan (termasuk 4.4 Perundangan terpakai)
5. Kawalan Dokumen & Rekod (termasuk 5.5 Saluran komunikasi & notifikasi)
6. Pengurusan Submittal
7. Kawalan Subkontraktor, Pasukan Dalaman & Pembekal
8. Inspection & Test Plan
9. Kawalan Bahan & Ujian
10. Kawalan Peralatan Ukur & Ujian
11. Non-Conformance & Tindakan Pembetulan
12. Audit Dalaman & Semakan Pengurusan
13. Latihan & Kompetensi
14. Penyerahan & DLP
15. Lampiran

# 15. Struktur SOM & SOP
PQP menyatakan **apa** dikawal dan **siapa** bertanggungjawab; SOM menyatakan **bagaimana**, langkah demi langkah, untuk pasukan tapak dan subkon. Templat: `Templat-SOM.docx` (contoh SOP-05 NCR, SOP-13 Kemasukan Pekerja dan SOP kosong).

**SOM:** 1 Tujuan & penggunaan · 2 Hubungan dengan PQP, MS & ITP · 3 Peranan & kuasa (RACI) · 4 Indeks SOP · 5 SOP-01 … SOP-15 · 6 Daftar perundangan · 7 Saluran komunikasi & notifikasi · 8 Borang & daftar · 9 Istilah & singkatan. SOP-11 hingga SOP-15 melaksanakan kewajipan JKKP, CIDB, Akta Kerja dan caruman majikan.

**Setiap SOP:** blok tajuk (no., revisi, pemilik, rujukan, perundangan) · 1 Tujuan · 2 Skop · 3 Tanggungjawab · 4 Prosedur (Bil · Tindakan · Oleh · Masa · Rekod) · 5 Carta alir · 6 Rekod & penyimpanan.

# 16. Bahasa & Penulisan
- Tulis dalam Bahasa Melayu; istilah teknikal Inggeris dikekalkan (method statement, shop drawing, T&C).
- Langkah prosedur bermula dengan kata kerja: "Kemukakan…", "Rekod…". Nyatakan siapa, bila (hari bekerja) dan rekod.
- Tarikh DD/MM/YYYY; wang RM9,999.00; masa 24 jam.
- Jawatan ikut carta organisasi kontrak sahaja (jawatan "Wakil Kualiti" telah dibuang dalam CP-QC Rev 5).
- Jangan teka – gunakan [PERLU SAHKAN]. Jangan guna ayat pasif tanpa pelaku. Takrifkan singkatan pada penggunaan pertama.

# 17. Fail & Penjanaan

| Fail | Kandungan |
|---|---|
| sistem-reka-bentuk/index.html | Sistem reka bentuk interaktif |
| Panduan-Gaya-Dokumen-KESB.docx | Panduan ini |
| Templat-PQP.docx | Templat PQP dengan gaya terpasang |
| Templat-SOM.docx | Templat SOM dengan contoh SOP |
| sumber/*.md | Sumber teks |
| sumber/_daftar-perundangan.md | Daftar perundangan (sumber tunggal) |
| sumber/_saluran-komunikasi.md | Saluran komunikasi & notifikasi (sumber tunggal) |
| scripts/build_reka_bentuk_docx.py | Penjana Word daripada .md; mengisi semula jadual bersama dalam index.html |

Jana semula: `py -3.12 scripts/build_reka_bentuk_docx.py`. Untuk dokumen baharu: salin templat .md, isi meta dan kandungan, kemudian `py -3.12 scripts/build_reka_bentuk_docx.py <fail.md> <fail.docx>`.
