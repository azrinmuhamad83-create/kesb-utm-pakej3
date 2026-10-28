---
name: quality-plan
description: Hasilkan Project Quality Plan (PQP) dalam Word dan ITP dalam Excel untuk projek pembinaan, berdasarkan drawing dan BQ. Guna bila user minta quality plan, PQP, pelan kualiti, atau ITP.
---

# Project Quality Plan (PQP) – Pembinaan

Aliran: drawing + BQ -> agent baca -> tulis PQP & ITP -> semak QA -> jana Word & Excel.

## Langkah 1: Semak input
1. Baca `CLAUDE.md`. Kalau maklumat penting masih [ ] (nama projek, klien, skop),
   tanya user dulu. Jangan teruskan dengan andaian.
2. Pastikan ada BQ dalam `docs/` dan drawing dalam `docs/drawings/`. Beritahu user apa yang tiada.

## Langkah 2: Baca dokumen (guna agent, jalankan serentak)
- Agent `bq-analyst` -> `output/bq_analisis.md` + `output/bq_summary.json`
  (BQ ialah dokumen kontrak: hasilnya termasuk **daftar kewajipan kualiti kontrak**)
- Agent `drawing-reader` -> `output/drawing_analisis.md`

Keutamaan dokumen kalau bercanggah: ikut susunan dalam `CLAUDE.md`. Catat setiap
percanggahan sebagai `[PERLU SAHKAN: percanggahan ...]`. Jangan pilih sendiri.

## Langkah 3: Tulis ITP -> `output/itp.json`
Satu entri untuk setiap trade yang wujud dalam BQ (abaikan trade yang tiada).
Ikut format dalam docstring `scripts/build_itp_xlsx.py`. Untuk setiap trade:
- Aktiviti mengikut urutan kerja: sebelum, semasa, selepas
- Kriteria penerimaan diambil dari kontrak/BQ, nota drawing atau spec, dengan rujukan
  (no. klausa / item BQ / no. drawing)
- Kekerapan ujian ikut kontrak kalau dinyatakan dalam BQ atau Preliminaries
- `titik`: H untuk kerja yang akan tertutup (tetulang sebelum konkrit, paip sebelum timbus,
  kalis air sebelum screed); W untuk ujian; R untuk semakan dokumen
- Isi `ujian_bahan` untuk semua bahan yang dikenal pasti oleh bq-analyst

## Langkah 4: Tulis PQP -> `output/CP-QC-draf.md`
Baris pertama: `# Project Quality Plan – <nama projek>`. Kemudian seksyen `##` ini:

1. Pengenalan – tujuan, skop (dari BQ & drawing), butiran kontrak (borang kontrak, tempoh, DLP, LAD)
   + jadual **Pematuhan Kewajipan Kontrak**: Kewajipan | Rujukan klausa | Seksyen PQP
   (setiap kewajipan dari `bq_analisis.md` mesti ada seksyen PQP yang memenuhinya)
2. Polisi & Objektif Kualiti – objektif SMART + KPI (QLASSIC, NCR ditutup ≤ 14 hari, lulus inspection kali pertama ≥ 90%)
3. Organisasi Projek – peranan & tanggungjawab (jadual)
4. Dokumen Rujukan – standard, spesifikasi, daftar drawing (dari drawing-reader)
5. Kawalan Dokumen & Rekod
6. Pengurusan Submittal – material approval, shop drawing, method statement (senaraikan ikut trade)
7. Kawalan Subkontraktor & Pembekal
8. Inspection & Test Plan – ringkasan trade dan hold point utama; rujuk "Lampiran A: ITP.xlsx"
9. Kawalan Bahan & Ujian – rujuk sheet "Ujian Bahan"
10. Kawalan Peralatan Ukur & Ujian – kalibrasi
11. Non-Conformance & Tindakan Pembetulan – aliran, tempoh, kuasa penutupan; rujuk sheet "Daftar NCR"
12. Audit Dalaman & Semakan Pengurusan
13. Latihan & Kompetensi
14. Penyerahan & DLP
15. Lampiran – senarai borang

Peraturan:
- Setiap kenyataan mesti boleh diaudit, iaitu ada siapa, bila dan rekod apa.
- Jangan reka nilai teknikal atau no. klausa. Kalau tak pasti, tulis `[PERLU SAHKAN: ...]`
  (script Word akan warnakan merah).
- Jadual guna format markdown `| ... |`. Guna `---` untuk page break.

## Langkah 5: Semakan QA
Jalankan agent `qa-reviewer`. Betulkan semua isu Major dalam `PQP-draf.md` dan `itp.json`.

## Langkah 6: Jana fail
```
py -3.12 scripts/build_itp_xlsx.py
py -3.12 scripts/build_pqp_docx.py
```

## Langkah 7: Lapor kepada user
- Laluan `output/Construction-Plan-QC.docx` dan `output/ITP.xlsx`
- Ringkasan: bilangan trade, aktiviti ITP, hold point
- Senarai semua `[PERLU SAHKAN]` yang user perlu isi
- Isu QA Minor yang belum dibetulkan
