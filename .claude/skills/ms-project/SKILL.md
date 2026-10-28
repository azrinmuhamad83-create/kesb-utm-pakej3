---
name: ms-project
description: Baca, analisis, bina dan kemas kini jadual Microsoft Project (.mpp) – ringkasan & tarikh siap vs kontrak, critical path, aktiviti lewat, lookahead, S-curve, semakan kualiti jadual, kemas kini progres, dan bina program kerja / recovery programme baharu. Guna bila user sebut MS Project, .mpp, program kerja, jadual kerja, critical path, S-curve, lookahead, progres jadual atau recovery programme.
---

# Microsoft Project – Program Kerja

Semua kerja .mpp melalui MS Project sendiri (COM) yang dipasang dalam komputer ini.
Python hanya membaca XML (MSPDI) yang dieksport oleh MS Project.

- Skrip: `.claude/skills/ms-project/scripts/`
- Python: `C:/Users/azrin/.claude/plugins/cache/construction-skills/construction/0.6.0/bin/construction-python` (ada openpyxl)
- Skrip `.ps1`: jalankan dalam alat **PowerShell** dengan `& "<laluan>\skrip.ps1" ...`
  (jangan guna `powershell -ExecutionPolicy Bypass -File` – disekat). Setiap panggilan
  membuka MS Project secara tersembunyi: ambil 10–60 saat; beri timeout 300000.
- Fail kerja (XML eksport, CSV sementara): `.construction/skills/ms-project/`
- Hasil untuk user (Excel, .mpp baharu): `output/programme/`
- **Jangan tulis ganti fail .mpp asal** dalam `docs/`. Skrip menolak -Out sama dengan -In.

## Fail program projek ini
- `docs/work program/UTM - Detailed Work Program (261020).mpp` – 541 aktiviti, berkos (RM9,090,675), ada baseline.
  Kalendar 7 hari. Siap 16/08/2028 – **2 hari melepasi tarikh siap SST 14/08/2028**.
- Master Programme REV00 (PDF sahaja – 3 versi: asal, "august start", "oktober start").
  PDF Gantt tidak boleh dibaca skrip; rasterize & baca dengan vision ikut construction-guide.
- Fail `output/programme/*.xml` lama ialah .mpp yang ditukar nama – **bukan** XML. Abaikan.
- Rujukan kontrak: tempoh 104 minggu, milik tapak 18/08/2026, siap 14/08/2028 (CLAUDE.md).

## Langkah 1: Tukar .mpp ke XML (sekali bagi setiap versi fail)
```powershell
& ".claude\skills\ms-project\scripts\mpp_convert.ps1" -In "docs\work program\<fail>.mpp" -Out ".construction\skills\ms-project\<nama>.xml"
```
Juga boleh: `-Out x.xlsx` (eksport Excel oleh MS Project) dan `.xml -> .mpp`.
Guna semula XML yang sudah ada jika .mpp tidak berubah (banding tarikh fail).

## Langkah 2: Analisis – `mspdi.py <subarahan> <xml>`
| Subarahan | Guna untuk | Contoh pilihan |
|---|---|---|
| `summary` | Ringkasan, tarikh siap vs kontrak, % kemajuan (kos), pecahan aras 2 | `--contract-finish 2028-08-14` |
| `critical` | Aktiviti kritikal (slack ≤ 0) | `--out x.xlsx` |
| `lookahead` | Aktiviti 2–4 minggu ke depan untuk mesyuarat tapak | `--from 2026-10-26 --weeks 3 --out x.xlsx` |
| `late` | Aktiviti lewat vs baseline pada status date | `--status-date 2026-10-26 --out x.xlsx` |
| `scurve` | S-curve rancang (baseline) vs forecast vs aktual + carta Excel | `--out x.xlsx --period month --weight cost` |
| `check` | Semakan kualiti jadual (tiada logik, kekangan, slack negatif, tugasan manual, tempoh panjang) | `--out x.xlsx` |
| `tasks` | Senarai penuh semua tugasan | `--out x.xlsx` atau `.csv` |

Tanpa `--out`, hasil dicetak ke skrin (untuk dibaca sendiri sebelum menjawab).
Bila `--out` ialah `.xlsx`, hantar fail kepada user.

## Langkah 3 (jika diminta): Kemas kini progres
1. Dapatkan progres dari user atau laporan tapak. Padankan setiap aktiviti kepada **UID**
   (lajur dari `tasks`). Nama aktiviti berulang ikut blok/ruang – sahkan dengan WBS/Induk,
   jangan teka. Senaraikan padanan kepada user sebelum menulis jika ada keraguan.
2. Tulis CSV (UTF-8) dalam `.construction/skills/ms-project/`:
   `UID,ActualStart,PercentComplete,ActualFinish,Notes`
3. Jalankan (fail baharu, asal tidak diubah):
```powershell
& ".claude\skills\ms-project\scripts\mpp_update.ps1" -In "docs\work program\<fail>.mpp" -Csv ".construction\skills\ms-project\progres.csv" -StatusDate 2026-11-15 -Out "output\programme\<fail>_kemaskini_20261115.mpp"
```
4. Tukar hasil ke XML dan jalankan `summary` + `late` untuk laporkan kesan pada tarikh siap.

## Langkah 4 (jika diminta): Bina program baharu / recovery programme
1. Tulis JSON ikut format dalam kepala `scripts/mpp_build.ps1`
   (tugasan, aras, tempoh hari bekerja, predecessor FS/SS/FF/SF + lag, kos, SNET, nota).
   Kalendar lalai Isnin–Sabtu 8 jam; tambah cuti umum Johor jika berkaitan.
2. Jalankan:
```powershell
& ".claude\skills\ms-project\scripts\mpp_build.ps1" -Json ".construction\skills\ms-project\program.json" -Out "output\programme\<nama>.mpp"
```
3. Tukar ke XML, jalankan `summary` dan `check`. Betulkan isu Tinggi/Sederhana sebelum serah.
4. Program untuk projek ini mesti siap ≤ 14/08/2028 dan ikut fasa penutupan ruang
   (Lampiran Preamble/1) dalam CLAUDE.md. Jika tidak boleh, nyatakan dengan jelas.

## Peraturan
- Tarikh siap & fasa diukur terhadap SST / CLAUDE.md, bukan terhadap jadual itu sendiri.
- `late` dan rancangan S-curve bergantung pada baseline. Jika `check` lapor "Tiada baseline",
  beritahu user – jangan anggap jadual semasa sebagai baseline.
- Kos dalam MSPDI disimpan dalam sen; skrip sudah menukar ke RM.
- % kemajuan ialah apa yang direkod dalam .mpp. Jangan isi atau anggar progres tanpa data user.
- Laporkan dalam Bahasa Melayu, dengan ID/WBS aktiviti sebagai rujukan.
