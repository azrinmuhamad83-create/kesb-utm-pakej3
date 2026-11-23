# UTM Pakej 3 – Cara Guna

Folder rasmi: **UTM PAKEJ 3 FINAL**. Folder lama PQP UTM PE KESB tidak lagi digunakan untuk dokumen baharu.

Input: lukisan (PDF) dan BQ. Output: Construction Plan (QC) dalam Word, ITP dan daftar-daftar dalam Excel, borang, method statement dan pakej serahan HQ.

## Penamaan fail
Fail hasil jana berakhir `-P3-FINAL`, contohnya `ITP-P3-FINAL.xlsx`. Semua nama ditetapkan dalam `scripts/penamaan.py`; jika mahu tukar tanda itu, ubah di situ sahaja.

| Hasil jana | Nama |
|---|---|
| ITP | `output/ITP-P3-FINAL.xlsx` |
| Construction Plan (QC) | `output/Construction-Plan-QC-P3-FINAL.docx` |
| Daftar Kelulusan Bahan | `output/Daftar-Kelulusan-Bahan-P3-FINAL.xlsx` |
| Daftar Item Terbuka | `output/Daftar-Item-Terbuka-P3-FINAL.xlsx` |
| Lampiran E – Jadual Loji | `output/Lampiran-E-Jadual-Loji-P3-FINAL.xlsx` |
| Lampiran K – Carta Alir | `output/Lampiran-K-Carta-Alir-P3-FINAL.pdf` |
| Pakej serahan HQ | `output/Serahan-HQ-CP-QC-Rev0-P3-FINAL/` (dan `.zip`) |

Tidak ditanda: borang `KESB-QF-xx`, method statement `MS-xx` dan surat, kerana nombor kawalan atau rujukan surat sudah cukup. Fail di dalam pakej HQ sudah bernombor dan ada Rev.

Daftar Kelulusan Bahan ialah rekod hidup yang dikemas kini Kerani Tapak. Jika folder ini masih ada fail bernama lama `Daftar-Kelulusan-Bahan.xlsx`, namakan semula dahulu kepada nama baharu. Skrip akan berhenti dan memberitahu jika nama lama dijumpai, supaya tiada fail kosong dijana di sebelahnya.

## Struktur
```
UTM PAKEJ 3 FINAL/
├── CLAUDE.md                 maklumat projek
├── docs/                     kontrak, BQ, lukisan, SST (tidak dikongsi ke repo)
├── output/                   semua hasil jana (tidak dikongsi ke repo)
├── scripts/                  kod
│   ├── penamaan.py              nama fail hasil jana
│   ├── bq_extract.py            BQ -> kumpul item ikut trade
│   ├── drawing_extract.py       lukisan PDF -> imej + teks
│   ├── build_itp_xlsx.py        itp.json -> ITP
│   ├── build_pqp_docx.py        CP-QC-draf.md -> Construction Plan (Word)
│   ├── build_borang.py          borang KESB-QF
│   ├── build_daftar_kelulusan.py, build_item_terbuka.py, build_lampiran_e.py, build_carta_alir.py
│   ├── build_surat_rfi.py, build_surat_whatsapp.py
│   └── build_pakej_hq.py        himpun pakej serahan HQ
└── .claude/
    ├── skills/                  quality-plan, ms-project
    └── agents/                  bq-analyst, drawing-reader, qa-reviewer
```

## Persediaan (sekali sahaja)
```
py -3.12 -m pip install openpyxl python-docx pymupdf pyyaml pywin32
```

## Langkah guna
1. Buka folder UTM PAKEJ 3 FINAL dalam VS Code.
2. Pastikan `CLAUDE.md` terkini.
3. Jalankan `/quality-plan` dalam panel Claude Code, atau skrip satu per satu, contohnya `py -3.12 scripts/build_itp_xlsx.py`.
4. Buka hasil dalam `output/` dan isi semua teks `[PERLU SAHKAN]`.
5. Untuk pakej HQ, jana dokumen sumber dahulu, kemudian `py -3.12 scripts/build_pakej_hq.py` (perlu Word dan Excel dipasang).
