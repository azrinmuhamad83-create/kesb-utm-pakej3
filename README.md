# PQP-Projek – Cara Guna

Input: **drawing (PDF)** + **BQ (Excel/PDF)**. Output: **PQP.docx (Word)** + **ITP.xlsx (Excel)**.

## Aliran
```
 docs/drawings/*.pdf          docs/BQ.xlsx
        |                          |
 [agent drawing-reader]     [agent bq-analyst]
        |                          |
        +-------------+------------+
                      |
          [skill /quality-plan]  -> PQP-draf.md + itp.json
                      |
             [agent qa-reviewer]  -> senarai isu, betulkan
                      |
              [code Python]
           /                 \
   output/PQP.docx      output/ITP.xlsx
```

## Struktur
```
PQP-Projek/
├── CLAUDE.md                        <- maklumat projek (isi dulu)
├── docs/
│   ├── BQ.xlsx                      <- BQ kontrak (termasuk Preliminaries)
│   ├── drawings/                    <- letak semua drawing PDF di sini
│   └── kontrak/                     <- (jika ada) Syarat Kontrak, Spesifikasi
├── output/                          <- semua hasil keluar di sini
├── scripts/                         <- CODE
│   ├── bq_extract.py                   BQ -> kumpul item ikut trade
│   ├── drawing_extract.py              drawing PDF -> imej + teks
│   ├── build_itp_xlsx.py               itp.json -> ITP.xlsx
│   └── build_pqp_docx.py               PQP-draf.md -> PQP.docx
└── .claude/
    ├── skills/quality-plan/SKILL.md <- SKILL: ketua aliran
    └── agents/
        ├── bq-analyst.md            <- AGENT: baca BQ
        ├── drawing-reader.md        <- AGENT: baca drawing
        └── qa-reviewer.md           <- AGENT: semak QA (baca sahaja)
```

## Persediaan (sekali sahaja)
```
py -3.12 -m pip install openpyxl python-docx pymupdf
```

## Langkah guna
1. Buka folder `PQP-Projek` dalam VS Code (File -> Open Folder).
2. Isi `CLAUDE.md`.
3. Letak BQ dalam `docs/`, dan drawing dalam `docs/drawings/`.
4. Buka panel Claude Code dan taip: `/quality-plan`
5. Claude akan jalankan semua langkah dan beri `output/PQP.docx` dan `output/ITP.xlsx`.
6. Buka Word dan isi semua teks merah `[PERLU SAHKAN]`.

## Boleh jalankan langkah satu per satu
- `Guna agent bq-analyst untuk baca BQ`
- `Guna agent drawing-reader untuk baca drawing`
- `Guna agent qa-reviewer untuk semak PQP`
- `py -3.12 scripts/build_pqp_docx.py` (jana semula Word selepas ubah draf)
