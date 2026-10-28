"""Baca BQ kontrak (Excel/CSV): kumpul item kerja ikut trade, dan ambil klausa
kontrak (Preliminaries / syarat) yang menyentuh kualiti.

Guna:  py -3.12 scripts/bq_extract.py docs/BQ.xlsx
Hasil: output/bq_summary.json
"""
import csv
import json
import re
import sys
from pathlib import Path

import openpyxl

# Kata kunci untuk kenal pasti trade. Tambah ikut keperluan projek.
TRADES = {
    "Kerja Awalan": ["preliminar", "awalan", "site office", "hoarding"],
    "Kerja Tanah": ["earthwork", "excavat", "kerja tanah", "korek", "backfill", "timbus", "fill"],
    "Cerucuk": ["pile", "cerucuk", "piling"],
    "Konkrit": ["concrete", "konkrit", "grade c", "lean", "blinding"],
    "Tetulang": ["reinforcement", "tetulang", "rebar", "fabric", "brc", "high yield", "mild steel bar"],
    "Acuan (Formwork)": ["formwork", "acuan", "shutter"],
    "Keluli Struktur": ["structural steel", "keluli struktur", "steel beam", "steel column", "purlin"],
    "Kerja Bata / Blok": ["brick", "bata", "block", "blok"],
    "Kalis Air": ["waterproof", "kalis air", "membrane", "damp proof"],
    "Bumbung": ["roof", "bumbung", "gutter", "flashing", "metal deck"],
    "Plaster & Kemasan Dinding": ["plaster", "render", "skim", "lepa"],
    "Kemasan Lantai": ["tile", "jubin", "screed", "terrazzo", "vinyl", "carpet", "granite", "marble"],
    "Siling": ["ceiling", "siling"],
    "Pintu & Tingkap": ["door", "pintu", "window", "tingkap", "ironmongery", "glazing", "louvre"],
    "Mengecat": ["paint", "cat ", "emulsion", "primer"],
    "Paip & Sanitari": ["plumbing", "pipe", "paip", "sanitary", "wc", "basin", "water tank", "tangki"],
    "Elektrik": ["electrical", "elektrik", "cable", "kabel", "lighting", "lampu", "socket", "db ", "switch"],
    "Mekanikal (ACMV / Fire)": ["acmv", "air cond", "fire", "kebakaran", "sprinkler", "hose reel", "lift"],
    "Kerja Luar & Saliran": ["drain", "longkang", "saliran", "road", "jalan", "premix", "kerb", "culvert", "fence", "pagar"],
    "Landskap": ["landscape", "landskap", "turf", "rumput", "tree", "pokok"],
}
HEADER_KEYS = ["description", "keterangan", "perihal", "butiran"]

# Klausa kontrak (Preliminaries / syarat) yang menyentuh kualiti
QUALITY_TOPICS = {
    "Ujian & Sampel": ["test", "ujian", "sample", "sampel", "laborator", "makmal"],
    "Pemeriksaan": ["inspect", "pemeriksaan", "periksa", "witness", "hold point"],
    "Kelulusan & Submittal": ["approv", "kelulusan", "lulus", "submit", "shop drawing", "method statement"],
    "Sistem Kualiti": ["quality", "kualiti", "iso 9001", "qlassic", "qa/qc", "quality plan"],
    "Kecacatan & DLP": ["defect", "kecacatan", "rectif", "baiki", "liability period", "dlp"],
    "Jaminan & Waranti": ["warrant", "waranti", "guarantee", "jaminan"],
    "Rekod & Dokumen": ["record", "rekod", "as-built", "as built", "o&m", "manual"],
    "Bahan & Mutu Kerja": ["material", "bahan", "workmanship", "mutu kerja", "standard", "sirim", "mpob"],
    "Penalti": ["liquidated", "l.a.d", "ganti rugi", "penalt"],
}
CLAUSE_MIN_LEN = 60  # baris tanpa unit/kuantiti yang lebih panjang = klausa, bukan tajuk
LUMP_SUM_UNITS = {"item", "ls", "l.s.", "sum", "lump sum", "prov"}


def read_rows(path: Path):
    if path.suffix.lower() == ".csv":
        with path.open(encoding="utf-8-sig") as f:
            yield path.stem, list(csv.reader(f))
        return
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    for ws in wb.worksheets:
        rows = [["" if c is None else str(c).strip() for c in r] for r in ws.iter_rows(values_only=True)]
        yield ws.title, rows


def find_columns(rows):
    """Cari baris header dan indeks lajur deskripsi, unit, kuantiti."""
    for i, row in enumerate(rows[:30]):
        low = [c.lower() for c in row]
        desc = next((j for j, c in enumerate(low) if any(k in c for k in HEADER_KEYS)), None)
        if desc is None:
            continue
        unit = next((j for j, c in enumerate(low) if c in ("unit", "unit ukuran") or c.startswith("unit")), None)
        qty = next((j for j, c in enumerate(low) if c.startswith(("qty", "quantity", "kuantiti"))), None)
        return i, desc, unit, qty
    return None


def match(text: str):
    """Trade yang kata kuncinya muncul paling awal dalam teks
    (cth. 'concrete in pile cap' -> Konkrit, bukan Cerucuk)."""
    t = text.lower()
    hits = [(t.find(k), trade) for trade, keys in TRADES.items() for k in keys if k in t]
    return min(hits)[1] if hits else None


def clause_topics(text: str):
    t = text.lower()
    return [topic for topic, keys in QUALITY_TOPICS.items() if any(k in t for k in keys)]


def classify(desc: str, section: str) -> str:
    # Deskripsi item diutamakan; tajuk seksyen hanya sebagai sandaran
    return match(desc) or match(section) or "Lain-lain (perlu semak)"


def main():
    if len(sys.argv) < 2:
        sys.exit("Guna: py -3.12 scripts/bq_extract.py <fail BQ .xlsx/.csv> [fail lain ...]")

    summary = {}
    clauses = []
    skipped = []
    for arg in sys.argv[1:]:
        path = Path(arg)
        for sheet, rows in read_rows(path):
            cols = find_columns(rows)
            if not cols:
                skipped.append(f"{path.name} / {sheet}")
                continue
            start, d, u, q = cols
            section = sheet
            for row in rows[start + 1:]:
                desc = row[d] if d < len(row) else ""
                if not desc:
                    continue
                desc = re.sub(r"\s+", " ", desc)
                unit = row[u] if u is not None and u < len(row) else ""
                qty = row[q] if q is not None and q < len(row) else ""
                source = f"{path.name} / {sheet}"
                topics = clause_topics(desc)
                # Baris tanpa unit dan kuantiti: tajuk seksyen (pendek) atau klausa kontrak (panjang)
                if not unit and not qty:
                    if len(desc) < CLAUSE_MIN_LEN:
                        section = desc
                    else:
                        clauses.append({"sumber": source, "seksyen": section, "klausa": desc,
                                        "topik_kualiti": topics})
                    continue
                # Item Preliminaries berharga (Item / LS) yang menyebut kewajipan kualiti
                if topics and unit.lower() in LUMP_SUM_UNITS:
                    clauses.append({"sumber": source, "seksyen": section, "klausa": desc,
                                    "topik_kualiti": topics})
                trade = classify(desc, section)
                summary.setdefault(trade, []).append({
                    "sumber": source,
                    "seksyen": section,
                    "item": desc[:200],
                    "unit": unit,
                    "kuantiti": qty,
                })

    quality = [c for c in clauses if c["topik_kualiti"]]
    out = Path("output/bq_summary.json")
    out.parent.mkdir(exist_ok=True)
    out.write_text(json.dumps({"trades": summary,
                               "klausa_kualiti": quality,
                               "klausa_lain": [c for c in clauses if not c["topik_kualiti"]],
                               "sheet_tidak_dibaca": skipped},
                              ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Siap: {out}")
    for trade, items in sorted(summary.items(), key=lambda x: -len(x[1])):
        print(f"  {trade:32} {len(items):4} item")
    print(f"Klausa kontrak: {len(clauses)} ({len(quality)} berkaitan kualiti)")
    if skipped:
        print("Sheet tanpa header deskripsi (tidak dibaca):", ", ".join(skipped))


if __name__ == "__main__":
    main()
