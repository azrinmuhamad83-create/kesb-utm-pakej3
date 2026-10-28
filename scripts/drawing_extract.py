"""Tukar drawing PDF jadi imej PNG dan ambil teks (title block, nota, legend).

Guna:  py -3.12 scripts/drawing_extract.py docs/drawings
Hasil: output/drawings/<nama>_p<N>.png  dan  output/drawing_text.json
"""
import json
import sys
from pathlib import Path

import pymupdf

DPI = 110  # cukup untuk baca title block; naikkan jika teks kecil


def main():
    if len(sys.argv) < 2:
        sys.exit("Guna: py -3.12 scripts/drawing_extract.py <folder atau fail PDF>")

    src = Path(sys.argv[1])
    pdfs = sorted(src.glob("**/*.pdf")) if src.is_dir() else [src]
    out_dir = Path("output/drawings")
    out_dir.mkdir(parents=True, exist_ok=True)

    result = []
    for pdf in pdfs:
        doc = pymupdf.open(pdf)
        for i, page in enumerate(doc, start=1):
            img = out_dir / f"{pdf.stem}_p{i}.png"
            page.get_pixmap(dpi=DPI).save(img)
            text = page.get_text().strip()
            result.append({
                "fail": pdf.name,
                "muka": i,
                "imej": str(img).replace("\\", "/"),
                "ada_teks": bool(text),  # False = drawing imbasan, perlu baca imej
                "teks": text[:4000],
            })
        doc.close()

    out = Path("output/drawing_text.json")
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    scanned = sum(1 for r in result if not r["ada_teks"])
    print(f"Siap: {len(result)} muka dari {len(pdfs)} PDF -> {out_dir}")
    if scanned:
        print(f"  {scanned} muka tiada lapisan teks (imbasan) - agent perlu baca imej")


if __name__ == "__main__":
    main()
