"""Penamaan fail hasil jana – satu tempat sahaja.

Semua skrip build ambil nama fail dari sini. Hujung `-P3-FINAL` menandakan fail itu datang
daripada folder rasmi UTM PAKEJ 3 FINAL, supaya tidak keliru dengan fail lama dalam
folder PQP UTM PE KESB.

Tidak dikenakan pada: borang KESB-QF-xx, method statement MS-xx dan surat (nombor kawalan /
rujukan surat sudah cukup sebagai pengenalan), serta fail di dalam pakej serahan HQ
(sudah bernombor dan ada Rev; folder pakej yang ditanda).
"""
import sys
from pathlib import Path

TANDA = "-P3-FINAL"


def nama(asas: str, ext: str) -> str:
    """nama("ITP", ".xlsx") -> "ITP-P3-FINAL.xlsx" """
    return f"{asas}{TANDA}{ext}"


def keluaran(asas: str, ext: str, folder: str = "output") -> Path:
    return Path(folder) / nama(asas, ext)


def semak_nama_lama(baru: Path, lama: str) -> None:
    """Untuk rekod hidup: jika fail bernama lama wujud tetapi yang baru belum ada,
    berhenti supaya tidak jana fail kosong di sebelah rekod yang sedang digunakan."""
    fail_lama = baru.parent / lama
    if fail_lama.exists() and not baru.exists():
        sys.exit(
            f"Jumpa {fail_lama} (nama lama) tetapi tiada {baru.name}.\n"
            f"Namakan semula fail itu kepada {baru.name} dahulu supaya kemas kini sedia ada dikekalkan."
        )
