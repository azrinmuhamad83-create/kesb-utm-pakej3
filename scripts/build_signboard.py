"""Jana signboard cetak A4 melintang (satu papan satu fail) untuk notice board pejabat tapak.

Guna:  py -3.12 scripts/build_signboard.py
Hasil: output/signboard-sampel.pdf
       output/signboard-dokumen-kontrak.pdf
"""
from pathlib import Path

from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

NAVY = HexColor("#1F3864")
AMBER = HexColor("#F2A900")

FONT_DIR = Path(r"C:\Windows\Fonts")
pdfmetrics.registerFont(TTFont("Arial", str(FONT_DIR / "arial.ttf")))
pdfmetrics.registerFont(TTFont("Arial-Bold", str(FONT_DIR / "arialbd.ttf")))

# Teks signboard (ubah di sini jika perlu)
BARIS_ATAS = "KARIM ENGINEERING SDN BHD"
PROJEK = ("UTM – Baikpulih Makmal P & P di Bangunan Akademik (RP1-RMKe-13)",
          "Pakej 3: Fakulti Komputeran (FKOM) & Fakulti Kejuruteraan Mekanikal (FKM)")

# (nama fail, baris tajuk, saiz maksimum tajuk dalam pt)
PAPAN = [
    ("signboard-sampel.pdf", ["SAMPEL"], 170),
    ("signboard-dokumen-kontrak.pdf", ["DOKUMEN", "KONTRAK"], 125),
]

OUT_DIR = Path(__file__).resolve().parent.parent / "output"


def teks_tengah(c, y, teks, font, saiz, warna):
    c.setFillColor(warna)
    c.setFont(font, saiz)
    c.drawCentredString(c._pagesize[0] / 2, y, teks)


def saiz_muat(teks, font, lebar_maks, saiz_maks):
    saiz = saiz_maks
    while pdfmetrics.stringWidth(teks, font, saiz) > lebar_maks:
        saiz -= 1
    return saiz


def bina(nama_fail, tajuk, saiz_maks):
    OUT_DIR.mkdir(exist_ok=True)
    out = OUT_DIR / nama_fail
    w, h = landscape(A4)
    c = canvas.Canvas(str(out), pagesize=(w, h))
    c.setTitle("Signboard – " + " ".join(tajuk).title())
    c.setAuthor("Karim Engineering Sdn Bhd")

    # Bingkai
    m = 9 * mm
    c.setStrokeColor(NAVY)
    c.setLineWidth(3 * mm)
    c.rect(m, m, w - 2 * m, h - 2 * m, fill=0, stroke=1)
    c.setStrokeColor(AMBER)
    c.setLineWidth(0.8 * mm)
    c.rect(m + 3.2 * mm, m + 3.2 * mm, w - 2 * (m + 3.2 * mm), h - 2 * (m + 3.2 * mm), fill=0, stroke=1)

    # Jalur atas: nama kontraktor
    jalur_h = 20 * mm
    jalur_y = h - m - 3.2 * mm - 0.4 * mm - jalur_h
    c.setFillColor(NAVY)
    c.rect(m + 3.6 * mm, jalur_y, w - 2 * (m + 3.6 * mm), jalur_h, fill=1, stroke=0)
    teks_tengah(c, jalur_y + jalur_h / 2 - 3.5 * mm, BARIS_ATAS, "Arial-Bold", 22, white)

    # Tajuk besar (saiz seragam ikut baris terpanjang)
    lebar_maks = w - 2 * 30 * mm
    saiz = min(saiz_muat(t, "Arial-Bold", lebar_maks, saiz_maks) for t in tajuk)
    jarak = saiz * 1.12

    # Tinggi blok kandungan (tajuk + garis + projek) dipusatkan dalam ruang di bawah jalur
    tinggi_tajuk = saiz * 0.72 + (len(tajuk) - 1) * jarak   # huruf besar baris 1 hingga garis dasar baris akhir
    tinggi_blok = tinggi_tajuk + 14 * mm + 19 * mm + 3.5 * mm
    ruang_atas = jalur_y
    ruang_bawah = m + 3.6 * mm
    blok_atas = ruang_bawah + (ruang_atas - ruang_bawah + tinggi_blok) / 2
    y = blok_atas - saiz * 0.72                              # garis dasar baris pertama
    for i, baris in enumerate(tajuk):
        teks_tengah(c, y - i * jarak, baris, "Arial-Bold", saiz, NAVY)

    # Garis pemisah
    gy = y - (len(tajuk) - 1) * jarak - 14 * mm
    c.setStrokeColor(AMBER)
    c.setLineWidth(1.4 * mm)
    c.line(w / 2 - 45 * mm, gy, w / 2 + 45 * mm, gy)

    # Nama projek
    teks_tengah(c, gy - 12 * mm, PROJEK[0], "Arial-Bold", 13, NAVY)
    teks_tengah(c, gy - 19 * mm, PROJEK[1], "Arial", 12, NAVY)

    c.showPage()
    c.save()
    print(out)


if __name__ == "__main__":
    for papan in PAPAN:
        bina(*papan)
