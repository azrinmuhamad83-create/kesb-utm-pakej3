"""Jana draf surat makluman prosedur Permintaan Maklumat (RFI) kepada S.O.

Guna:  py -3.12 scripts/build_surat_rfi.py
Hasil: output/surat/Surat-SO-Prosedur-RFI.md/.docx/.pdf
Prosedur selari dengan CP-QC Seksyen 5.1 dan Lampiran K Carta K9. Surat ialah DRAF – dikeluarkan hanya selepas diluluskan PM.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import build_pqp_docx  # noqa: E402
from build_surat_whatsapp import OUT, PROJEK, to_pdf  # noqa: E402


def surat():
    return f"""# Draf Surat – Makluman Prosedur Permintaan Maklumat (RFI)

**[KEPALA SURAT KARIM ENGINEERING SDN BHD]**

Ruj. Kami : KESB/UTM-P3/SURAT/[no.]
Ruj. Tuan : UTM.J.17.03/16.11/667(4)
Tarikh : [tarikh]

Pengarah
Bahagian Pengurusan Aset
Universiti Teknologi Malaysia
81310 Johor Bahru, Johor

Tuan,

**{PROJEK.upper()}**

**NO. KONTRAK: TN/2026-21/D-1/SP/BPA**

**MAKLUMAN PROSEDUR PERMINTAAN MAKLUMAT (REQUEST FOR INFORMATION – RFI)**

Dengan hormatnya perkara di atas adalah dirujuk.

**2.** Dokumen kontrak menetapkan bahawa:
   a) sebarang percanggahan dalam dokumen kontrak hendaklah dirujuk kepada S.O. secepat mungkin, dan S.O. akan memutuskan yang mana perlu diikuti (BQ A Preamble 15.3);
   b) penyerahan oleh Kontraktor dibuat secara bertulis kepada S.O. atau Wakil S.O. (BQ A Preliminaries 3.0 (ii)); dan
   c) sebarang arahan, penjelasan atau keterangan S.O. kepada Site Agent dianggap telah diberikan kepada Kontraktor (BQ A Preliminaries 36.0, Klausa 59).

Walau bagaimanapun, dokumen kontrak tidak menetapkan borang, penomboran atau tempoh jawapan bagi pertanyaan tersebut. Semakan awal KESB terhadap lukisan kontrak dan BQ juga mendapati beberapa perkara yang memerlukan penjelasan atau keputusan S.O. sebelum kerja berkaitan dapat dilaksanakan.

**3.** Sehubungan itu, dimaklumkan bahawa **mulai surat ini, semua pertanyaan teknikal KESB** – termasuk percanggahan atau maklumat yang tidak lengkap dalam lukisan, BQ dan spesifikasi – **akan dikemukakan melalui saluran Permintaan Maklumat (RFI)** sahaja. Prosedur RFI yang akan digunakan sepanjang tempoh kontrak, selaras dengan Pelan Kualiti Pembinaan (CP-QC) projek, adalah seperti berikut:
   a) **Borang:** setiap RFI dikemukakan menggunakan borang **KESB-QF-16 Permintaan Maklumat (RFI)** (contoh di Lampiran A), satu perkara bagi setiap RFI;
   b) **Penomboran:** RFI bernombor berturutan mengikut format **KESB/UTM-P3/RFI/xxx**;
   c) **Kandungan:** perkara, rujukan lukisan / BQ / spesifikasi, huraian percanggahan atau maklumat yang tidak lengkap, cadangan kontraktor, tarikh jawapan diperlukan dan lampiran petikan lukisan berkaitan;
   d) **Penghantaran:** dikemukakan oleh Pengurus Projek kepada S.O. bersama Transmittal Dokumen (KESB-QF-15), secara salinan keras dan e-mel kepada pegawai yang dinamakan oleh S.O. (perenggan 4(a));
   e) **Makluman segera (WhatsApp):** untuk mempercepat proses, sebaik RFI dikemukakan KESB akan membuat makluman (blast) dalam kumpulan WhatsApp utama projek, menyatakan no. RFI, perkara, lokasi dan tarikh jawapan diperlukan, beserta salinan RFI dan lampirannya (PDF). Makluman ini tidak menggantikan penghantaran rasmi dalam perenggan 3(d), dan jawapan S.O. tetap dibuat secara bertulis seperti perenggan 3(f);
   f) **Jawapan:** S.O. dipohon menjawab dalam ruang "Jawapan S.O." pada borang yang sama dengan tandatangan dan tarikh, atau melalui surat, sebelum tarikh jawapan diperlukan yang dinyatakan dalam RFI. Tempoh yang dicadangkan ialah **tujuh (7) hari bekerja** dari tarikh penerimaan, kecuali bagi perkara yang lebih mendesak;
   g) **Arahan perubahan:** jawapan RFI yang melibatkan perubahan skop, kuantiti atau spesifikasi dipohon dikeluarkan sebagai Arahan S.O. bertulis selaras dengan peruntukan Kontrak. Jawapan RFI semata-mata tidak dianggap sebagai arahan perubahan kerja;
   h) **Kerja berkaitan:** kerja yang terlibat secara langsung dengan sesuatu RFI tidak dimulakan sehingga jawapan S.O. diterima. Kerja lain diteruskan mengikut program;
   i) **Daftar RFI:** KESB menyelenggara Daftar RFI (no., tarikh hantar, perkara, tarikh jawapan diperlukan, tarikh jawapan diterima dan status) dan membentangkannya dalam setiap mesyuarat tapak; dan
   j) **Kawalan dokumen:** jawapan S.O., arahan dan revisi lukisan yang diterima didaftarkan, diedarkan kepada pasukan tapak dan subkontraktor, dan revisi lama dicop "BATAL".

Aliran prosedur diringkaskan di **Lampiran B**.

**4.** Bagi melaksanakan prosedur ini, KESB dengan hormatnya memohon jasa baik pihak tuan untuk:
   a) menamakan pegawai yang akan menerima dan menjawab RFI bagi pihak S.O., serta alamat e-mel untuk penghantaran;
   b) mengesahkan tempoh jawapan yang dicadangkan dalam perenggan 3(f), atau memaklumkan tempoh lain yang sesuai; dan
   c) memaklumkan sekiranya UTM mempunyai borang atau prosedur RFI tersendiri yang perlu digunakan.

**5.** Untuk makluman tuan, RFI yang berikutnya akan dikemukakan mengikut prosedur ini bermula dengan **KESB/UTM-P3/RFI/001**. Untuk rekod, sebarang kelewatan jawapan RFI yang menjejaskan program kerja akan direkodkan, dan KESB berhak mengemukakan tuntutan lanjutan masa dan/atau kos tertakluk kepada peruntukan Kontrak.

**6.** Kerjasama dan perhatian pihak tuan dalam perkara ini amatlah dihargai dan didahului dengan ucapan terima kasih.

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

## Lampiran A: Borang KESB-QF-16 Permintaan Maklumat (RFI)

Contoh borang dilampirkan berasingan (fail "KESB-QF-16 Permintaan Maklumat (RFI)").

| Bahagian borang | Diisi oleh |
|---|---|
| No. RFI, tarikh, perkara, rujukan, tarikh jawapan diperlukan | KESB |
| Soalan / percanggahan | KESB |
| Cadangan kontraktor | KESB |
| Jawapan S.O. | S.O. / wakil S.O. |
| Tandatangan Pengurus Projek (KESB) | KESB |
| Tandatangan S.O. / wakil, nama & tarikh | S.O. / wakil S.O. |

## Lampiran B: Aliran Prosedur RFI

| Langkah | Tindakan | Oleh | Rekod |
|---|---|---|---|
| 1 | Percanggahan atau maklumat tidak lengkap dikenal pasti dalam lukisan, BQ atau spesifikasi | Pasukan tapak KESB | Daftar isu |
| 2 | RFI disediakan dengan rujukan lukisan / BQ dan cadangan kontraktor | Pengurus Projek | KESB-QF-16 |
| 3 | RFI dikemukakan kepada S.O.; kerja berkaitan tidak dimulakan | Pengurus Projek | KESB-QF-15 Transmittal; Daftar RFI |
| 3a | Makluman (blast) RFI beserta salinan PDF dalam kumpulan WhatsApp utama projek untuk mempercepat proses | Pengurus Projek / Kerani Tapak | Tangkap layar dalam fail RFI |
| 4 | S.O. menjawab dalam borang / surat; Arahan S.O. bertulis jika melibatkan perubahan | S.O. / wakil S.O. | KESB-QF-16 (Jawapan S.O.) |
| 5 | Jawapan, arahan atau revisi lukisan diterima dan didaftarkan | Kerani Tapak | Daftar Dokumen; Daftar RFI |
| 6 | Impak ke atas kerja, bahan dan program dinilai; dokumen dikemas kini dan diedarkan; revisi lama dicop "BATAL" | Pengurus Projek | Daftar Dokumen; KESB-QF-15 |
| 7 | Kerja diteruskan mengikut maklumat terkini; RFI ditutup dalam Daftar RFI | Site Supervisor / Pengurus Projek | Daftar RFI |
"""


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    md = OUT / "Surat-SO-Prosedur-RFI.md"
    md.write_text(surat(), encoding="utf-8")
    build_pqp_docx.main(md, md.with_suffix(".docx"))
    to_pdf([md.with_suffix(".docx")])
    print("Siap:", md.with_suffix(".pdf"))


if __name__ == "__main__":
    main()
