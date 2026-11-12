"""Jana Papan Pemuka Projek (HTML mudah alih) daripada rekod hidup projek.

Guna:  py -3.12 scripts/build_papan_pemuka.py
Hasil: output/Papan-Pemuka.html – satu fail, tanpa internet; buka dari OneDrive di telefon.

Sumber (baca sahaja): output/Jadual-Tindakan-PM.xlsx, output/Daftar-Kelulusan-Bahan.xlsx,
output/Daftar-Item-Terbuka.xlsx, output/log_isu.json. Jalankan semula selepas daftar dikemas kini.
Nilai wang (RM...) disembunyikan kerana fail ini boleh dibuka oleh staf tapak.
Warna & fon ikut sistem reka bentuk dokumen KESB (navy 1F4E78, Arial).
"""
import html
import json
import re
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

import openpyxl

MULA = date(2026, 8, 18)   # milik tapak (SST para 4)
SIAP = date(2028, 8, 14)   # tarikh siap kontrak (SST para 13)
FASA = [  # Seni bina ikut Master Programme REV00
    ("Fasa 1", date(2026, 9, 1), date(2026, 12, 1)),
    ("Fasa 2", date(2027, 1, 1), date(2027, 5, 14)),
    ("Fasa 3", date(2027, 7, 1), date(2027, 11, 25)),
    ("Fasa 4", date(2028, 1, 1), date(2028, 4, 29)),
]


def cari_root():
    for d in Path(__file__).resolve().parents:
        if (d / "output" / "log_isu.json").exists():
            return d
    raise SystemExit("Folder output/ dengan log_isu.json tidak dijumpai.")


def tarikh(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str) and (m := re.search(r"(\d{2})/(\d{2})/(\d{4})", v)):
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    return None


def bersih(t):
    """Teks selamat untuk HTML; sembunyikan nilai wang."""
    t = re.sub(r"RM\s?[\d,]+(?:\.\d+)?", "RM•••", str(t or ""))
    return html.escape(t)


def fmt(d):
    return d.strftime("%d/%m/%Y") if d else "–"


def baca(root, today):
    out = root / "output"
    # --- Tindakan PM
    wb = openpyxl.load_workbook(out / "Jadual-Tindakan-PM.xlsx", read_only=True, data_only=True)
    tindakan = []
    for r in wb["Tindakan"].iter_rows(min_row=5, values_only=True):
        if isinstance(r[0], (int, float)) and r[3]:
            st = str(r[7] or "Belum mula")
            selesai = st.lower().startswith(("ditutup", "selesai", "siap"))
            t = tarikh(r[1])
            tindakan.append(dict(no=int(r[0]), sasaran=t, keutamaan=r[2] or "", tindakan=r[3],
                                 siapa=r[5] or "", status=st, selesai=selesai,
                                 lewat=(not selesai and t is not None and t < today)))
    penting = []
    for r in wb["Tarikh Penting"].iter_rows(min_row=5, values_only=True):
        if r[0] and r[1]:
            penting.append((tarikh(r[0]), str(r[0]), r[1]))
    # --- Kelulusan bahan
    wb = openpyxl.load_workbook(out / "Daftar-Kelulusan-Bahan.xlsx", read_only=True, data_only=True)
    bahan = []
    for r in wb["Senarai Induk"].iter_rows(min_row=4, values_only=True):
        if isinstance(r[0], (int, float)) and r[2]:
            st = r[9] or "Belum dikemukakan"
            perlu = tarikh(r[5])
            lulus = str(st).startswith("Lulus") or st == "Tidak berkaitan"
            bahan.append(dict(trade=r[1] or "Lain-lain", item=r[2], status=st, perlu=perlu,
                              lulus=lulus, dikemuka=st != "Belum dikemukakan",
                              lewat=(not lulus and perlu is not None and perlu < today)))
    # --- Item terbuka (perlu disahkan)
    wb = openpyxl.load_workbook(out / "Daftar-Item-Terbuka.xlsx", read_only=True, data_only=True)
    item = [dict(kategori=r[1] or "", status=r[7] or "Terbuka")
            for r in wb["Item Terbuka"].iter_rows(min_row=5, values_only=True) if r[0]]
    # --- Isu pengguna
    isu = json.loads((out / "log_isu.json").read_text(encoding="utf-8"))["isu"]
    return tindakan, penting, bahan, item, isu


def cincin(pct, warna, saiz=120, tebal=12, label="", sub=""):
    r = (saiz - tebal) / 2
    c = 2 * 3.14159 * r
    return f"""<svg class="ring" viewBox="0 0 {saiz} {saiz}" width="{saiz}" height="{saiz}" aria-label="{label} {pct:.0f}%">
<circle cx="{saiz/2}" cy="{saiz/2}" r="{r}" fill="none" stroke="var(--trek)" stroke-width="{tebal}"/>
<circle class="arc" cx="{saiz/2}" cy="{saiz/2}" r="{r}" fill="none" stroke="{warna}" stroke-width="{tebal}" stroke-linecap="round"
 stroke-dasharray="{c:.1f}" stroke-dashoffset="{c:.1f}" data-off="{c * (1 - pct / 100):.1f}" transform="rotate(-90 {saiz/2} {saiz/2})"/>
<text x="50%" y="48%" text-anchor="middle" class="ring-n">{label}</text>
<text x="50%" y="66%" text-anchor="middle" class="ring-s">{sub}</text></svg>"""


def bina(root):
    today = date.today()
    tindakan, penting, bahan, item, isu = baca(root, today)

    hari_lalu = max(0, (today - MULA).days)
    jumlah_hari = (SIAP - MULA).days
    baki = max(0, (SIAP - today).days)
    pct_masa = min(100, 100 * hari_lalu / jumlah_hari)
    minggu = hari_lalu // 7 + 1

    seterusnya = sorted([p for p in penting if p[0] and p[0] >= today], key=lambda p: p[0])
    nxt = seterusnya[0] if seterusnya else None

    t_selesai = sum(t["selesai"] for t in tindakan)
    t_lewat = [t for t in tindakan if t["lewat"]]
    t_minggu = [t for t in tindakan if not t["selesai"] and not t["lewat"] and t["sasaran"]
                and 0 <= (t["sasaran"] - today).days <= 7]
    pct_tindakan = 100 * t_selesai / len(tindakan) if tindakan else 0

    b_lulus = sum(b["lulus"] for b in bahan)
    b_kemuka = sum(b["dikemuka"] and not b["lulus"] for b in bahan)
    b_lewat = sum(b["lewat"] for b in bahan)
    pct_bahan = 100 * b_lulus / len(bahan) if bahan else 0

    i_tutup = sum(not str(i["status"]).startswith("Terbuka") for i in item)
    isu_buka = [x for x in isu if str(x.get("status", "")).startswith("Terbuka")]

    # --- Fasa
    fasa_html = []
    for nama, a, b in FASA:
        if today > b:
            k, pct, nota = "siap", 100, "Siap"
        elif today >= a:
            pct = 100 * (today - a).days / max(1, (b - a).days)
            k, nota = "aktif", f"{(b - today).days} hari lagi"
        else:
            k, pct, nota = "akan", 0, f"Mula {fmt(a)}"
        fasa_html.append(f"""<div class="fasa {k}"><div class="fasa-h"><b>{nama}</b><span>{nota}</span></div>
<div class="bar"><i style="--w:{pct:.0f}%"></i></div><small>{fmt(a)} – {fmt(b)}</small></div>""")

    # --- Pencapaian (lencana)
    lencana = [
        ("🚀", "Projek bermula", "Milik tapak 18/08/2026", True),
        ("📘", "CP-QC Rev 5 siap", "Dihantar untuk semakan HQ", True),
        ("🦺", "Toolbox talk pertama", "28/10/2026", (root / "output/laporan-toolbox").exists()),
        ("✅", "Tindakan pertama ditutup", f"{t_selesai} tindakan selesai", t_selesai >= 1),
        ("📨", "RFA pertama dikemukakan", "Kelulusan bahan bermula", b_kemuka + b_lulus >= 1),
        ("🏅", "10 bahan diluluskan", f"{b_lulus}/10", b_lulus >= 10),
        ("🧹", "Tiada tindakan lewat", "Semua ikut jadual", not t_lewat),
        ("🏁", "Seni Bina Fasa 1 siap", "Sasaran 01/12/2026", today > FASA[0][2]),
    ]
    dapat = sum(l[3] for l in lencana)
    lencana_html = "".join(
        f'<div class="badge {"on" if on else "off"}"><span class="em">{e}</span><b>{bersih(n)}</b><small>{bersih(s)}</small></div>'
        for e, n, s, on in lencana)

    # --- Senarai tindakan
    def kad_tindakan(t):
        if t["selesai"]:
            chip = '<span class="chip ok">Selesai</span>'
        elif t["lewat"]:
            chip = f'<span class="chip bad">Lewat {(today - t["sasaran"]).days} hari</span>'
        else:
            d = (t["sasaran"] - today).days if t["sasaran"] else None
            chip = f'<span class="chip wait">{"Hari ini" if d == 0 else f"{d} hari lagi" if d is not None else "–"}</span>'
        kp = str(t["keutamaan"]).lower()
        return f"""<li class="kad k-{kp}" data-q="{bersih(t['tindakan']).lower()} {bersih(t['siapa']).lower()}" data-s="{'selesai' if t['selesai'] else 'lewat' if t['lewat'] else 'buka'}">
<div class="kad-h"><span class="no">#{t['no']}</span><span class="kp">{bersih(t['keutamaan'])}</span>{chip}</div>
<p>{bersih(t['tindakan'])}</p><small>🎯 {fmt(t['sasaran'])}{' · 👥 ' + bersih(t['siapa']) if t['siapa'] else ''}</small></li>"""

    susun = sorted(tindakan, key=lambda t: (t["selesai"], not t["lewat"], t["sasaran"] or date.max))
    tindakan_html = "".join(kad_tindakan(t) for t in susun)

    # --- Bahan ikut trade
    per_trade = defaultdict(Counter)
    for b in bahan:
        per_trade[b["trade"]]["jumlah"] += 1
        per_trade[b["trade"]]["lulus"] += b["lulus"]
        per_trade[b["trade"]]["kemuka"] += b["dikemuka"] and not b["lulus"]
        per_trade[b["trade"]]["lewat"] += b["lewat"]
    trade_html = "".join(
        f"""<li class="trade"><div class="kad-h"><b>{bersih(tr)}</b><span>{c['lulus']}/{c['jumlah']}</span></div>
<div class="bar stack"><i class="g" style="--w:{100*c['lulus']/c['jumlah']:.0f}%"></i><i class="y" style="--w:{100*c['kemuka']/c['jumlah']:.0f}%"></i></div>
{f'<small class="warn">⚠ {c["lewat"]} lepas tarikh perlu</small>' if c['lewat'] else ''}</li>"""
        for tr, c in sorted(per_trade.items(), key=lambda kv: (-kv[1]["lewat"], kv[0])))

    # --- Isu pengguna
    isu_html = "".join(
        f"""<li class="kad" data-s="{'buka' if str(x.get('status','')).startswith('Terbuka') else 'selesai'}">
<div class="kad-h"><span class="no">{bersih(x['no'])}</span><span class="kp">{bersih(x.get('jenis',''))}</span>
<span class="chip {'bad' if str(x.get('status','')).startswith('Terbuka') else 'ok'}">{bersih(x.get('status',''))}</span></div>
<p>{bersih(x.get('isu',''))}</p><small>📍 {bersih(x.get('lokasi',''))} · 🕒 {bersih(x.get('tarikh_masa',''))}</small></li>"""
        for x in reversed(isu)) or "<li class='kad'><p>Tiada isu direkod. 🎉</p></li>"

    # --- Item terbuka ikut kategori
    kat = Counter(i["kategori"] for i in item if str(i["status"]).startswith("Terbuka"))
    kat_html = "".join(f"<li><span>{bersih(k)}</span><b>{n}</b></li>" for k, n in kat.most_common())

    # --- Mesej motivasi ikut keadaan
    if t_lewat:
        sorakan = f"{len(t_lewat)} tindakan lewat. Selesaikan satu hari ini dan papan ini jadi lebih hijau 💪"
    elif t_minggu:
        sorakan = f"Tiada yang lewat. {len(t_minggu)} tindakan menunggu minggu ini. Teruskan! 🔥"
    else:
        sorakan = "Semua tindakan ikut jadual. Kerja hebat, pasukan! 🏆"

    nxt_html = (f"""<div class="next"><small>PERISTIWA SETERUSNYA</small><b class="big">{(nxt[0]-today).days}</b><span>hari lagi</span>
<p>{bersih(nxt[2])}</p><small>{bersih(nxt[1])}</small></div>""" if nxt else "")

    dijana = datetime.now().strftime("%d/%m/%Y %H:%M")
    k_lewat = "bad" if t_lewat else "ok"
    k_isu = "bad" if isu_buka else "ok"
    pill_lewat = f'<b class="pill">{len(t_lewat)}</b>' if t_lewat else ""
    pill_isu = f'<b class="pill">{len(isu_buka)}</b>' if isu_buka else ""
    return TEMPLAT.format(**locals(), ring_masa=cincin(pct_masa, "var(--navy-l)", 132, 13, f"{pct_masa:.0f}%", "masa kontrak"),
                          ring_tindakan=cincin(pct_tindakan, "var(--hijau)", 92, 10, f"{t_selesai}/{len(tindakan)}", "tindakan"),
                          ring_bahan=cincin(pct_bahan, "var(--kuning)", 92, 10, f"{b_lulus}/{len(bahan)}", "bahan lulus"),
                          n_tindakan=len(tindakan), n_lewat=len(t_lewat), n_minggu=len(t_minggu),
                          n_bahan=len(bahan), n_item=len(item) - i_tutup, n_isu=len(isu_buka),
                          fasa="".join(fasa_html), n_lencana=len(lencana))


TEMPLAT = """<!doctype html>
<html lang="ms"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#1F4E78">
<title>Papan Pemuka – UTM Pakej 3 (KESB)</title>
<style>
:root{{--navy:#1F4E78;--navy-l:#2E75B6;--pucat:#DDEBF7;--hijau:#3A9D5D;--kuning:#E0A800;--merah:#C00000;
--bg:#EEF2F6;--kad:#fff;--teks:#1b1f24;--lembut:#5b6470;--trek:#E3E8EE;--garis:#dde3ea}}
@media (prefers-color-scheme:dark){{:root{{--bg:#0E1620;--kad:#16212D;--teks:#E8EEF4;--lembut:#9AA7B4;--trek:#24313F;--garis:#24313F;--pucat:#1D2F42}}}}
*{{box-sizing:border-box;-webkit-tap-highlight-color:transparent}}
body{{margin:0;font-family:Arial,Helvetica,sans-serif;background:var(--bg);color:var(--teks);padding-bottom:84px}}
header{{background:linear-gradient(135deg,#1F4E78,#163A5A 60%,#2E75B6);color:#fff;padding:18px 18px 64px;border-radius:0 0 26px 26px}}
header small{{opacity:.8;letter-spacing:.08em;font-size:11px}}
header h1{{margin:4px 0 2px;font-size:20px}}
header p{{margin:0;opacity:.85;font-size:13px}}
.wrap{{max-width:560px;margin:-50px auto 0;padding:0 14px}}
.card{{background:var(--kad);border-radius:18px;padding:16px;margin-bottom:14px;box-shadow:0 4px 16px rgba(16,40,70,.08)}}
.hero{{display:flex;align-items:center;gap:14px}}
.hero .info b{{font-size:30px;display:block;line-height:1}}
.hero .info span{{color:var(--lembut);font-size:13px}}
.ring .arc{{transition:stroke-dashoffset 1.4s cubic-bezier(.2,.8,.2,1)}}
.ring-n{{font-size:20px;font-weight:bold;fill:var(--teks)}} .ring-s{{font-size:10px;fill:var(--lembut)}}
.rings{{display:flex;justify-content:space-around}}
.cheer{{background:var(--pucat);border-radius:14px;padding:12px 14px;font-size:14px;line-height:1.4}}
.next{{text-align:center;padding:6px 0}}
.next small{{color:var(--lembut);font-size:11px;letter-spacing:.08em}}
.next .big{{display:block;font-size:56px;color:var(--navy-l);line-height:1.05}}
.next p{{margin:6px 0 2px;font-weight:bold}}
.tiles{{display:grid;grid-template-columns:repeat(2,1fr);gap:10px}}
.tile{{background:var(--kad);border-radius:16px;padding:14px;box-shadow:0 4px 16px rgba(16,40,70,.08);text-decoration:none;color:inherit}}
.tile b{{font-size:28px;display:block}} .tile span{{font-size:12px;color:var(--lembut)}}
.tile.bad b{{color:var(--merah)}} .tile.ok b{{color:var(--hijau)}}
h2{{font-size:15px;margin:20px 4px 10px;display:flex;justify-content:space-between;align-items:baseline}}
h2 small{{font-weight:normal;color:var(--lembut);font-size:12px}}
.fasa{{margin-bottom:12px}} .fasa-h{{display:flex;justify-content:space-between;font-size:14px}}
.fasa-h span{{font-size:12px;color:var(--lembut)}} .fasa small{{font-size:11px;color:var(--lembut)}}
.fasa.aktif .fasa-h b::after{{content:" ● sedang berjalan";color:var(--hijau);font-size:11px;font-weight:normal}}
.bar{{height:10px;background:var(--trek);border-radius:6px;overflow:hidden;margin:6px 0 3px;display:flex}}
.bar i{{display:block;height:100%;width:0;background:var(--navy-l);border-radius:6px;transition:width 1.2s ease}}
.bar.stack i{{border-radius:0}} .bar i.g{{background:var(--hijau)}} .bar i.y{{background:var(--kuning)}}
.go .bar i{{width:var(--w)}}
.fasa.siap .bar i{{background:var(--hijau)}}
.badges{{display:grid;grid-template-columns:repeat(4,1fr);gap:8px}}
.badge{{text-align:center;padding:10px 4px;border-radius:14px;background:var(--pucat);font-size:11px}}
.badge .em{{font-size:26px;display:block}} .badge b{{display:block;margin:4px 0 2px;font-size:11px}}
.badge small{{color:var(--lembut);font-size:10px}}
.badge.off{{filter:grayscale(1);opacity:.45}}
.badge.on{{animation:pop .6s ease both}}
@keyframes pop{{0%{{transform:scale(.6);opacity:0}}70%{{transform:scale(1.08)}}100%{{transform:scale(1);opacity:1}}}}
ul{{list-style:none;margin:0;padding:0}}
.kad{{background:var(--kad);border-radius:14px;padding:12px 14px;margin-bottom:8px;border-left:5px solid var(--navy-l);box-shadow:0 2px 8px rgba(16,40,70,.06)}}
.kad.k-kritikal{{border-left-color:var(--merah)}} .kad.k-tinggi{{border-left-color:var(--kuning)}}
.kad[data-s=selesai]{{opacity:.6;border-left-color:var(--hijau)}}
.kad-h{{display:flex;gap:8px;align-items:center;font-size:12px}}
.kad-h .no{{font-weight:bold;color:var(--navy-l)}} .kad-h .kp{{color:var(--lembut);flex:1}}
.kad p{{margin:6px 0 4px;font-size:14px;line-height:1.4}} .kad small{{font-size:12px;color:var(--lembut)}}
.chip{{font-size:11px;font-weight:bold;padding:3px 8px;border-radius:20px;white-space:nowrap}}
.chip.ok{{background:#E2EFDA;color:#2D6B3F}} .chip.bad{{background:#F8CBAD;color:#8A1F11}} .chip.wait{{background:#FFF2CC;color:#7A5B00}}
.trade{{background:var(--kad);border-radius:14px;padding:12px 14px;margin-bottom:8px}}
.trade .kad-h{{justify-content:space-between;font-size:13px}} .warn{{color:var(--merah);font-size:12px}}
.legend{{display:flex;gap:14px;font-size:12px;color:var(--lembut);margin:0 4px 10px}}
.legend i{{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:4px;vertical-align:-1px}}
.kat li{{display:flex;justify-content:space-between;padding:9px 2px;border-bottom:1px solid var(--garis);font-size:14px}}
.filters{{display:flex;gap:6px;overflow-x:auto;margin:0 0 10px;padding-bottom:2px}}
.filters button{{border:1px solid var(--garis);background:var(--kad);color:var(--teks);border-radius:20px;padding:7px 13px;font-size:13px;white-space:nowrap}}
.filters button.on{{background:var(--navy);color:#fff;border-color:var(--navy)}}
input[type=search]{{width:100%;padding:11px 14px;border-radius:12px;border:1px solid var(--garis);background:var(--kad);color:var(--teks);font-size:15px;margin-bottom:8px}}
section.tab{{display:none}} section.tab:not(#t-utama){{padding-top:58px}} section.tab.on{{display:block;animation:fade .3s}}
@keyframes fade{{from{{opacity:0;transform:translateY(6px)}}to{{opacity:1}}}}
nav{{position:fixed;bottom:0;left:0;right:0;background:var(--kad);display:flex;border-top:1px solid var(--garis);padding-bottom:env(safe-area-inset-bottom);z-index:5}}
nav button{{flex:1;border:0;background:none;color:var(--lembut);padding:9px 0 8px;font-size:11px}}
nav button span{{display:block;font-size:21px}}
nav button.on{{color:var(--navy-l);font-weight:bold}}
.pill{{background:var(--merah);color:#fff;border-radius:10px;padding:0 6px;font-size:10px;margin-left:2px}}
footer{{text-align:center;font-size:11px;color:var(--lembut);padding:10px 20px 20px}}
</style></head>
<body>
<header>
  <small>KARIM ENGINEERING SDN BHD · PAPAN PEMUKA PROJEK</small>
  <h1>UTM Pakej 3 · FKOM &amp; FKM</h1>
  <p>Minggu {minggu} daripada 104 · {today:%d/%m/%Y}</p>
</header>
<div class="wrap">

<section class="tab on" id="t-utama">
  <div class="card hero">{ring_masa}
    <div class="info"><b>{baki}</b><span>hari ke tarikh siap kontrak<br>14/08/2028</span></div></div>
  <div class="card">{nxt_html}</div>
  <div class="cheer">{sorakan}</div>
  <h2>Hari ini</h2>
  <div class="tiles">
    <a class="tile {k_lewat}" href="#" data-go="t-tindakan" data-f="lewat"><b>{n_lewat}</b><span>tindakan lewat</span></a>
    <a class="tile" href="#" data-go="t-tindakan" data-f="buka"><b>{n_minggu}</b><span>perlu siap 7 hari ini</span></a>
    <a class="tile {k_isu}" href="#" data-go="t-isu"><b>{n_isu}</b><span>isu pengguna terbuka</span></a>
    <a class="tile" href="#" data-go="t-bahan"><b>{b_lewat}</b><span>bahan lepas tarikh perlu</span></a>
  </div>
  <h2>Kemajuan</h2>
  <div class="card rings">{ring_tindakan}{ring_bahan}</div>
  <h2>Fasa seni bina <small>Master Programme REV00</small></h2>
  <div class="card">{fasa}</div>
  <h2>Pencapaian pasukan <small>{dapat}/{n_lencana} dibuka</small></h2>
  <div class="card badges">{lencana_html}</div>
</section>

<section class="tab" id="t-tindakan">
  <h2>Tindakan PM <small>{t_selesai}/{n_tindakan} selesai</small></h2>
  <input type="search" placeholder="Cari tindakan atau nama…" data-cari="l-tindakan">
  <div class="filters" data-senarai="l-tindakan">
    <button class="on" data-f="semua">Semua</button><button data-f="lewat">Lewat ({n_lewat})</button>
    <button data-f="buka">Belum siap</button><button data-f="selesai">Selesai</button></div>
  <ul id="l-tindakan">{tindakan_html}</ul>
</section>

<section class="tab" id="t-bahan">
  <h2>Kelulusan bahan <small>{b_lulus} lulus · {b_kemuka} dikemukakan · {n_bahan} item</small></h2>
  <div class="legend"><span><i style="background:var(--hijau)"></i>Lulus</span><span><i style="background:var(--kuning)"></i>Dikemukakan</span><span><i style="background:var(--trek)"></i>Belum</span></div>
  <ul class="go">{trade_html}</ul>
  <h2>Perkara perlu disahkan <small>{n_item} terbuka</small></h2>
  <div class="card"><ul class="kat">{kat_html}</ul></div>
</section>

<section class="tab" id="t-isu">
  <h2>Isu pengguna <small>{n_isu} terbuka</small></h2>
  <ul>{isu_html}</ul>
</section>

<footer>Dijana {dijana} daripada Jadual Tindakan PM, Daftar Kelulusan Bahan, Daftar Item Terbuka &amp; Log Isu.<br>
Paparan baca sahaja – kemas kini dibuat dalam daftar asal. Nilai wang disembunyikan.</footer>
</div>

<nav>
  <button class="on" data-go="t-utama"><span>🏠</span>Utama</button>
  <button data-go="t-tindakan"><span>✅</span>Tindakan{pill_lewat}</button>
  <button data-go="t-bahan"><span>🧱</span>Bahan</button>
  <button data-go="t-isu"><span>💬</span>Isu{pill_isu}</button>
</nav>

<script>
function animate(root){{
  root.querySelectorAll('.arc').forEach(a=>requestAnimationFrame(()=>a.style.strokeDashoffset=a.dataset.off));
  root.querySelectorAll('.card,.go').forEach(c=>c.classList.add('go'));
}}
function go(id,f){{
  document.querySelectorAll('section.tab').forEach(s=>s.classList.toggle('on',s.id===id));
  document.querySelectorAll('nav button').forEach(b=>b.classList.toggle('on',b.dataset.go===id));
  if(f){{const b=document.querySelector('#'+id+' .filters [data-f="'+f+'"]'); if(b) b.click();}}
  window.scrollTo({{top:0,behavior:'smooth'}}); setTimeout(()=>animate(document.getElementById(id)),50);
}}
document.querySelectorAll('[data-go]').forEach(el=>el.addEventListener('click',e=>{{e.preventDefault();go(el.dataset.go,el.dataset.f);}}));
document.querySelectorAll('.filters').forEach(fl=>fl.addEventListener('click',e=>{{
  const b=e.target.closest('button'); if(!b) return;
  fl.querySelectorAll('button').forEach(x=>x.classList.toggle('on',x===b));
  document.querySelectorAll('#'+fl.dataset.senarai+' > li').forEach(li=>{{
    const f=b.dataset.f; li.style.display=(f==='semua'||li.dataset.s===f)?'':'none';}});
}}));
document.querySelectorAll('[data-cari]').forEach(inp=>inp.addEventListener('input',()=>{{
  const q=inp.value.toLowerCase();
  document.querySelectorAll('#'+inp.dataset.cari+' > li').forEach(li=>li.style.display=li.dataset.q.includes(q)?'':'none');
}}));
window.addEventListener('load',()=>animate(document));
</script>
</body></html>"""


if __name__ == "__main__":
    root = cari_root()
    out = root / "output" / "Papan-Pemuka.html"
    out.write_text(bina(root), encoding="utf-8")
    print(f"Siap: {out}")
