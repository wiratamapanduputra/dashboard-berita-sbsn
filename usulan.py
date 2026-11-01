"""Usulan perubahan berita dari dashboard online, menunggu persetujuan admin.

Alur:
  1. Pengguna dashboard online mengklik ✏️ pada kartu berita, mengisi nama dan usulannya
     (koreksi tone/topik/seri/judul/situs/tanggal, atau usulan hapus beserta alasan).
  2. Usulan dicatat di Google Sheet TERPISAH ("Usulan Berita SBSN", tab "Usulan").
     Data berita belum berubah sama sekali.
  3. Admin membuka menu 📥 Usulan di aplikasi admin, lalu menyetujui atau menolak.
     Usulan yang disetujui diterapkan ke file lokal (sumber data utama), lalu ikut
     diteruskan ke Google Sheet data dashboard.

Sheet usulan sengaja dipisah dari Sheet data: akun pembaca (dipakai dashboard online)
hanya boleh MENULIS di Sheet usulan dan tetap hanya bisa MEMBACA Sheet data.

Pengaturan di komputer admin: isi ID_SHEET_USULAN di bawah. Kunci penulis (kunci_google.json)
harus punya akses Editor ke Sheet usulan.
"""

import json
import os
import random
from datetime import datetime

import pandas as pd

import scraper_sbsn_chrome as sc
import sinkron_sheets as sk

# ID atau URL lengkap Google Sheet usulan (BUKAN Sheet data berita)
ID_SHEET_USULAN = "https://docs.google.com/spreadsheets/d/1X_B1QndkCk7QCFD5RrR1Df6XT7RTxERPxauLn-M6TH4/edit?usp=sharing"

TAB_USULAN = "Usulan"
KOLOM_USULAN = ["ID", "Diajukan", "Pengusul", "Jenis", "Judul saat ini", "Link", "Ringkasan usulan",
                "Alasan", "Status", "Diputuskan", "Oleh", "Catatan admin", "Data"]
KOLOM_EDIT = ["Judul", "Situs", "Tanggal", "Topik", "Tone", "Ritel"]
LABEL = {"Judul": "Judul", "Situs": "Situs", "Tanggal": "Tanggal", "Topik": "Kata kunci",
         "Tone": "Tone", "Ritel": "Seri ritel"}
MENUNGGU, DISETUJUI, DITOLAK = "Menunggu", "Disetujui", "Ditolak"
BATAS = {"nama": 60, "teks": 300, "alasan": 500}


def id_sheet_usulan():
    return os.environ.get("SBSN_ID_SHEET_USULAN") or ID_SHEET_USULAN


def aktif():
    """True bila fitur usulan sudah diatur di komputer admin."""
    return bool(id_sheet_usulan()) and sk.FILE_KUNCI.exists()


def buka(id_sheet=None, info=None):
    return sk.buka(id_sheet or id_sheet_usulan(), info=info, baca_saja=False)


def nilai_baku(kolom, v):
    """Bentuk baku nilai satu kolom, supaya perbandingan lama/baru konsisten."""
    if kolom == "Tanggal":
        t = pd.to_datetime(v, errors="coerce")
        return t.strftime("%Y-%m-%d") if pd.notna(t) else ""
    if kolom == "Ritel":
        return ", ".join(sc.daftar_seri(sk._teks(v)))
    return sk._teks(v)


def ringkas(perubahan, hapus):
    if hapus:
        return "Hapus berita"
    return "; ".join(f"{LABEL.get(k, k)}: {lama or '–'} → {baru or '–'}"
                     for k, (lama, baru) in perubahan.items())


def kirim_usulan(sh, pengusul, link, judul_saat_ini, perubahan=None, hapus=False, alasan=""):
    """Catat satu usulan. perubahan = {kolom: (nilai_lama, nilai_baru)}. Mengembalikan ID."""
    perubahan = {k: (nilai_baku(k, a), nilai_baku(k, b)) for k, (a, b) in (perubahan or {}).items()
                 if k in KOLOM_EDIT}
    perubahan = {k: (a, b[:BATAS["teks"]]) for k, (a, b) in perubahan.items() if a != b}
    pengusul = sk._teks(pengusul)[:BATAS["nama"]]
    alasan = sk._teks(alasan)[:BATAS["alasan"]]
    if len(pengusul) < 3:
        raise ValueError("Nama pengusul wajib diisi (minimal 3 huruf).")
    if hapus and not alasan:
        raise ValueError("Alasan wajib diisi untuk usulan hapus.")
    if not hapus and not perubahan:
        raise ValueError("Belum ada isian yang diubah.")
    ws = sk._tab(sh, TAB_USULAN, KOLOM_USULAN)
    uid = f"U{datetime.now():%y%m%d%H%M%S}{random.randint(100, 999)}"
    data = {"perubahan": {k: list(v) for k, v in perubahan.items()}, "hapus": bool(hapus)}
    ws.append_rows([[uid, sk.sekarang_wib(), pengusul, "Hapus" if hapus else "Ubah",
                     sk._teks(judul_saat_ini)[:BATAS["teks"]], sk._teks(link),
                     ringkas(perubahan, hapus), alasan, MENUNGGU, "", "", "",
                     json.dumps(data, ensure_ascii=False)]], value_input_option="RAW")
    return uid


def daftar_usulan(sh):
    """Semua usulan sebagai DataFrame (kolom KOLOM_USULAN + _data berisi dict)."""
    try:
        ws = sh.worksheet(TAB_USULAN)
    except Exception:
        return pd.DataFrame(columns=KOLOM_USULAN + ["_data"])
    df = sk.baca_tab(ws)
    for k in KOLOM_USULAN:
        if k not in df.columns:
            df[k] = ""
    df = df[df["ID"].str.strip() != ""].copy()

    def _data(s):
        try:
            d = json.loads(s)
            return {"perubahan": {k: tuple(v) for k, v in d.get("perubahan", {}).items()},
                    "hapus": bool(d.get("hapus"))}
        except Exception:
            return {"perubahan": {}, "hapus": False}
    df["_data"] = df["Data"].map(_data)
    df["Status"] = df["Status"].str.strip().replace("", MENUNGGU)
    return df.reset_index(drop=True)


def link_menunggu(df_usulan):
    """Kunci link berita yang punya usulan berstatus Menunggu."""
    d = df_usulan[df_usulan["Status"] == MENUNGGU]
    return set(d["Link"].map(sc.kunci_link))


def putuskan(sh, uid, status, oleh="", catatan=""):
    """Tandai usulan sebagai Disetujui / Ditolak."""
    ws = sh.worksheet(TAB_USULAN)
    nilai = ws.get_all_values()
    for i, r in enumerate(nilai[1:], start=2):
        if r and r[0] == uid:
            ws.update(range_name=f"I{i}", values=[[status, sk.sekarang_wib(), sk._teks(oleh)[:BATAS["nama"]],
                                                   sk._teks(catatan)[:BATAS["alasan"]]]],
                      value_input_option="RAW")
            return True
    raise KeyError(f"Usulan {uid} tidak ditemukan di Sheet usulan.")
