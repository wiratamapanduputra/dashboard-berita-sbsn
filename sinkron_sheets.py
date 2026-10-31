"""Sinkronisasi data berita SBSN ke Google Sheets, sumber data dashboard online.

Isi Google Sheet (dibuat otomatis saat unggah pertama):
    Berita   : Tanggal | Situs | Judul | Link | Topik | Tone | Ritel | Diunggah
    Dihapus  : Link | Judul | Alasan   -> berita di sini disembunyikan dari dashboard
                                          dan tidak akan diunggah ulang
    Info     : waktu pembaruan terakhir dan jumlah berita

Koreksi data (tone, topik, seri) bisa langsung di tab "Berita". Untuk menghapus berita,
potong (cut) barisnya dari tab "Berita" lalu tempel ke tab "Dihapus".

Cara pakai dari terminal (venv aktif):
    python sinkron_sheets.py --cek                 # uji koneksi ke Google Sheet
    python sinkron_sheets.py --semua               # unggah semua file di folder hasil
    python sinkron_sheets.py hasil/berita_sbsn_2026-10.xlsx   # unggah file tertentu

Pengaturan: isi ID_SHEET di bawah dan taruh file kunci service account (JSON)
di folder ini dengan nama kunci_google.json.
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd

import scraper_sbsn_chrome as sc

# ---------------------------------------------------------------------------
# PENGATURAN
# ---------------------------------------------------------------------------

FOLDER = Path(__file__).resolve().parent
# ID atau URL lengkap Google Sheet, misalnya
# "https://docs.google.com/spreadsheets/d/1AbC...xyz/edit"  atau cukup "1AbC...xyz"
ID_SHEET = "https://docs.google.com/spreadsheets/d/1djyewT1kP8yJ0gvzrAJcprSpqQIYoSyFS2v-Sj0b7Zk/edit?usp=sharing"
# Kunci service account PENULIS (akses Editor ke Sheet). Jangan dibagikan / diunggah ke GitHub.
FILE_KUNCI = FOLDER / "kunci_google.json"

TAB_BERITA, TAB_DIHAPUS, TAB_INFO = "Berita", "Dihapus", "Info"
KOLOM_SHEET = ["Tanggal", "Situs", "Judul", "Link", "Topik", "Tone", "Ritel", "Diunggah"]
KOLOM_DIHAPUS = ["Link", "Judul", "Alasan"]

SCOPE_TULIS = ["https://www.googleapis.com/auth/spreadsheets"]
SCOPE_BACA = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
WIB = timezone(timedelta(hours=7))


def sekarang_wib():
    return datetime.now(WIB).strftime("%Y-%m-%d %H:%M WIB")


# ---------------------------------------------------------------------------
# KONEKSI
# ---------------------------------------------------------------------------

def id_dari_url(teks):
    """Terima ID atau URL Google Sheet, kembalikan ID-nya."""
    teks = str(teks or "").strip()
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", teks)
    return m.group(1) if m else teks


def email_service_account(info=None):
    try:
        if info is None:
            info = json.loads(FILE_KUNCI.read_text(encoding="utf-8"))
        return info.get("client_email", "")
    except Exception:
        return ""


def buka(id_sheet=None, info=None, baca_saja=False):
    """Buka Google Sheet. info = isi kunci service account (dict); kosong = pakai FILE_KUNCI."""
    import gspread

    id_sheet = id_dari_url(id_sheet or os.environ.get("SBSN_ID_SHEET") or ID_SHEET)
    if not id_sheet:
        raise ValueError("ID Google Sheet belum diisi. Isi ID_SHEET di sinkron_sheets.py.")
    scopes = SCOPE_BACA if baca_saja else SCOPE_TULIS
    if info:
        gc = gspread.service_account_from_dict(dict(info), scopes=scopes)
    else:
        if not FILE_KUNCI.exists():
            raise FileNotFoundError(
                f"File kunci tidak ditemukan: {FILE_KUNCI.name}. Unduh kunci JSON service account "
                "dari Google Cloud dan simpan di folder ini dengan nama tersebut.")
        gc = gspread.service_account(filename=str(FILE_KUNCI), scopes=scopes)
    try:
        return gc.open_by_key(id_sheet)
    except gspread.exceptions.SpreadsheetNotFound:
        raise PermissionError(
            "Google Sheet tidak ditemukan atau belum dibagikan. Buka Sheet, klik Bagikan/Share, "
            f"lalu tambahkan email service account: {email_service_account(info) or '(lihat file kunci)'}")
    except gspread.exceptions.APIError as e:
        raise PermissionError(f"Google menolak akses ke Sheet: {e}. Pastikan Google Sheets API "
                              "sudah diaktifkan dan Sheet dibagikan ke email service account.")


def _tab(sh, nama, header):
    """Ambil tab; buat beserta judul kolom bila belum ada (hanya untuk penulis)."""
    import gspread
    try:
        ws = sh.worksheet(nama)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=nama, rows=1000, cols=max(len(header), 3))
    if not any(ws.row_values(1)):
        ws.update(range_name="A1", values=[header], value_input_option="RAW")
        try:
            ws.freeze(rows=1)
        except Exception:
            pass
    return ws


def _tab_ada(sh, nama):
    import gspread
    try:
        return sh.worksheet(nama)
    except gspread.exceptions.WorksheetNotFound:
        return None


def baca_tab(ws):
    """Isi tab sebagai DataFrame teks (baris pertama = judul kolom)."""
    if ws is None:
        return pd.DataFrame()
    nilai = ws.get_all_values()
    if not nilai:
        return pd.DataFrame()
    header = [str(h).strip() for h in nilai[0]]
    baris = [list(r) + [""] * (len(header) - len(r)) for r in nilai[1:]]
    df = pd.DataFrame([r[:len(header)] for r in baris], columns=header)
    df = df.loc[:, [h for h in df.columns if h]]
    return df.fillna("").astype(str)


# ---------------------------------------------------------------------------
# UNGGAH (dipakai komputer scraping)
# ---------------------------------------------------------------------------

def _rss(link):
    return "news.google.com" in str(link)


def _js(judul, situs):
    return (str(judul).strip().lower(), str(situs).strip().lower())


def _teks(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    return str(v).strip()


def ke_baris(r, waktu):
    t = pd.to_datetime(r.get("Tanggal"), errors="coerce")
    return [t.strftime("%Y-%m-%d") if pd.notna(t) else "",
            _teks(r.get("Situs")), _teks(r.get("Judul")), _teks(r.get("Link")),
            _teks(r.get("Topik")), _teks(r.get("Tone")),
            ", ".join(sc.daftar_seri(_teks(r.get("Ritel")))), waktu]


def pilih_baru(df, lama, dihapus):
    """Baris df yang belum ada di Sheet (link sama, atau versi RSS dari judul+situs yang sama)
    dan tidak ada di tab Dihapus. Mengembalikan list baris siap tulis."""
    kunci_ada = set(lama["Link"].map(sc.kunci_link)) if "Link" in lama else set()
    if "Link" in dihapus:
        kunci_ada |= set(dihapus["Link"].map(sc.kunci_link))
    js_ada = {}
    if {"Judul", "Situs", "Link"} <= set(lama.columns):
        for j, s, l in zip(lama["Judul"], lama["Situs"], lama["Link"]):
            js_ada[_js(j, s)] = js_ada.get(_js(j, s), False) or _rss(l)
    js_hapus = set()
    if {"Judul"} <= set(dihapus.columns):
        js_hapus = {str(j).strip().lower() for j in dihapus["Judul"] if str(j).strip()}

    waktu = sekarang_wib()
    baru = []
    for r in df.to_dict("records"):
        link = _teks(r.get("Link"))
        if not link:
            continue
        k = sc.kunci_link(link)
        if k in kunci_ada:
            continue
        js = _js(_teks(r.get("Judul")), _teks(r.get("Situs")))
        # Berita yang sama bisa datang dari halaman Google (link asli) dan RSS (link pengalihan)
        if js in js_ada and (_rss(link) or js_ada[js]):
            continue
        if _rss(link) and js[0] in js_hapus:
            continue
        kunci_ada.add(k)
        js_ada[js] = js_ada.get(js, False) or _rss(link)
        baru.append(ke_baris(r, waktu))
    return baru


def perbarui_info(sh, jumlah):
    ws = _tab(sh, TAB_INFO, ["Keterangan", "Nilai"])
    ws.update(range_name="A2", values=[
        ["Terakhir diperbarui", sekarang_wib()],
        ["Jumlah berita", str(jumlah)],
        ["Catatan", "Diisi otomatis oleh skrip. Jangan diubah."]], value_input_option="RAW")


def unggah(df, sh=None, log=print):
    """Tambahkan berita baru ke tab Berita. Baris yang sudah ada TIDAK ditimpa, jadi koreksi
    manual di Sheet tetap aman. Mengembalikan jumlah berita yang ditambahkan."""
    sh = sh or buka()
    ws = _tab(sh, TAB_BERITA, KOLOM_SHEET)
    wsd = _tab(sh, TAB_DIHAPUS, KOLOM_DIHAPUS)
    lama, dihapus = baca_tab(ws), baca_tab(wsd)
    baru = pilih_baru(df if df is not None else pd.DataFrame(), lama, dihapus)
    if baru:
        ws.append_rows(baru, value_input_option="RAW")
        try:  # urutkan dari tanggal terbaru (baris judul dibekukan, tidak ikut diurutkan)
            akhir = len(lama) + len(baru) + 1
            ws.sort((1, "des"), range=f"A2:{chr(64 + len(KOLOM_SHEET))}{akhir}")
        except Exception:
            pass
    perbarui_info(sh, len(lama) + len(baru))
    log(f"Google Sheet: {len(baru)} berita baru ditambahkan (total {len(lama) + len(baru)}).")
    return len(baru)


# ---------------------------------------------------------------------------
# BACA (dipakai dashboard online)
# ---------------------------------------------------------------------------

def _tanggal(s):
    s = s.fillna("").astype(str).str.strip()
    t = pd.to_datetime(s, format="%Y-%m-%d", errors="coerce")
    sisa = t.isna() & s.ne("")
    if sisa.any():  # tanggal yang diketik manual, misalnya 31/10/2026
        t[sisa] = pd.to_datetime(s[sisa], dayfirst=True, errors="coerce")
    return t


def rapikan_baca(df, dihapus=None):
    """Ubah isi tab Berita menjadi DataFrame dengan kolom sc.KOLOM (seperti file Excel hasil)."""
    df = df.copy() if df is not None else pd.DataFrame()
    for k in KOLOM_SHEET:
        if k not in df.columns:
            df[k] = ""
    df = df.fillna("").astype(str)
    df = df[(df["Link"].str.strip() != "") | (df["Judul"].str.strip() != "")]
    if dihapus is not None and len(dihapus) and "Link" in dihapus:
        buang = set(dihapus["Link"].map(sc.kunci_link))
        df = df[~df["Link"].map(sc.kunci_link).isin(buang)]
    if df.empty:
        return pd.DataFrame(columns=sc.KOLOM)
    df["Tanggal"] = _tanggal(df["Tanggal"])
    tone = df["Tone"].str.strip().str.capitalize()
    df["Tone"] = tone.where(tone.isin(["Positif", "Netral", "Negatif"]), "Netral")
    df["Topik"] = df["Topik"].str.strip()
    df["Ritel"] = df["Ritel"].map(lambda v: ", ".join(sc.daftar_seri(v)))
    df["Bulan"] = df["Tanggal"].apply(sc.nama_bulan)
    df = sc.buang_duplikat(df)
    df = df.sort_values("Tanggal", ascending=False, na_position="last").reset_index(drop=True)
    df["No"] = range(1, len(df) + 1)
    return df[sc.KOLOM]


def baca_untuk_dashboard(sh):
    """(DataFrame berita, dict info) dari Google Sheet."""
    berita = baca_tab(_tab_ada(sh, TAB_BERITA))
    dihapus = baca_tab(_tab_ada(sh, TAB_DIHAPUS))
    info = {}
    t = baca_tab(_tab_ada(sh, TAB_INFO))
    if len(t.columns) >= 2:
        info = dict(zip(t.iloc[:, 0], t.iloc[:, 1]))
    return rapikan_baca(berita, dihapus), info


# ---------------------------------------------------------------------------
# JALUR TERMINAL
# ---------------------------------------------------------------------------

def baca_excel_hasil(files):
    data = []
    for p in files:
        data.append(pd.read_excel(p))
    return pd.concat(data, ignore_index=True) if data else pd.DataFrame(columns=sc.KOLOM)


def main():
    ap = argparse.ArgumentParser(description="Unggah hasil scraping ke Google Sheet dashboard.")
    ap.add_argument("file", nargs="*", help="File Excel hasil yang diunggah")
    ap.add_argument("--semua", action="store_true", help="Unggah semua file di folder hasil")
    ap.add_argument("--cek", action="store_true", help="Hanya uji koneksi ke Google Sheet")
    a = ap.parse_args()

    try:
        sh = buka()
    except Exception as e:
        print(f"GAGAL membuka Google Sheet: {e}")
        sys.exit(1)
    print(f"Terhubung ke Google Sheet: {sh.title}")
    if a.cek:
        df, info = baca_untuk_dashboard(sh)
        print(f"Isi saat ini: {len(df)} berita. Terakhir diperbarui: "
              f"{info.get('Terakhir diperbarui', '-')}")
        return

    files = [Path(f) for f in a.file]
    if a.semua:
        files += sorted(sc.FOLDER_HASIL.glob("berita_sbsn_*.xlsx"))
    if not files:
        ap.error("Sebutkan file yang diunggah, atau pakai --semua.")
    print(f"Membaca {len(files)} file…")
    try:
        unggah(baca_excel_hasil(files), sh)
    except Exception as e:
        print(f"GAGAL mengunggah: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
