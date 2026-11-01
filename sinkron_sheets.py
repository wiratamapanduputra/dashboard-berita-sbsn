"""Sinkronisasi data berita SBSN ke Google Sheets, sumber data dashboard online.

Isi Google Sheet (dibuat otomatis saat unggah pertama):
    Berita   : Tanggal | Situs | Judul | Link | Topik | Tone | Ritel | Diunggah
    Dihapus  : Link | Judul | Alasan   -> berita di sini disembunyikan dari dashboard
                                          dan tidak akan diunggah ulang
    Info     : waktu pembaruan terakhir dan jumlah berita

Koreksi data (tone, topik, seri) bisa langsung di tab "Berita". Untuk menghapus berita,
potong (cut) barisnya dari tab "Berita" lalu tempel ke tab "Dihapus".
Koreksi dan penghapusan lewat aplikasi admin (app.py) diteruskan otomatis ke Sheet; bila
sedang gagal (misalnya tidak ada internet), perubahan disimpan di antrean_sinkron.json dan
dikirim ulang oleh jadwal scraping atau unggah berikutnya.

Cara pakai dari terminal (venv aktif):
    python sinkron_sheets.py --cek                 # uji koneksi ke Google Sheet
    python sinkron_sheets.py --semua               # unggah semua file di folder hasil
    python sinkron_sheets.py hasil/berita_sbsn_2026-10.xlsx   # unggah file tertentu

Pengaturan (sekali saja, TIDAK tertimpa saat berkas program diperbarui):
  - pengaturan_online.ini di folder ini, berisi:
        [google_sheet]
        id_sheet = https://docs.google.com/spreadsheets/d/xxxx/edit
        id_sheet_usulan = https://docs.google.com/spreadsheets/d/yyyy/edit
  - kunci_google.json (kunci service account penulis) di folder ini.

    python sinkron_sheets.py --samakan   # samakan Sheet dengan data lokal (data lokal = sumber utama)
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
# Isi ID/URL Sheet di pengaturan_online.ini (disarankan). ID_SHEET di bawah hanya cadangan
# untuk versi lama; nilainya ikut hilang bila berkas ini ditimpa versi baru.
ID_SHEET = ""
FILE_PENGATURAN = FOLDER / "pengaturan_online.ini"
# Kunci service account PENULIS (akses Editor ke Sheet). Jangan dibagikan / diunggah ke GitHub.
FILE_KUNCI = FOLDER / "kunci_google.json"
# Antrean koreksi/penghapusan dari aplikasi admin yang belum terkirim ke Sheet
FILE_ANTREAN = FOLDER / "antrean_sinkron.json"

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


def baca_pengaturan(kunci):
    """Nilai dari pengaturan_online.ini, bagian [google_sheet]. Kosong bila tidak ada."""
    import configparser
    cp = configparser.RawConfigParser()
    try:
        cp.read(FILE_PENGATURAN, encoding="utf-8-sig")
        return cp.get("google_sheet", kunci, fallback="").strip().strip('"').strip("'")
    except Exception:
        return ""


def id_sheet_data():
    return os.environ.get("SBSN_ID_SHEET") or baca_pengaturan("id_sheet") or ID_SHEET


def status_online():
    """(aktif, alasan). alasan kosong bila dashboard online memang tidak dipakai di komputer ini."""
    ada_id, ada_kunci = bool(id_sheet_data()), FILE_KUNCI.exists()
    if ada_id and ada_kunci:
        return True, ""
    if not ada_id and not ada_kunci:
        return False, ""
    if not ada_id:
        return False, f"ID Sheet data belum diisi di {FILE_PENGATURAN.name} (baris id_sheet)."
    return False, f"File kunci {FILE_KUNCI.name} tidak ditemukan di folder scraper."


def aktif():
    """True bila dashboard online sudah diatur di komputer ini (ID Sheet dan kunci penulis ada)."""
    return status_online()[0]


def buka(id_sheet=None, info=None, baca_saja=False):
    """Buka Google Sheet. info = isi kunci service account (dict); kosong = pakai FILE_KUNCI."""
    import gspread

    id_sheet = id_dari_url(id_sheet or id_sheet_data())
    if not id_sheet:
        raise ValueError(f"ID Google Sheet belum diisi. Isi id_sheet di {FILE_PENGATURAN.name}.")
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
    """Baris df yang belum ada di Sheet dan tidak ada di tab Dihapus. Mengembalikan list baris
    siap tulis.

    Aturan "berita sama" dibuat SAMA PERSIS dengan aplikasi admin (sc.buang_duplikat), supaya
    jumlah berita di dashboard online sama dengan di komputer:
      - link sama (setelah dibakukan) = berita sama;
      - link pengalihan Google (news.google.com / google.com/goto) dianggap sama dengan berita
        lain hanya bila judulnya cocok (minimal 25 huruf), situsnya sama, dan tanggalnya
        berselisih paling lama 7 hari.
    Dulu cukup judul + situs sama, sehingga dua berita berbeda dengan judul pendek yang sama
    (misalnya dua halaman berjudul "BCA Sekuritas") hanya terunggah satu.
    """
    if df is None or df.empty:
        return []
    kolom = ["Tanggal", "Situs", "Judul", "Link", "Topik", "Tone", "Ritel"]

    def siapkan(d, tanda_lama):
        d = d.copy()
        for k in kolom:
            if k not in d.columns:
                d[k] = ""
        d = d[kolom].copy()
        for k in kolom[1:]:
            d[k] = d[k].map(_teks)
        d["Tanggal"] = (_tanggal(d["Tanggal"].astype(str).str.slice(0, 10)) if tanda_lama
                        else pd.to_datetime(d["Tanggal"], errors="coerce"))
        d["_lama"] = tanda_lama
        return d[d["Link"] != ""]

    baru_d = siapkan(df, False)
    lama_d = siapkan(lama, True) if len(lama) else baru_d.iloc[0:0]
    kunci_ada = set(lama_d["Link"].map(sc.kunci_link))
    if "Link" in dihapus:
        kunci_ada |= set(dihapus["Link"].map(sc.kunci_link))
    judul_hapus = set()
    if "Judul" in dihapus:
        judul_hapus = {sc.kunci_judul(j) for j in dihapus["Judul"]}
        judul_hapus = {j for j in judul_hapus if len(j) >= 25}

    gab = sc.buang_duplikat(pd.concat([lama_d, baru_d], ignore_index=True))
    pilih = gab[~gab["_lama"].astype(bool)]
    pilih = pilih[~pilih["Link"].map(sc.kunci_link).isin(kunci_ada)]
    # Versi link pengalihan dari berita yang sudah dihapus (judul panjang sama) tidak diunggah lagi
    sudah_dihapus = [bool(sc.adalah_pengalihan(l)) and sc.kunci_judul(j) in judul_hapus
                     for l, j in zip(pilih["Link"], pilih["Judul"])]
    pilih = pilih[[not x for x in sudah_dihapus]]
    waktu = sekarang_wib()
    return [ke_baris(r, waktu) for r in pilih.to_dict("records")]


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
# KOREKSI & PENGHAPUSAN DARI APLIKASI ADMIN
# ---------------------------------------------------------------------------

def _baca_antrean():
    try:
        return json.loads(FILE_ANTREAN.read_text(encoding="utf-8"))
    except Exception:
        return []


def _simpan_antrean(items):
    if items:
        FILE_ANTREAN.write_text(json.dumps(items, ensure_ascii=False, indent=1), encoding="utf-8")
    else:
        FILE_ANTREAN.unlink(missing_ok=True)


def jumlah_antrean():
    return len(_baca_antrean())


def catat_perubahan(ubah=(), hapus=()):
    """Masukkan koreksi (list dict berkolom Link, Judul, Situs, Tanggal, Topik, Tone, Ritel) dan
    penghapusan (list dict berkolom Link, Judul) ke antrean, lalu langsung coba kirim.
    Mengembalikan (berhasil, pesan)."""
    items = _baca_antrean()
    waktu = sekarang_wib()
    for r in ubah:
        baris = ke_baris(r, waktu)
        items.append({"aksi": "ubah", "waktu": waktu,
                      "data": dict(zip(KOLOM_SHEET[:7], baris[:7]))})
    for h in hapus:
        items.append({"aksi": "hapus", "waktu": waktu, "link": _teks(h.get("Link")),
                      "judul": _teks(h.get("Judul"))})
    _simpan_antrean(items)
    return kirim_antrean_aman()


def kirim_antrean_aman(log=lambda *_: None):
    """Seperti kirim_antrean, tetapi tidak melempar error. Mengembalikan (berhasil, pesan)."""
    try:
        n = kirim_antrean(log=log)
        return True, f"{n} perubahan diteruskan ke dashboard online."
    except Exception as e:
        return False, (f"Perubahan tersimpan di komputer, tetapi belum terkirim ke dashboard online "
                       f"({e}). Akan dicoba lagi otomatis pada jadwal berikutnya.")


def kirim_antrean(sh=None, log=print):
    """Terapkan antrean ke Sheet: koreksi memperbarui baris di tab Berita; penghapusan dicatat di
    tab Dihapus dan barisnya dibuang dari tab Berita. Mengembalikan jumlah perubahan terkirim."""
    items = _baca_antrean()
    if not items:
        return 0
    sh = sh or buka()
    ws = _tab(sh, TAB_BERITA, KOLOM_SHEET)
    wsd = _tab(sh, TAB_DIHAPUS, KOLOM_DIHAPUS)
    nilai = ws.get_all_values()
    header = [str(h).strip() for h in (nilai[0] if nilai else KOLOM_SHEET)]
    if header[:7] != KOLOM_SHEET[:7]:
        raise ValueError("Urutan kolom tab Berita berubah. Kembalikan judul kolom baris pertama "
                         f"menjadi: {', '.join(KOLOM_SHEET)}")
    baris_ke = {}   # kunci link -> nomor baris di Sheet (mulai 2)
    for i, r in enumerate(nilai[1:], start=2):
        if len(r) > 3 and str(r[3]).strip():
            baris_ke.setdefault(sc.kunci_link(r[3]), []).append(i)

    # Perubahan terakhir per berita yang berlaku; penghapusan mengalahkan koreksi
    ubah, hapus = {}, {}
    for it in items:
        if it.get("aksi") == "hapus":
            k = sc.kunci_link(it.get("link", ""))
            hapus[k] = it
            ubah.pop(k, None)
        elif it.get("aksi") == "ubah":
            k = sc.kunci_link(it["data"].get("Link", ""))
            if k not in hapus:
                ubah[k] = it["data"]

    pembaruan = []
    for k, d in ubah.items():
        for row in baris_ke.get(k, []):
            isi = [d.get(c, "") for c in KOLOM_SHEET[:7]]
            isi[3] = nilai[row - 1][3]   # link di Sheet tidak diubah
            pembaruan.append({"range": f"A{row}:G{row}", "values": [isi]})
    if pembaruan:
        ws.batch_update(pembaruan, value_input_option="RAW")

    tercatat = set()
    dh = baca_tab(wsd)
    if "Link" in dh:
        tercatat = set(dh["Link"].map(sc.kunci_link))
    catat = [[h["link"], h.get("judul", ""), f"Dihapus lewat aplikasi admin, {h.get('waktu', '')}"]
             for k, h in hapus.items() if k not in tercatat and h.get("link")]
    if catat:
        wsd.append_rows(catat, value_input_option="RAW")
    buang = sorted({row for k in hapus for row in baris_ke.get(k, [])}, reverse=True)
    if buang:   # dari bawah ke atas supaya nomor baris tidak bergeser
        sh.batch_update({"requests": [
            {"deleteDimension": {"range": {"sheetId": ws.id, "dimension": "ROWS",
                                           "startIndex": r - 1, "endIndex": r}}} for r in buang]})
    perbarui_info(sh, max(len(nilai) - 1 - len(buang), 0))
    _simpan_antrean(_baca_antrean()[len(items):])   # sisakan yang masuk saat pengiriman berjalan
    log(f"Google Sheet: {len(pembaruan)} baris dikoreksi, {len(hapus)} berita dihapus "
        "(dari aplikasi admin).")
    return len(items)


def samakan(df_lokal, hapus_links=(), sh=None, log=print):
    """Jadikan baris di tab Berita sama dengan data lokal (sumber utama) untuk berita yang ada
    di keduanya, dan hapus dari Sheet berita di hapus_links. Berguna untuk menyusulkan koreksi
    yang dulu tidak terkirim. Koreksi yang dibuat LANGSUNG di Sheet ikut tertimpa.
    Mengembalikan (jumlah koreksi, jumlah penghapusan)."""
    sh = sh or buka()
    ws = _tab(sh, TAB_BERITA, KOLOM_SHEET)
    nilai = ws.get_all_values()
    di_sheet = {}
    for r in nilai[1:]:
        if len(r) > 3 and str(r[3]).strip():
            di_sheet.setdefault(sc.kunci_link(r[3]), list(r) + [""] * 8)
    waktu = sekarang_wib()
    items, sudah, n_ubah, n_hapus = _baca_antrean(), set(), 0, 0
    for r in df_lokal.to_dict("records"):
        k = sc.kunci_link(_teks(r.get("Link")))
        if k in sudah or k not in di_sheet:
            continue
        sudah.add(k)
        baru = ke_baris(r, waktu)[:7]
        lama = [str(x).strip() for x in di_sheet[k][:7]]
        if any(baru[i] != lama[i] for i in (0, 1, 2, 4, 5, 6)):
            items.append({"aksi": "ubah", "waktu": waktu, "data": dict(zip(KOLOM_SHEET[:7], baru))})
            n_ubah += 1
    for link in hapus_links:
        if sc.kunci_link(link) in di_sheet:
            items.append({"aksi": "hapus", "waktu": waktu, "link": _teks(link), "judul": ""})
            n_hapus += 1
    _simpan_antrean(items)
    kirim_antrean(sh, log=log)
    return n_ubah, n_hapus


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
    ap.add_argument("--samakan", action="store_true",
                    help="Samakan isi Sheet dengan data lokal (menyusulkan koreksi yang belum terkirim)")
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
        print(f"Antrean perubahan dari aplikasi admin yang belum terkirim: {jumlah_antrean()}")
        return

    try:
        kirim_antrean(sh)
    except Exception as e:
        print(f"PERINGATAN: antrean koreksi/penghapusan belum terkirim: {e}")
    if a.samakan:
        lokal = baca_excel_hasil(sorted(sc.FOLDER_HASIL.glob("berita_sbsn_*.xlsx")))
        hapus = []
        try:
            import usulan
            if usulan.aktif():
                u = usulan.daftar_usulan(usulan.buka())
                hapus = list(u.loc[(u["Status"] == usulan.DISETUJUI)
                                   & u["_data"].map(lambda d: d["hapus"]), "Link"])
        except Exception as e:
            print(f"PERINGATAN: daftar usulan hapus tidak terbaca: {e}")
        n_u, n_h = samakan(lokal, hapus, sh)
        print(f"Selesai: {n_u} berita dikoreksi, {n_h} berita dihapus di Google Sheet.")
        return
    files = [Path(f) for f in a.file]
    if a.semua:
        files += sorted(sc.FOLDER_HASIL.glob("berita_sbsn_*.xlsx"))
    if not files:
        ap.error("Sebutkan file yang diunggah, atau pakai --semua / --samakan.")
    print(f"Membaca {len(files)} file…")
    try:
        unggah(baca_excel_hasil(files), sh)
    except Exception as e:
        print(f"GAGAL mengunggah: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
