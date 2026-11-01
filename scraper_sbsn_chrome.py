"""Scraper tab "Berita" (News) Google Search untuk topik SBSN, memakai Chrome.

Versi 2: tone berita, susunan kolom baru, dan label seri SBSN Ritel.

Cara pakai (dari terminal, venv aktif):
    python scraper_sbsn_chrome.py                          # bulan berjalan
    python scraper_sbsn_chrome.py 2026-09                  # bulan tertentu
    python scraper_sbsn_chrome.py --dari 2026-10-01 --sampai 2026-10-20
    python scraper_sbsn_chrome.py --seri SR025             # + pencarian khusus seri
    python scraper_sbsn_chrome.py --seri SR025 ST015       # beberapa seri sekaligus
    python scraper_sbsn_chrome.py --seri SR025 ST015 --hanya-seri  # berita seri itu saja
    python scraper_sbsn_chrome.py --situs kontan.co.id bisnis.com   # + pencarian khusus situs
    python scraper_sbsn_chrome.py --situs kontan.co.id --hanya-situs  # hanya situs tersebut
    python scraper_sbsn_chrome.py --pemantauan             # + situs Lapis 2 di daftar_situs.csv

Situs tujuan bawaan diambil dari daftar_situs.csv (Lapis 1), bila file itu ada.

Atau lewat platform:  streamlit run app.py

Chrome akan terbuka sendiri. Jangan ditutup selama proses berjalan.
Jika Google menampilkan CAPTCHA, selesaikan di jendela Chrome itu;
skrip menunggu dan otomatis melanjutkan setelah verifikasi selesai.
"""

import argparse
import io
import random
import re
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import urlencode

import pandas as pd

# ---------------------------------------------------------------------------
# PENGATURAN
# ---------------------------------------------------------------------------

FOLDER = Path(__file__).resolve().parent
FOLDER_HASIL = FOLDER / "hasil"
PROFIL_CHROME = FOLDER / "chrome_profile"

# Kata kunci pencarian -> nilai kolom "Topik"
KATA_KUNCI = {
    "SBSN Ritel": "SBSN Ritel",
    "Lelang SBSN": "Lelang SBSN",
    "Proyek SBSN": "Proyek SBSN",
}

# Pencarian tambahan bila seri diisi ({seri} diganti, misalnya SR025)
KATA_KUNCI_SERI = ["Sukuk Ritel {seri}", "{seri} SBSN"]
# True  -> berita hasil pencarian khusus seri (misalnya "Sukuk Ritel SR025") otomatis diberi label
#          seri itu, walau kodenya tidak tertulis di judul/cuplikan.
# False -> label seri hanya diberikan bila kode seri benar-benar tertulis di judul/cuplikan.
LABEL_DARI_PENCARIAN = True

MAKS_HALAMAN = 10      # tiap halaman berisi sekitar 10 berita
MAKS_HALAMAN_SITUS = 5 # batas halaman untuk pencarian khusus situs tujuan

# Situs media tujuan (opsional). Kosong = semua situs yang muncul di Google Berita.
# Contoh: SITUS_TUJUAN = ["kontan.co.id", "cnbcindonesia.com", "bisnis.com"]
# Bila dibiarkan kosong dan file daftar_situs.csv ada, situs Lapis 1 di file itu yang dipakai.
SITUS_TUJUAN = []
# Situs pemantauan (Lapis 2): dicari berkelompok hanya bila diminta (pilihan di aplikasi,
# atau python jadwal_scraping.py --pemantauan). Kosong = diambil dari daftar_situs.csv.
SITUS_PEMANTAUAN = []
# Situs yang tidak ingin diambil beritanya. Contoh: ["contoh-agregator.com"]
# (Situs Lapis X di daftar_situs.csv juga dibuang dari hasil, tanpa perlu ditulis di sini.)
SITUS_DIKECUALIKAN = []
# Daftar situs berlapis (bisa diedit di Excel). Kolom wajib: Situs;Lapis
#   Lapis 1 = prioritas, dicari khusus setiap pengambilan
#   Lapis 2 = pemantauan, dicari khusus bila diminta (misalnya seminggu sekali)
#   Lapis 3 = rujukan saja, tidak dicari khusus (tetap tertangkap lewat pencarian umum)
#   Lapis X = dikecualikan (toko, broker, iklan), beritanya tidak disimpan
FILE_DAFTAR_SITUS = FOLDER / "daftar_situs.csv"
# Cara mencari situs tujuan:
#   "rss"    -> lewat RSS Google News, satu permintaan ringan per situs (DISARANKAN; tidak membuka
#               Chrome dan jauh lebih jarang diblokir)
#   "google" -> lewat halaman pencarian Google di Chrome, beberapa situs digabung (site:a OR site:b);
#               kueri seperti ini lebih mudah dianggap otomatis dan diblokir Google (error 403)
SUMBER_SITUS = "rss"
# Untuk SUMBER_SITUS = "google": jumlah situs per pencarian gabungan
SITUS_PER_KUERI = 6
# Untuk SUMBER_SITUS = "rss": jumlah situs per permintaan RSS (site:a OR site:b ...).
# 1 = satu situs per permintaan (cara lama). 8 membuat 49 situs cukup 7 permintaan per kata kunci.
SITUS_PER_RSS = 8
# Selain halaman Google, setiap kata kunci juga dicari lewat RSS Google News (satu permintaan
# ringan). Ini menambah cakupan dan memberi TANGGAL TERBIT PASTI untuk berita yang di halaman
# Google hanya tertulis "3 minggu lalu" / "1 bulan lalu".
RSS_UMUM = True
RSS_URL = "https://news.google.com/rss/search"
JEDA_RSS = (2, 4)      # jeda acak (detik) antar permintaan RSS

# Aturan duplikasi:
#   False -> berita dianggap sama HANYA jika link-nya sama. Link berbeda = berita berbeda,
#            semuanya ikut dihitung dalam rekapitulasi (walau judul & situsnya sama).
#   True  -> selain link sama, berita dengan judul DAN situs yang sama juga dianggap duplikat.
GABUNG_JUDUL_SAMA = False
TUNGGU_CAPTCHA = 600   # detik maksimum menunggu verifikasi Google

KOLOM = ["No", "Tanggal", "Situs", "Judul", "Link", "Topik", "Bulan", "Tone", "Ritel"]

NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
              "Agustus", "September", "Oktober", "November", "Desember"]

BULAN = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "mei": 5, "jun": 6,
         "jul": 7, "agu": 8, "agt": 8, "sep": 9, "okt": 10, "nov": 11, "des": 12}

# ---------------------------------------------------------------------------
# KAMUS TONE  (silakan tambah/kurangi sesuai kebutuhan)
# Frasa panjang dicek lebih dulu, jadi "bebas risiko gagal bayar" dihitung
# positif dan tidak ikut terhitung sebagai "gagal bayar" (negatif).
# ---------------------------------------------------------------------------

KATA_POSITIF = [
    "bebas risiko", "tanpa risiko", "dijamin negara", "dijamin pemerintah",
    "dijamin undang-undang", "melampaui target", "melebihi target", "lampaui target",
    "capai target", "mencapai target", "oversubscribed", "kelebihan permintaan",
    "laris", "diminati", "diburu", "antusias", "minat tinggi", "minat investor tinggi",
    "rekor", "tertinggi", "terbesar", "tembus", "meningkat", "peningkatan", "naik",
    "tumbuh", "pertumbuhan", "sukses", "berhasil", "menarik", "menguntungkan",
    "untung", "cuan", "kompetitif", "optimis", "optimistis", "positif",
    "mendukung", "dukungan", "bermanfaat", "manfaat", "rampung", "selesai dibangun",
    "diresmikan", "peresmian", "beroperasi", "resmi dibuka", "inklusif", "berkah",
    "stabil", "solid", "kuat", "apresiasi", "prestasi", "penghargaan",
]

KATA_NEGATIF = [
    "gagal bayar", "di bawah target", "tidak capai target", "tak capai target",
    "meleset", "kurang diminati", "sepi peminat", "sepi", "lesu", "melemah",
    "turun", "menurun", "penurunan", "anjlok", "merosot", "jeblok", "rugi",
    "kerugian", "gagal", "defisit", "utang menumpuk", "beban utang", "membebani",
    "kritik", "dikritik", "sorotan", "mangkrak", "terbengkalai", "tertunda",
    "molor", "terlambat", "korupsi", "penyimpangan", "penyelewengan", "bermasalah",
    "masalah", "khawatir", "kekhawatiran", "waspada", "ancaman", "negatif",
    "batal", "dibatalkan", "ditolak", "tekanan", "gejolak", "krisis", "tersangka",
]

# Pola kode seri SBSN Ritel: SR025, SR023T3, ST014, ST012T2, CWLS SW005, dll.
POLA_SERI = re.compile(r"\b(SR|ST|SW|SWR)[\s-]?(\d{3})(T\d{1,2})?\b", re.IGNORECASE)

# ---------------------------------------------------------------------------
# PENGOLAHAN TEKS
# ---------------------------------------------------------------------------


def _pola(daftar):
    urut = sorted(set(daftar), key=len, reverse=True)
    return re.compile(r"\b(" + "|".join(re.escape(k) for k in urut) + r")", re.IGNORECASE)


_POLA_POS = _pola(KATA_POSITIF)
_POLA_NEG = _pola(KATA_NEGATIF)
_POLA_SEMUA = _pola(KATA_POSITIF + KATA_NEGATIF)
_SET_POS = {k.lower() for k in KATA_POSITIF}


def skor_teks(teks):
    """Skor +1 per kata positif, -1 per kata negatif (frasa terpanjang menang)."""
    skor = 0
    for m in _POLA_SEMUA.finditer(teks or ""):
        skor += 1 if m.group(1).lower() in _SET_POS else -1
    return skor


def tentukan_tone(judul, cuplikan=""):
    """Positif / Negatif / Netral. Judul diberi bobot dua kali lipat cuplikan."""
    skor = 2 * skor_teks(judul) + skor_teks(cuplikan)
    if skor > 0:
        return "Positif"
    if skor < 0:
        return "Negatif"
    return "Netral"


def deteksi_seri(teks):
    """Ambil semua kode seri di teks, misalnya 'SR025, ST014'."""
    hasil = []
    for m in POLA_SERI.finditer(teks or ""):
        kode = (m.group(1) + m.group(2) + (m.group(3) or "")).upper()
        if kode not in hasil:
            hasil.append(kode)
    return ", ".join(hasil)


def daftar_seri(seri):
    """'SR025, st015 SR024' atau ['SR025','ST015'] -> ['SR025', 'ST015', 'SR024']."""
    if not seri:
        return []
    if isinstance(seri, str):
        seri = re.split(r"[,;\s]+", seri)
    hasil = []
    for s in seri:
        kode = re.sub(r"[\s-]", "", str(s)).upper()
        if kode and kode not in hasil:
            hasil.append(kode)
    return hasil


def nama_bulan(tgl):
    if pd.isna(tgl):
        return ""
    return f"{NAMA_BULAN[tgl.month - 1]} {tgl.year}"


def parse_tanggal(teks, sekarang):
    """Ubah '3 hari lalu', 'kemarin', atau '12 Okt 2026' menjadi tanggal."""
    t = teks.lower().strip()
    if t == "kemarin":
        return sekarang - timedelta(days=1)
    m = re.fullmatch(r"(\d+)\s+(detik|menit|jam|hari|minggu|bulan|tahun)\s+(?:yang\s+)?lalu", t)
    if m:
        n = int(m.group(1))
        delta = {"detik": timedelta(seconds=n), "menit": timedelta(minutes=n),
                 "jam": timedelta(hours=n), "hari": timedelta(days=n),
                 "minggu": timedelta(weeks=n), "bulan": timedelta(days=30 * n),
                 "tahun": timedelta(days=365 * n)}[m.group(2)]
        return sekarang - delta
    m = re.fullmatch(r"(\d{1,2})\s+([a-z]{3,9})\.?(?:\s+(\d{4}))?", t)
    if m and m.group(2)[:3] in BULAN:
        tahun = int(m.group(3)) if m.group(3) else sekarang.year
        try:
            return datetime(tahun, BULAN[m.group(2)[:3]], int(m.group(1)))
        except ValueError:
            return None
    return None


# ---------------------------------------------------------------------------
# SCRAPING
# ---------------------------------------------------------------------------

# Dijalankan di dalam halaman: ambil setiap tautan berita beserta teks di dalamnya.
# Tidak bergantung pada nama class CSS Google (yang sering berubah).
JS_AMBIL = """
const hasil = [];
const seen = new Set();
document.querySelectorAll('a').forEach(a => {
  const h = a.querySelector('[role="heading"]');
  if (!h) return;
  let href = a.href;
  // Google kadang memakai link pengalihan (google.com/url?q=... atau /goto?url=...).
  // Ambil link asli bila tersedia.
  try {
    const u = new URL(href);
    if (u.hostname.endsWith('google.com') && u.pathname === '/url') {
      href = u.searchParams.get('q') || u.searchParams.get('url') || href;
    }
  } catch (e) {}
  const alt = a.getAttribute('data-href') || a.getAttribute('data-url');
  if (alt && /^https?:/.test(alt) && !alt.includes('google.')) href = alt;
  if (!href || href.includes('google.com/search') || seen.has(href)) return;
  seen.add(href);
  const baris = a.innerText.split('\\n').map(s => s.trim()).filter(Boolean);
  hasil.push({href: href, judul: h.innerText.trim(), baris: baris});
});
return hasil;
"""


def rentang_bulan(bulan_str):
    """'2026-10' -> (date(2026,10,1), date(2026,10,31))."""
    tahun, bulan = map(int, bulan_str.split("-"))
    awal = date(tahun, bulan, 1)
    akhir = (awal.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
    return awal, akhir


def buat_url(kata_kunci, start, dari, sampai):
    """URL pencarian tab News, dibatasi ke rentang tanggal."""
    tbs = (f"cdr:1,cd_min:{dari.month}/{dari.day}/{dari.year},"
           f"cd_max:{sampai.month}/{sampai.day}/{sampai.year}")
    params = {"q": kata_kunci, "tbm": "nws", "hl": "id", "gl": "id",
              "start": start, "tbs": tbs}
    return "https://www.google.com/search?" + urlencode(params)


def olah_item(item, topik, sekarang, seri_query=""):
    judul = item["judul"]
    baris = item["baris"]

    tanggal, teks_tanggal = None, None
    for b in reversed(baris):
        tgl = parse_tanggal(b, sekarang)
        if tgl:
            tanggal, teks_tanggal = tgl, b
            break

    kandidat = [b for b in baris if b != judul and b != teks_tanggal]
    situs = kandidat[0] if kandidat else ""
    cuplikan = " ".join(kandidat[1:])

    ritel = deteksi_seri(f"{judul} {cuplikan}")
    if LABEL_DARI_PENCARIAN and seri_query and seri_query not in ritel.split(", "):
        # Berita hasil pencarian khusus seri tetap diberi label seri tersebut
        ritel = ", ".join(filter(None, [seri_query, ritel]))

    tgl = pd.Timestamp(tanggal.date()) if tanggal else pd.NaT
    # "3 minggu lalu", "1 bulan lalu", "1 tahun lalu" hanya perkiraan kasar
    kira = bool(teks_tanggal and re.search(r"minggu|bulan|tahun", teks_tanggal.lower()))
    return {
        "Tanggal": tgl,
        "Situs": situs,
        "Judul": judul,
        "Link": item["href"],
        "Topik": topik,
        "Bulan": nama_bulan(tgl),
        "Tone": tentukan_tone(judul, cuplikan),
        "Ritel": ritel,
        "_kira": kira,
    }


def tunggu_verifikasi(driver, log):
    """Tunggu sampai CAPTCHA / halaman persetujuan Google diselesaikan."""
    def terhalang():
        u = driver.current_url
        return "/sorry/" in u or "consent.google" in u

    if not terhalang():
        return
    log("Google meminta verifikasi. Selesaikan di jendela Chrome; proses akan lanjut otomatis.")
    batas = time.time() + TUNGGU_CAPTCHA
    while terhalang():
        if time.time() > batas:
            raise RuntimeError("Verifikasi Google tidak diselesaikan dalam batas waktu.")
        time.sleep(3)
    time.sleep(2)


class GoogleMemblokir(RuntimeError):
    """Google menolak permintaan (HTTP 403/429): akses dibatasi sementara."""


def diblokir(driver):
    """True bila halaman berisi penolakan Google (403 Forbidden / 429)."""
    try:
        judul = (driver.title or "").lower()
        isi = driver.page_source[:3000].lower()
    except Exception:
        return False
    return ("does not have permission to get url" in isi or judul.startswith("error 403")
            or "403. that" in isi or "429. that" in isi or "unusual traffic" in isi)


PESAN_BLOKIR = ("Google menolak permintaan (error 403): akses dari komputer/jaringan ini dibatasi "
                "sementara karena dianggap terlalu banyak pencarian otomatis. Pengambilan via "
                "halaman Google dihentikan; data yang sudah terambil tetap disimpan. Tunggu beberapa "
                "jam sebelum mencoba lagi, dan kurangi jumlah pencarian.")


def daftar_situs(teks):
    """'https://www.Kontan.co.id/berita, cnbcindonesia.com' -> ['kontan.co.id', 'cnbcindonesia.com']."""
    if not teks:
        return []
    if isinstance(teks, str):
        teks = re.split(r"[,;\s]+", teks)
    hasil = []
    for t in teks:
        t = str(t).strip().lower()
        t = re.sub(r"^[a-z]+://", "", t).split("/")[0].removeprefix("www.")
        if t and "." in t and t not in hasil:
            hasil.append(t)
    return hasil


def muat_daftar_situs(path=None):
    """Baca daftar_situs.csv -> {"1": [...], "2": [...], "3": [...], "X": [...]}.
    File boleh dipisah titik koma atau koma (CSV dari Excel Indonesia memakai titik koma)."""
    path = Path(path or FILE_DAFTAR_SITUS)
    hasil = {"1": [], "2": [], "3": [], "X": []}
    if not path.exists():
        return hasil
    try:
        d = pd.read_csv(path, sep=None, engine="python", dtype=str, encoding="utf-8-sig")
    except Exception as e:
        print(f"Peringatan: {path.name} tidak dapat dibaca ({e}).")
        return hasil
    kol = {c.strip().lower(): c for c in d.columns}
    if "situs" not in kol or "lapis" not in kol:
        print(f"Peringatan: {path.name} harus punya kolom 'Situs' dan 'Lapis'.")
        return hasil
    for s, l in zip(d[kol["situs"]].fillna(""), d[kol["lapis"]].fillna("")):
        l = str(l).strip().upper().removeprefix("LAPIS").strip()
        for x in daftar_situs(s):
            if l in hasil and x not in hasil[l]:
                hasil[l].append(x)
    return hasil


DAFTAR_SITUS = muat_daftar_situs()
if not SITUS_TUJUAN:
    SITUS_TUJUAN = DAFTAR_SITUS["1"]
if not SITUS_PEMANTAUAN:
    SITUS_PEMANTAUAN = DAFTAR_SITUS["2"]
# Lapis X hanya disaring dari hasil (tidak ditulis sebagai -site: di kueri, supaya kueri tetap pendek)
SITUS_DISARING = DAFTAR_SITUS["X"]


def situs_dengan_pemantauan(situs=None, pemantauan=False):
    """Situs tujuan (bawaan: SITUS_TUJUAN) ditambah situs pemantauan bila diminta."""
    situs = daftar_situs(SITUS_TUJUAN if situs is None else situs)
    if pemantauan:
        situs += [x for x in SITUS_PEMANTAUAN if x not in situs]
    return situs


def _kelompok(situs, ukuran):
    """Bagi situs ke kelompok yang ukurannya rata (49 situs, maks 8 -> 7 kelompok isi 7)."""
    ukuran = max(1, int(ukuran or 1))
    if situs:
        ukuran = -(-len(situs) // -(-len(situs) // ukuran))
    for i in range(0, len(situs), ukuran):
        grup = situs[i:i + ukuran]
        yield (f"site:{grup[0]}" if len(grup) == 1
               else "(" + " OR ".join(f"site:{x}" for x in grup) + ")")


def susun_kueri(kata_kunci, situs=None, kecualikan=None, hanya_situs=False, sumber_situs=None):
    """Daftar pencarian untuk satu kata kunci: list (kueri, maks_halaman, jenis).

    jenis "google" = halaman pencarian Google di Chrome, "rss" = RSS Google News.
    - pencarian umum (semua situs) lewat Chrome, kecuali bila hanya_situs=True
    - pencarian situs tujuan: lewat RSS dengan SITUS_PER_RSS situs digabung per permintaan
      (bawaan), atau lewat Chrome dengan SITUS_PER_KUERI situs digabung (SUMBER_SITUS = "google")
    Situs dikecualikan ditambahkan sebagai -site: pada setiap kueri.
    """
    sumber_situs = sumber_situs or SUMBER_SITUS
    situs = daftar_situs(situs)
    minus = "".join(f" -site:{x}" for x in daftar_situs(kecualikan))
    kueri = []
    if not (hanya_situs and situs):
        kueri.append((f"{kata_kunci}{minus}", MAKS_HALAMAN, "google"))
        if RSS_UMUM:
            kueri.append((f"{kata_kunci}{minus}", 1, "rss"))
    if sumber_situs == "rss":
        kueri += [(f"{kata_kunci} {g}", 1, "rss") for g in _kelompok(situs, SITUS_PER_RSS)]
    else:
        kueri += [(f"{kata_kunci} {g}{minus}", MAKS_HALAMAN_SITUS, "google")
                  for g in _kelompok(situs, SITUS_PER_KUERI)]
    return kueri


def perkiraan(n_kata_kunci, n_seri=0, n_situs=0, hanya_situs=False, sumber_situs=None):
    """Perkiraan beban: (pencarian Google, halaman Google, permintaan RSS, menit)."""
    sumber_situs = sumber_situs or SUMBER_SITUS
    umum = 0 if (hanya_situs and n_situs) else 1
    n_seri_q = n_seri * len(KATA_KUNCI_SERI)
    g_kueri = n_kata_kunci * umum + (0 if (hanya_situs and n_situs) else n_seri_q)
    g_hal = n_kata_kunci * umum * 4 + (0 if (hanya_situs and n_situs) else n_seri_q * 2)
    rss = (n_kata_kunci * umum + (0 if (hanya_situs and n_situs) else n_seri_q)) if RSS_UMUM else 0
    if sumber_situs == "rss":
        n_grup = -(-n_situs // max(1, SITUS_PER_RSS)) if n_situs else 0
        rss += n_kata_kunci * n_grup + (n_seri_q * n_grup if hanya_situs else 0)
    else:
        n_grup = -(-n_situs // SITUS_PER_KUERI) if n_situs else 0
        g_kueri += n_kata_kunci * n_grup
        g_hal += n_kata_kunci * n_grup * 2
    menit = g_hal * 7 / 60 + g_kueri * 4.5 / 60 + rss * 4 / 60
    return g_kueri, g_hal, rss, round(menit)


def ambil_rss(kueri, topik, dari, sampai, seri_query="", log=print):
    """Ambil berita dari RSS Google News untuk satu kueri (misalnya 'SBSN Ritel site:kontan.co.id')."""
    import requests
    from bs4 import BeautifulSoup
    from email.utils import parsedate_to_datetime

    q = f"{kueri} after:{dari.isoformat()} before:{(sampai + timedelta(days=1)).isoformat()}"
    r = requests.get(RSS_URL, params={"q": q, "hl": "id", "gl": "ID", "ceid": "ID:id"},
                     headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}, timeout=30)
    if r.status_code in (403, 429):
        raise GoogleMemblokir(f"RSS Google News menolak permintaan (HTTP {r.status_code}).")
    r.raise_for_status()
    hasil = []
    for item in BeautifulSoup(r.content, "xml").find_all("item"):
        judul = item.title.get_text(strip=True) if item.title else ""
        src = item.find("source")
        situs = src.get_text(strip=True) if src else ""
        if situs and judul.lower().endswith(f" - {situs.lower()}"):
            judul = judul[: -len(f" - {situs}")]
        elif " - " in judul and len(judul.rsplit(" - ", 1)[1]) <= 40:
            judul = judul.rsplit(" - ", 1)[0]   # nama media di akhir judul RSS
        try:
            t = parsedate_to_datetime(item.pubDate.get_text(strip=True))
            t = (t - (t.utcoffset() or timedelta(0))).replace(tzinfo=None) + timedelta(hours=7)
            tgl = pd.Timestamp(t.date())
        except Exception:
            tgl = pd.NaT
        ritel = deteksi_seri(judul)
        if LABEL_DARI_PENCARIAN and seri_query and seri_query not in ritel.split(", "):
            ritel = ", ".join(filter(None, [seri_query, ritel]))
        domain = urlsplit_host(src.get("url", "")) if src is not None else ""
        hasil.append({"Tanggal": tgl, "Situs": situs, "Judul": judul,
                      "Link": item.link.get_text(strip=True) if item.link else "",
                      "Topik": topik, "Bulan": nama_bulan(tgl), "Tone": tentukan_tone(judul),
                      "Ritel": ritel, "_domain": domain, "_rss": True, "_kira": False})
    log(f"  RSS: {len(hasil)} berita")
    return hasil


def halaman_ada_berikutnya(driver):
    """True bila Google masih menampilkan tautan ke halaman hasil berikutnya."""
    try:
        return bool(driver.execute_script(
            "return !!(document.querySelector('a#pnnext') || "
            "document.querySelector('a[aria-label*=\"berikutnya\" i], a[aria-label*=\"next\" i]'))"))
    except Exception:
        return False


def scrape_kata_kunci(driver, kata_kunci, topik, dari, sampai, sekarang, log, seri_query="",
                      maks_halaman=None):
    hasil, sudah = [], set()
    for halaman in range(maks_halaman or MAKS_HALAMAN):
        driver.get(buat_url(kata_kunci, halaman * 10, dari, sampai))
        time.sleep(random.uniform(3, 6))
        tunggu_verifikasi(driver, log)
        if diblokir(driver):
            raise GoogleMemblokir(PESAN_BLOKIR)

        items = driver.execute_script(JS_AMBIL)
        baru = [i for i in items if i["href"] not in sudah]
        if not baru and halaman_ada_berikutnya(driver):
            # Halaman mungkin belum selesai dimuat: tunggu lalu baca sekali lagi
            time.sleep(random.uniform(4, 7))
            items = driver.execute_script(JS_AMBIL)
            baru = [i for i in items if i["href"] not in sudah]
        log(f"  halaman {halaman + 1}: {len(baru)} berita")
        if not baru:
            break
        for i in baru:
            sudah.add(i["href"])
            hasil.append(olah_item(i, topik, sekarang, seri_query))
    return hasil


def buka_chrome():
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options

    opsi = Options()
    opsi.add_argument("--lang=id")
    opsi.add_argument("--disable-blink-features=AutomationControlled")
    # Profil Chrome khusus skrip ini (terpisah dari Chrome pribadi kamu)
    opsi.add_argument(f"--user-data-dir={PROFIL_CHROME}")
    return webdriver.Chrome(options=opsi)


def jalankan_scraping(dari, sampai, seri="", hanya_seri=False, kata_kunci=None, log=print,
                      gabung_judul_sama=None, situs=None, kecualikan=None, hanya_situs=False,
                      sumber_situs=None):
    """Jalankan seluruh proses dan kembalikan DataFrame dengan kolom KOLOM.

    seri boleh satu atau beberapa kode: "SR025", "SR025, ST015", atau ["SR025", "ST015"].
    Peringatan (misalnya Google memblokir) tersedia di df.attrs["peringatan"].
    """
    seri = daftar_seri(seri)
    kata_kunci = dict(kata_kunci or KATA_KUNCI)
    situs = daftar_situs(SITUS_TUJUAN if situs is None else situs)
    kecualikan = daftar_situs(SITUS_DIKECUALIKAN if kecualikan is None else kecualikan)
    sekarang = datetime.now()

    daftar = []   # (kueri, topik, seri, maks_halaman, jenis)
    for kw, topik in kata_kunci.items():
        daftar += [(q, topik, "", h, j) for q, h, j in
                   susun_kueri(kw, situs, kecualikan, hanya_situs, sumber_situs)]
    for s_ in seri:
        for pola in KATA_KUNCI_SERI:
            daftar += [(q, "SBSN Ritel", s_, h, j) for q, h, j in
                       susun_kueri(pola.format(seri=s_), situs if hanya_situs else [], kecualikan,
                                   hanya_situs, sumber_situs)]
    via_google = [d for d in daftar if d[4] == "google"]
    via_rss = [d for d in daftar if d[4] == "rss"]
    log(f"{len(via_google)} pencarian Google (Chrome) dan {len(via_rss)} permintaan RSS.")

    semua, peringatan = [], []
    if via_google:
        driver = buka_chrome()
        try:
            for kw, topik, seri_q, maks, _ in via_google:
                log(f"Mengambil: {kw}")
                try:
                    hasil = scrape_kata_kunci(driver, kw, topik, dari, sampai, sekarang, log,
                                              seri_q, maks)
                    for h in hasil:
                        h["_domain"], h["_rss"] = urlsplit_host(h["Link"]), False
                    semua.extend(hasil)
                except GoogleMemblokir as e:
                    log(f"⚠️ {e}")
                    peringatan.append(str(e))
                    break
                except RuntimeError:
                    raise
                except Exception as e:  # satu kata kunci gagal, lanjut yang lain
                    log(f"  gagal: {e}")
                time.sleep(random.uniform(3, 6))
        finally:
            driver.quit()

    for kw, topik, seri_q, _, _ in via_rss:
        log(f"Mengambil (RSS): {kw}")
        try:
            semua.extend(ambil_rss(kw, topik, dari, sampai, seri_q, log))
        except GoogleMemblokir as e:
            pesan = f"{e} Pengambilan RSS dihentikan; tunggu beberapa jam sebelum mencoba lagi."
            log(f"⚠️ {pesan}")
            peringatan.append(pesan)
            break
        except Exception as e:
            log(f"  gagal: {e}")
        time.sleep(random.uniform(*JEDA_RSS))

    df = pd.DataFrame(semua)
    if len(df):
        saring = kecualikan + [x for x in SITUS_DISARING if x not in kecualikan]
        if saring:
            df = df[~df["_domain"].map(lambda h: cocok_domain(h, saring))]
        if hanya_situs and situs:
            df = df[df["_domain"].map(lambda h: cocok_domain(h, situs))]
    if len(df):
        log(f"Terkumpul {int((~df['_rss']).sum())} berita dari halaman Google dan "
            f"{int(df['_rss'].sum())} dari RSS (sebelum duplikat dibuang).")
    hasil = rapikan(df, dari, sampai, seri, hanya_seri, gabung_judul_sama)
    log(f"Setelah duplikat dibuang: {len(hasil)} berita unik.")
    hasil.attrs["peringatan"] = peringatan
    return hasil


def urlsplit_host(url):
    from urllib.parse import urlsplit
    return urlsplit(str(url)).netloc.lower().removeprefix("www.")


def cocok_domain(host, daftar):
    """True bila host sama dengan / subdomain dari salah satu domain di daftar."""
    return any(host == d or host.endswith("." + d) for d in daftar)


_SLD_ID = {"co", "go", "ac", "or", "web", "my", "sch", "net", "biz", "ponpes", "desa", "mil"}


def domain_dasar(host):
    """'keuangan.kontan.co.id' -> 'kontan.co.id'; 'm.detik.com' -> 'detik.com'."""
    bagian = str(host or "").lower().removeprefix("www.").split(".")
    if len(bagian) >= 3 and bagian[-1] == "id" and bagian[-2] in _SLD_ID:
        return ".".join(bagian[-3:])
    return ".".join(bagian[-2:])


def kunci_judul(judul):
    """Judul baku (huruf/angka saja) untuk mencocokkan berita yang sama dari sumber berbeda.

    Akhiran "..." dibuang karena Google sering memotong judul yang panjang, dan nama media di
    akhir judul ("... - Kabar Baitullah") juga dibuang.
    """
    j = str(judul or "").lower().strip()
    j = re.sub(r"\s*(\.\.\.|…)\s*$", "", j)
    if " - " in j and len(j.rsplit(" - ", 1)[1]) <= 40:
        j = j.rsplit(" - ", 1)[0]
    return re.sub(r"[^a-z0-9]", "", j)


def judul_cocok(a, b, min_huruf=25):
    """Judul sama, atau salah satunya potongan awal dari yang lain (judul terpotong Google)."""
    if not a or not b or min(len(a), len(b)) < min_huruf:
        return a == b and bool(a)
    return a.startswith(b) or b.startswith(a)


def _norm_situs(s):
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def situs_cocok(a, b):
    a, b = _norm_situs(a), _norm_situs(b)
    return not a or not b or a == b or a in b or b in a


def adalah_rss(link):
    return "news.google.com" in str(link)


_POLA_PENGALIHAN = re.compile(r"^https?://(www\.)?(news\.google\.[a-z.]+/|google\.[a-z.]+/(goto|url)\b)",
                              re.IGNORECASE)


def adalah_pengalihan(link):
    """True untuk link pengalihan Google, bukan link artikel asli:
    news.google.com/rss/articles/..., google.com/goto?url=..., google.com/url?q=...
    Link seperti ini berbeda-beda untuk artikel yang SAMA (tergantung pencarian), jadi tidak
    bisa dipakai untuk mengenali duplikat."""
    return bool(_POLA_PENGALIHAN.match(str(link).strip()))


_PARAM_PELACAK = re.compile(r"^(utm_\w+|fbclid|gclid|ocid|cmpid|ref|src|amp|page|oc)$", re.IGNORECASE)


def kunci_link(url):
    """Bentuk baku link untuk membandingkan duplikat.

    Hanya membuang bagian yang tidak mengubah artikel: http/https, huruf besar pada nama domain,
    "www."/"m."/"amp." di depan domain, versi AMP (/amp), garis miring di akhir, #bagian, dan
    parameter seperti utm_source, fbclid, page=all. Link ke artikel berbeda tetap dianggap berbeda.
    """
    from urllib.parse import parse_qs, parse_qsl, urlsplit, urlunsplit
    u = urlsplit(str(url).strip())
    if re.match(r"(www\.)?google\.[a-z.]+$", u.netloc.lower()) and u.path == "/url":
        asli = (parse_qs(u.query).get("q") or parse_qs(u.query).get("url") or [""])[0]
        if asli.startswith("http"):
            u = urlsplit(asli)
    host = u.netloc.lower()
    for awalan in ("www.", "m.", "amp.", "mobile."):
        host = host.removeprefix(awalan)
    path = re.sub(r"(^/amp(?=/)|/amp/?$|\.amp$)", "", u.path).rstrip("/")
    query = urlencode([(k, v) for k, v in parse_qsl(u.query, keep_blank_values=True)
                       if not _PARAM_PELACAK.match(k)])
    return urlunsplit(("", host, path, query, ""))


def _gabung_ritel(*nilai):
    """Gabungkan label seri: ('SR025', 'ST015, SR025', '') -> 'SR025, ST015'."""
    return ", ".join(dict.fromkeys(
        x for v in nilai if isinstance(v, str) or pd.notna(v)
        for x in daftar_seri(str(v or ""))))


def _satukan(df, kunci):
    """Buang baris dengan kunci sama (baris pertama dipertahankan), label Ritel semua baris
    dengan kunci itu digabung ke baris yang dipertahankan."""
    kunci = pd.Series(list(kunci), index=df.index)
    ganda = kunci.duplicated(keep="first")
    if not ganda.any():
        return df
    label = df["Ritel"].groupby(kunci, sort=False).agg(lambda s: _gabung_ritel(*s))
    df = df[~ganda].copy()
    df["Ritel"] = kunci[df.index].map(label)
    return df


def buang_duplikat(df, gabung_judul_sama=None):
    """Buang duplikat.

    1. Link sama (setelah dibakukan oleh kunci_link) -> satu berita. Label Ritel yang berbeda
       (misalnya satu salinan "SR025", salinan lain kosong atau "ST015") digabung.
    2. Berita dengan link PENGALIHAN Google (news.google.com/..., google.com/goto?url=...,
       google.com/url?q=...) yang judul + situsnya sama dengan berita lain -> satu berita. Link
       pengalihan berbeda untuk setiap pencarian walau artikelnya sama, jadi BUKAN berita berbeda. Versi dengan link asli diutamakan; tetapi bila versi RSS adalah data
       lama (kolom _lama) yang mungkin sudah dikoreksi, data lama dipertahankan dan link-nya
       diganti dengan link asli.
    3. Bila gabung_judul_sama (atau GABUNG_JUDUL_SAMA) aktif: judul + situs sama -> satu berita,
       walau kedua link-nya link asli.
    """
    if gabung_judul_sama is None:
        gabung_judul_sama = GABUNG_JUDUL_SAMA
    if df.empty:
        return df
    # 1) Link sama -> satu berita; label Ritel dari semua salinannya digabung
    df = _satukan(df.copy(), df["Link"].map(kunci_link))

    idx = list(df.index)
    jd = {i: kunci_judul(df.at[i, "Judul"]) for i in idx}
    rss = {i: adalah_pengalihan(df.at[i, "Link"]) for i in idx}
    lama = {i: bool(df.at[i, "_lama"]) if "_lama" in df.columns else False for i in idx}
    tgl = {i: df.at[i, "Tanggal"] for i in idx}

    def dekat(a, b):
        return pd.isna(a) or pd.isna(b) or abs((pd.Timestamp(a) - pd.Timestamp(b)).days) <= 7

    def sama(i, j):
        return (judul_cocok(jd[i], jd[j]) and situs_cocok(df.at[i, "Situs"], df.at[j, "Situs"])
                and dekat(tgl[i], tgl[j]))

    buang, rss_simpan = set(), []
    asli = [i for i in idx if not rss[i]]
    for i in idx:
        if not rss[i]:
            continue
        pasangan = next((j for j in asli if j not in buang and sama(i, j)), None)
        if pasangan is not None:
            j = pasangan
            if lama[i] and not lama[j]:
                df.at[i, "Link"] = df.at[j, "Link"]          # data lama tetap, link asli dipakai
                df.at[i, "Ritel"] = _gabung_ritel(df.at[i, "Ritel"], df.at[j, "Ritel"])
                buang.add(j)
                rss_simpan.append(i)
            else:
                df.at[j, "Ritel"] = _gabung_ritel(df.at[j, "Ritel"], df.at[i, "Ritel"])
                kira_j = bool(df.at[j, "_kira"]) if "_kira" in df.columns else False
                if pd.isna(df.at[j, "Tanggal"]) or kira_j:
                    if pd.notna(tgl[i]):
                        df.at[j, "Tanggal"] = tgl[i]      # tanggal pasti dari RSS
                buang.add(i)
            continue
        pasangan = next((j for j in rss_simpan if sama(i, j)), None)
        if pasangan is not None:
            df.at[pasangan, "Ritel"] = _gabung_ritel(df.at[pasangan, "Ritel"], df.at[i, "Ritel"])
            buang.add(i)
        else:
            rss_simpan.append(i)
    df = df.drop(index=list(buang))

    if gabung_judul_sama:
        df = _satukan(df, zip(df["Judul"].map(kunci_judul), df["Situs"].map(_norm_situs)))
    return df


def rapikan(df, dari, sampai, seri="", hanya_seri=False, gabung_judul_sama=None):
    if df.empty:
        return pd.DataFrame(columns=KOLOM)

    df = df.copy()
    for k, v in [("_rss", False), ("_kira", False), ("_domain", "")]:
        if k not in df.columns:
            df[k] = v
    df["_rss"] = df["_rss"].fillna(False).astype(bool)
    df["_kira"] = df["_kira"].fillna(False).astype(bool)
    df["_domain"] = df["_domain"].fillna("")
    # 1) Tanggal perkiraan dari halaman Google ("1 bulan lalu") diganti tanggal pasti dari RSS
    #    bila beritanya sama (judul + situs) -> dilakukan di buang_duplikat() di bawah. Agar
    #    berita itu tidak terbuang karena tanggal perkiraannya meleset, langkah 2 menyesuaikannya.

    # 2) Pencarian sudah dibatasi Google ke rentang tanggal, jadi tanggal perkiraan yang meleset
    #    keluar rentang TIDAK dibuang, melainkan disesuaikan ke batas rentang terdekat.
    awal, akhir = pd.Timestamp(dari), pd.Timestamp(sampai)
    kira = df["_kira"]
    df.loc[kira & (df["Tanggal"] < awal), "Tanggal"] = awal
    df.loc[kira & (df["Tanggal"] > akhir), "Tanggal"] = akhir

    # 3) Tanggal pasti di luar rentang dibuang (baris tanpa tanggal terbaca tetap dipertahankan)
    tgl = df["Tanggal"]
    df = df[tgl.isna() | ((tgl >= awal) & (tgl <= akhir))]
    df = df.assign(Bulan=df["Tanggal"].map(nama_bulan))

    # Gabungkan label seri sebelum membuang duplikat, supaya label tidak hilang
    kunci = df["Link"].map(kunci_link)
    label = df.groupby(kunci)["Ritel"].agg(
        lambda s: ", ".join(dict.fromkeys(x for v in s for x in v.split(", ") if x)))
    df = df.assign(Ritel=kunci.map(label))

    # Berita yang sama dari halaman Google (link asli) dan dari RSS (link pengalihan Google)
    # dijadikan satu; versi link asli diutamakan dan memakai tanggal pasti dari RSS.
    df = df.sort_values("_rss", kind="stable")
    df = buang_duplikat(df, gabung_judul_sama)

    seri = daftar_seri(seri)
    if seri and hanya_seri:
        df = df[df["Ritel"].str.split(", ").apply(lambda xs: any(s in xs for s in seri))]

    df = df.sort_values("Tanggal", ascending=True, na_position="last", kind="stable").reset_index(drop=True)
    df.insert(0, "No", range(1, len(df) + 1))
    return df[KOLOM]


# ---------------------------------------------------------------------------
# SIMPAN EXCEL
# ---------------------------------------------------------------------------

LEBAR = {"No": 6, "Tanggal": 13, "Situs": 24, "Judul": 80, "Link": 50,
         "Topik": 14, "Bulan": 16, "Tone": 10, "Ritel": 14}


def urutkan_tanggal(df):
    """Urutan baku semua keluaran: tanggal terlama -> terbaru (tanpa tanggal di akhir)."""
    df = df.copy()
    df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
    return df.sort_values("Tanggal", ascending=True, na_position="last", kind="stable").reset_index(drop=True)


def excel_bytes(df):
    from openpyxl.styles import Alignment, Font, PatternFill

    df = urutkan_tanggal(df)
    df["No"] = range(1, len(df) + 1)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df[KOLOM].to_excel(writer, index=False, sheet_name="Berita")
        ws = writer.sheets["Berita"]
        for i, kolom in enumerate(KOLOM, start=1):
            huruf = ws.cell(row=1, column=i).column_letter
            ws.column_dimensions[huruf].width = LEBAR[kolom]
            sel = ws.cell(row=1, column=i)
            sel.font = Font(bold=True, color="FFFFFF")
            sel.fill = PatternFill("solid", fgColor="4B2E83")
            sel.alignment = Alignment(horizontal="center")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        warna = {"Positif": "D9F2E3", "Negatif": "F9D9D9", "Netral": "DCEAF7"}   # Netral = biru muda
        k_tgl = KOLOM.index("Tanggal") + 1
        k_link = KOLOM.index("Link") + 1
        k_tone = KOLOM.index("Tone") + 1
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=k_tgl).number_format = "DD/MM/YYYY"
            link = ws.cell(row=r, column=k_link)
            if link.value:
                link.hyperlink = link.value
                link.font = Font(color="0563C1", underline="single")
            tone = ws.cell(row=r, column=k_tone)
            if tone.value in warna:
                tone.fill = PatternFill("solid", fgColor=warna[tone.value])
    return buf.getvalue()


def nama_file(dari, sampai, seri=""):
    if dari.day == 1 and sampai == rentang_bulan(dari.strftime("%Y-%m"))[1]:
        periode = dari.strftime("%Y-%m")
    else:
        periode = f"{dari:%Y%m%d}-{sampai:%Y%m%d}"
    seri = daftar_seri(seri)
    akhiran = ("_" + "_".join(seri)) if seri else ""
    return f"berita_sbsn_{periode}{akhiran}.xlsx"


class FileTerkunci(PermissionError):
    """File hasil tidak bisa ditulis, biasanya karena sedang dibuka di Excel."""

    def __init__(self, path):
        self.path = Path(path)
        super().__init__(
            f"File {self.path.name} sedang dibuka di program lain (biasanya Microsoft Excel) atau "
            "sedang disinkronkan OneDrive, sehingga tidak bisa disimpan. Tutup file tersebut di "
            "Excel, lalu ulangi.")


def tulis_file(path, data, coba=4, jeda=1.5):
    """Tulis file; bila terkunci (Windows: file dibuka di Excel), coba lagi beberapa kali."""
    path = Path(path)
    for k in range(coba):
        try:
            path.write_bytes(data)
            return path
        except PermissionError:
            if k == coba - 1:
                raise FileTerkunci(path) from None
            time.sleep(jeda)


def gabung_dengan_lama(lama, baru, gabung_judul_sama=None):
    """Gabungkan hasil baru ke data lama. Data lama (termasuk koreksi tone/seri) diutamakan.

    Mengembalikan (gabungan, jumlah_berita_baru).
    """
    lama = lama.copy()
    lama["Tanggal"] = pd.to_datetime(lama["Tanggal"], errors="coerce")
    for k in ["Situs", "Judul", "Link", "Topik", "Bulan", "Tone", "Ritel"]:
        lama[k] = lama[k].fillna("").astype(str)
    gabungan = pd.concat([lama.assign(_lama=True), baru.assign(_lama=False)], ignore_index=True)
    gabungan = buang_duplikat(gabungan, gabung_judul_sama)
    n_baru = int((~gabungan["_lama"].astype(bool)).sum())
    gabungan = gabungan.sort_values("Tanggal", ascending=True, na_position="last", kind="stable").reset_index(drop=True)
    gabungan["Bulan"] = gabungan["Tanggal"].map(nama_bulan)
    gabungan["No"] = range(1, len(gabungan) + 1)
    return gabungan[KOLOM], n_baru


def bersihkan_file(path, gabung_judul_sama=None):
    """Buang duplikat di dalam satu file hasil yang sudah tersimpan. Mengembalikan jumlah dibuang."""
    df = pd.read_excel(path)
    df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
    for k in ["Situs", "Judul", "Link", "Topik", "Bulan", "Tone", "Ritel"]:
        df[k] = df[k].fillna("").astype(str)
    # Urutkan link asli lebih dulu supaya versi link asli yang dipertahankan
    df = df.assign(_r=df["Link"].map(adalah_pengalihan)).sort_values("_r", kind="stable").drop(columns="_r")
    bersih = buang_duplikat(df.assign(_lama=True), gabung_judul_sama)
    dibuang = len(df) - len(bersih)
    urutan_berubah = not df["Tanggal"].reset_index(drop=True).equals(
        urutkan_tanggal(df)["Tanggal"].reset_index(drop=True))
    if dibuang or urutan_berubah:
        bersih = bersih.sort_values("Tanggal", ascending=True, na_position="last", kind="stable").reset_index(drop=True)
        tulis_file(path, excel_bytes(bersih))
    return dibuang


def simpan_excel(df, dari, sampai, seri="", gabung=True, gabung_judul_sama=None):
    """Simpan hasil ke Excel.

    gabung=True (bawaan): bila file periode yang sama sudah ada, hasil baru DITAMBAHKAN ke file
    itu (berita yang sudah ada tidak digandakan, koreksi lama tetap). Jadi pengambilan berulang
    untuk periode yang sama makin melengkapi data, bukan menimpanya.
    Mengembalikan (path, jumlah_berita_baru, jumlah_total).

    Bila file tujuan sedang dibuka di Excel (terkunci), hasil TIDAK hilang: disimpan ke file
    cadangan bernama sama + jam (misalnya ..._cadangan_143005.xlsx) dan peringatan ditambahkan ke
    df.attrs["peringatan"].
    """
    FOLDER_HASIL.mkdir(exist_ok=True)
    peringatan = df.attrs.setdefault("peringatan", [])
    path = FOLDER_HASIL / nama_file(dari, sampai, seri)
    n_baru = len(df)
    if gabung and path.exists():
        lama = pd.read_excel(path)
        df, n_baru = gabung_dengan_lama(lama, df, gabung_judul_sama)
    data = excel_bytes(df)
    try:
        tulis_file(path, data)
    except FileTerkunci:
        cadangan = path.with_name(f"{path.stem}_cadangan_{datetime.now():%H%M%S}{path.suffix}")
        tulis_file(cadangan, data)
        peringatan.append(
            f"File {path.name} sedang dibuka di Excel (atau sedang disinkronkan OneDrive), jadi hasil "
            f"disimpan ke file cadangan {cadangan.name}. Data tidak hilang dan tetap tampil di "
            "Dashboard. Lain kali tutup file hasil di Excel sebelum mengambil berita.")
        path = cadangan
    return path, n_baru, len(df)


# ---------------------------------------------------------------------------
# JALUR TERMINAL
# ---------------------------------------------------------------------------


def main():
    ap = argparse.ArgumentParser(description="Scraper berita SBSN dari tab Berita Google.")
    ap.add_argument("bulan", nargs="?", help="Bulan YYYY-MM (bawaan: bulan berjalan)")
    ap.add_argument("--dari", help="Tanggal awal YYYY-MM-DD")
    ap.add_argument("--sampai", help="Tanggal akhir YYYY-MM-DD")
    ap.add_argument("--seri", nargs="+", default=[],
                    help="Satu atau beberapa kode seri, misalnya: --seri SR025 ST015")
    ap.add_argument("--situs", nargs="+", default=None,
                    help="Situs tujuan tambahan, misalnya: --situs kontan.co.id bisnis.com")
    ap.add_argument("--pemantauan", action="store_true",
                    help="Sertakan situs pemantauan (Lapis 2 di daftar_situs.csv)")
    ap.add_argument("--situs-via-google", action="store_true",
                    help="Cari situs tujuan lewat halaman Google (bukan RSS); lebih mudah diblokir")
    ap.add_argument("--hanya-situs", action="store_true",
                    help="Hanya ambil berita dari situs tujuan")
    ap.add_argument("--kecualikan", nargs="+", default=None,
                    help="Situs yang tidak diambil beritanya")
    ap.add_argument("--gabung-judul-sama", action="store_true",
                    help="Anggap judul & situs yang sama sebagai duplikat walau link berbeda")
    ap.add_argument("--timpa", action="store_true",
                    help="Timpa file periode yang sama (bawaan: hasil baru ditambahkan ke file lama)")
    ap.add_argument("--hanya-seri", action="store_true",
                    help="Hanya simpan berita yang berlabel salah satu seri tersebut")
    a = ap.parse_args()

    if a.dari or a.sampai:
        if not (a.dari and a.sampai):
            ap.error("--dari dan --sampai harus diisi keduanya.")
        dari = date.fromisoformat(a.dari)
        sampai = date.fromisoformat(a.sampai)
    else:
        dari, sampai = rentang_bulan(a.bulan or datetime.now().strftime("%Y-%m"))

    df = jalankan_scraping(dari, sampai, a.seri, a.hanya_seri,
                           gabung_judul_sama=a.gabung_judul_sama or None,
                           situs=situs_dengan_pemantauan(a.situs, a.pemantauan),
                           kecualikan=a.kecualikan, hanya_situs=a.hanya_situs,
                           sumber_situs="google" if a.situs_via_google else None)
    for p in df.attrs.get("peringatan", []):
        print("PERINGATAN:", p)
    if df.empty:
        print("Tidak ada data yang terambil.")
        return
    path, n_baru, n_total = simpan_excel(df, dari, sampai, a.seri, gabung=not a.timpa)
    print(f"\nSelesai: {len(df)} berita terambil, {n_baru} berita baru. "
          f"Total {n_total} berita di {path}")
    print(df["Tone"].value_counts().to_string())


if __name__ == "__main__":
    main()
