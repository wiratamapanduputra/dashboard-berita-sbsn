"""Dashboard Berita SBSN — versi online, BACA SAJA (untuk pimpinan dan rekan kerja).

Memakai ulang tampilan app.py (tema, grafik, Dashboard, Ringkasan pimpinan, laporan PDF/Word)
tanpa mengubah app.py, tetapi:
  - data dibaca dari Google Sheet (bukan folder hasil),
  - menu Ambil berita dan Tinjau & kelola tidak ditampilkan,
  - berita tidak bisa diedit atau dihapus langsung dari dashboard. Bila Sheet usulan diatur,
    pengguna bisa MENGUSULKAN perubahan (✏️ pada kartu berita, wajib mengisi nama). Usulan baru
    berlaku setelah disetujui admin di aplikasi admin (menu 📥 Usulan).

Jalankan:
  - Online  : file utama (main file) di Streamlit Community Cloud = viewer_app.py
  - Lokal   : streamlit run viewer_app.py
              Tanpa secrets Google Sheet, dashboard memakai file di folder hasil (mode pratinjau).

Isi secrets (Streamlit Cloud > App settings > Secrets, atau .streamlit/secrets.toml):
    id_sheet = "https://docs.google.com/spreadsheets/d/xxxx/edit"
    # kata_sandi = "opsional, bila aplikasi dibuat publik"
    # id_sheet_usulan = "https://docs.google.com/spreadsheets/d/yyyy/edit"   # aktifkan fitur usulan
    # kode_usulan = "opsional: kode yang wajib diisi saat mengirim usulan"
    [gcp_service_account]
    type = "service_account"
    ... (salin seluruh isi kunci JSON service account PEMBACA)
"""

import hmac
import html
from datetime import date
from pathlib import Path

import streamlit as st

FOLDER = Path(__file__).resolve().parent
FILE_APP = FOLDER / "app.py"
PENANDA = "# SUSUNAN HALAMAN"
MENIT_SEGAR = 10
MAKS_USULAN_PER_SESI = 20

# Teks di app.py yang tidak berlaku pada versi baca-saja (dilewati bila tidak ditemukan)
GANTI = [
    ('st.caption("Klik ✏️ di samping kartu untuk mengedit atau menghapus berita.")',
     "st.caption(TEKS_KARTU) if TEKS_KARTU else None"),
    ('Ambil berita dulu di menu **🔎 Ambil berita** "\n                "atau ganti sumber data.',
     'Silakan ganti sumber data."\n                "'),
]


# ---------------------------------------------------------------------------
# MEMUAT BAGIAN DEFINISI app.py
# ---------------------------------------------------------------------------

@st.cache_resource(show_spinner=False)
def _kode_app(mtime):
    teks = FILE_APP.read_text(encoding="utf-8")
    i = teks.find(PENANDA)
    if i < 0:
        raise RuntimeError(f"Penanda '{PENANDA}' tidak ditemukan di app.py.")
    bagian = teks[:i]
    for lama, baru in GANTI:
        bagian = bagian.replace(lama, baru)
    return compile(bagian, str(FILE_APP), "exec")


# Menjalankan semua definisi, tema, dan pengaturan halaman dari app.py (tanpa menampilkan halaman)
G = {"__name__": "app_viewer", "__file__": str(FILE_APP)}
exec(_kode_app(FILE_APP.stat().st_mtime), G)
sc, ss = G["sc"], G["ss"]
G["TEKS_KARTU"] = ""


# ---------------------------------------------------------------------------
# AKSES & DATA
# ---------------------------------------------------------------------------

def rahasia(kunci, bawaan=None):
    try:
        return st.secrets[kunci] if kunci in st.secrets else bawaan
    except Exception:  # belum ada file secrets
        return bawaan


def cek_kata_sandi():
    """Lapisan tambahan opsional: hanya aktif bila 'kata_sandi' diisi di secrets."""
    sandi = rahasia("kata_sandi")
    if not sandi or ss.get("sandi_ok"):
        return
    G["hero"]()
    with st.form("masuk"):
        isian = st.text_input("Kata sandi", type="password")
        if st.form_submit_button("Masuk", type="primary"):
            if hmac.compare_digest(isian, str(sandi)):
                ss["sandi_ok"] = True
                st.rerun()
            st.error("Kata sandi salah.")
    st.stop()


@st.cache_data(ttl=MENIT_SEGAR * 60, show_spinner="Memuat data berita terbaru…")
def muat_dari_sheet(id_sheet, info):
    import sinkron_sheets as sk
    sh = sk.buka(id_sheet, info=dict(info), baca_saja=True)
    return sk.baca_untuk_dashboard(sh)


def muat_data():
    """(DataFrame, info, sumber). Tanpa secrets: pratinjau dari folder hasil."""
    id_sheet, akun = rahasia("id_sheet"), rahasia("gcp_service_account")
    if id_sheet and akun:
        try:
            df, info = muat_dari_sheet(str(id_sheet), dict(akun))
            return df, info, "sheet"
        except Exception as e:
            G["hero"]()
            st.error(f"Data tidak dapat dimuat dari Google Sheet: {e}")
            st.stop()
    files = sorted(sc.FOLDER_HASIL.glob("berita_sbsn_*.xlsx")) if sc.FOLDER_HASIL.exists() else []
    df = G["gabung"](files)
    return df, {"Terakhir diperbarui": "pratinjau lokal"}, "lokal"


# ---------------------------------------------------------------------------
# PENGGANTI FUNGSI app.py UNTUK VERSI BACA-SAJA
# ---------------------------------------------------------------------------

def sumber_data_online(kunci):
    semua = G["gabung"]()
    nama_periode, saring = G["nama_periode"], G["saring_periode"]
    hari_ini = date.today()
    b_dari, b_sampai = sc.rentang_bulan(hari_ini.strftime("%Y-%m"))
    opsi = [f"Bulan ini ({nama_periode(b_dari, b_sampai)})", "Bulan tertentu", "Rentang tanggal",
            "Semua data"]
    if ss.get(f"src_{kunci}") not in range(len(opsi)):
        ss[f"src_{kunci}"] = 0
    with st.container(border=True):
        c1, c2 = st.columns([2, 3])
        mode = c1.selectbox("Sumber data", range(len(opsi)), format_func=lambda i: opsi[i],
                            key=f"src_{kunci}")
        if mode == 0:
            dari, sampai = b_dari, b_sampai
            df = saring(semua, dari, sampai)
            c2.caption("Menampilkan berita bulan berjalan. Pilih **Sumber data** lain untuk "
                       "melihat bulan atau rentang tanggal lain.")
        elif mode == 1:
            t = semua["Tanggal"].dropna()
            bulan = sorted({x.strftime("%Y-%m") for x in t}, reverse=True) or [b_dari.strftime("%Y-%m")]
            k = f"rng_bln_{kunci}"
            if ss.get(k) not in bulan:
                ss[k] = bulan[0]
            pilih = c2.selectbox("Bulan", bulan, key=k,
                                 format_func=lambda s: f"{sc.NAMA_BULAN[int(s[5:]) - 1]} {s[:4]}")
            dari, sampai = sc.rentang_bulan(pilih)
            df = saring(semua, dari, sampai)
        elif mode == 2:
            ss.setdefault(f"rng_{kunci}", (b_dari, hari_ini))
            r = c2.date_input("Rentang tanggal", key=f"rng_{kunci}", format="DD/MM/YYYY")
            dari, sampai = (r if len(r) == 2 else (r[0], r[0]))
            df = saring(semua, dari, sampai)
        else:
            df = semua
            t = df["Tanggal"].dropna()
            dari, sampai = (t.min().date(), t.max().date()) if len(t) else (None, None)
    df = df.reset_index(drop=True)
    df["No"] = range(1, len(df) + 1)
    return df, dari, sampai, mode


def tabel_baca(df, kunci):
    st.dataframe(
        df[sc.KOLOM], hide_index=True, width="stretch", key=f"tabel_{kunci}",
        column_config={
            "Link": st.column_config.LinkColumn("Link", display_text="buka"),
            "Tanggal": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Judul": st.column_config.TextColumn(width="large"),
            "Topik": st.column_config.TextColumn("Kata kunci"),
        })


_kartu_asli = G["kartu_berita"]


def kartu_baca(df, kunci, per_halaman=15, edit=True):
    return _kartu_asli(df, kunci, per_halaman, edit=False)


# ---------------------------------------------------------------------------
# USULAN PERUBAHAN (menunggu persetujuan admin)
# ---------------------------------------------------------------------------

def usulan_aktif():
    return bool(rahasia("id_sheet_usulan") and rahasia("gcp_service_account"))


def _sheet_usulan():
    import usulan as us
    return us.buka(str(rahasia("id_sheet_usulan")), info=dict(rahasia("gcp_service_account")))


@st.cache_data(ttl=60, show_spinner=False)
def menunggu_usulan(_kunci_cache):
    """Kunci link berita yang sedang punya usulan menunggu persetujuan."""
    import usulan as us
    try:
        return us.link_menunggu(us.daftar_usulan(_sheet_usulan()))
    except Exception:
        return set()


@st.dialog("✏️ Usulkan perubahan berita", width="large")
def dialog_usulan(r):
    import usulan as us
    TONE = G["TONE"]
    st.caption("Usulan **tidak langsung mengubah data**. Data berubah setelah disetujui admin.")
    st.markdown(f"**{html.escape(r['Judul'])}**  \n{html.escape(r['Situs'] or '—')} · "
                f"[buka berita]({r['Link']})")
    c1, c2 = st.columns(2)
    nama = c1.text_input("Nama pengusul *", value=ss.get("nama_pengusul", ""), max_chars=60)
    kode_wajib = rahasia("kode_usulan")
    kode = c2.text_input("Kode usulan *", type="password") if kode_wajib else ""
    hapus = st.toggle("Usulkan berita ini **dihapus** (tidak relevan / ganda)")
    perubahan = {}
    if not hapus:
        a, b, c = st.columns(3)
        tone = a.selectbox("Tone", TONE, index=TONE.index(r["Tone"]) if r["Tone"] in TONE else 1)
        opsi = sorted(set(SEMUA["Topik"]) | set(sc.KATA_KUNCI.values()) | {r["Topik"]})
        topik = b.selectbox("Kata kunci", opsi, index=opsi.index(r["Topik"]), accept_new_options=True)
        ritel = c.text_input("Seri ritel", r["Ritel"], placeholder="mis. SR025, ST015")
        with st.expander("Ubah judul, situs, atau tanggal"):
            judul = st.text_input("Judul", r["Judul"])
            d1, d2 = st.columns(2)
            situs = d1.text_input("Situs media", r["Situs"])
            tgl = d2.date_input("Tanggal", r["Tanggal"].date() if G["pd"].notna(r["Tanggal"]) else None,
                                format="DD/MM/YYYY")
        baru = {"Judul": judul, "Situs": situs, "Tanggal": tgl, "Topik": topik or r["Topik"],
                "Tone": tone, "Ritel": ritel}
        perubahan = {k: (r[k], baru[k]) for k in us.KOLOM_EDIT}
    alasan = st.text_area("Alasan / catatan" + (" *" if hapus else " (opsional)"), max_chars=500)
    if st.button("📨 Kirim usulan", type="primary"):
        if kode_wajib and not hmac.compare_digest(kode, str(kode_wajib)):
            st.error("Kode usulan salah.")
            return
        if ss.get("n_usulan", 0) >= MAKS_USULAN_PER_SESI:
            st.error("Batas usulan untuk sesi ini sudah tercapai. Muat ulang halaman untuk lanjut.")
            return
        try:
            uid = us.kirim_usulan(_sheet_usulan(), nama, r["Link"], r["Judul"], perubahan, hapus, alasan)
        except ValueError as e:
            st.error(str(e))
            return
        except Exception as e:
            st.error(f"Usulan gagal dikirim: {e}")
            return
        ss["nama_pengusul"] = nama.strip()
        ss["n_usulan"] = ss.get("n_usulan", 0) + 1
        menunggu_usulan.clear()
        ss["toast"] = f"Usulan {uid} terkirim dan menunggu persetujuan admin. Terima kasih!"
        st.rerun()


def kartu_usulan(df, kunci, per_halaman=15, edit=True):
    """Kartu berita dengan tombol ✏️ untuk mengusulkan perubahan."""
    if df.empty:
        st.info("Tidak ada berita yang cocok.")
        return
    total_hal = max(1, -(-len(df) // per_halaman))
    hal = 1
    if total_hal > 1:
        k = f"pg_{kunci}"
        if ss.get(k, 1) > total_hal:
            ss[k] = 1
        ss.setdefault(k, 1)
        hal = st.number_input(f"Halaman (dari {total_hal})", 1, total_hal, key=k)
    menunggu = menunggu_usulan(str(rahasia("id_sheet_usulan")))
    for i, r in df.iloc[(hal - 1) * per_halaman: hal * per_halaman].iterrows():
        kartu = G["html_kartu"](r).rstrip()
        if sc.kunci_link(r["Link"]) in menunggu and kartu.endswith("</div>"):
            kartu = (kartu[:-len("</div>")] + '<span class="badge b-kosong">⏳ Ada usulan menunggu '
                     'persetujuan admin</span></div>')
        c1, c2 = st.columns([24, 1], vertical_alignment="center")
        c1.markdown(kartu, unsafe_allow_html=True)
        if c2.button("✏️", key=f"us_{kunci}_{i}", help="Usulkan perubahan untuk berita ini"):
            dialog_usulan(r.to_dict())


# ---------------------------------------------------------------------------
# SUSUNAN HALAMAN
# ---------------------------------------------------------------------------

cek_kata_sandi()
SEMUA, INFO, SUMBER = muat_data()

G["gabung"] = lambda files=None: SEMUA.copy()
G["daftar_file"] = lambda: []
G["sumber_data"] = sumber_data_online
G["tabel_edit"] = tabel_baca
G["kartu_berita"] = kartu_baca
USULAN = usulan_aktif()
if USULAN:
    G["kartu_berita"] = kartu_usulan
    G["TEKS_KARTU"] = ("Klik ✏️ di samping kartu untuk mengusulkan koreksi atau penghapusan. "
                       "Usulan berlaku setelah disetujui admin.")

MENU = G["MENU"]
MENU_VIEWER = [MENU[1], MENU[2]]          # Dashboard, Ringkasan pimpinan
if ss.get("menu") not in MENU_VIEWER:
    ss["menu"] = MENU_VIEWER[0]

G["hero"]()
if SUMBER == "lokal":
    st.warning("Mode pratinjau: data dibaca dari folder **hasil** di komputer ini karena Google "
               "Sheet belum diatur di secrets.")
st.radio("Menu", MENU_VIEWER, key="menu", horizontal=True, label_visibility="collapsed")
if "toast" in ss:
    st.toast(ss.pop("toast"), icon="✅")

if ss["menu"] == MENU[2]:
    G["halaman_pimpinan"]()
else:
    G["halaman_dashboard"]()

with st.sidebar:
    st.markdown("### Data berita")
    st.markdown(f"**{len(SEMUA)} berita**  \nTerakhir diperbarui: "
                f"**{INFO.get('Terakhir diperbarui', '-')}**")
    st.caption(f"Data diambil otomatis setiap hari kerja. Dashboard memuat ulang data setiap "
               f"{MENIT_SEGAR} menit.")
    if st.button("🔄 Muat ulang data sekarang"):
        muat_dari_sheet.clear()
        st.rerun()
    st.markdown("### Tentang")
    if USULAN:
        st.caption("Data hanya diubah oleh tim pengelola Direktorat Pembiayaan Syariah. Usulan "
                   "koreksi lewat tombol ✏️ ditinjau dulu sebelum diterapkan.")
    else:
        st.caption("Versi baca-saja. Koreksi atau penghapusan berita dilakukan oleh tim pengelola "
                   "Direktorat Pembiayaan Syariah.")
