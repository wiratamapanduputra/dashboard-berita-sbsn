"""Platform Pemantauan Berita SBSN (Streamlit) — tampilan DJPPR.

Jalankan dari folder scraper-sbsn dengan venv aktif:
    streamlit run app.py
atau klik dua kali jalankan_platform.bat

Logo resmi (opsional): taruh file berikut di subfolder "assets":
    assets/logo_kemenkeu.png   dan   assets/logo_djppr.png

Menu: 1. Ambil berita  2. Dashboard  3. Ringkasan pimpinan  4. Tinjau & kelola
"""

import base64
import html
import re
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path

import altair as alt
import pandas as pd
import plotly.express as px
import streamlit as st

import laporan
import scraper_sbsn_chrome as sc

st.set_page_config(page_title="Pemantauan Berita SBSN", page_icon="📰", layout="wide")
ss = st.session_state

# ---------------------------------------------------------------------------
# NAVIGASI & PENYIMPANAN PILIHAN
# ---------------------------------------------------------------------------

MENU = ["🔎 Ambil berita", "📊 Dashboard", "📋 Ringkasan pimpinan", "✏️ Tinjau & kelola"]
SUB_DASH = ["Ikhtisar", "Daftar berita", "Situs media"]
FILTER_KEYS = ["f_tone", "f_topik", "f_situs", "f_seri", "f_cari", "f_tgl"]

# Pilihan di halaman yang sedang tidak dibuka tetap diingat
for _k in list(ss.keys()):
    if str(_k).startswith(("ab_", "src_", "rng_", "fil_", "f_", "dv_", "pg_")):
        ss[_k] = ss[_k]

# Perpindahan halaman yang diminta dari klik grafik / kartu (diterapkan sebelum widget dibuat)
_nav = ss.pop("nav", None)
if _nav:
    if _nav.pop("_reset_filter", False):
        for _k in FILTER_KEYS:
            ss.pop(_k, None)
    for _k, _v in _nav.items():
        ss[_k] = _v
ss.setdefault("menu", MENU[1] if sc.FOLDER_HASIL.exists() and any(sc.FOLDER_HASIL.glob("*.xlsx"))
              else MENU[0])
ss.setdefault("dash_view", SUB_DASH[0])
ss.setdefault("ver", 0)


def pergi(menu=None, view=None, reset_filter=False, **filter_):
    """Minta pindah halaman; dipakai sebagai callback atau sebelum st.rerun()."""
    nav = {}
    if menu:
        nav["menu"] = menu
    if view:
        nav["dash_view"] = view
    if reset_filter:
        nav["_reset_filter"] = True
    nav.update({f"f_{k}": v for k, v in filter_.items()})
    ss["nav"] = nav
    ss["ver"] = ss.get("ver", 0) + 1   # kosongkan pilihan pada grafik


def ke_daftar(**filter_):
    pergi(menu=MENU[1], view="Daftar berita", reset_filter=True, **filter_)


# ---------------------------------------------------------------------------
# TEMA DJPPR
# ---------------------------------------------------------------------------

UNGU, EMAS, MAGENTA, PLUM = "#6C007D", "#FFC000", "#D0208A", "#7C3F84"
GRADIEN = "linear-gradient(135deg, #6B1F6B 0%, #8B5890 50%, #C0208A 100%)"
TONE = ["Positif", "Netral", "Negatif"]
WARNA_TONE = {"Positif": "#2FA888", "Netral": "#3B82C4", "Negatif": "#E0533D"}   # Netral = biru
PALET_TOPIK = [UNGU, EMAS, MAGENTA, "#002060", PLUM, "#2FA888", "#B4C7E7"]

_CONFIG = Path(__file__).resolve().parent / ".streamlit" / "config.toml"
if not _CONFIG.exists():
    _CONFIG.parent.mkdir(exist_ok=True)
    _CONFIG.write_text(
        '[theme.light]\nprimaryColor = "#6C007D"\nbackgroundColor = "#F5F6F8"\n'
        'secondaryBackgroundColor = "#EDE7F0"\n\n'
        '[theme.dark]\nprimaryColor = "#C77DD6"\n\n'
        '[browser]\ngatherUsageStats = false\n\n'
        '[client]\ntoolbarMode = "viewer"\n',
        encoding="utf-8")

try:
    GELAP = st.context.theme.type == "dark"
except Exception:
    GELAP = False

KARTU_BG = "#26212C" if GELAP else "#FFFFFF"
TEKS = "#EDEDED" if GELAP else "#2A2A2A"
TEKS_PUDAR = "#B9B3BF" if GELAP else "#6B6B6B"
GARIS = "#3A3340" if GELAP else "#E6E1EA"
UNGU_TEKS = "#E7C6EE" if GELAP else UNGU

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&display=swap');
.block-container {{ padding-top: 3.2rem; }}
h1, h2, h3 {{ font-family: 'Oswald', 'Arial Narrow', sans-serif !important; }}

.hero {{ background: {GRADIEN}; border-radius: 14px; padding: 18px 26px;
        border-bottom: 5px solid {EMAS}; color: #fff; display: flex;
        align-items: center; justify-content: space-between; gap: 16px; margin-bottom: 10px; }}
.hero h1 {{ color: #fff !important; margin: 0; font-size: 1.9rem; letter-spacing: .5px; padding: 0; }}
.hero p {{ margin: 2px 0 0; opacity: .9; }}
.hero .logo {{ display: flex; gap: 14px; align-items: center; }}
.hero .logo img {{ height: 44px; background: #fff; border-radius: 6px; padding: 4px; }}
.hero .wordmark {{ font-family: 'Oswald', sans-serif; font-weight: 700; font-size: 1.3rem;
                  border: 2px solid {EMAS}; border-radius: 999px; padding: 2px 14px; }}

/* Menu utama sebagai tab */
.st-key-menu div[role="radiogroup"] {{ gap: 4px; flex-wrap: wrap; border-bottom: 3px solid {EMAS}; }}
.st-key-menu label {{ padding: 9px 18px !important; margin: 0 !important; border-radius: 10px 10px 0 0;
                     cursor: pointer; background: {'#2E2734' if GELAP else '#EDE7F0'}; }}
.st-key-menu label > div:first-child, .st-key-menu [data-testid="stRadioOption"] > div > div:first-child:not([data-testid]) {{ display: none; }}
.st-key-menu label p {{ font-family: 'Oswald', sans-serif; font-size: 1.05rem; color: {TEKS}; }}
.st-key-menu label:has(input:checked) {{ background: {GRADIEN}; }}
.st-key-menu label:has(input:checked) p {{ color: #fff !important; }}

/* Sub-menu dashboard sebagai pil */
.st-key-dash_view div[role="radiogroup"] {{ gap: 6px; }}
.st-key-dash_view label {{ padding: 5px 16px !important; margin: 0 !important; border-radius: 999px;
                          border: 1.5px solid {UNGU_TEKS}; cursor: pointer; }}
.st-key-dash_view label > div:first-child, .st-key-dash_view [data-testid="stRadioOption"] > div > div:first-child:not([data-testid]) {{ display: none; }}
.st-key-dash_view label p {{ color: {UNGU_TEKS}; font-weight: 600; }}
.st-key-dash_view label:has(input:checked) {{ background: {UNGU}; border-color: {UNGU}; }}
.st-key-dash_view label:has(input:checked) p {{ color: #fff !important; }}

.sec {{ font-family: 'Oswald', sans-serif; font-size: 1.2rem; font-weight: 600;
        color: {TEKS}; margin: 14px 0 6px; padding-bottom: 6px; position: relative; }}
.sec::before {{ content: ''; display: inline-block; width: 10px; height: 10px;
               background: {EMAS}; margin-right: 10px; vertical-align: middle; }}
.sec::after {{ content: ''; position: absolute; left: 0; bottom: 0; width: 100%;
              height: 2px; background: linear-gradient(90deg, {MAGENTA}, {PLUM}, transparent); }}

.kpi {{ background: {KARTU_BG}; border-radius: 12px; padding: 14px 18px;
        border-left: 6px solid var(--c); box-shadow: 0 1px 4px rgba(0,0,0,.08);
        transition: transform .12s ease, box-shadow .12s ease; }}
.kpi .lbl {{ color: {TEKS_PUDAR}; font-size: .82rem; text-transform: uppercase; letter-spacing: .5px; }}
.kpi .val {{ font-family: 'Oswald', sans-serif; font-size: 2.1rem; font-weight: 700;
             color: {TEKS}; line-height: 1.1; }}
.kpi .dlt {{ font-size: .82rem; color: {TEKS_PUDAR}; }}
.kpi .naik {{ color: #2FA888; }} .kpi .turun {{ color: #E0533D; }}
.kpi .lihat {{ font-size: .78rem; color: {MAGENTA}; font-weight: 600; }}

/* Kartu KPI yang bisa diklik: tombol transparan menutupi kartu */
div[class*="st-key-kpi_"] {{ position: relative; gap: 0; }}
div[class*="st-key-kpi_"] [data-testid="stElementContainer"]:has(.stButton) {{
    position: absolute; inset: 0; z-index: 5; }}
div[class*="st-key-kpi_"] .stButton, div[class*="st-key-kpi_"] .stButton > div,
div[class*="st-key-kpi_"] .stButton button {{ width: 100%; height: 100%; }}
div[class*="st-key-kpi_"] .stButton button {{ opacity: 0; cursor: pointer; }}
div[class*="st-key-kpi_"]:hover .kpi {{ transform: translateY(-3px);
    box-shadow: 0 8px 18px rgba(108,0,125,.22); }}

.card {{ background: {KARTU_BG}; border-radius: 12px; padding: 12px 18px; margin-bottom: 2px;
         border-left: 6px solid var(--c); box-shadow: 0 1px 4px rgba(0,0,0,.07); }}
.card a.jdl {{ font-weight: 700; font-size: 1.03rem; color: {TEKS}; text-decoration: none; }}
.card a.jdl:hover {{ color: {MAGENTA}; text-decoration: underline; }}
.card .meta {{ color: {TEKS_PUDAR}; font-size: .84rem; margin: 3px 0 8px; }}
.badge {{ display: inline-block; font-size: .76rem; border-radius: 999px;
          padding: 2px 10px; margin: 0 6px 2px 0; }}
.badge b {{ font-weight: 700; }}
.b-tone {{ color: #fff; background: var(--c); }}
.b-topik {{ color: {UNGU_TEKS}; border: 1.5px solid {UNGU_TEKS}; }}
.b-seri {{ color: #2A2A2A; background: {EMAS}; }}
.b-kosong {{ color: {TEKS_PUDAR}; border: 1.5px dashed {GARIS}; }}

.chip {{ display: inline-block; background: {'#3A2F40' if GELAP else '#F1E6F4'}; color: {UNGU_TEKS};
         border-radius: 999px; padding: 3px 12px; margin: 0 6px 6px 0; font-size: .85rem; }}

.awan {{ background: {KARTU_BG}; border-radius: 12px; padding: 18px; line-height: 2.2;
         text-align: center; box-shadow: 0 1px 4px rgba(0,0,0,.07); }}
.awan span {{ margin: 0 8px; font-family: 'Oswald', sans-serif; white-space: nowrap; }}

.stButton button[kind="primary"], .stDownloadButton button {{
    background: {GRADIEN}; border: none; color: #fff; }}
section[data-testid="stSidebar"] {{ border-right: 4px solid {EMAS}; }}
</style>
""", unsafe_allow_html=True)


def sec(judul):
    st.markdown(f'<div class="sec">{html.escape(judul)}</div>', unsafe_allow_html=True)


def _logo(nama):
    p = sc.FOLDER / "assets" / nama
    if p.exists():
        return f'<img src="data:image/png;base64,{base64.b64encode(p.read_bytes()).decode()}">'
    return ""


def hero():
    logo = _logo("logo_kemenkeu.png") + _logo("logo_djppr.png")
    kanan = logo or '<span class="wordmark">DJPPR</span>'
    st.markdown(f"""
    <div class="hero">
      <div><h1>PEMANTAUAN BERITA SBSN</h1>
      <p>Direktorat Pembiayaan Syariah · DJPPR Kementerian Keuangan</p></div>
      <div class="logo">{kanan}</div>
    </div>""", unsafe_allow_html=True)


def kpi_html(label, nilai, warna, delta=None, kebalikan=False, lihat=""):
    dlt = ""
    if delta is not None:
        if delta == 0:
            dlt = '<div class="dlt">sama dengan periode lalu</div>'
        else:
            baik = (delta > 0) != kebalikan
            dlt = (f'<div class="dlt"><span class="{"naik" if baik else "turun"}">'
                   f'{"▲" if delta > 0 else "▼"} {abs(delta)}</span> vs periode lalu</div>')
    if lihat:
        dlt += f'<div class="lihat">{html.escape(lihat)} →</div>'
    return (f'<div class="kpi" style="--c:{warna}"><div class="lbl">{html.escape(label)}</div>'
            f'<div class="val">{nilai}</div>{dlt}</div>')


def kpi(kolom, *a, **k):
    kolom.markdown(kpi_html(*a, **k), unsafe_allow_html=True)


def kpi_klik(kolom, kunci, label, nilai, warna, aksi, args=(), kwargs=None, lihat="Lihat berita"):
    """Kartu KPI yang bisa diklik (tombol transparan menutupi kartu)."""
    with kolom.container(key=f"kpi_{kunci}"):
        st.markdown(kpi_html(label, nilai, warna, lihat=lihat), unsafe_allow_html=True)
        st.button(f"{lihat}: {label}", key=f"btn_kpi_{kunci}", on_click=aksi, args=args,
                  kwargs=kwargs or {})


def gaya_plotly(fig, tinggi=320):
    fig.update_layout(
        height=tinggi, margin=dict(l=10, r=10, t=20, b=10),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Arial", color=TEKS), legend_title_text="",
        legend=dict(orientation="h", y=-0.18), hoverlabel=dict(font_family="Arial"))
    fig.update_xaxes(gridcolor=GARIS, title=None)
    fig.update_yaxes(gridcolor=GARIS, title=None)
    return fig


def gaya_alt(ch, tinggi=300):
    return (ch.properties(height=tinggi, background="transparent")
            .configure_view(stroke=None)
            .configure_axis(labelColor=TEKS_PUDAR, titleColor=TEKS_PUDAR, gridColor=GARIS,
                            domainColor=GARIS, tickColor=GARIS, labelFont="Arial", titleFont="Arial",
                            labelFontSize=11)
            .configure_legend(labelColor=TEKS, titleColor=TEKS, orient="bottom", labelFont="Arial",
                              titleFont="Arial", labelFontSize=11, symbolType="circle")
            .configure_text(font="Arial"))


# ---------------------------------------------------------------------------
# DATA
# ---------------------------------------------------------------------------

def daftar_file():
    sc.FOLDER_HASIL.mkdir(exist_ok=True)
    return sorted(sc.FOLDER_HASIL.glob("berita_sbsn_*.xlsx"),
                  key=lambda p: p.stat().st_mtime, reverse=True)


@st.cache_data(show_spinner=False)
def _baca(path_str, mtime):
    df = pd.read_excel(path_str)
    for k in sc.KOLOM:
        if k not in df.columns:
            df[k] = ""
    df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
    for k in ["Situs", "Judul", "Link", "Topik", "Bulan", "Tone", "Ritel"]:
        df[k] = df[k].fillna("").astype(str)
    return df[sc.KOLOM]


def baca_excel(path):
    return _baca(str(path), path.stat().st_mtime).copy()


def gabung(files):
    """Gabungkan beberapa file; file terbaru menang bila berita sama (hasil koreksi)."""
    if not files:
        return pd.DataFrame(columns=sc.KOLOM)
    df = pd.concat([baca_excel(p) for p in files], ignore_index=True)
    df = sc.buang_duplikat(df, ss.get("gabung_judul_sama", sc.GABUNG_JUDUL_SAMA))
    df = df.sort_values("Tanggal", ascending=True, na_position="last", kind="stable").reset_index(drop=True)
    df["No"] = range(1, len(df) + 1)
    return df


def terapkan_perubahan(edit=None, hapus=()):
    """Simpan koreksi/penghapusan ke SEMUA file hasil yang memuat link tersebut."""
    ubah = {}
    if edit is not None and len(edit):
        for r in edit.to_dict("records"):
            ubah[sc.kunci_link(r["Link"])] = r
    hapus_k = {sc.kunci_link(x) for x in hapus}
    n_file = 0
    for path in daftar_file():
        df = baca_excel(path)
        k = df["Link"].map(sc.kunci_link)
        berubah = False
        if hapus_k and k.isin(hapus_k).any():
            df, k = df[~k.isin(hapus_k)], k[~k.isin(hapus_k)]
            berubah = True
        for i, kk in k.items():
            if kk in ubah:
                for kol in ["Judul", "Situs", "Tanggal", "Topik", "Tone", "Ritel"]:
                    if kol in ubah[kk]:
                        df.at[i, kol] = ubah[kk][kol]
                berubah = True
        if berubah:
            df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
            df["Ritel"] = df["Ritel"].fillna("").astype(str).map(lambda v: ", ".join(sc.daftar_seri(v)))
            df["Bulan"] = df["Tanggal"].apply(sc.nama_bulan)
            df = df.sort_values("Tanggal", ascending=True, na_position="last", kind="stable").reset_index(drop=True)
            path.write_bytes(sc.excel_bytes(df))
            n_file += 1
    return n_file


def daftar_seri_df(df):
    return sorted({x for v in df["Ritel"] for x in v.split(", ") if x})


# ---------------------------------------------------------------------------
# PERIODE & SUMBER DATA
# ---------------------------------------------------------------------------

def periode_ambil():
    if "periode_ambil" not in ss:
        ss["periode_ambil"] = sc.rentang_bulan(date.today().strftime("%Y-%m"))
    return ss["periode_ambil"]


def bulan_penuh(dari, sampai):
    return dari.day == 1 and sampai == sc.rentang_bulan(dari.strftime("%Y-%m"))[1]


def nama_periode(dari, sampai):
    if dari is None:
        return "-"
    if bulan_penuh(dari, sampai):
        return f"{sc.NAMA_BULAN[dari.month - 1]} {dari.year}"
    if dari == sampai:
        return laporan.tgl_id(pd.Timestamp(dari))
    return f"{laporan.tgl_id(pd.Timestamp(dari))} – {laporan.tgl_id(pd.Timestamp(sampai))}"


def periode_lalu(dari, sampai):
    if bulan_penuh(dari, sampai):
        return sc.rentang_bulan((dari - timedelta(days=1)).strftime("%Y-%m"))
    panjang = (sampai - dari).days + 1
    return dari - timedelta(days=panjang), dari - timedelta(days=1)


def saring_periode(df, dari, sampai):
    t = df["Tanggal"].dt.date
    return df[(t >= dari) & (t <= sampai)]


def sumber_data(kunci):
    """Pemilih sumber data per halaman. Bawaan: periode dari menu Ambil berita."""
    files = daftar_file()
    p_dari, p_sampai = periode_ambil()
    opsi = [f"Periode Ambil berita ({nama_periode(p_dari, p_sampai)})", "Rentang tanggal lain",
            "Pilih file hasil", "Semua data"]
    ss.setdefault(f"src_{kunci}", 0)
    with st.container(border=True):
        c1, c2 = st.columns([2, 3])
        mode = c1.selectbox("Sumber data", range(4), format_func=lambda i: opsi[i], key=f"src_{kunci}")
        semua = gabung(files)
        if mode == 0:
            dari, sampai = p_dari, p_sampai
            df = saring_periode(semua, dari, sampai)
            c2.caption("Mengikuti periode yang dipilih di menu **🔎 Ambil berita**. "
                       "Ganti sumber data di sebelah kiri bila perlu.")
        elif mode == 1:
            ss.setdefault(f"rng_{kunci}", (p_dari, p_sampai))
            r = c2.date_input("Rentang tanggal", key=f"rng_{kunci}", format="DD/MM/YYYY")
            dari, sampai = (r if len(r) == 2 else (r[0], r[0]))
            df = saring_periode(semua, dari, sampai)
        elif mode == 2:
            nama = [p.name for p in files]
            ss.setdefault(f"fil_{kunci}", nama[:1])
            ss[f"fil_{kunci}"] = [n for n in ss[f"fil_{kunci}"] if n in nama]
            pilih = c2.multiselect("File", nama, key=f"fil_{kunci}", placeholder="Pilih file")
            df = gabung([sc.FOLDER_HASIL / n for n in pilih])
            dari = sampai = None
        else:
            df = semua
            dari = sampai = None
        if dari is None:
            t = df["Tanggal"].dropna()
            dari, sampai = (t.min().date(), t.max().date()) if len(t) else (None, None)
    df = df.reset_index(drop=True)
    df["No"] = range(1, len(df) + 1)
    return df, dari, sampai, mode


# ---------------------------------------------------------------------------
# GRAFIK
# ---------------------------------------------------------------------------

def donut_alt(df, kolom, warna_map, nama_sel, klik=True, tinggi=290):
    hitung = df[kolom].replace("", "(kosong)").value_counts().reset_index()
    hitung.columns = [kolom, "Jumlah"]
    total = int(hitung["Jumlah"].sum())
    hitung["Persen"] = hitung["Jumlah"] / max(total, 1)
    urutan = list(warna_map) if isinstance(warna_map, dict) else list(hitung[kolom])
    warna = ([warna_map.get(x, "#ABAAAA") for x in urutan] if isinstance(warna_map, dict)
             else [PALET_TOPIK[i % len(PALET_TOPIK)] for i in range(len(urutan))])
    hitung = hitung[hitung["Jumlah"] > 0]
    dasar = alt.Chart(hitung).encode(
        theta=alt.Theta("Jumlah:Q", stack=True),
        color=alt.Color(f"{kolom}:N", scale=alt.Scale(domain=urutan, range=warna),
                        legend=alt.Legend(title=None, columns=2)),
        order=alt.Order("Jumlah:Q", sort="descending"),
        tooltip=[alt.Tooltip(f"{kolom}:N"), alt.Tooltip("Jumlah:Q"),
                 alt.Tooltip("Persen:Q", format=".0%")])
    busur = dasar.mark_arc(innerRadius=62, outerRadius=100, stroke=KARTU_BG, strokeWidth=2,
                           cursor="pointer" if klik else "default")
    if klik:
        sel = alt.selection_point(fields=[kolom], name=nama_sel)
        busur = busur.add_params(sel)
    label = dasar.mark_text(radius=116, fontSize=11, color=TEKS, fontWeight="bold").encode(
        text=alt.Text("Persen:Q", format=".0%"))
    tengah = alt.Chart(pd.DataFrame({"t": [f"{total}"]})).mark_text(
        fontSize=26, fontWeight="bold", color=TEKS, font="Arial", dy=-6).encode(text="t:N")
    bawah = alt.Chart(pd.DataFrame({"t": ["berita"]})).mark_text(
        fontSize=12, color=TEKS_PUDAR, dy=16).encode(text="t:N")
    return gaya_alt(alt.layer(busur, label, tengah, bawah).properties(
        padding={"top": 18, "bottom": 4, "left": 4, "right": 4}), tinggi)


def pilihan_alt(ev, nama, kolom):
    """Ambil nilai yang diklik dari hasil on_select Altair."""
    try:
        sel = ev.selection.get(nama) or []
        if isinstance(sel, dict):
            sel = [sel]
        return [s[kolom] for s in sel if kolom in s]
    except Exception:
        return []


def grafik_tren(df):
    d = df.dropna(subset=["Tanggal"])
    if d.empty:
        return None
    harian = d.groupby([d["Tanggal"].dt.date, "Tone"]).size().rename("Jumlah").reset_index()
    fig = px.bar(harian, x="Tanggal", y="Jumlah", color="Tone",
                 category_orders={"Tone": TONE}, color_discrete_map=WARNA_TONE)
    fig.update_traces(hovertemplate="%{x|%d/%m/%Y}<br>%{y} berita<extra></extra>")
    fig.update_layout(barmode="stack", clickmode="event+select", bargap=.15)
    fig.update_xaxes(tickformat="%d/%m")
    return gaya_plotly(fig)


def bar_tone_topik(df):
    t = df.groupby(["Topik", "Tone"]).size().rename("Jumlah").reset_index()
    sel = alt.selection_point(fields=["Topik", "Tone"], name="pilih_tt")
    ch = alt.Chart(t).mark_bar(cursor="pointer", cornerRadiusEnd=3).encode(
        y=alt.Y("Topik:N", title=None, sort="-x"),
        x=alt.X("Jumlah:Q", title=None, stack=True, axis=alt.Axis(tickMinStep=1, format="d")),
        color=alt.Color("Tone:N", scale=alt.Scale(domain=TONE, range=[WARNA_TONE[x] for x in TONE]),
                        legend=alt.Legend(title=None)),
        order=alt.Order("urut:Q"),
        tooltip=["Topik", "Tone", "Jumlah"]).transform_calculate(
        urut="indexof(['Positif','Netral','Negatif'], datum.Tone)").add_params(sel)
    return gaya_alt(ch, 260)


def bar_situs(df, n=8):
    s = df["Situs"].replace("", "(tanpa nama)").value_counts().head(n).reset_index()
    s.columns = ["Situs", "Jumlah"]
    sel = alt.selection_point(fields=["Situs"], name="pilih_situs")
    ch = alt.Chart(s).mark_bar(cursor="pointer", cornerRadiusEnd=4, color=UNGU).encode(
        y=alt.Y("Situs:N", sort="-x", title=None),
        x=alt.X("Jumlah:Q", title=None, axis=alt.Axis(tickMinStep=1, format="d")),
        tooltip=["Situs", "Jumlah"]).add_params(sel)
    return gaya_alt(ch, 290)


def rekap_situs(df):
    d = df.assign(Situs=df["Situs"].replace("", "(tanpa nama)"))
    t = pd.crosstab(d["Situs"], d["Tone"]).reindex(columns=TONE, fill_value=0)
    t["Jumlah berita"] = t.sum(axis=1)
    t["Indeks sentimen"] = ((t["Positif"] - t["Negatif"]) / t["Jumlah berita"] * 100).round().astype(int)
    t = t.sort_values(["Jumlah berita", "Positif"], ascending=False).reset_index()
    t.insert(0, "No", range(1, len(t) + 1))
    return t[["No", "Situs", "Jumlah berita", "Positif", "Netral", "Negatif", "Indeks sentimen"]]


def grafik_4_situs(rekap):
    top = rekap.head(4).melt(id_vars=["Situs"], value_vars=TONE, var_name="Tone", value_name="Jumlah")
    sel = alt.selection_point(fields=["Situs", "Tone"], name="pilih_4")
    urutan = list(rekap.head(4)["Situs"])
    batang = alt.Chart(top).mark_bar(cursor="pointer", cornerRadiusEnd=3).encode(
        x=alt.X("Situs:N", sort=urutan, title=None, axis=alt.Axis(labelAngle=0, labelFontSize=12)),
        xOffset=alt.XOffset("Tone:N", sort=TONE),
        y=alt.Y("Jumlah:Q", title="Jumlah berita", axis=alt.Axis(tickMinStep=1, format="d")),
        color=alt.Color("Tone:N", scale=alt.Scale(domain=TONE, range=[WARNA_TONE[x] for x in TONE]),
                        legend=alt.Legend(title=None)),
        tooltip=["Situs", "Tone", "Jumlah"]).add_params(sel)
    teks = alt.Chart(top).mark_text(dy=-8, fontSize=11, color=TEKS).encode(
        x=alt.X("Situs:N", sort=urutan), xOffset=alt.XOffset("Tone:N", sort=TONE),
        y="Jumlah:Q", text="Jumlah:Q")
    return gaya_alt(alt.layer(batang, teks), 340)


def kata_teratas(judul, n=25):
    return laporan.kata_teratas(judul, n)


def awan_kata(df):
    kata = kata_teratas(df["Judul"])
    if not kata:
        st.write("Belum ada data.")
        return
    maks = kata[0][1]
    warna = [UNGU, MAGENTA, PLUM, "#B8860B", "#002060"]
    isi = "".join(
        f'<span style="font-size:{0.9 + 1.6 * n / maks:.2f}rem;'
        f'color:{"#E7C6EE" if GELAP and i % 5 == 0 else warna[i % 5]}" '
        f'title="{n} judul">{html.escape(w)}</span>'
        for i, (w, n) in enumerate(sorted(kata, key=lambda x: x[0])))
    st.markdown(f'<div class="awan">{isi}</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# DAFTAR BERITA: KARTU, TABEL, EDIT
# ---------------------------------------------------------------------------

def tgl_teks(t):
    return f"{t.day} {sc.NAMA_BULAN[t.month - 1][:3]} {t.year}" if pd.notna(t) else "tanggal tidak terbaca"


def html_kartu(r):
    warna = WARNA_TONE.get(r["Tone"], WARNA_TONE["Netral"])
    seri = [s for s in str(r["Ritel"]).split(", ") if s]
    ritel = (f'<span class="badge b-seri">Ritel: <b>{html.escape(", ".join(seri))}</b></span>' if seri
             else '<span class="badge b-kosong">Ritel: –</span>')
    return f"""
    <div class="card" style="--c:{warna}">
      <a class="jdl" href="{html.escape(r['Link'])}" target="_blank">{html.escape(r['Judul'])}</a>
      <div class="meta">{html.escape(r['Situs'] or '—')} · {tgl_teks(r['Tanggal'])}</div>
      <span class="badge b-tone" style="--c:{warna}">Tone: <b>{html.escape(r['Tone'])}</b></span>
      <span class="badge b-topik">Kata kunci: <b>{html.escape(r['Topik'])}</b></span>{ritel}
    </div>"""


@st.dialog("✏️ Edit berita", width="large")
def dialog_edit(r, opsi_topik):
    st.caption(r["Link"])
    judul = st.text_input("Judul", r["Judul"])
    c1, c2 = st.columns(2)
    situs = c1.text_input("Situs media", r["Situs"])
    tgl = c2.date_input("Tanggal", r["Tanggal"].date() if pd.notna(r["Tanggal"]) else None,
                        format="DD/MM/YYYY")
    c3, c4, c5 = st.columns(3)
    tone = c3.selectbox("Tone", TONE, index=TONE.index(r["Tone"]) if r["Tone"] in TONE else 1)
    opsi = sorted(set(opsi_topik) | {r["Topik"]})
    topik = c4.selectbox("Kata kunci", opsi, index=opsi.index(r["Topik"]), accept_new_options=True)
    ritel = c5.text_input("Ritel (seri)", r["Ritel"], placeholder="mis. SR025, ST015",
                          help="Pisahkan dengan koma bila lebih dari satu.")
    if st.button("💾 Simpan", type="primary"):
        terapkan_perubahan(pd.DataFrame([{
            "Link": r["Link"], "Judul": judul, "Situs": situs,
            "Tanggal": pd.Timestamp(tgl) if tgl else pd.NaT, "Topik": topik or r["Topik"],
            "Tone": tone, "Ritel": ", ".join(sc.daftar_seri(ritel))}]))
        ss["toast"] = "Perubahan berita disimpan"
        st.rerun()
    st.divider()
    yakin = st.checkbox("Saya ingin menghapus berita ini dari data hasil")
    if st.button("🗑️ Hapus berita", disabled=not yakin):
        terapkan_perubahan(hapus=[r["Link"]])
        ss["toast"] = "Berita dihapus"
        st.rerun()


def kartu_berita(df, kunci, per_halaman=15, edit=True):
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
    potong = df.iloc[(hal - 1) * per_halaman: hal * per_halaman]
    opsi_topik = sorted(set(df["Topik"]) | set(sc.KATA_KUNCI.values()))
    for i, r in potong.iterrows():
        if edit:
            c1, c2 = st.columns([24, 1], vertical_alignment="center")
            c1.markdown(html_kartu(r), unsafe_allow_html=True)
            if c2.button("✏️", key=f"ed_{kunci}_{i}", help="Edit berita ini"):
                dialog_edit(r.to_dict(), opsi_topik)
        else:
            st.markdown(html_kartu(r), unsafe_allow_html=True)


def tabel_edit(df, kunci):
    st.caption("Klik sel untuk mengubah. Centang **Hapus** untuk membuang berita. "
               "Lalu klik **Simpan perubahan tabel**.")
    asli = df[sc.KOLOM].reset_index(drop=True)
    edit = st.data_editor(
        asli.assign(Hapus=False)[["Hapus"] + sc.KOLOM], hide_index=True, width="stretch",
        num_rows="fixed", disabled=["No", "Link", "Bulan"], key=f"editor_{kunci}_{ss['ver']}",
        column_config={
            "Hapus": st.column_config.CheckboxColumn("Hapus", width="small"),
            "Link": st.column_config.LinkColumn("Link", display_text="buka"),
            "Tanggal": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Judul": st.column_config.TextColumn(width="large"),
            "Tone": st.column_config.SelectboxColumn(options=TONE, required=True),
            "Topik": st.column_config.SelectboxColumn(
                "Kata kunci", options=sorted(set(asli["Topik"]) | set(sc.KATA_KUNCI.values()))),
        })
    edit["Hapus"] = edit["Hapus"].fillna(False).astype(bool)
    kolom = ["Judul", "Situs", "Tanggal", "Topik", "Tone", "Ritel"]
    def _baku(x):
        x = x[kolom].copy()
        x["Tanggal"] = pd.to_datetime(x["Tanggal"], errors="coerce").dt.strftime("%Y-%m-%d")
        return x.fillna("").astype(str)
    a, b = _baku(asli), _baku(edit)
    berubah = (a != b).any(axis=1) & ~edit["Hapus"]
    n_ubah, n_hapus = int(berubah.sum()), int(edit["Hapus"].sum())
    label = "💾 Simpan perubahan tabel"
    if n_ubah or n_hapus:
        label += f" ({n_ubah} diubah, {n_hapus} dihapus)"
    if st.button(label, type="primary", disabled=not (n_ubah or n_hapus), key=f"simpan_{kunci}"):
        e = edit[berubah].copy()
        e["Tanggal"] = pd.to_datetime(e["Tanggal"], errors="coerce")
        terapkan_perubahan(e, hapus=edit.loc[edit["Hapus"], "Link"].tolist())
        ss["toast"] = f"{n_ubah} berita diubah, {n_hapus} dihapus"
        ss["ver"] += 1
        st.rerun()


def csv_bytes(df):
    """CSV untuk Excel Indonesia: pemisah titik koma, UTF-8 dengan BOM, tanggal DD/MM/YYYY."""
    d = sc.urutkan_tanggal(df[[k for k in sc.KOLOM if k in df.columns]])
    d["No"] = range(1, len(d) + 1)
    d["Tanggal"] = pd.to_datetime(d["Tanggal"], errors="coerce").dt.strftime("%d/%m/%Y")
    return d.to_csv(index=False, sep=";").encode("utf-8-sig")


def panel_laporan(df, periode, df_lalu=None, nama_lalu=None, kunci="lap", catatan="",
                  nama_file="laporan_berita_sbsn", rentang=None):
    a1, a2, a3 = st.columns([2, 1.3, 1.5])
    judul = a1.text_input("Judul laporan", "Laporan Pemantauan Berita SBSN", key=f"judul_{kunci}")
    fmt = a2.pills("Format", ["PDF", "Word"], selection_mode="multi",
                   default=["PDF", "Word"], key=f"fmt_{kunci}") or []
    lamp = a3.checkbox("Sertakan lampiran daftar berita", True, key=f"lamp_{kunci}")
    tanda = (periode, judul, lamp, catatan, len(df), tuple(df["Link"].head(50)))
    if st.button("Buat laporan", type="primary", key=f"buat_{kunci}", disabled=not fmt):
        hasil = {}
        with st.spinner("Menyusun laporan…"):
            arg = dict(df=df, periode=periode, df_lalu=df_lalu, nama_lalu=nama_lalu,
                       judul=judul, catatan=catatan, lampiran=lamp, rentang=rentang)
            if "PDF" in fmt:
                hasil["PDF"] = laporan.buat_laporan_pdf(**arg)
            if "Word" in fmt:
                hasil["Word"] = laporan.buat_laporan_docx(**arg)
        ss[f"laporan_{kunci}"] = (tanda, hasil)
    simpan = ss.get(f"laporan_{kunci}")
    b1, b2, b3, _ = st.columns([1, 1, 1, 1.5])
    if simpan and simpan[0] == tanda:
        if "PDF" in simpan[1]:
            b1.download_button("⬇️ Unduh PDF", simpan[1]["PDF"], file_name=f"{nama_file}.pdf",
                               mime="application/pdf", key=f"unduh_pdf_{kunci}")
        if "Word" in simpan[1]:
            b2.download_button(
                "⬇️ Unduh Word", simpan[1]["Word"], file_name=f"{nama_file}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"unduh_docx_{kunci}")
    elif simpan:
        st.caption("Pengaturan atau data berubah. Klik **Buat laporan** lagi.")
    b3.download_button("⬇️ Data (CSV)", csv_bytes(df), file_name=f"{nama_file}_data.csv",
                       mime="text/csv", key=f"csv_{kunci}")


# ===========================================================================
# HALAMAN 1: AMBIL BERITA
# ===========================================================================

def halaman_ambil():
    sec("Pengaturan pencarian")
    c1, c2 = st.columns(2)
    hari_ini = date.today()
    with c1:
        ss.setdefault("ab_mode", "Per bulan")
        mode = st.segmented_control("Periode", ["Per bulan", "Rentang tanggal"], key="ab_mode") \
            or "Per bulan"
        if mode == "Per bulan":
            pilihan = [f"{t}-{b:02d}" for t in range(hari_ini.year, hari_ini.year - 3, -1)
                       for b in range(12, 0, -1) if date(t, b, 1) <= hari_ini]
            ss.setdefault("ab_bulan", pilihan[0])
            bulan_str = st.selectbox("Bulan", pilihan, key="ab_bulan",
                                     format_func=lambda s: f"{sc.NAMA_BULAN[int(s[5:]) - 1]} {s[:4]}")
            dari, sampai = sc.rentang_bulan(bulan_str)
        else:
            ss.setdefault("ab_rentang", (hari_ini.replace(day=1), hari_ini))
            r = st.date_input("Dari – sampai", key="ab_rentang", format="DD/MM/YYYY")
            dari, sampai = (r if len(r) == 2 else (r[0], r[0]))
        ss.setdefault("ab_kw", "\n".join(sc.KATA_KUNCI))
        teks_kw = st.text_area("Kata kunci (satu per baris)", key="ab_kw", height=110)
        kata_kunci = {x.strip(): x.strip() for x in teks_kw.splitlines() if x.strip()}
    with c2:
        ss.setdefault("ab_seri", "")
        seri = sc.daftar_seri(st.text_input(
            "Seri SBSN Ritel (opsional)", key="ab_seri", placeholder="mis. SR025, ST015",
            help="Boleh lebih dari satu, pisahkan dengan koma atau spasi."))
        if seri:
            st.markdown(" ".join(f'<span class="badge b-seri">{html.escape(s)}</span>' for s in seri),
                        unsafe_allow_html=True)
        ss.setdefault("ab_hanya", False)
        hanya_seri = st.checkbox("Hanya simpan berita seri tersebut", key="ab_hanya", disabled=not seri)
        ss.setdefault("ab_label_ketat", not sc.LABEL_DARI_PENCARIAN)
        label_ketat = st.checkbox(
            "Beri label seri hanya bila kode seri tertulis di judul/cuplikan", key="ab_label_ketat",
            disabled=not seri,
            help="Bila tidak dicentang, semua berita yang muncul di pencarian khusus seri (misalnya "
                 "'Sukuk Ritel SR025') diberi label seri itu, termasuk berita umum seperti jadwal SBN "
                 "2026 yang muncul di pencarian setiap seri.")
        st.info("Periode ini juga menjadi **periode bawaan Dashboard dan Ringkasan pimpinan**.\n\n"
                "Saat proses berjalan, jendela Chrome akan terbuka sendiri. Jangan ditutup. "
                "Jika muncul CAPTCHA, selesaikan di jendela itu; proses lanjut otomatis.")
    ss["periode_ambil"] = (dari, sampai)

    with st.expander("🌐 Situs media tujuan (opsional)", expanded=bool(ss.get("ab_situs"))):
        st.caption("Kosongkan untuk mengambil dari semua situs yang muncul di Google Berita. "
                   "Situs yang diisi akan dicari secara khusus (site:…) **sebagai tambahan**, "
                   "sehingga berita dari situs tersebut lebih lengkap.")
        s1, s2 = st.columns(2)
        ss.setdefault("ab_situs", "\n".join(sc.SITUS_TUJUAN))
        situs = sc.daftar_situs(s1.text_area(
            "Situs tujuan (satu per baris)", key="ab_situs", height=150,
            placeholder="kontan.co.id\ncnbcindonesia.com\nbisnis.com\nantaranews.com"))
        ss.setdefault("ab_kecuali", "\n".join(sc.SITUS_DIKECUALIKAN))
        kecuali = sc.daftar_situs(s2.text_area(
            "Situs dikecualikan (satu per baris)", key="ab_kecuali", height=150,
            placeholder="contoh-agregator.com"))
        ss.setdefault("ab_hanya_situs", False)
        hanya_situs = st.checkbox("Hanya ambil berita dari situs tujuan", key="ab_hanya_situs",
                                  disabled=not situs)
        n_q, n_hal, n_rss, menit = sc.perkiraan(len(kata_kunci), len(seri), len(situs), hanya_situs)
        info = (f"{len(situs)} situs tujuan · {len(kecuali)} dikecualikan → **{n_q} pencarian Google** "
                f"(±{n_hal} halaman, lewat Chrome)"
                + (f" + **{n_rss} permintaan RSS** (ringan, tanpa Chrome)" if n_rss else "")
                + f" · perkiraan **±{max(menit, 1)} menit**.")
        if n_rss and not n_q:
            info += " Chrome tidak perlu dibuka."
        if n_hal > 60:
            st.warning(info + " Jumlah ini besar: Google kemungkinan meminta CAPTCHA berkali-kali. "
                       "Kurangi situs atau kata kunci, atau bagi menjadi beberapa kali pengambilan.")
        else:
            st.info(info)

    st.markdown(f"**Ringkasan:** {nama_periode(dari, sampai)} ({dari:%d/%m/%Y} – {sampai:%d/%m/%Y}) · "
                f"{len(kata_kunci)} kata kunci" + (f" · seri {', '.join(seri)}" if seri else "")
                + (f" · {len(situs)} situs tujuan" + (" (hanya situs ini)" if hanya_situs else "")
                   if situs else ""))
    ss.setdefault("ab_timpa", False)
    timpa = st.checkbox("Timpa file periode yang sama (bawaan: hasil baru ditambahkan ke file lama)",
                        key="ab_timpa",
                        help="Bila tidak dicentang, pengambilan berulang untuk periode yang sama akan "
                             "menambah berita yang belum ada, tanpa menghapus berita dan koreksi lama.")
    b1, b2, _ = st.columns([1.2, 1.2, 3])
    mulai = b1.button("🚀 Mulai ambil berita", type="primary", disabled=not kata_kunci)
    b2.button("📊 Lihat data periode ini", on_click=pergi,
              kwargs=dict(menu=MENU[1], view=SUB_DASH[0]))

    if mulai:
        with st.status("Mengambil berita…", expanded=True) as status:
            sc.LABEL_DARI_PENCARIAN = not label_ketat
            try:
                hasil = sc.jalankan_scraping(
                    dari, sampai, seri, hanya_seri, kata_kunci, log=st.write,
                    gabung_judul_sama=ss.get("gabung_judul_sama", sc.GABUNG_JUDUL_SAMA),
                    situs=situs, kecualikan=kecuali, hanya_situs=hanya_situs)
            except Exception as e:
                status.update(label=f"Gagal: {e}", state="error")
                st.stop()
            peringatan = hasil.attrs.get("peringatan", [])
            if hasil.empty:
                status.update(label="Tidak ada berita yang terambil.", state="error")
                for pesan in peringatan:
                    st.error(pesan)
                st.stop()
            path, n_baru, n_total = sc.simpan_excel(
                hasil, dari, sampai, seri, gabung=not timpa,
                gabung_judul_sama=ss.get("gabung_judul_sama", sc.GABUNG_JUDUL_SAMA))
            ss["file_aktif"] = path.name
            status.update(label=(f"Selesai: {len(hasil)} berita terambil · {n_baru} berita baru · "
                                 f"total {n_total} berita di {path.name}"),
                          state="complete", expanded=False)
        for pesan in peringatan:
            st.warning(f"⚠️ {pesan}")
        if not peringatan:
            st.balloons()
        st.success(f"{len(hasil)} berita terambil pada pengambilan ini; **{n_baru} di antaranya baru**. "
                   f"File {path.name} kini berisi **{n_total} berita**. Klik **📊 Lihat data periode ini** "
                   "untuk membuka Dashboard.")
        kartu_berita(hasil.assign(Tanggal=pd.to_datetime(hasil["Tanggal"])), "hasil_baru", edit=False)


# ===========================================================================
# HALAMAN 2: DASHBOARD (Ikhtisar · Daftar berita · Situs media)
# ===========================================================================

def dash_ikhtisar(df):
    hitung = df["Tone"].value_counts()
    k = st.columns(5)
    kpi_klik(k[0], "total", "Total berita", len(df), UNGU, ke_daftar, lihat="Lihat semua")
    for kol, t in zip(k[1:4], TONE):
        kpi_klik(kol, t.lower(), t, int(hitung.get(t, 0)), WARNA_TONE[t], ke_daftar,
                 kwargs={"tone": [t]})
    kpi_klik(k[4], "situs", "Situs media", df["Situs"].replace("", pd.NA).nunique(), EMAS,
             pergi, kwargs={"menu": MENU[1], "view": "Situs media"}, lihat="Lihat rincian")
    st.caption("💡 Klik kartu atau bagian grafik untuk membuka daftar berita yang sesuai.")

    v = ss["ver"]
    g1, g2, g3 = st.columns(3)
    with g1:
        sec("Komposisi tone")
        ev = st.altair_chart(donut_alt(df, "Tone", WARNA_TONE, "pilih_tone"), on_select="rerun",
                             key=f"d_tone_{v}", width="stretch")
        pil = pilihan_alt(ev, "pilih_tone", "Tone")
        if pil:
            ke_daftar(tone=pil)
            st.rerun()
    with g2:
        sec("Berita per kata kunci")
        ev = st.altair_chart(donut_alt(df, "Topik", None, "pilih_topik"), on_select="rerun",
                             key=f"d_topik_{v}", width="stretch")
        pil = pilihan_alt(ev, "pilih_topik", "Topik")
        if pil:
            ke_daftar(topik=pil)
            st.rerun()
    with g3:
        sec("Situs teratas")
        ev = st.altair_chart(bar_situs(df), on_select="rerun", key=f"b_situs_{v}", width="stretch")
        pil = pilihan_alt(ev, "pilih_situs", "Situs")
        if pil:
            ke_daftar(situs=pil)
            st.rerun()

    c1, c2 = st.columns([2, 1])
    with c1:
        sec("Tren harian")
        fig = grafik_tren(df)
        if fig is not None:
            ev = st.plotly_chart(fig, width="stretch", on_select="rerun", selection_mode="points",
                                 key=f"tren_{v}", config={"displaylogo": False})
            try:
                hari = sorted({str(p["x"])[:10] for p in ev.selection.points})
            except Exception:
                hari = []
            if hari:
                d0, d1 = date.fromisoformat(hari[0]), date.fromisoformat(hari[-1])
                ke_daftar(tgl=(d0, d1))
                st.rerun()
    with c2:
        sec("Tone per kata kunci")
        ev = st.altair_chart(bar_tone_topik(df), on_select="rerun", key=f"tt_{v}", width="stretch")
        try:
            sel = ev.selection.get("pilih_tt") or []
        except Exception:
            sel = []
        if sel:
            s0 = sel[0] if isinstance(sel, list) else sel
            ke_daftar(topik=[s0["Topik"]], tone=[s0["Tone"]])
            st.rerun()


def dash_daftar(df):
    t_all = df["Tanggal"].dropna()
    t_min = t_all.min().date() if len(t_all) else date.today()
    t_max = t_all.max().date() if len(t_all) else date.today()
    # nilai awal & validasi filter
    ss.setdefault("f_tone", [])
    for k, kol in [("f_topik", "Topik"), ("f_situs", "Situs")]:
        ss[k] = [x for x in ss.get(k, []) if x in set(df[kol])]
    ss["f_seri"] = [x for x in ss.get("f_seri", []) if x in daftar_seri_df(df)]
    ss.setdefault("f_cari", "")
    tg = ss.get("f_tgl")
    if not (isinstance(tg, (tuple, list)) and len(tg) == 2 and t_min <= tg[0] <= tg[1] <= t_max):
        ss["f_tgl"] = (t_min, t_max)

    st.button("← Kembali ke ikhtisar", on_click=pergi, kwargs={"view": "Ikhtisar"})
    with st.container(border=True):
        a, b, c = st.columns([1.2, 1.5, 1.5])
        a.pills("Tone", TONE, selection_mode="multi", key="f_tone")
        b.multiselect("Kata kunci", sorted(df["Topik"].unique()), key="f_topik", placeholder="Semua")
        c.multiselect("Situs media", sorted(df["Situs"].unique()), key="f_situs", placeholder="Semua",
                      format_func=lambda s: s or "(tanpa nama)")
        d, e, f, g = st.columns([1.2, 1.5, 1.5, .6], vertical_alignment="bottom")
        d.multiselect("Seri ritel", daftar_seri_df(df), key="f_seri", placeholder="Semua")
        if t_min < t_max:
            e.date_input("Rentang tanggal", key="f_tgl", min_value=t_min, max_value=t_max,
                         format="DD/MM/YYYY")
        f.text_input("Cari di judul", key="f_cari")
        g.button("Reset", on_click=pergi, kwargs={"reset_filter": True}, help="Hapus semua filter")

    d = df
    if ss["f_tone"]:
        d = d[d["Tone"].isin(ss["f_tone"])]
    if ss["f_topik"]:
        d = d[d["Topik"].isin(ss["f_topik"])]
    if ss["f_situs"]:
        d = d[d["Situs"].isin(ss["f_situs"])]
    if ss["f_seri"]:
        d = d[d["Ritel"].apply(lambda v: any(s in v.split(", ") for s in ss["f_seri"]))]
    if ss["f_cari"]:
        d = d[d["Judul"].str.contains(ss["f_cari"], case=False, regex=False)]
    tg = ss["f_tgl"]
    if isinstance(tg, (tuple, list)) and len(tg) == 2 and (tg[0] > t_min or tg[1] < t_max):
        t = d["Tanggal"].dt.date
        d = d[(t >= tg[0]) & (t <= tg[1])]

    chip = []
    for lbl, k in [("Tone", "f_tone"), ("Kata kunci", "f_topik"), ("Situs", "f_situs"),
                   ("Seri", "f_seri")]:
        if ss[k]:
            chip.append(f"{lbl}: {', '.join(ss[k])}")
    if ss["f_cari"]:
        chip.append(f'Judul memuat "{ss["f_cari"]}"')
    if isinstance(tg, (tuple, list)) and len(tg) == 2 and (tg[0] > t_min or tg[1] < t_max):
        chip.append(f"Tanggal: {nama_periode(tg[0], tg[1])}")
    st.markdown(f"**{len(d)} berita**" + (" · " + "".join(
        f'<span class="chip">{html.escape(c)}</span>' for c in chip) if chip else ""),
        unsafe_allow_html=True)

    m1, m2 = st.columns([1, 1])
    ss.setdefault("dv_mode", "Kartu")
    mode = m1.segmented_control("Tampilan", ["Kartu", "Tabel"], key="dv_mode") or "Kartu"
    ss.setdefault("dv_urut", "Tanggal terlama")
    if ss["dv_urut"] not in ["Tanggal terlama", "Tanggal terbaru", "Negatif dulu", "Situs A–Z"]:
        ss["dv_urut"] = "Tanggal terlama"
    urut = m2.selectbox("Urutkan", ["Tanggal terlama", "Tanggal terbaru", "Negatif dulu", "Situs A–Z"],
                        key="dv_urut")
    if urut == "Tanggal terbaru":
        d = d.sort_values("Tanggal", ascending=False, na_position="last")
    elif urut == "Negatif dulu":
        d = d.assign(_k=d["Tone"].map({"Negatif": 0, "Netral": 1, "Positif": 2})).sort_values(
            ["_k", "Tanggal"], ascending=[True, True]).drop(columns="_k")
    elif urut == "Situs A–Z":
        d = d.sort_values(["Situs", "Tanggal"], ascending=[True, True])
    else:
        d = d.sort_values("Tanggal", ascending=True, na_position="last", kind="stable")
    d = d.reset_index(drop=True)
    d["No"] = range(1, len(d) + 1)

    if mode == "Kartu":
        st.caption("Klik ✏️ di samping kartu untuk mengedit atau menghapus berita.")
        kartu_berita(d, "daftar")
    else:
        tabel_edit(d, "daftar")

    u1, u2, _ = st.columns([1, 1, 3])
    nama = f"berita_sbsn_{datetime.now():%Y%m%d_%H%M}"
    u1.download_button("⬇️ Unduh Excel", sc.excel_bytes(d), file_name=f"{nama}.xlsx", key="xlsx_daftar")
    u2.download_button("⬇️ Unduh CSV", csv_bytes(d), file_name=f"{nama}.csv", mime="text/csv",
                       key="csv_daftar")
    with st.expander("🖨️ Laporan siap cetak untuk daftar ini (PDF / Word)"):
        t = d["Tanggal"].dropna()
        per = nama_periode(t.min().date(), t.max().date()) if len(t) else "-"
        panel_laporan(d, per, kunci="lapdaftar", catatan=("Data disaring: " + "; ".join(chip) + ".")
                      if chip else "", nama_file=f"laporan_berita_sbsn_{datetime.now():%Y%m%d}")


def dash_situs(df):
    st.button("← Kembali ke ikhtisar", on_click=pergi, kwargs={"view": "Ikhtisar"}, key="kembali_situs")
    if df.empty:
        st.info("Belum ada data.")
        return
    rek = rekap_situs(df)
    k = st.columns(4)
    kpi(k[0], "Jumlah situs media", len(rek), EMAS)
    kpi(k[1], "Total berita", int(rek["Jumlah berita"].sum()), UNGU)
    kpi(k[2], "Situs paling aktif", html.escape(rek.iloc[0]["Situs"]), PLUM)
    pos = rek[rek["Jumlah berita"] >= 2].sort_values("Indeks sentimen", ascending=False)
    kpi(k[3], "Paling positif (≥2 berita)", html.escape(pos.iloc[0]["Situs"]) if len(pos) else "–",
        WARNA_TONE["Positif"])

    v = ss["ver"]
    sec("4 situs teratas menurut tone pemberitaan")
    st.caption("Klik batang untuk melihat berita dari situs dan tone tersebut.")
    ev = st.altair_chart(grafik_4_situs(rek), on_select="rerun", key=f"g4_{v}", width="stretch")
    try:
        sel = ev.selection.get("pilih_4") or []
    except Exception:
        sel = []
    if sel:
        s0 = sel[0] if isinstance(sel, list) else sel
        ke_daftar(situs=[s0["Situs"]], tone=[s0["Tone"]])
        st.rerun()

    sec("Daftar situs media")
    st.caption("Klik satu baris untuk melihat daftar berita dari situs tersebut.")
    maks = int(rek["Jumlah berita"].max())
    ev = st.dataframe(
        rek, hide_index=True, width="stretch", on_select="rerun", selection_mode="single-row",
        key=f"tbl_situs_{v}",
        column_config={
            "Jumlah berita": st.column_config.ProgressColumn("Jumlah berita", format="%d",
                                                             min_value=0, max_value=maks),
            "Indeks sentimen": st.column_config.NumberColumn(
                format="%+d", help="(Positif − Negatif) / jumlah berita × 100"),
        })
    try:
        baris = ev.selection.rows
    except Exception:
        baris = []
    if baris:
        situs = rek.iloc[baris[0]]["Situs"]
        ke_daftar(situs=[situs] if situs != "(tanpa nama)" else [""])
        st.rerun()
    st.download_button("⬇️ Unduh rekap situs (CSV)",
                       rek.to_csv(index=False, sep=";").encode("utf-8-sig"),
                       file_name=f"rekap_situs_media_{datetime.now():%Y%m%d}.csv", mime="text/csv")


def halaman_dashboard():
    df, dari, sampai, _ = sumber_data("dash")
    if df.empty:
        st.info("Tidak ada berita pada sumber data ini. Ambil berita dulu di menu **🔎 Ambil berita** "
                "atau ganti sumber data.")
        return
    st.markdown(f"**Periode data:** {nama_periode(dari, sampai)} · {len(df)} berita")
    st.radio("Tampilan dashboard", SUB_DASH, key="dash_view", horizontal=True,
             label_visibility="collapsed")
    view = ss["dash_view"]
    if view == "Daftar berita":
        dash_daftar(df)
    elif view == "Situs media":
        dash_situs(df)
    else:
        dash_ikhtisar(df)


# ===========================================================================
# HALAMAN 3: RINGKASAN PIMPINAN
# ===========================================================================

def halaman_pimpinan():
    kini, dari, sampai, _ = sumber_data("pimp")
    if kini.empty or dari is None:
        st.info("Tidak ada berita bertanggal pada sumber data ini.")
        return
    nama_kini = nama_periode(dari, sampai)
    l_dari, l_sampai = periode_lalu(dari, sampai)
    lalu = saring_periode(gabung(daftar_file()), l_dari, l_sampai)
    nama_lalu = nama_periode(l_dari, l_sampai)
    ada_lalu = not lalu.empty
    st.caption(f"Periode: **{nama_kini}** · dibandingkan dengan **{nama_lalu}**" if ada_lalu
               else f"Periode: **{nama_kini}** · tidak ada data {nama_lalu} untuk pembanding.")

    hk, hl = kini["Tone"].value_counts(), lalu["Tone"].value_counts()

    def d(t=None):
        if not ada_lalu:
            return None
        if t is None:
            return len(kini) - len(lalu)
        return int(hk.get(t, 0)) - int(hl.get(t, 0))

    k = st.columns(4)
    kpi(k[0], "Total berita", len(kini), UNGU, d())
    kpi(k[1], "Positif", int(hk.get("Positif", 0)), WARNA_TONE["Positif"], d("Positif"))
    kpi(k[2], "Netral", int(hk.get("Netral", 0)), WARNA_TONE["Netral"], d("Netral"))
    kpi(k[3], "Negatif", int(hk.get("Negatif", 0)), WARNA_TONE["Negatif"], d("Negatif"), kebalikan=True)

    with st.container(border=True):
        st.markdown(f"**🖨️ Laporan siap cetak: {nama_kini}** (PDF / Word, berisi pembahasan dan "
                    "insight per kata kunci)")
        panel_laporan(kini, nama_kini, lalu if ada_lalu else None, nama_lalu, kunci="pimp",
                      nama_file=f"laporan_berita_sbsn_{dari:%Y%m%d}_{sampai:%Y%m%d}",
                      rentang=(dari, sampai))

    A = laporan.analisis(kini, nama_kini, lalu if ada_lalu else None, nama_lalu)
    sec("Insight utama")
    st.markdown("\n".join(f"- {x}" for x in A["insight_utama"]))

    g1, g2, g3 = st.columns([1, 1, 1.4])
    with g1:
        sec("Komposisi tone")
        st.altair_chart(donut_alt(kini, "Tone", WARNA_TONE, "p_tone", klik=False), width="stretch")
    with g2:
        sec("Berita per kata kunci")
        st.altair_chart(donut_alt(kini, "Topik", None, "p_topik", klik=False), width="stretch")
    with g3:
        sec("Kata yang paling sering muncul di judul")
        awan_kata(kini)

    sec("Pembahasan dan insight per kata kunci")
    for i, t in enumerate(A["topik"]):
        with st.expander(f"{t['nama']} · {t['n']} berita · indeks sentimen {t['indeks']:+d} "
                         f"({t['label']})", expanded=(i == 0)):
            st.write(t["pembahasan"])
            k1, k2 = st.columns(2)
            k1.markdown("**Insight**\n" + "\n".join(f"- {x}" for x in t["insight"]))
            k2.markdown("**Rekomendasi**\n" + "\n".join(f"- {x}" for x in t["rekomendasi"]))

    sec("Rekap per seri SBSN Ritel")
    if A["rekap_seri"] is None:
        st.write("Tidak ada berita berlabel seri pada periode ini.")
    else:
        rek = A["rekap_seri"]
        st.dataframe(rek, width="stretch", column_config={"Total": st.column_config.ProgressColumn(
            "Total", format="%d", min_value=0, max_value=int(rek["Total"].max()))})

    sec("Perlu perhatian: berita bertone negatif")
    neg = kini[kini["Tone"] == "Negatif"].sort_values("Tanggal", ascending=True, na_position="last")
    if neg.empty:
        st.success("Tidak ada berita bertone negatif pada periode ini.")
    else:
        kartu_berita(neg, "neg", per_halaman=10)

    sec("Teks ringkasan (siap salin)")
    teks = [A["ringkasan"], "", "Insight utama:"] + [f"- {x}" for x in A["insight_utama"]]
    st.code("\n".join(teks), language=None, wrap_lines=True)


# ===========================================================================
# HALAMAN 4: TINJAU & KELOLA
# ===========================================================================

def halaman_kelola():
    files = daftar_file()
    sec("Tinjau dan koreksi per file")
    if not files:
        st.write("Belum ada hasil. Ambil berita dulu.")
    else:
        nama = [p.name for p in files]
        aktif = ss.get("file_aktif")
        pilih = st.selectbox("File hasil", nama, index=nama.index(aktif) if aktif in nama else 0)
        df = baca_excel(sc.FOLDER_HASIL / pilih)
        f1, f2, f3 = st.columns(3)
        r_tone = f1.pills("Tone", TONE, selection_mode="multi", key="r_tone") or []
        r_seri = f2.multiselect("Seri", daftar_seri_df(df), key="r_seri", placeholder="Semua")
        r_cari = f3.text_input("Cari judul", key="r_cari")
        r_kembar = st.checkbox("Tampilkan hanya berita berjudul sama (link berbeda) untuk diperiksa",
                               key="r_kembar")
        tampil = df
        if r_kembar:
            jd = df["Judul"].str.strip().str.lower()
            tampil = tampil[jd.duplicated(keep=False)].sort_values(["Judul", "Situs"])
        if r_tone:
            tampil = tampil[tampil["Tone"].isin(r_tone)]
        if r_seri:
            tampil = tampil[tampil["Ritel"].apply(lambda v: any(s in v.split(", ") for s in r_seri))]
        if r_cari:
            tampil = tampil[tampil["Judul"].str.contains(r_cari, case=False, regex=False)]
        tabel_edit(tampil, f"kelola_{pilih}")
        b1, b2, _ = st.columns([1, 1, 3])
        b1.download_button("⬇️ Unduh Excel", data=(sc.FOLDER_HASIL / pilih).read_bytes(), file_name=pilih,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        b2.download_button("⬇️ Unduh CSV", data=csv_bytes(df), file_name=pilih.replace(".xlsx", ".csv"),
                           mime="text/csv", key="csv_review")

    sec("Kelola file hasil scraping")
    if files:
        with st.container(border=True):
            st.markdown("**🧹 Rapikan file yang sudah tersimpan (urutan tanggal & duplikat)**")
            st.caption("Mengurutkan ulang isi file dari tanggal terlama ke terbaru, dan menyatukan berita "
                       "yang tercatat dua kali karena satu versi memakai link asli dan "
                       "versi lain memakai link pengalihan Google (news.google.com), atau versi AMP/mobile "
                       "dari artikel yang sama. Versi link asli dan koreksi yang sudah ada dipertahankan.")
            if st.button("Rapikan semua file", key="bersih_dup"):
                total, rincian = 0, []
                for p in files:
                    n = sc.bersihkan_file(p, ss.get("gabung_judul_sama", sc.GABUNG_JUDUL_SAMA))
                    total += n
                    if n:
                        rincian.append(f"{p.name}: {n}")
                ss["toast"] = ((f"{total} berita ganda dibuang ({'; '.join(rincian)}). " if total
                                else "Tidak ada duplikat. ") + "Semua file kini urut dari tanggal terlama.")
                st.rerun()
    sampah_dir = sc.FOLDER_HASIL / "_sampah"
    if files:
        hapus_file = st.multiselect("Pilih file yang akan dihapus", [p.name for p in files],
                                    placeholder="Pilih satu atau beberapa file")
        if hapus_file:
            yakin = st.checkbox(f"Ya, pindahkan {len(hapus_file)} file ke Sampah")
            if st.button("🗑️ Pindahkan ke Sampah", disabled=not yakin):
                sampah_dir.mkdir(parents=True, exist_ok=True)
                for n in hapus_file:
                    tujuan = sampah_dir / n
                    if tujuan.exists():
                        tujuan = sampah_dir / f"{tujuan.stem}_{datetime.now():%Y%m%d%H%M%S}.xlsx"
                    (sc.FOLDER_HASIL / n).replace(tujuan)
                ss.pop("file_aktif", None)
                ss["toast"] = f"{len(hapus_file)} file dipindahkan ke Sampah"
                st.rerun()

    isi_sampah = sorted(sampah_dir.glob("*.xlsx")) if sampah_dir.exists() else []
    with st.expander(f"Sampah ({len(isi_sampah)} file)"):
        if not isi_sampah:
            st.write("Sampah kosong.")
        else:
            st.caption("File di Sampah tidak tampil di dashboard dan bisa dipulihkan.")
            pilih_s = st.multiselect("Pilih file", [p.name for p in isi_sampah], key="pilih_sampah",
                                     placeholder="Pilih file")
            s1, s2 = st.columns(2)
            if s1.button("↩️ Pulihkan", disabled=not pilih_s):
                for n in pilih_s:
                    tujuan = sc.FOLDER_HASIL / n
                    if tujuan.exists():
                        tujuan = sc.FOLDER_HASIL / f"{tujuan.stem}_pulih.xlsx"
                    (sampah_dir / n).replace(tujuan)
                ss["toast"] = f"{len(pilih_s)} file dipulihkan"
                st.rerun()
            permanen = s2.checkbox("Saya paham file akan hilang permanen", key="yakin_permanen")
            if s2.button("❌ Hapus permanen", disabled=not (pilih_s and permanen)):
                for n in pilih_s:
                    (sampah_dir / n).unlink(missing_ok=True)
                ss["toast"] = f"{len(pilih_s)} file dihapus permanen"
                st.rerun()


# ===========================================================================
# SUSUNAN HALAMAN
# ===========================================================================

hero()
st.radio("Menu", MENU, key="menu", horizontal=True, label_visibility="collapsed")
if "toast" in ss:
    st.toast(ss.pop("toast"), icon="✅")

halaman = ss["menu"]
if halaman == MENU[0]:
    halaman_ambil()
elif halaman == MENU[1]:
    halaman_dashboard()
elif halaman == MENU[2]:
    halaman_pimpinan()
else:
    halaman_kelola()

with st.sidebar:
    st.markdown("### Periode Ambil berita")
    p0, p1 = periode_ambil()
    st.markdown(f"**{nama_periode(p0, p1)}**  \n{p0:%d/%m/%Y} – {p1:%d/%m/%Y}")
    st.caption("Menjadi periode bawaan Dashboard dan Ringkasan pimpinan.")
    st.markdown("### Aturan duplikasi")
    st.toggle("Judul & situs sama = duplikat", value=sc.GABUNG_JUDUL_SAMA, key="gabung_judul_sama",
              help="Mati (bawaan): berita dianggap sama hanya jika link-nya sama; link berbeda "
                   "dihitung sebagai berita berbeda. Hidup: berita dengan judul dan situs yang sama "
                   "juga dianggap satu berita walau link-nya berbeda.")
    st.caption("Saat ini: " + ("link sama **atau** judul+situs sama dianggap duplikat."
                               if ss.get("gabung_judul_sama") else
                               "hanya **link sama** yang dianggap duplikat."))
