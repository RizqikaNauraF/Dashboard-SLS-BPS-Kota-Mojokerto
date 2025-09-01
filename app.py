import os
import io
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
from datetime import datetime


# PDF export
try:
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    HAS_PDF = True
except Exception:
    HAS_PDF = False

st.set_page_config(
    page_title="Dashboard SLS",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)
# ===================== Mapping ===================== #
COLMAP = {
    "id sls": "id_sls", "id_sls": "id_sls", "idsls": "id_sls", "id": "id_sls",
    "nama sls": "nama_sls", "nama_sls": "nama_sls", "nama": "nama_sls",
    "jumlah usaha plkumkm": "plkumkm", "plkumkm": "plkumkm", "jumlah plkumkm": "plkumkm",
    "jumlah usaha kdm": "kdm", "kdm": "kdm", "jumlah kdm": "kdm",
    "selisih jumlah usaha": "selisih", "selisih": "selisih",
}
REQUIRED = ["id_sls", "nama_sls", "plkumkm", "kdm", "selisih"]

KECAMATAN_MAP = {
    "3576010": "Prajurit Kulon",
    "3576020": "Magersari",
    "3576021": "Kranggan",
}

KELURAHAN_MAP = {
    "3576010001": "Surodinawan",
    "3576010004": "Prajuritkulon",
    "3576010005": "Blooto",
    "3576010006": "Mentikan",
    "3576010007": "Kauman",
    "3576010008": "Pulorejo",
    "3576020002": "Gunung Gedangan",
    "3576020003": "Kedundung",
    "3576020004": "Balongsari",
    "3576020008": "Gedongan",
    "3576020009": "Magersari",
    "3576020010": "Wates",
    "3576021004": "Miji",
    "3576021001": "Kranggan",
    "3576021006": "Purwotengah",
    "3576021005": "Sentanan",
    "3576021003": "Jagalan",
    "3576021002": "Meri",
}

# ===================== Load Data ===================== #
@st.cache_data
def load_data(src):
    df = pd.read_excel(src, engine="openpyxl")
    df.columns = [str(c).strip().lower() for c in df.columns]
    keep = {}
    for c in df.columns:
        if c in COLMAP:
            keep[c] = COLMAP[c]
    df = df[list(keep.keys())].rename(columns=keep)

    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"Kolom wajib tidak lengkap: {missing}. Harus ada: {REQUIRED}")

    for c in ["plkumkm", "kdm", "selisih"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df[["plkumkm", "kdm", "selisih"]] = df[["plkumkm", "kdm", "selisih"]].fillna(0).astype(int)

    df["parsed_id"] = df["nama_sls"].astype(str).str.extract(r"^\s*\[?(\d+)\]?", expand=False)

    # Tambahkan kolom kecamatan & kelurahan
    df["id_sls"] = df["id_sls"].astype(str)
    df["kecamatan"] = df["id_sls"].str[:7].map(KECAMATAN_MAP)
    df["kelurahan"] = df["id_sls"].str[:10].map(KELURAHAN_MAP)

    df["kategori"] = np.select(
        [df["selisih"].lt(0), df["selisih"].eq(0), df["selisih"].gt(0)],
        ["Hijau (Over/Bagus)", "Kuning (Match)", "Merah (Kurang)"],
        default="Kuning (Match)",
    )
    return df

def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "SLS") -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return out.getvalue()

def to_pdf_bytes(title: str, table_df: pd.DataFrame) -> bytes:
    if not HAS_PDF:
        raise RuntimeError("ReportLab belum terpasang.")
    buff = io.BytesIO()
    doc = SimpleDocTemplate(buff, pagesize=landscape(A4),
                            leftMargin=24, rightMargin=24, topMargin=18, bottomMargin=18)
    styles = getSampleStyleSheet()
    story = [Paragraph(title, styles["Title"]), Spacer(1, 6),
             Paragraph(datetime.now().strftime("Dibuat: %d %b %Y %H:%M"), styles["Normal"]),
             Spacer(1, 12)]
    data = [list(table_df.columns)] + table_df.astype(str).values.tolist()
    t = Table(data, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.whitesmoke, colors.white]),
    ]))
    story.append(t)
    doc.build(story)
    buff.seek(0)
    return buff.read()

# ===================== Sidebar ===================== #
st.sidebar.header("📂 Data")
default_path = "Data KDM SLS.xlsx"
up = st.sidebar.file_uploader("Upload Excel (opsional)", type=["xlsx"]) 
source = up if up is not None else (default_path if os.path.exists(default_path) else None)
if source is None:
    st.warning(f"Letakkan file **{default_path}** di folder kerja, atau upload dari sidebar.")
    st.stop()

# Load data
try:
    df = load_data(source)
except Exception as e:
    st.error(f"Gagal memuat data: {e}")
    st.stop()

# ---- Filter Kecamatan di Sidebar ---- #
st.sidebar.markdown("---")
st.sidebar.subheader("🏘️ Filter Wilayah")

kecamatan_pilihan = st.sidebar.selectbox(
    "Pilih Kecamatan", 
    options=["(Semua)"] + sorted(df["kecamatan"].dropna().unique().tolist())
)

# ===================== Main Page ===================== #
st.write(f"📅 Data terakhir diperbarui pada: Senin, 1 September 2025, pukul: 05.00")
st.title("📊 Dashboard SLS — BPS Kota Mojokerto")
st.header("PLKUMKM vs KDM")
st.caption("Definisi: Selisih = KDM − PLKUMKM. 0 = Match, <0 = Over/Bagus, >0 = Kurang.")

# ---- Filter Kelurahan (menyesuaikan pilihan kecamatan) ---- #
if kecamatan_pilihan == "(Semua)":
    kelurahan_options = df["kelurahan"].dropna().unique().tolist()
else:
    kelurahan_options = (
        df[df["kecamatan"] == kecamatan_pilihan]["kelurahan"]
        .dropna()
        .unique()
        .tolist()
    )

kelurahan_pilihan = st.multiselect("🏠 Filter Kelurahan", options=kelurahan_options)

# ---- Terapkan filter awal ---- #
view = df.copy()
if kecamatan_pilihan != "(Semua)":
    view = view[view["kecamatan"] == kecamatan_pilihan]
if kelurahan_pilihan:
    view = view[view["kelurahan"].isin(kelurahan_pilihan)]


# ---- KPI ---- #
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total PLKUMKM", f"{int(view['plkumkm'].sum()):,}")
col2.metric("Total KDM", f"{int(view['kdm'].sum()):,}")
col3.metric("Σ Selisih", f"{int(view['selisih'].sum()):,}")
col4.metric("Jumlah SLS", f"{len(view):,}")

col5, col6, col7 = st.columns(3)
col5.metric("Match (=0)", int((view['selisih'] == 0).sum()))
col6.metric("Kurang (>0)", int((view['selisih'] > 0).sum()))
col7.metric("Bagus/Over (<0)", int((view['selisih'] < 0).sum()))

# ---- Pilih Urutan Ranking ---- #
sort_dir = st.radio(
    "Urutkan Ranking Selisih",
    ["Terkecil → Terbesar", "Terbesar → Terkecil"],
    index=0
)
ascending = True if sort_dir.startswith("Terkecil") else False

# ---- Warna untuk Selisih ---- #
COLOR_HIJAU, COLOR_KUNING, COLOR_MERAH = "#e9f7ef", "#fff9db", "#fdecea"

def row_style(row):
    if row["selisih"] < 0:
        return [f"background-color: {COLOR_HIJAU}" for _ in row]
    elif row["selisih"] == 0:
        return [f"background-color: {COLOR_KUNING}" for _ in row]
    else:
        return [f"background-color: {COLOR_MERAH}" for _ in row]

# ---- Akumulasi per Kelurahan ---- #
st.markdown("### 🏠 Akumulasi per Kelurahan")

if not view.empty:
    kel_summary = (
        view.groupby("kelurahan")
        .agg({
            "plkumkm": "sum",
            "kdm": "sum",
            "selisih": "sum",
            "id_sls": "count"
        })
        .rename(columns={"id_sls": "jumlah_sls"})
        .reset_index()
        .sort_values("selisih", ascending=ascending)
        .reset_index(drop=True)
    )

    kel_summary["ranking_selisih"] = kel_summary["selisih"].rank(
        method="min", ascending=ascending
    ).astype(int)

    # urutin ranking_selisih ke paling kiri
    cols = ["ranking_selisih"] + [c for c in kel_summary.columns if c != "ranking_selisih"]
    kel_summary = kel_summary[cols]

    styled_kel = kel_summary.style.apply(row_style, axis=1)
    st.dataframe(styled_kel, use_container_width=True, hide_index=True)

else:
    st.info("Tidak ada data untuk ditampilkan pada akumulasi kelurahan.")


# ---- Controls lain ---- #
st.markdown("---")
colA, colB = st.columns([3, 3])   # FIXED: 2 kolom saja
with colA:
    q = st.text_input("🔎 Cari Nama/ID SLS", placeholder="Ketik Nama SLS atau ID…")
with colB:
    kategori_pilihan = st.multiselect(
        "🟩🟨🟥 Filter Kategori", 
        options=["Hijau (Over/Bagus)", "Kuning (Match)", "Merah (Kurang)"],
        default=["Hijau (Over/Bagus)", "Kuning (Match)", "Merah (Kurang)"]
    )

# ---- Apply Filter awal ---- #
view = df.copy()
if q:
    s = q.strip().lower()
    view = view[
        view["nama_sls"].astype(str).str.lower().str.contains(s, na=False) |
        view["id_sls"].astype(str).str.contains(s, na=False) |
        view["parsed_id"].fillna("").str.contains(s, na=False)
    ]
if kategori_pilihan:
    view = view[view["kategori"].isin(kategori_pilihan)]
if kecamatan_pilihan != "(Semua)":
    view = view[view["kecamatan"] == kecamatan_pilihan]
if kelurahan_pilihan:
    view = view[view["kelurahan"].isin(kelurahan_pilihan)]

# ---- Sort data ---- #
view_sorted = view.sort_values("selisih", ascending=ascending).reset_index(drop=True)
view_sorted["ranking_selisih"] = view_sorted["selisih"].rank(
    method="min", ascending=ascending).astype(int)

# ---- Tabel SLS ---- #
st.subheader("📋 Tabel SLS")
show_cols = ["ranking_selisih", "id_sls", "nama_sls", "kecamatan", "kelurahan", "plkumkm", "kdm", "selisih", "kategori"]

styled_sls = view_sorted[show_cols].style.apply(row_style, axis=1)
st.dataframe(styled_sls, use_container_width=True, hide_index=True)

# ---- Grafik ---- #
st.markdown("---")
st.subheader("📊 Grafik Ranking Selisih")
if not view_sorted.empty:
    fig = px.bar(view_sorted, x="nama_sls", y="selisih",
                 hover_data=["id_sls", "kecamatan", "kelurahan", "plkumkm", "kdm", "kategori"],
                 title=f"Ranking Selisih (urut: {'Naik' if ascending else 'Turun'})")
    fig.update_layout(xaxis_tickangle=-45, height=500)
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Tidak ada data untuk divisualisasikan.")

# ---- Unduhan ---- #
st.markdown("---")
st.subheader("⬇️ Unduh Hasil (sesuai filter)")
colx, coly = st.columns(2)
with colx:
    st.download_button("💾 Download Excel",
        data=to_excel_bytes(view_sorted[show_cols]),
        file_name=f"SLS_filtered_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with coly:
    if HAS_PDF:
        try:
            pdf_bytes = to_pdf_bytes("Tabel SLS (Hasil Filter)", view_sorted[show_cols])
            st.download_button("🧾 Download PDF", data=pdf_bytes,
                file_name=f"SLS_filtered_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf")
        except Exception as e:
            st.warning(f"PDF gagal dibuat: {e}")
    else:
        st.info("Export PDF butuh paket 'reportlab'. Install: pip install reportlab")

st.caption("Gunakan pencarian & filter kategori/kecamatan/kelurahan untuk fokus. Urutkan Selisih agar terlihat mana yang match, over (bagus), atau masih kurang.")

st.markdown("""
                <hr style="border: 0.5px solid #ccc;" />
                <center><small>&copy; 2025 BPS Kota Mojokerto</small></center>
                """, unsafe_allow_html=True)
