"""Catalogo interactivo de inmuebles - aplicacion principal (Streamlit)."""
from __future__ import annotations

import io
import math
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from data_loader import (
    DATA_DIR,
    DEFAULT_SOURCE_CANDIDATES,
    format_currency,
    get_property_image,
    load_properties,
)
from pdf_generator import generate_catalog_pdf

SAMPLE_PATH = DATA_DIR / "propiedades_ejemplo.xlsx"
DEFAULT_PATH = DATA_DIR / "propiedades.xlsx"
TMP_DIR = Path(__file__).parent / "assets" / "tmp"

st.set_page_config(page_title="Catalogo de Inmuebles", page_icon="🏠", layout="wide")

st.markdown(
    """
    <style>
    .property-card {
        border: 1px solid #e2e8f0; border-radius: 14px; padding: 14px;
        background: #ffffff; box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        margin-bottom: 18px; height: 100%;
    }
    .property-price { font-size: 1.35rem; font-weight: 700; color: #c9a227; margin: 2px 0 6px 0; }
    .property-name { font-size: 1.05rem; font-weight: 700; color: #1e3a5f; margin-bottom: 2px; }
    .property-loc { color: #64748b; font-size: 0.85rem; margin-bottom: 8px; }
    .badge { display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 0.72rem;
             font-weight: 600; color: white; margin-bottom: 8px; }
    .badge-dev { background: #c9a227; }
    .badge-market { background: #1e3a5f; }
    .spec-row { font-size: 0.85rem; color: #334155; margin-bottom: 2px; }
    </style>
    """,
    unsafe_allow_html=True,
)

if "selected" not in st.session_state:
    st.session_state.selected = {}


@st.cache_data(show_spinner="Leyendo Excel...")
def _load_from_path_cached(path_str: str, mtime: float) -> pd.DataFrame:
    return load_properties(path_str)


def _load_from_path(path: Path) -> pd.DataFrame:
    return _load_from_path_cached(str(path), path.stat().st_mtime)


@st.cache_data(show_spinner="Leyendo Excel...")
def _load_from_bytes(file_bytes: bytes) -> pd.DataFrame:
    return load_properties(io.BytesIO(file_bytes))


@st.cache_data(show_spinner=False)
def _cached_image_path(url: str | None) -> str:
    return str(get_property_image(url))


def _cached_gallery_paths(urls) -> list[str]:
    if not urls:
        return [str(get_property_image(None))]
    return [_cached_image_path(u) for u in urls]


def _fmt_m2(value) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "N/D"
    return f"{value:,.0f} m²"


# ---------------------------------------------------------------- Sidebar --
st.sidebar.title("🏠 Catalogo de Inmuebles")
st.sidebar.caption("La Paz, Baja California Sur")

fuente_opciones = ["Archivo por defecto", "Datos de ejemplo", "Subir mi propio Excel"]
if not DEFAULT_PATH.exists():
    fuente_opciones.remove("Archivo por defecto")
fuente = st.sidebar.radio("Fuente de datos", fuente_opciones, index=0)

uploaded_file = None
if fuente == "Subir mi propio Excel":
    uploaded_file = st.sidebar.file_uploader("Selecciona un archivo .xlsx", type=["xlsx"])

if fuente == "Subir mi propio Excel" and uploaded_file is not None:
    df = _load_from_bytes(uploaded_file.getvalue())
elif fuente == "Datos de ejemplo":
    df = _load_from_path(SAMPLE_PATH)
elif fuente == "Archivo por defecto":
    df = _load_from_path(DEFAULT_PATH)
else:
    df = pd.DataFrame()

if df.empty:
    st.info("Sube un archivo Excel (.xlsx) desde el panel lateral para comenzar, o genera datos de ejemplo con `python generate_sample_data.py`.")
    st.stop()

st.sidebar.markdown("---")
st.sidebar.subheader("Filtros")

ubicaciones = sorted(u for u in df["ubicacion_filtro"].unique() if u)
sel_ubicaciones = st.sidebar.multiselect("Zona / Colonia / Ciudad", ubicaciones, default=[])

precios_validos = df["precio"].dropna()
if len(precios_validos) > 0:
    precio_min_abs, precio_max_abs = int(precios_validos.min()), int(precios_validos.max())
    if precio_min_abs == precio_max_abs:
        precio_max_abs += 1
    sel_precio = st.sidebar.slider(
        "Rango de precio (MXN)", min_value=precio_min_abs, max_value=precio_max_abs,
        value=(precio_min_abs, precio_max_abs), step=max(1, (precio_max_abs - precio_min_abs) // 100),
    )
else:
    sel_precio = None

estados_disponibles = sorted(df["estado"].dropna().unique())
sel_estados = st.sidebar.multiselect("Estado de la propiedad", estados_disponibles, default=estados_disponibles)

umbral_opciones = {"Cualquiera": 0, "1+": 1, "2+": 2, "3+": 3, "4+": 4}
c1, c2, c3 = st.sidebar.columns(3)
sel_recamaras = c1.selectbox("Recamaras", list(umbral_opciones.keys()), key="rec")
sel_banos = c2.selectbox("Banos", list(umbral_opciones.keys()), key="ban")
sel_estac = c3.selectbox("Cocheras", list(umbral_opciones.keys()), key="est")

# ------------------------------------------------------------- Filtrado ---
filtered = df.copy()
if sel_ubicaciones:
    filtered = filtered[filtered["ubicacion_filtro"].isin(sel_ubicaciones)]
if sel_precio is not None:
    filtered = filtered[
        filtered["precio"].isna() | filtered["precio"].between(sel_precio[0], sel_precio[1])
    ]
if sel_estados:
    filtered = filtered[filtered["estado"].isin(sel_estados)]
filtered = filtered[filtered["recamaras"] >= umbral_opciones[sel_recamaras]]
filtered = filtered[filtered["banos"] >= umbral_opciones[sel_banos]]
filtered = filtered[filtered["estacionamientos"] >= umbral_opciones[sel_estac]]

st.sidebar.markdown("---")
n_sel = len(st.session_state.selected)
st.sidebar.subheader(f"Seleccionadas: {n_sel}")
if n_sel > 0:
    with st.sidebar.expander("Ver seleccionadas"):
        for pid, p in st.session_state.selected.items():
            st.write(f"- {p['nombre']}")
    if st.sidebar.button("Limpiar seleccion"):
        st.session_state.selected = {}
        st.rerun()

    if st.sidebar.button("📄 Exportar Seleccionados a PDF", type="primary", width="stretch"):
        TMP_DIR.mkdir(parents=True, exist_ok=True)
        props_for_pdf = []
        for pid, row in st.session_state.selected.items():
            props_for_pdf.append(
                {
                    "nombre": row["nombre"],
                    "precio_display": format_currency(row["precio"], row.get("moneda")) if pd.notna(row["precio"]) else (row.get("precio_original") or "Precio no disponible"),
                    "ubicacion": row["ubicacion_filtro"],
                    "estado": row["estado"],
                    "recamaras": row["recamaras"],
                    "banos": row["banos"],
                    "estacionamientos": row["estacionamientos"],
                    "m2_construccion_display": _fmt_m2(row["m2_construccion"]),
                    "m2_terreno_display": _fmt_m2(row["m2_terreno"]),
                    "descripcion": row.get("descripcion") or "",
                    "enlace_externo": row.get("enlace_externo"),
                    "image_path": _cached_image_path(row.get("imagen_url")),
                }
            )
        out_path = TMP_DIR / f"catalogo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        generate_catalog_pdf(props_for_pdf, out_path)
        pdf_bytes = out_path.read_bytes()
        out_path.unlink(missing_ok=True)
        st.session_state["pdf_ready"] = pdf_bytes
        st.session_state["pdf_name"] = out_path.name

if st.session_state.get("pdf_ready"):
    st.sidebar.download_button(
        "⬇️ Descargar catalogo.pdf", data=st.session_state["pdf_ready"],
        file_name=st.session_state.get("pdf_name", "catalogo.pdf"), mime="application/pdf",
        width="stretch",
    )

# ---------------------------------------------------------------- Main ----
st.title("Catalogo de Inmuebles")
st.caption(f"Mostrando {len(filtered)} de {len(df)} propiedades")

if filtered.empty:
    st.warning("Ningun inmueble coincide con los filtros seleccionados.")
    st.stop()

COLS_PER_ROW = 3
rows = list(filtered.iterrows())
for i in range(0, len(rows), COLS_PER_ROW):
    cols = st.columns(COLS_PER_ROW)
    for col, (idx, row) in zip(cols, rows[i : i + COLS_PER_ROW]):
        with col:
            with st.container():
                st.markdown('<div class="property-card">', unsafe_allow_html=True)
                gallery = _cached_gallery_paths(row.get("imagenes_urls"))
                st.image(gallery[0], width="stretch")
                if len(gallery) > 1:
                    thumb_cols = st.columns(min(len(gallery) - 1, 5))
                    for tcol, thumb in zip(thumb_cols, gallery[1:6]):
                        tcol.image(thumb, width="stretch")

                badge_class = "badge-dev" if row["estado"] == "Desarrollo" else "badge-market"
                st.markdown(f'<span class="badge {badge_class}">{row["estado"]}</span>', unsafe_allow_html=True)
                st.markdown(f'<div class="property-name">{row["nombre"]}</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="property-loc">📍 {row["ubicacion_filtro"]}</div>', unsafe_allow_html=True)

                precio_text = (
                    format_currency(row["precio"], row.get("moneda"))
                    if pd.notna(row["precio"])
                    else (row.get("precio_original") or "Precio no disponible")
                )
                st.markdown(f'<div class="property-price">{precio_text}</div>', unsafe_allow_html=True)

                st.markdown(
                    f'<div class="spec-row">🛏️ {int(row["recamaras"])} rec &nbsp;|&nbsp; '
                    f'🛁 {int(row["banos"])} banos &nbsp;|&nbsp; 🚗 {int(row["estacionamientos"])}</div>'
                    f'<div class="spec-row">📐 Construccion: {_fmt_m2(row["m2_construccion"])} &nbsp;|&nbsp; '
                    f'Terreno: {_fmt_m2(row["m2_terreno"])}</div>',
                    unsafe_allow_html=True,
                )

                if row.get("enlace_externo"):
                    st.markdown(f'<a href="{row["enlace_externo"]}" target="_blank">Ver anuncio original ↗</a>', unsafe_allow_html=True)

                pid = row["id"]
                checked = st.checkbox(
                    "Seleccionar para PDF", value=(pid in st.session_state.selected), key=f"chk_{pid}_{row['hoja_origen']}"
                )
                if checked:
                    st.session_state.selected[pid] = row.to_dict()
                else:
                    st.session_state.selected.pop(pid, None)

                st.markdown("</div>", unsafe_allow_html=True)
