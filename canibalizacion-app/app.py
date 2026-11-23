"""
App de canibalización de tiendas (Streamlit)
- Capa DANE (personas y viviendas por manzana)
- Tiendas operando + ventas desde OneDrive/SharePoint (Microsoft Graph) o carga manual
- Isócronas de 300 m (caminando real con OpenRouteService o círculo simulado)
- Canibalización sugerida y editable, impacto en ventas y contribución (3 meses)
- Gráfico de ventas 6 meses, comentarios y PDF final con mapa estilo OpenStreetMap
"""
import base64
import io
import os
import re
import tempfile
from datetime import date

import contextily as cx
import folium
import geopandas as gpd
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import msal
import numpy as np
import pandas as pd
import requests
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from shapely.geometry import Point, shape
from streamlit_folium import st_folium

st.set_page_config(page_title="Canibalización de tiendas", page_icon="🏪", layout="wide")
CRS_M = 3116  # MAGNA-SIRGAS / Colombia Bogotá (metros)


# ----------------------------------------------------------------------------
# Utilidades generales
# ----------------------------------------------------------------------------
def has_secret(k):
    try:
        return k in st.secrets
    except Exception:
        return False


def cop(x):
    return "—" if x is None or pd.isna(x) else f"${x:,.0f}".replace(",", ".")


def guess(cols, keys):
    low = [str(c).lower() for c in cols]
    # Primero prioriza coincidencias exactas: en el Book, Y=latitud y X=longitud.
    for k in keys:
        for i, c in enumerate(low):
            if c.strip() == k.lower():
                return i
    for k in keys:
        for i, c in enumerate(low):
            if k in c:
                return i
    return 0


def parse_coord(txt):
    """Acepta '4.5878, -74.1034', '(4.5878; -74.1034)', etc. Devuelve (lat, lon) o None."""
    nums = re.findall(r"-?\d+\.\d+|-?\d+", txt or "")
    if len(nums) != 2:
        return None
    lat, lon = float(nums[0]), float(nums[1])
    return (lat, lon) if -90 <= lat <= 90 and -180 <= lon <= 180 else None


def to_num(s):
    return pd.to_numeric(s.astype(str).str.replace(",", ".", regex=False), errors="coerce")


def distance_m(lat1, lon1, lat2, lon2):
    """Distancia geodésica aproximada en metros, robusta para cualquier zona de Colombia."""
    lat1, lon1, lat2, lon2 = map(np.radians, [lat1, lon1, lat2, lon2])
    dlat, dlon = lat2 - lat1, lon2 - lon1
    a = np.sin(dlat / 2) ** 2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon / 2) ** 2
    return float(2 * 6371000 * np.arcsin(np.sqrt(a)))


def to_m(geom, from_crs=4326):
    return gpd.GeoSeries([geom], crs=from_crs).to_crs(CRS_M).iloc[0]


def to_ll(geom):
    return gpd.GeoSeries([geom], crs=CRS_M).to_crs(4326).iloc[0]


# ----------------------------------------------------------------------------
# Conexión a OneDrive / SharePoint (Microsoft Graph)
# ----------------------------------------------------------------------------
def graph_download(share_url: str) -> bytes:
    az = st.secrets["azure"]
    cca = msal.ConfidentialClientApplication(
        az["client_id"],
        authority=f"https://login.microsoftonline.com/{az['tenant_id']}",
        client_credential=az["client_secret"],
    )
    tok = cca.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in tok:
        raise RuntimeError(tok.get("error_description", "No se pudo obtener token"))
    enc = "u!" + base64.urlsafe_b64encode(share_url.encode()).decode().rstrip("=")
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{enc}/driveItem/content",
        headers={"Authorization": f"Bearer {tok['access_token']}"},
        timeout=120,
    )
    r.raise_for_status()
    return r.content


def candidate_urls(u: str):
    c = [u + ("&" if "?" in u else "?") + "download=1"]
    m = re.match(r"(https://[^/]+)/:[a-z]:/[a-z]/(personal/[^/]+)/([^/?]+)", u)
    if m:
        host, site, tok = m.groups()
        c.append(f"{host}/{site}/_layouts/15/download.aspx?share={tok}")
    return c


def anon_download(share_url: str) -> bytes:
    """Descarga directa de un link 'Cualquier persona con el vínculo'. Prueba varias formas y reporta qué pasó."""
    log = []
    for u in candidate_urls(share_url.strip()):
        try:
            r = requests.get(u, timeout=120, allow_redirects=True, headers={"User-Agent": "Mozilla/5.0"})
            if r.content[:2] == b"PK":  # un .xlsx es un zip
                return r.content
            log.append(f"HTTP {r.status_code} · {r.headers.get('content-type', '?')[:40]} · {r.url.split('?')[0][:80]}")
        except Exception as e:  # noqa
            log.append(f"Error de red: {e}")
    raise RuntimeError("SharePoint no entregó el Excel (probablemente el link pide inicio de sesión). Intentos:\n- " + "\n- ".join(log))


@st.cache_data(ttl=900, show_spinner="Descargando desde OneDrive...")
def cached_download(url: str) -> bytes:
    try:
        return anon_download(url)
    except Exception as e:
        if has_secret("azure"):  # respaldo opcional con Microsoft Graph
            return graph_download(url)
        raise e


@st.cache_data(ttl=3600, show_spinner="Descargando capa de manzanas desde OneDrive...")
def cached_layer_download(url: str) -> bytes:
    """Descarga la capa remota y la conserva en caché durante una hora."""
    return cached_download(url)


def secret_url(flat_key: str, old_key: str) -> str:
    """Lee ONEDRIVE_BOOK_URL / ONEDRIVE_SALES_URL (o el formato anterior [onedrive])."""
    if has_secret(flat_key):
        return str(st.secrets[flat_key])
    if has_secret("onedrive"):
        return str(st.secrets["onedrive"].get(old_key, ""))
    return ""


def get_excel(label: str, secret_key: str, key: str):
    """Devuelve DataFrame desde OneDrive (link en secrets) o archivo subido."""
    flat = {"tiendas_url": "ONEDRIVE_BOOK_URL", "ventas_url": "ONEDRIVE_SALES_URL"}[secret_key]
    default_url = secret_url(flat, secret_key)
    opts = (["OneDrive (secrets)"] if default_url else []) + ["Subir archivo"]
    src = st.radio(f"Fuente: {label}", opts, key=f"src_{key}", horizontal=True)
    data = None
    try:
        if src == "OneDrive (secrets)":
            data = cached_download(default_url)
            st.caption("✅ Conectado a OneDrive")
            if st.button("🔄 Refrescar", key=f"ref_{key}"):
                cached_download.clear()
                st.rerun()
        else:
            up = st.file_uploader(label, type=["xlsx", "xls", "csv"], key=f"up_{key}")
            if up:
                data = up.getvalue()
                if up.name.endswith(".csv"):
                    return pd.read_csv(io.BytesIO(data))
    except Exception as e:  # noqa
        st.error(f"No pude leer {label}: {e}")
        return None
    if data is None:
        return None
    xl = pd.ExcelFile(io.BytesIO(data))
    sheet = st.selectbox("Hoja", xl.sheet_names, key=f"sh_{key}")
    return xl.parse(sheet)


# ----------------------------------------------------------------------------
# Capa DANE
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner="Cargando manzanas DANE...")
def read_manzanas(content: bytes, name: str, bbox=None) -> gpd.GeoDataFrame:
    suffix = os.path.splitext(name)[1]
    if suffix.lower() == ".lyrx":
        raise ValueError(
            "Un .lyrx es una definición de capa de ArcGIS, no contiene los polígonos. "
            "Exporta la capa como GeoPackage (.gpkg), GeoJSON (.geojson) o ZIP de shapefile "
            "incluyendo geometría y atributos."
        )
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as f:
        f.write(content)
        path = f.name
    g = gpd.read_file(f"zip://{path}" if suffix.lower() == ".zip" else path, bbox=bbox)
    if g.empty:
        raise ValueError("La capa no contiene geometrías.")
    if g.crs is None:
        st.warning("La capa no trae CRS; asumiré EPSG:4326 (longitud/latitud).")
        g = g.set_crs(4326)
    return g.to_crs(CRS_M)


def bbox_around(point_m, radius_m):
    """Bounding box WGS84 para leer solo las manzanas cercanas al análisis."""
    roi = gpd.GeoSeries([point_m.buffer(radius_m)], crs=CRS_M).to_crs(4326).iloc[0]
    return tuple(roi.bounds)


def remote_layer_url():
    return str(st.secrets.get("ONEDRIVE_LAYER_URL", "")) if has_secret("ONEDRIVE_LAYER_URL") else ""


def load_manzanas():
    local = [f for f in os.listdir("data") if f.lower().endswith((".gpkg", ".geojson", ".json", ".zip", ".shp"))] if os.path.isdir("data") else []
    if local:
        p = os.path.join("data", local[0])
        if os.path.getsize(p) > 100 * 1024 * 1024:
            st.info("La capa local es grande; se leerá solo alrededor de la tienda nueva.")
            st.session_state["heavy_layer_path"] = p
            return None
        st.caption(f"Usando capa local: {p}")
        with open(p, "rb") as fh:
            return read_manzanas(fh.read(), p)
    remote_url = remote_layer_url()
    if remote_url:
        st.info("La capa de OneDrive se cargará después de ingresar la coordenada, solo para el área cercana.")
        return None
    up = st.file_uploader("Capa de manzanas (.gpkg, .geojson, .json o .zip shapefile)", type=["gpkg", "geojson", "json", "zip", "lyrx"], key="manz")
    if not up:
        return None
    try:
        g = read_manzanas(up.getvalue(), up.name)
        st.success(f"Capa cargada: {len(g):,} manzanas")
        return g
    except Exception as e:
        st.error(f"No pude leer la capa geográfica: {e}")
        return None


def load_manzanas_near(point_m, radius_m):
    """Carga una ventana local/remota, evitando leer la capa nacional completa."""
    bbox = bbox_around(point_m, radius_m)
    path = st.session_state.get("heavy_layer_path")
    try:
        if path:
            return gpd.read_file(f"zip://{path}" if path.lower().endswith(".zip") else path, bbox=bbox).to_crs(CRS_M)
        url = remote_layer_url()
        if url:
            content = cached_layer_download(url)
            return read_manzanas(content, "VIV&HAB.zip", bbox=bbox)
    except Exception as e:
        st.error(f"No pude cargar las manzanas cercanas: {e}")
    return None


def pop_in(poly, manz, c_pers, c_viv):
    """Población y viviendas dentro de un polígono (ponderado por área)."""
    if manz is None:
        return np.nan, np.nan
    idx = manz.sindex.query(poly, predicate="intersects")
    sub = manz.iloc[idx]
    if sub.empty:
        return 0.0, 0.0
    areas = sub.geometry.buffer(0).area.replace(0, np.nan)
    frac = (sub.geometry.buffer(0).intersection(poly).area / areas).fillna(0).clip(0, 1)
    return float((frac * sub[c_pers]).sum()), float((frac * sub[c_viv]).sum())


def manzanas_in(poly, manz, c_pers, c_viv):
    """Devuelve la capa recortada con población/viviendas ponderadas por área."""
    if manz is None:
        return None
    idx = manz.sindex.query(poly, predicate="intersects")
    sub = manz.iloc[idx].copy()
    if sub.empty:
        return sub
    sub["__fraccion"] = (sub.geometry.buffer(0).intersection(poly).area / sub.geometry.buffer(0).area.replace(0, np.nan)).fillna(0).clip(0, 1)
    sub["Personas en isócrona"] = sub[c_pers] * sub["__fraccion"]
    sub["Viviendas en isócrona"] = sub[c_viv] * sub["__fraccion"]
    return sub


# ----------------------------------------------------------------------------
# Isócronas
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def ors_iso(lat, lon, radio, key):
    r = requests.post(
        "https://api.openrouteservice.org/v2/isochrones/foot-walking",
        headers={"Authorization": key, "Content-Type": "application/json"},
        json={"locations": [[lon, lat]], "range": [radio], "range_type": "distance"},
        timeout=30,
    )
    r.raise_for_status()
    return r.json()["features"][0]["geometry"]


def make_iso(lat, lon, radio, mode):
    pt_m = to_m(Point(lon, lat))
    if mode.startswith("Caminando"):
        key = st.secrets.get("ors", {}).get("api_key", "") if has_secret("ors") else ""
        if key:
            try:
                return to_m(shape(ors_iso(lat, lon, radio, key))).buffer(0)
            except Exception as e:  # noqa
                st.warning(f"ORS falló ({e}). Uso círculo simulado.")
        else:
            st.warning("Sin API key de ORS: uso círculo simulado de 300 m.")
    return pt_m.buffer(radio)


# ----------------------------------------------------------------------------
# Mapa estático (para el informe) y gráfico
# ----------------------------------------------------------------------------
def static_map(new_pt, iso_new, rows, new_name):
    fig, ax = plt.subplots(figsize=(8, 6.5))
    gpd.GeoSeries([iso_new], crs=CRS_M).to_crs(3857).plot(ax=ax, color="#e63946", alpha=0.25, edgecolor="#e63946", linewidth=2)
    xs, ys = [], []
    for r in rows:
        gpd.GeoSeries([r["iso"]], crs=CRS_M).to_crs(3857).plot(ax=ax, color="#1d6fd6", alpha=0.25, edgecolor="#1d6fd6", linewidth=2)
        a = gpd.GeoSeries([new_pt, r["pt"]], crs=CRS_M).to_crs(3857)
        ax.plot([a[0].x, a[1].x], [a[0].y, a[1].y], "k--", lw=1.2)
        ax.annotate(f"{r['dist']:.0f} m", ((a[0].x + a[1].x) / 2, (a[0].y + a[1].y) / 2), fontsize=9, weight="bold",
                    bbox=dict(boxstyle="round", fc="white", alpha=0.85))
        ax.scatter(a[1].x, a[1].y, c="#1d6fd6", s=90, zorder=5, edgecolor="white")
        ax.annotate(r["nombre"], (a[1].x, a[1].y), xytext=(6, 6), textcoords="offset points", fontsize=8, weight="bold")
    n = gpd.GeoSeries([new_pt], crs=CRS_M).to_crs(3857).iloc[0]
    ax.scatter(n.x, n.y, c="#e63946", s=140, marker="*", zorder=6, edgecolor="white")
    ax.annotate(f"NUEVA: {new_name}", (n.x, n.y), xytext=(6, -14), textcoords="offset points", fontsize=8, weight="bold", color="#e63946")
    tb = gpd.GeoSeries([iso_new] + [r["iso"] for r in rows], crs=CRS_M).to_crs(3857).total_bounds
    pad = 120
    ax.set_xlim(tb[0] - pad, tb[2] + pad)
    ax.set_ylim(tb[1] - pad, tb[3] + pad)
    try:
        cx.add_basemap(ax, source=cx.providers.OpenStreetMap.Mapnik, attribution_size=6)
    except Exception:
        pass
    ax.set_axis_off()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def sales_chart(piv, pct_map):
    loss = (piv * pd.Series(pct_map)).sum(axis=1) if not piv.empty else pd.Series(dtype=float)
    tot = piv.sum(axis=1)
    fig, ax = plt.subplots(figsize=(8, 3.8))
    x = [str(p) for p in tot.index]
    ax.bar(x, tot - loss, color="#1d6fd6", label="Ventas tras canibalización")
    ax.bar(x, loss, bottom=tot - loss, color="#e63946", hatch="//", label="Pérdida estimada")
    for i, v in enumerate(tot):
        ax.text(i, v, f"{v/1e6:,.1f} M", ha="center", va="bottom", fontsize=8)
    ax.set_ylabel("Ventas (COP)")
    ax.set_title("Ventas últimos 6 meses – tiendas impactadas")
    ax.yaxis.set_major_formatter(lambda v, _: f"{v/1e6:,.0f} M")
    ax.legend(fontsize=8)
    ax.spines[["top", "right"]].set_visible(False)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# PDF
# ----------------------------------------------------------------------------
def tbl(data, widths=None, header=True):
    t = Table(data, colWidths=widths, repeatRows=1 if header else 0)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1d6fd6")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 8), ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f5fa")]),
    ]))
    return t


def build_pdf(c):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.6 * cm, rightMargin=1.6 * cm, topMargin=1.5 * cm, bottomMargin=1.5 * cm)
    s = getSampleStyleSheet()
    el = [Paragraph("Informe de canibalización – nueva tienda", s["Title"]),
          Paragraph(f"<b>Tienda propuesta:</b> {c['nombre']} &nbsp; | &nbsp; <b>Coordenada:</b> {c['lat']:.6f}, {c['lon']:.6f} "
                    f"&nbsp; | &nbsp; <b>Fecha:</b> {date.today():%d/%m/%Y}", s["Normal"]),
          Paragraph(f"Isócrona: {c['radio']} m ({c['modo']}). Tasa de migración base: {c['mig']:.0%}.", s["Normal"]),
          Spacer(1, 8), Image(io.BytesIO(c["map"]), width=16 * cm, height=13 * cm), Spacer(1, 8),
          Paragraph("Tiendas operando en traslape", s["Heading2"])]
    d = [["Tienda", "Dist. (m)", "Pob. isócrona", "Pob. compartida", "% traslape", "% sugerido", "% final"]]
    for _, r in c["stores"].iterrows():
        d.append([r["Nombre"], f"{r['Distancia (m)']:.0f}", f"{r['Pob. isócrona existente']:,.0f}", f"{r['Pob. compartida']:,.0f}",
                  f"{r['% traslape']:.1f}%", f"{r['% sugerido']:.1f}%", f"{r['% final (editable)']:.1f}%"])
    el += [tbl(d), Spacer(1, 8)]
    if c["imp"] is not None:
        el.append(Paragraph("Impacto últimos 3 meses", s["Heading2"]))
        d = [["Tienda", "Ventas 3M", "Impacto ventas", "Ventas 3M post", "Contrib. 3M", "Impacto contrib.", "Contrib. post"]]
        for _, r in c["imp"].iterrows():
            d.append([r["Nombre"], cop(r["Ventas 3M"]), cop(-r["Impacto ventas"]), cop(r["Ventas 3M post"]),
                      cop(r["Contribución 3M"]), cop(-r["Impacto contribución"]), cop(r["Contribución 3M post"])])
        t = c["tot"]
        d.append(["TOTAL", cop(t["v"]), cop(-t["iv"]), cop(t["v"] - t["iv"]), cop(t["c"]), cop(-t["ic"]), cop(t["c"] - t["ic"])])
        el += [tbl(d), Spacer(1, 6),
               Paragraph(f"<b>Ventas 3M:</b> {cop(t['v'])} → {cop(t['v']-t['iv'])} ({-t['iv']/t['v']:.1%}). "
                         f"<b>Contribución 3M:</b> {cop(t['c'])} → {cop(t['c']-t['ic'])} ({-t['ic']/t['c']:.1%})." if t["v"] and t["c"] else "", s["Normal"]),
               Spacer(1, 8), Image(io.BytesIO(c["chart"]), width=16 * cm, height=7.6 * cm)]
    el += [Paragraph("Comentarios", s["Heading2"]), Paragraph((c["coment"] or "Sin comentarios.").replace("\n", "<br/>"), s["Normal"])]
    doc.build(el)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
st.title("🏪 Simulador de canibalización de tiendas")

with st.sidebar:
    with st.expander("🔎 Diagnóstico de secrets"):
        for k in ("ONEDRIVE_BOOK_URL", "ONEDRIVE_SALES_URL"):
            st.write(("✅ " if secret_url(k, "") else "❌ ") + k)
        try:
            st.caption("Claves detectadas: " + ", ".join(st.secrets.keys()))
        except Exception:
            st.caption("No hay secrets configurados.")
    st.header("1. Datos")
    with st.expander("Tiendas operando", expanded=True):
        df_t = get_excel("Tiendas operando (Book.xlsx)", "tiendas_url", "tiendas")
    with st.expander("Histórico de ventas", expanded=True):
        df_v = get_excel("Ventas (LibroVentas.xlsx)", "ventas_url", "ventas")
    with st.expander("Capa DANE", expanded=True):
        manz = load_manzanas()
    st.header("2. Parámetros")
    radio = st.number_input("Radio isócrona (m)", 100, 1000, 300, 50)
    modo = st.radio("Tipo de isócrona", ["Caminando real (ORS)", "Círculo simulado"])
    mig = st.slider("Tasa de migración esperada", 5, 80, 30, 5, help="% de los clientes compartidos que se pasarían a la nueva tienda") / 100

if df_t is None:
    st.info("👈 Conecta o sube la base de **tiendas operando** para empezar.")
    st.stop()

# --- mapeo de columnas ---
with st.sidebar.expander("Mapeo de columnas", expanded=True):
    ct = list(df_t.columns)
    c_id = st.selectbox("Tienda: ID", ct, guess(ct, ["id", "cod", "tienda"]))
    c_nom = st.selectbox("Tienda: nombre", ct, guess(ct, ["nombre", "name", "tienda"]))
    c_lat = st.selectbox("Latitud (Y)", ct, guess(ct, ["y", "lat", "latitude", "latitud"]))
    c_lon = st.selectbox("Longitud (X)", ct, guess(ct, ["x", "lon", "lng", "long", "longitude", "longitud"]))
    if manz is not None:
        cm_ = [c for c in manz.columns if c != "geometry"]
        c_pers = st.selectbox("DANE: personas", cm_, guess(cm_, ["tp27_perso", "personas_s", "perso", "pob"]))
        c_viv = st.selectbox("DANE: viviendas", cm_, guess(cm_, ["tvivienda", "vivienda", "viv"]))
        manz[c_pers] = to_num(manz[c_pers]).fillna(0)
        manz[c_viv] = to_num(manz[c_viv]).fillna(0)
        st.caption(f"Capa lista: **{len(manz):,} manzanas** · **{manz[c_pers].sum():,.0f} personas** · **{manz[c_viv].sum():,.0f} viviendas**")
    else:
        c_pers = c_viv = None
    if df_v is not None:
        cv = list(df_v.columns)
        v_id = st.selectbox("Ventas: ID tienda", cv, guess(cv, ["id", "cod", "tienda"]))
        v_fecha = st.selectbox("Ventas: fecha", cv, guess(cv, ["fecha", "date", "mes", "periodo"]))
        v_venta = st.selectbox("Ventas: valor venta", cv, guess(cv, ["venta", "sales"]))
        v_contr = st.selectbox("Ventas: contribución", ["(ninguna)"] + cv, 0)

tiendas = df_t[[c_id, c_nom, c_lat, c_lon]].copy()
tiendas.columns = ["id", "nombre", "lat", "lon"]
tiendas["id"] = tiendas["id"].astype(str)
tiendas["lat"], tiendas["lon"] = to_num(tiendas["lat"]), to_num(tiendas["lon"])
# Si el Book trae X/Y invertidas, corrige automáticamente usando sus rangos geográficos.
if tiendas["lat"].abs().max() > 90 and tiendas["lon"].abs().max() <= 90:
    tiendas[["lat", "lon"]] = tiendas[["lon", "lat"]]
st.caption(f"Coordenadas usadas: **Y → latitud** (`{c_lat}`), **X → longitud** (`{c_lon}`).")
tiendas = tiendas.dropna(subset=["lat", "lon"])
tiendas = tiendas[(tiendas.lat.between(-90, 90)) & (tiendas.lon.between(-180, 180))]

# --- tienda nueva ---
st.subheader("Tienda nueva propuesta")
c1, c2 = st.columns([1, 2])
sel = c1.text_input("Nombre", "Propuesta 1")
coord = c2.text_input("Coordenadas (lat, lon)", "4.587808725598518, -74.1034563011646",
                      help="Pega la coordenada tal cual la copias de Google Maps")
parsed = parse_coord(coord)
if not parsed:
    st.error("Formato no válido. Usa: 4.587808725598518, -74.1034563011646")
    st.stop()
nw = pd.Series(dict(lat=parsed[0], lon=parsed[1]))
new_pt = to_m(Point(nw.lon, nw.lat))

# Las capas nacionales se cargan después de la coordenada y solo en una ventana local.
if manz is None and (remote_layer_url() or st.session_state.get("heavy_layer_path")):
    manz = load_manzanas_near(new_pt, max(radio * 3, 1000))
    if manz is not None and not manz.empty:
        c_pers = next((c for c in manz.columns if "tp27_perso" in str(c).lower() or "personas" in str(c).lower()), None)
        c_viv = next((c for c in manz.columns if "tvivienda" in str(c).lower() or "vivienda" in str(c).lower()), None)
        if c_pers and c_viv:
            manz[c_pers] = to_num(manz[c_pers]).fillna(0)
            manz[c_viv] = to_num(manz[c_viv]).fillna(0)
            st.success(f"Cargadas {len(manz):,} manzanas cercanas desde la capa nacional.")
        else:
            st.error("La capa no contiene TP27_PERSO y TVIVIENDA.")
            manz = None
iso_new = make_iso(nw.lat, nw.lon, radio, modo)

# --- tiendas cercanas ---
tiendas["pt"] = [to_m(Point(x, y)) for x, y in zip(tiendas.lon, tiendas.lat)]
tiendas["dist"] = [distance_m(nw.lat, nw.lon, y, x) for x, y in zip(tiendas.lon, tiendas.lat)]
near5 = tiendas.sort_values("dist").head(5)
if near5.empty:
    st.error("La base de tiendas no tiene coordenadas válidas.")
    st.stop()
n0 = near5.iloc[0]
st.success(f"📍 Tienda operando más cercana: **{n0.nombre}** (ID {n0.id}) a **{n0.dist:.0f} m**")
st.dataframe(near5[["id", "nombre", "dist"]].rename(columns={"id": "ID", "nombre": "Nombre", "dist": "Distancia (m)"})
             .style.format({"Distancia (m)": "{:,.0f}"}), use_container_width=True, hide_index=True)
en_radio = tiendas[tiendas.dist <= radio].sort_values("dist")
cerca = en_radio
st.caption(f"Tiendas dentro del radio de {radio:,.0f} m: **{len(en_radio)}**")
if en_radio.empty:
    st.warning("No hay tiendas del archivo dentro del radio seleccionado. Aumenta el radio o revisa las coordenadas.")

up_real = st.file_uploader("(Opcional) Isócronas reales de encuesta O-D (GeoJSON con columna 'id')", type=["geojson"])
real = gpd.read_file(up_real).to_crs(CRS_M) if up_real else None

rows = []
for _, t in cerca.iterrows():
    iso = make_iso(t.lat, t.lon, radio, modo)
    if real is not None and "id" in real.columns and (real["id"].astype(str) == t.id).any():
        iso = real[real["id"].astype(str) == t.id].geometry.iloc[0]
    inter = iso.intersection(iso_new)
    if inter.is_empty:
        continue
    pe, ve = pop_in(iso, manz, c_pers, c_viv)
    ps, vs = pop_in(inter, manz, c_pers, c_viv)
    share = (ps / pe) if pe and not np.isnan(pe) else inter.area / iso.area
    rows.append(dict(id=t.id, nombre=t.nombre, pt=t.pt, iso=iso, dist=t.dist, pe=pe, ve=ve, ps=ps, vs=vs, share=share))

# --- mapa interactivo ---
st.subheader("Mapa")
m = folium.Map(location=[nw.lat, nw.lon], zoom_start=17, tiles="OpenStreetMap")
if manz is not None:
    near = manz[manz.intersects(new_pt.buffer(radio * 3))]
    if not near.empty:
        n4326 = near.to_crs(4326)[[c_pers, c_viv, "geometry"]]
        folium.GeoJson(n4326, name="Manzanas DANE", style_function=lambda f: dict(color="#777", weight=1, fillOpacity=0.05),
                       tooltip=folium.GeoJsonTooltip(fields=[c_pers, c_viv], aliases=["Personas", "Viviendas"])).add_to(m)
folium.GeoJson(to_ll(iso_new).__geo_interface__, name="Isócrona nueva",
               style_function=lambda f: dict(color="#e63946", fillColor="#e63946", fillOpacity=0.25)).add_to(m)
folium.Circle([nw.lat, nw.lon], radius=radio, color="#e63946", weight=2, dash_array="6",
              fill=False, tooltip=f"Radio de referencia: {radio:.0f} m").add_to(m)
folium.Marker([nw.lat, nw.lon], tooltip=f"NUEVA: {sel}", icon=folium.Icon(color="red", icon="star")).add_to(m)
row_by_id = {str(r["id"]): r for r in rows}
# Ubicar todas las tiendas del Book; las dentro del radio quedan resaltadas.
map_stores = tiendas.sort_values("dist")
for _, t in map_stores.iterrows():
    r = row_by_id.get(str(t.id))
    inside = t.dist <= radio
    color = "blue" if inside else "gray"
    line_color = "#1d6fd6" if inside else "#999999"
    line_style = "6" if inside else "2"
    label = "dentro del radio" if inside else "fuera del radio"
    folium.Marker(
        [t.lat, t.lon],
        tooltip=f"{t.nombre} · {t.dist:.0f} m · {label}",
        popup=folium.Popup(f"<b>{t.nombre}</b><br>Distancia: {t.dist:.0f} m<br>{label.title()}", max_width=260),
        icon=folium.Icon(color=color, icon="shopping-cart"),
    ).add_to(m)
    folium.PolyLine(
        [[nw.lat, nw.lon], [t.lat, t.lon]],
        color=line_color, weight=3 if inside else 1, dash_array=line_style,
        tooltip=f"Distancia a {t.nombre}: {t.dist:.0f} m",
    ).add_to(m)
    mid_lat, mid_lon = (nw.lat + t.lat) / 2, (nw.lon + t.lon) / 2
    folium.Marker(
        [mid_lat, mid_lon],
        icon=folium.DivIcon(html=f'<div style="font-size:11px;font-weight:700;color:{line_color};background:white;border:1px solid {line_color};border-radius:4px;padding:2px 4px;white-space:nowrap">{t.dist:.0f} m</div>'),
    ).add_to(m)
    if r:
        folium.GeoJson(to_ll(r["iso"]).__geo_interface__, name=f"Isócrona {t.nombre}",
                       style_function=lambda f: dict(color="#1d6fd6", fillColor="#1d6fd6", fillOpacity=0.2)).add_to(m)
folium.LayerControl().add_to(m)
st_folium(m, height=520, use_container_width=True, returned_objects=[])

if not rows:
    st.success("No hay tiendas operando con isócronas traslapadas con la nueva: canibalización ≈ 0.")
    st.stop()

# --- tabla editable ---
st.subheader("Canibalización sugerida (editable)")
tab = pd.DataFrame([{
    "ID": r["id"], "Nombre": r["nombre"], "Distancia (m)": round(r["dist"]),
    "Pob. isócrona existente": r["pe"], "Viv. isócrona existente": r["ve"],
    "Pob. compartida": r["ps"], "Viv. compartida": r["vs"], "% traslape": r["share"] * 100,
    "% sugerido": r["share"] * mig * 100, "% final (editable)": round(r["share"] * mig * 100, 1)} for r in rows])
tab = st.data_editor(tab, use_container_width=True, disabled=[c for c in tab.columns if c != "% final (editable)"], key=f"tab_{sel}_{radio}_{mig}")
pct_map = dict(zip(tab["ID"], tab["% final (editable)"] / 100))

st.subheader("Habitantes y viviendas alcanzados")
new_pe, new_ve = pop_in(iso_new, manz, c_pers, c_viv)
new_blocks = manzanas_in(iso_new, manz, c_pers, c_viv)
a, b, c_ = st.columns(3)
a.metric("Nueva · habitantes", f"{new_pe:,.0f}")
b.metric("Nueva · viviendas", f"{new_ve:,.0f}")
c_.metric("Manzanas tocadas", f"{len(new_blocks) if new_blocks is not None else 0:,}")

detalle = []
for r in rows:
    detalle.append({"Tienda operando": r["nombre"], "Habitantes en isócrona": r["pe"], "Viviendas en isócrona": r["ve"],
                    "Habitantes compartidos": r["ps"], "Viviendas compartidas": r["vs"]})
if detalle:
    st.dataframe(pd.DataFrame(detalle).style.format("{:,.0f}", subset=pd.IndexSlice[:, 1:]), use_container_width=True, hide_index=True)

# --- impacto ---
imp = tot = chart = None
if df_v is not None:
    sv = df_v[[v_id, v_fecha, v_venta] + ([v_contr] if v_contr != "(ninguna)" else [])].copy()
    sv.columns = ["id", "fecha", "venta"] + (["contr"] if v_contr != "(ninguna)" else [])
    sv["id"] = sv["id"].astype(str)
    sv["fecha"] = pd.to_datetime(sv["fecha"], errors="coerce")
    sv["venta"] = to_num(sv["venta"])
    if "contr" not in sv:
        sv["contr"] = 0.0
    sv["contr"] = to_num(sv["contr"])
    sv = sv.dropna(subset=["fecha"])
    sv["mes"] = sv["fecha"].dt.to_period("M")
    ult = sv["mes"].max()
    m3, m6 = [ult - i for i in range(3)], [ult - i for i in range(5, -1, -1)]
    s3 = sv[sv.mes.isin(m3) & sv.id.isin(pct_map)].groupby("id")[["venta", "contr"]].sum()
    out = []
    for _, r in tab.iterrows():
        v, c = (s3.loc[r["ID"], "venta"], s3.loc[r["ID"], "contr"]) if r["ID"] in s3.index else (0.0, 0.0)
        p = pct_map[r["ID"]]
        out.append({"Nombre": r["Nombre"], "% canib.": p * 100, "Ventas 3M": v, "Impacto ventas": v * p, "Ventas 3M post": v * (1 - p),
                    "Contribución 3M": c, "Impacto contribución": c * p, "Contribución 3M post": c * (1 - p)})
    imp = pd.DataFrame(out)
    tot = dict(v=imp["Ventas 3M"].sum(), iv=imp["Impacto ventas"].sum(), c=imp["Contribución 3M"].sum(), ic=imp["Impacto contribución"].sum())
    st.subheader(f"Impacto últimos 3 meses ({m3[-1]} a {m3[0]})")
    st.dataframe(imp.style.format({c: "{:,.0f}" for c in imp.columns if c not in ("Nombre", "% canib.")} | {"% canib.": "{:.1f}%"}), use_container_width=True)
    a, b, c_ = st.columns(3)
    a.metric("Ventas 3M", cop(tot["v"]), f"-{cop(tot['iv'])}", delta_color="inverse")
    b.metric("Contribución 3M", cop(tot["c"]), f"-{cop(tot['ic'])}", delta_color="inverse")
    c_.metric("% ventas afectado", f"{(tot['iv']/tot['v'] if tot['v'] else 0):.1%}")
    piv = sv[sv.mes.isin(m6) & sv.id.isin(pct_map)].pivot_table(index="mes", columns="id", values="venta", aggfunc="sum").reindex(m6).fillna(0)
    chart = sales_chart(piv, pct_map)
    st.image(chart)
else:
    st.info("Conecta la base de ventas para ver el impacto en ventas y contribución.")

# --- informe ---
st.subheader("Informe")
coment = st.text_area("Comentarios (se imprimen en el informe)", height=120)
if st.button("📄 Generar informe PDF", type="primary"):
    with st.spinner("Generando mapa e informe..."):
        png = static_map(new_pt, iso_new, rows, sel)
        pdf = build_pdf(dict(nombre=sel, lat=nw.lat, lon=nw.lon, radio=radio, modo=modo, mig=mig, map=png, stores=tab,
                             imp=imp, tot=tot, chart=chart, coment=coment))
    st.image(png, caption="Mapa incluido en el informe (© OpenStreetMap contributors)")
    st.download_button("⬇️ Descargar informe", pdf, f"informe_canibalizacion_{sel}.pdf", "application/pdf")
