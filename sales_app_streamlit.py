import streamlit as st
import pandas as pd
import os
from datetime import date
from openpyxl import load_workbook

# =========================
# Config
# =========================
EXCEL_FILE = "mimamuni sales datta+.xlsx"
VENTAS_SHEET = "Ventas"
RAW_SHEET = "Sheet1"

# Default catalog tiles (name -> price). Edit or add your own here.
DEFAULT_CATALOG = {
    "bolsa": 120.0,
    "jeans": 50.0,
    "t-shirt": 25.0,
    "jacket": 120.0,
    "cinturón": 20.0,
}

st.set_page_config(page_title="Ventas - Tienda de Ropa", page_icon="🛍️", layout="centered")
st.title("🛍️ Registro de Ventas")
st.caption("Sube tu Excel una vez. Usa los cuadrados para elegir artículos; el precio se rellena solo.")

# -------------------------
# Session state helpers
# -------------------------
if "catalog" not in st.session_state:
    st.session_state.catalog = DEFAULT_CATALOG.copy()
if "articulo" not in st.session_state:
    st.session_state.articulo = ""
if "precio_unit" not in st.session_state:
    st.session_state.precio_unit = 0.0

def ensure_excel_exists() -> bool:
    return os.path.exists(EXCEL_FILE)

@st.cache_data(show_spinner=False)
def load_table() -> pd.DataFrame:
    if not ensure_excel_exists():
        return pd.DataFrame(columns=["Fecha","Cantidad","Nombre del Artículo","Método de Pago","Precio Unitario","Venta Total","Comentarios"])
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=VENTAS_SHEET)
        df.columns = [str(c) for c in df.columns]
        return df
    except Exception:
        try:
            raw = pd.read_excel(EXCEL_FILE, sheet_name=RAW_SHEET, header=None)
            header_row_idx = None
            for i in range(len(raw)):
                vals = raw.iloc[i].astype(str).tolist()
                if {"Fecha","Cantidad","Nombre del Artículo"}.issubset(set(vals)):
                    header_row_idx = i; break
            if header_row_idx is None:
                raise RuntimeError("Cabeceras no encontradas en RAW.")
            headers = raw.iloc[header_row_idx].tolist()
            table = raw.iloc[header_row_idx+1:].copy()
            table.columns = headers
            table = table.dropna(axis=1, how="all")
            key = [c for c in ["Fecha","Cantidad","Nombre del Artículo"] if c in table.columns]
            table = table.dropna(subset=key, how="all")
            cols = ["Fecha","Cantidad","Nombre del Artículo","Método de Pago","Precio Unitario","Venta Total","Comentarios"]
            df = table[[c for c in cols if c in table.columns]].copy()
            df = df.loc[:, ~df.columns.duplicated()]
            if "Fecha" in df: df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.date
            if "Cantidad" in df: df["Cantidad"] = pd.to_numeric(df["Cantidad"], errors="coerce").astype("Int64")
            for c in ["Precio Unitario","Venta Total"]:
                if c in df: df[c] = pd.to_numeric(df[c], errors="coerce")
            for c in cols:
                if c not in df: df[c] = None
            return df[cols]
        except Exception:
            return pd.DataFrame(columns=["Fecha","Cantidad","Nombre del Artículo","Método de Pago","Precio Unitario","Venta Total","Comentarios"])

def save_append_row(row: dict):
    if ensure_excel_exists():
        try:
            book = load_workbook(EXCEL_FILE)
            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                writer.book = book
                if VENTAS_SHEET in writer.book.sheetnames:
                    existing = pd.read_excel(EXCEL_FILE, sheet_name=VENTAS_SHEET)
                    existing.columns = [str(c) for c in existing.columns]
                    new_df = pd.concat([existing, pd.DataFrame([row])], ignore_index=True)
                    idx = writer.book.sheetnames.index(VENTAS_SHEET)
                    ws = writer.book.worksheets[idx]
                    writer.book.remove(ws)
                    writer.book.create_sheet(VENTAS_SHEET, idx)
                    new_df.to_excel(writer, sheet_name=VENTAS_SHEET, index=False)
                else:
                    pd.DataFrame([row]).to_excel(writer, sheet_name=VENTAS_SHEET, index=False)
            return
        except Exception as e:
            st.error(f"Error guardando en Excel: {e}")
            return
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        pd.DataFrame([row]).to_excel(writer, sheet_name=VENTAS_SHEET, index=False)

# -------------------------
# Upload / download Excel
# -------------------------
st.subheader("📂 Tu archivo Excel")
uploaded = st.file_uploader("Sube tu Excel (.xlsx) — se guardará como “mimamuni sales datta+.xlsx”", type=["xlsx"])
if uploaded is not None:
    with open(EXCEL_FILE, "wb") as f:
        f.write(uploaded.getbuffer())
    st.success("Excel guardado. Ya puedes usar el formulario y los cuadrados de artículos.")
    st.cache_data.clear()

if ensure_excel_exists():
    with open(EXCEL_FILE, "rb") as f:
        st.download_button("⬇️ Descargar Excel actualizado", f, file_name=EXCEL_FILE, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Aún no has subido el archivo. Sube tu Excel arriba para empezar.")

# -------------------------
# Catalog management (sidebar)
# -------------------------
with st.sidebar:
    st.header("🗂️ Catálogo de artículos")
    st.caption("Añade/edita tus tiles (nombre + precio).")
    with st.form("add_item"):
        new_name = st.text_input("Nombre del artículo", placeholder="bolsa")
        new_price = st.number_input("Precio", min_value=0.0, step=1.0, value=0.0, format="%.2f")
        add_btn = st.form_submit_button("Agregar/Actualizar")
        if add_btn and new_name.strip():
            st.session_state.catalog[new_name.strip()] = float(new_price)
            st.success(f"Guardado: {new_name.strip()} → {float(new_price):.2f}")

    if st.session_state.catalog:
        df_cat = pd.DataFrame(
            [{"Artículo": k, "Precio": v} for k, v in sorted(st.session_state.catalog.items())]
        )
        st.dataframe(df_cat, use_container_width=True, hide_index=True)

# -------------------------
# Data preview
# -------------------------
df = load_table()
with st.expander("📊 Ver datos actuales", expanded=False):
    st.dataframe(df, use_container_width=True)

# -------------------------
# Catalog tiles (squares)
# -------------------------
st.subheader("🧱 Elige un artículo")
tiles_per_row = 3
items = list(st.session_state.catalog.items())
for i in range(0, len(items), tiles_per_row):
    cols = st.columns(tiles_per_row)
    for col, (name, price) in zip(cols, items[i:i+tiles_per_row]):
        with col:
            clicked = st.button(f"{name}\n${price:.2f}", key=f"tile_{name}", use_container_width=True)
            if clicked:
                st.session_state.articulo = name
                st.session_state.precio_unit = float(price)

st.divider()

# -------------------------
# Entry form
# -------------------------
st.subheader("➕ Añadir una venta")

col1, col2 = st.columns(2)
with col1:
    fecha = st.date_input("Fecha", value=date.today())
    cantidad = st.number_input("Cantidad", min_value=1, step=1, value=1)
    # Text input prefilled from clicked tile
    articulo = st.text_input("Nombre del Artículo", value=st.session_state.articulo)
with col2:
    metodo = st.radio("Método de Pago", ["E", "T"], horizontal=True, help="E=Efectivo, T=Tarjeta")
    precio_unit = st.number_input(
        "Precio Unitario",
        min_value=0.0, step=1.0,
        value=float(st.session_state.precio_unit) if st.session_state.precio_unit else 0.0,
        format="%.2f",
    )
    venta_total = st.number_input(
        "Venta Total (auto)",
        min_value=0.0, step=1.0,
        value=float(cantidad) * float(precio_unit),
        format="%.2f",
        help="Se calcula automáticamente; puedes ajustarlo si lo necesitas."
    )

comentarios = st.text_area("Comentarios (opcional)", value="")

disabled = (not ensure_excel_exists()) or (not articulo) or (precio_unit <= 0)
if st.button("Guardar en Excel", type="primary", use_container_width=True, disabled=disabled):
    new_row = {
        "Fecha": fecha,
        "Cantidad": int(cantidad),
        "Nombre del Artículo": articulo,
        "Método de Pago": metodo,
        "Precio Unitario": float(precio_unit),
        "Venta Total": float(venta_total),
        "Comentarios": comentarios.strip() or None,
    }
    save_append_row(new_row)
    st.success("✅ Venta guardada.")
    # Clear quick-fill so next entry doesn't keep old values
    st.session_state.articulo = ""
    st.session_state.precio_unit = 0.0
    st.cache_data.clear()
    st.rerun()

st.caption("Consejo: usa el panel lateral para agregar más tiles (artículo + precio). Haz clic en un tile para rellenar el formulario al instante.")
