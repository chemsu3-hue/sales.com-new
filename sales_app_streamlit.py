import streamlit as st
import pandas as pd
import os
from datetime import date, datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

EXCEL_FILE = "mimamuni sales datta+.xlsx"
RAW_SHEET = "Sheet1"

# Quick tiles: name -> price (edit or add more)
DEFAULT_CATALOG = {
    "bolsa": 120.0,
    "jeans": 50.0,
    "t-shirt": 25.0,
    "jacket": 120.0,
    "cinturón": 20.0,
}

st.set_page_config(page_title="Ventas - Tienda de Ropa", page_icon="🛍️", layout="centered")
st.title("🛍️ Registro de Ventas")
st.caption("Sube tu Excel. Cada venta se añade en la tabla principal de Sheet1.")

# ----------------- helpers for Sheet1 table -----------------
EXPECTED = ["Fecha","Cantidad","Nombre del Artículo","Método de Pago","Precio Unitario","Venta Total","Comentarios"]

def ensure_excel_exists() -> bool:
    return os.path.exists(EXCEL_FILE)

def find_header_row_and_map(ws):
    """
    Find the row in Sheet1 that contains the headers (Fecha, Cantidad, Nombre del Artículo...).
    Return (header_row_index, {header_name: column_index})
    """
    # scan first ~200 rows/columns to find a row that contains the 3 core headers
    max_rows = min(ws.max_row, 200)
    max_cols = min(ws.max_column, 30)
    for r in range(1, max_rows+1):
        values = [str(ws.cell(r, c).value).strip() if ws.cell(r, c).value is not None else "" for c in range(1, max_cols+1)]
        if {"Fecha","Cantidad","Nombre del Artículo"}.issubset(set(values)):
            # build header->col map
            col_map = {}
            for c in range(1, max_cols+1):
                val = ws.cell(r, c).value
                if val is not None:
                    name = str(val).strip()
                    if name in EXPECTED:
                        col_map[name] = c
            return r, col_map
    return None, {}

def find_next_empty_data_row(ws, header_row, key_col):
    """
    Starting from the row after header_row, find the first 'empty' row based on key_col (e.g., Fecha or Cantidad).
    If key_col is missing, fall back to scanning until a block of empties appears, else ws.max_row + 1.
    """
    start = header_row + 1
    # If the sheet has totals or other blocks below, we still append at the first empty line we find.
    # We'll define "empty" as all Expected columns empty in that row.
    r = start
    while r <= ws.max_row:
        empty = True
        for name, col in key_col.items():
            if ws.cell(r, col).value not in (None, ""):
                empty = False
                break
        if empty:
            return r
        r += 1
    return ws.max_row + 1

def append_sale_to_sheet1(row_dict):
    """
    Append one sale row into the detected table on Sheet1.
    """
    if not ensure_excel_exists():
        raise FileNotFoundError("Excel no encontrado. Sube tu archivo primero.")

    wb = load_workbook(EXCEL_FILE)
    if RAW_SHEET not in wb.sheetnames:
        raise ValueError(f"No se encontró la hoja {RAW_SHEET}.")

    ws = wb[RAW_SHEET]
    header_row, col_map = find_header_row_and_map(ws)
    if not header_row:
        raise RuntimeError("No se encontraron las cabeceras (Fecha/Cantidad/Nombre del Artículo) en Sheet1.")

    # Build a dict of columns that we will write (only ones that exist)
    write_cols = {k: v for k, v in col_map.items() if k in EXPECTED}

    # For 'Venta Total' default to Cantidad * Precio Unitario if missing
    if ("Venta Total" in write_cols) and (("Venta Total" not in row_dict) or (row_dict["Venta Total"] in (None, ""))):
        try:
            row_dict["Venta Total"] = float(row_dict.get("Cantidad", 0)) * float(row_dict.get("Precio Unitario", 0))
        except Exception:
            row_dict["Venta Total"] = None

    # Decide which columns to test for emptiness to find the next row
    key_cols_for_empty = {k: write_cols[k] for k in write_cols if k in ["Fecha","Cantidad","Nombre del Artículo"]}
    if not key_cols_for_empty:
        key_cols_for_empty = write_cols  # fallback

    next_row = find_next_empty_data_row(ws, header_row, key_cols_for_empty)

    # Write values
    for header, col_idx in write_cols.items():
        val = row_dict.get(header, None)
        if header == "Fecha" and isinstance(val, (date, datetime)):
            ws.cell(next_row, col_idx).value = datetime.combine(val, datetime.min.time())
        else:
            ws.cell(next_row, col_idx).value = val

    wb.save(EXCEL_FILE)

@st.cache_data(show_spinner=False)
def read_current_table_as_df():
    if not ensure_excel_exists():
        return pd.DataFrame(columns=EXPECTED)
    try:
        # read Sheet1 as a dataframe by detecting header row
        from openpyxl import load_workbook
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = wb[RAW_SHEET]
        header_row, col_map = find_header_row_and_map(ws)
        if not header_row:
            return pd.DataFrame(columns=EXPECTED)
        # build records until first empty block
        records = []
        r = header_row + 1
        while r <= ws.max_row:
            # Stop if row is completely empty across the expected cols
            if all(ws.cell(r, col_map.get(h, 0)).value in (None, "") for h in col_map):
                break
            rec = {}
            for name, col in col_map.items():
                rec[name] = ws.cell(r, col).value
            records.append(rec)
            r += 1
        df = pd.DataFrame(records)
        # Normalize types
        if "Fecha" in df.columns:
            df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.date
        for c in ["Cantidad", "Precio Unitario", "Venta Total"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")
        for c in EXPECTED:
            if c not in df.columns:
                df[c] = None
        return df[EXPECTED]
    except Exception:
        return pd.DataFrame(columns=EXPECTED)

# ----------------- Upload / Download -----------------
st.subheader("📂 Tu archivo Excel (principal)")
uploaded = st.file_uploader("Sube tu Excel (.xlsx). El app escribirá en la tabla de Sheet1.", type=["xlsx"])
if uploaded is not None:
    with open(EXCEL_FILE, "wb") as f:
        f.write(uploaded.getbuffer())
    st.success("Excel guardado. (Consejo: guarda una copia de seguridad también).")
    st.cache_data.clear()

if ensure_excel_exists():
    with open(EXCEL_FILE, "rb") as f:
        st.download_button("⬇️ Descargar Excel actualizado", f, file_name=EXCEL_FILE, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Aún no has subido el archivo.")

# ----------------- Sidebar: catalog tiles config -----------------
if "catalog" not in st.session_state:
    st.session_state.catalog = DEFAULT_CATALOG.copy()
with st.sidebar:
    st.header("🗂️ Catálogo de artículos")
    with st.form("add_item"):
        n = st.text_input("Artículo", placeholder="bolsa")
        p = st.number_input("Precio", min_value=0.0, step=1.0, value=0.0, format="%.2f")
        ok = st.form_submit_button("Agregar/Actualizar")
        if ok and n.strip():
            st.session_state.catalog[n.strip()] = float(p)
            st.success(f"Guardado: {n.strip()} → {p:.2f}")
    if st.session_state.catalog:
        st.dataframe(pd.DataFrame(sorted(st.session_state.catalog.items()), columns=["Artículo","Precio"]), hide_index=True, use_container_width=True)

# ----------------- Preview current table -----------------
df = read_current_table_as_df()
with st.expander("📊 Ver datos actuales (Sheet1)", expanded=False):
    st.dataframe(df, use_container_width=True)

# ----------------- Tiles -----------------
if "articulo" not in st.session_state: st.session_state.articulo = ""
if "precio_unit" not in st.session_state: st.session_state.precio_unit = 0.0

st.subheader("🧱 Elige un artículo")
cols_per_row = 3
items = list(st.session_state.catalog.items())
for i in range(0, len(items), cols_per_row):
    cols = st.columns(cols_per_row)
    for col, (name, price) in zip(cols, items[i:i+cols_per_row]):
        with col:
            if st.button(f"{name}\n${price:.2f}", key=f"tile_{name}", use_container_width=True):
                st.session_state.articulo = name
                st.session_state.precio_unit = float(price)

st.divider()

# ----------------- Entry form -----------------
st.subheader("➕ Añadir una venta a Sheet1")
c1, c2 = st.columns(2)
with c1:
    fecha = st.date_input("Fecha", value=date.today())
    cantidad = st.number_input("Cantidad", min_value=1, step=1, value=1)
    articulo = st.text_input("Nombre del Artículo", value=st.session_state.articulo)
with c2:
    metodo = st.radio("Método de Pago", ["E", "T"], horizontal=True, help="E=Efectivo, T=Tarjeta")
    precio_unit = st.number_input("Precio Unitario", min_value=0.0, step=1.0, value=float(st.session_state.precio_unit), format="%.2f")
    venta_total = st.number_input("Venta Total (auto)", min_value=0.0, step=1.0, value=float(cantidad)*float(precio_unit), format="%.2f")

comentarios = st.text_area("Comentarios (opcional)", value="")

disabled = (not ensure_excel_exists()) or (not articulo) or (precio_unit <= 0)
if st.button("Guardar en la tabla de Sheet1", type="primary", use_container_width=True, disabled=disabled):
    try:
        append_sale_to_sheet1({
            "Fecha": fecha,
            "Cantidad": int(cantidad),
            "Nombre del Artículo": articulo,
            "Método de Pago": metodo,
            "Precio Unitario": float(precio_unit),
            "Venta Total": float(venta_total),
            "Comentarios": comentarios.strip() or None,
        })
        st.success("✅ Venta añadida en la tabla principal de Sheet1.")
        st.session_state.articulo = ""
        st.session_state.precio_unit = 0.0
        st.cache_data.clear()
        st.rerun()
    except Exception as e:
        st.error(f"No se pudo escribir en Sheet1: {e}")

st.caption("El app detecta la fila de cabeceras en Sheet1 y escribe nuevas filas debajo sin tocar otras celdas del dashboard.")
