import streamlit as st
import pandas as pd
import os
from datetime import date, datetime
from openpyxl import load_workbook

# =========================
# Config
# =========================
EXCEL_FILE = "mimamuni sales datta+.xlsx"
RAW_SHEET = "Sheet1"     # donde está tu tabla principal de ventas
CAT_SHEET = "Catalogo"   # hoja nueva para guardar Artículo/Precio de forma persistente

# Catálogo por defecto si no existe aún en el Excel
DEFAULT_CATALOG = [
    {"Artículo": "bolsa", "Precio": 120.0},
    {"Artículo": "jeans", "Precio": 50.0},
    {"Artículo": "t-shirt", "Precio": 25.0},
    {"Artículo": "jacket", "Precio": 120.0},
    {"Artículo": "cinturón", "Precio": 20.0},
]

EXPECTED = ["Fecha","Cantidad","Nombre del Artículo","Método de Pago","Precio Unitario","Venta Total","Comentarios"]

st.set_page_config(page_title="Ventas - Tienda de Ropa", page_icon="🛍️", layout="wide")
st.title("🛍️ Registro de Ventas")
st.caption("El catálogo se guarda en la hoja 'Catalogo' del mismo Excel. Las ventas se agregan a la tabla de Sheet1.")

# ----------------- utilidades Excel -----------------
def ensure_excel_exists() -> bool:
    return os.path.exists(EXCEL_FILE)

def open_wb():
    if not ensure_excel_exists():
        raise FileNotFoundError("Excel no encontrado. Sube tu archivo primero.")
    return load_workbook(EXCEL_FILE)

def write_sheet_replace(df: pd.DataFrame, sheet_name: str):
    """Reemplaza completamente una hoja por el contenido de df (creándola si no existe)."""
    wb = open_wb()
    from openpyxl.utils.dataframe import dataframe_to_rows

    # Si existe, elimínala para escribir limpia
    if sheet_name in wb.sheetnames:
        ws_old = wb[sheet_name]
        wb.remove(ws_old)
    ws = wb.create_sheet(sheet_name)

    # Escribir encabezados + filas
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    wb.save(EXCEL_FILE)

def load_catalog_df() -> pd.DataFrame:
    """Carga el catálogo desde la hoja Catalogo; si no existe, lo crea por defecto."""
    if not ensure_excel_exists():
        # Excel aún no subido: devolver df vacío con columnas correctas
        return pd.DataFrame(DEFAULT_CATALOG, columns=["Artículo","Precio"])
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=CAT_SHEET)
        # normalizar columnas
        df = df.rename(columns={"Articulo":"Artículo","precio":"Precio"})
        # forzar tipos
        if "Artículo" not in df.columns or "Precio" not in df.columns:
            df = pd.DataFrame(DEFAULT_CATALOG, columns=["Artículo","Precio"])
        df["Artículo"] = df["Artículo"].astype(str)
        df["Precio"] = pd.to_numeric(df["Precio"], errors="coerce").fillna(0.0)
        return df[["Artículo","Precio"]]
    except Exception:
        # crear por primera vez
        df = pd.DataFrame(DEFAULT_CATALOG, columns=["Artículo","Precio"])
        try:
            write_sheet_replace(df, CAT_SHEET)
        except Exception:
            pass
        return df

def save_catalog_df(df: pd.DataFrame):
    """Guarda el catálogo (limpio, sin duplicados vacíos) en la hoja Catalogo."""
    clean = df.copy()
    clean["Artículo"] = clean["Artículo"].astype(str).str.strip()
    clean["Precio"] = pd.to_numeric(clean["Precio"], errors="coerce").fillna(0.0)
    # eliminar filas totalmente vacías
    clean = clean[clean["Artículo"] != ""]
    # opcional: dejar solo la última ocurrencia por nombre
    clean = clean.drop_duplicates(subset=["Artículo"], keep="last")
    write_sheet_replace(clean[["Artículo","Precio"]], CAT_SHEET)

# ---- detectar cabeceras en Sheet1 y añadir venta debajo ----
def find_header_row_and_map(ws):
    max_rows = min(ws.max_row, 200)
    max_cols = min(ws.max_column, 30)
    for r in range(1, max_rows+1):
        vals = [str(ws.cell(r, c).value).strip() if ws.cell(r,c).value is not None else "" for c in range(1, max_cols+1)]
        if {"Fecha","Cantidad","Nombre del Artículo"}.issubset(set(vals)):
            col_map = {}
            for c in range(1, max_cols+1):
                val = ws.cell(r, c).value
                if val is not None:
                    name = str(val).strip()
                    if name in EXPECTED:
                        col_map[name] = c
            return r, col_map
    return None, {}

def find_next_empty_data_row(ws, header_row, key_cols):
    start = header_row + 1
    r = start
    while r <= ws.max_row:
        empty = True
        for col in key_cols.values():
            if ws.cell(r, col).value not in (None, ""):
                empty = False
                break
        if empty:
            return r
        r += 1
    return ws.max_row + 1

def append_sale_to_sheet1(row_dict):
    wb = open_wb()
    if RAW_SHEET not in wb.sheetnames:
        raise ValueError(f"No se encontró la hoja {RAW_SHEET}.")
    ws = wb[RAW_SHEET]

    header_row, col_map = find_header_row_and_map(ws)
    if not header_row:
        raise RuntimeError("No se encontraron las cabeceras (Fecha/Cantidad/Nombre del Artículo) en Sheet1.")

    # calcular Venta Total si falta
    if "Venta Total" in col_map and (("Venta Total" not in row_dict) or row_dict["Venta Total"] in (None, "")):
        try:
            row_dict["Venta Total"] = float(row_dict.get("Cantidad", 0)) * float(row_dict.get("Precio Unitario", 0))
        except Exception:
            row_dict["Venta Total"] = None

    key_cols = {k: col_map[k] for k in col_map if k in ["Fecha","Cantidad","Nombre del Artículo"]}
    if not key_cols:
        key_cols = col_map
    next_row = find_next_empty_data_row(ws, header_row, key_cols)

    for h, c in col_map.items():
        val = row_dict.get(h, None)
        if h == "Fecha" and isinstance(val, (date, datetime)):
            ws.cell(next_row, c).value = datetime.combine(val, datetime.min.time())
        else:
            ws.cell(next_row, c).value = val

    wb.save(EXCEL_FILE)

# =========================
# Subir/descargar Excel
# =========================
st.subheader("📂 Tu archivo Excel")
uploaded = st.file_uploader("Sube tu Excel (.xlsx). Se usará la tabla de 'Sheet1' y se guardará el catálogo en 'Catalogo'.", type=["xlsx"])
if uploaded is not None:
    with open(EXCEL_FILE, "wb") as f:
        f.write(uploaded.getbuffer())
    st.success("Excel guardado.")
    st.cache_data.clear()

if ensure_excel_exists():
    with open(EXCEL_FILE, "rb") as f:
        st.download_button("⬇️ Descargar Excel actualizado", f, file_name=EXCEL_FILE, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Aún no has subido el archivo.")

# =========================
# Catálogo: CRUD persistente
# =========================
st.divider()
st.subheader("🗂️ Catálogo de Artículo y Precio (se guarda en la hoja 'Catalogo')")
catalog_df = load_catalog_df()

# Agregar columna "Eliminar" para marcar filas a borrar en la UI (no se guarda)
if "Eliminar" not in catalog_df.columns:
    catalog_df["Eliminar"] = False

edited_df = st.data_editor(
    catalog_df,
    num_rows="dynamic",           # permite añadir filas vacías desde la UI
    use_container_width=True,
    column_config={
        "Artículo": st.column_config.TextColumn(required=True),
        "Precio": st.column_config.NumberColumn(min_value=0.0, step=1.0, format="%.2f"),
        "Eliminar": st.column_config.CheckboxColumn(help="Marca para borrar esta fila"),
    },
    hide_index=True,
    key="catalog_editor"
)

colA, colB, colC = st.columns([1,1,2])
with colA:
    if st.button("💾 Guardar catálogo"):
        # borrar marcadas y guardar
        to_save = edited_df.copy()
        if "Eliminar" in to_save.columns:
            to_save = to_save[to_save["Eliminar"] == False].drop(columns=["Eliminar"])
        save_catalog_df(to_save)
        st.success("Catálogo guardado en la hoja 'Catalogo'.")
with colB:
    if st.button("↩️ Deshacer cambios (recargar)"):
        st.cache_data.clear()
        st.rerun()

# =========================
# Cuadrados (tiles) desde el catálogo guardado
# =========================
tiles_df = edited_df.copy()
if "Eliminar" in tiles_df.columns:
    tiles_df = tiles_df[tiles_df["Eliminar"] == False]
tiles_df = tiles_df.dropna(subset=["Artículo"])
tiles_df["Precio"] = pd.to_numeric(tiles_df["Precio"], errors="coerce").fillna(0.0)

if "articulo_sel" not in st.session_state: st.session_state.articulo_sel = ""
if "precio_sel" not in st.session_state: st.session_state.precio_sel = 0.0

st.subheader("🧱 Elige un artículo")
cols_per_row = 4
items = list(tiles_df[["Artículo","Precio"]].itertuples(index=False, name=None))
for i in range(0, len(items), cols_per_row):
    cols = st.columns(cols_per_row)
    for col, (name, price) in zip(cols, items[i:i+cols_per_row]):
        with col:
            if st.button(f"{name}\n${float(price):.2f}", key=f"tile_{name}", use_container_width=True):
                st.session_state.articulo_sel = name
                st.session_state.precio_sel = float(price)

# =========================
# Formulario de venta -> Sheet1
# =========================
st.divider()
st.subheader("➕ Añadir venta a la tabla de Sheet1")

c1, c2 = st.columns(2)
with c1:
    fecha = st.date_input("Fecha", value=date.today())
    cantidad = st.number_input("Cantidad", min_value=1, step=1, value=1)
    articulo = st.text_input("Nombre del Artículo", value=st.session_state.articulo_sel)
with c2:
    metodo = st.radio("Método de Pago", ["E", "T"], horizontal=True, help="E=Efectivo, T=Tarjeta")
    precio_unit = st.number_input("Precio Unitario", min_value=0.0, step=1.0, value=float(st.session_state.precio_sel), format="%.2f")
    venta_total = st.number_input("Venta Total (auto)", min_value=0.0, step=1.0, value=float(cantidad)*float(precio_unit), format="%.2f")

comentarios = st.text_area("Comentarios (opcional)", value="")

disabled = (not ensure_excel_exists()) or (not articulo) or (precio_unit <= 0)
if st.button("Guardar en Sheet1", type="primary", use_container_width=True, disabled=disabled):
    try:
        append_sale_to_sheet1({
            "Fecha": fecha,
            "Cantidad": int(cantidad),
            "Nombre del Artículo": articulo,
            "Método de Pago": metodo,
            "Precio Unitario": float(precio_unit),
            "Venta Total": float(venta_total),
            "Comentarios": (comentarios or "").strip() or None,
        })
        st.success("✅ Venta agregada en la tabla principal de Sheet1.")
        # limpiar selección para el siguiente registro
        st.session_state.articulo_sel = ""
        st.session_state.precio_sel = 0.0
        st.cache_data.clear()
        st.rerun()
    except Exception as e:
        st.error(f"No se pudo escribir en Sheet1: {e}")

st.caption("En 'Catálogo de Artículo y Precio' puedes añadir, editar o borrar entradas y luego pulsa 'Guardar catálogo'.")
