# sales_app_streamlit.py  — robust write to your main Excel
import streamlit as st
import pandas as pd
import os
from datetime import date, datetime
from openpyxl import load_workbook
import unicodedata

EXCEL_FILE = "mimamuni sales datta+.xlsx"   # your workbook filename
DEFAULT_SHEET = "Sheet1"                    # default target sheet
CAT_SHEET = "Catalogo"                      # persistent catalog

EXPECTED = ["Fecha","Cantidad","Nombre del Artículo","Método de Pago","Precio Unitario","Venta Total","Comentarios"]

DEFAULT_CATALOG = [
    {"Artículo": "bolsa",   "Precio": 120.0},
    {"Artículo": "jeans",   "Precio": 50.0},
    {"Artículo": "t-shirt", "Precio": 25.0},
    {"Artículo": "jacket",  "Precio": 120.0},
    {"Artículo": "cinturón","Precio": 20.0},
]

st.set_page_config(page_title="Ventas - Tienda de Ropa", page_icon="🛍️", layout="wide")

# --------- small style polish ---------
st.markdown("""
<style>
button[kind="secondary"]{border-radius:16px;padding:16px 12px;min-height:80px;white-space:pre-line;font-weight:600}
.block-container{padding-top:1.2rem}
</style>
""", unsafe_allow_html=True)

st.title("🛍️ Registro de Ventas")
st.caption("Añade/edita artículos, elige con un clic y guarda la venta en tu Excel principal.")

# ================== helpers ==================
def ensure_excel_exists() -> bool:
    return os.path.exists(EXCEL_FILE)

def open_wb():
    if not ensure_excel_exists():
        raise FileNotFoundError("Excel no encontrado. Sube tu archivo primero.")
    return load_workbook(EXCEL_FILE)

def strip_accents_lower(s:str) -> str:
    if s is None: return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode("ascii")
    return s.strip().lower()

# Map many header variants -> standard header
HEADER_SYNONYMS = {
    "fecha": "Fecha",
    "cantidad": "Cantidad",
    "nombre del articulo": "Nombre del Artículo",
    "nombre del artículo": "Nombre del Artículo",
    "articulo": "Nombre del Artículo",
    "artículo": "Nombre del Artículo",
    "metodo de pago": "Método de Pago",
    "metodo pago": "Método de Pago",
    "precio unitario": "Precio Unitario",
    "venta total": "Venta Total",
    "comentarios": "Comentarios",
    "comentario": "Comentarios",
}

def write_sheet_replace(df: pd.DataFrame, sheet_name: str):
    from openpyxl.utils.dataframe import dataframe_to_rows
    wb = open_wb()
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    ws = wb.create_sheet(sheet_name)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    wb.save(EXCEL_FILE)

@st.cache_data(show_spinner=False)
def load_catalog_df() -> pd.DataFrame:
    if not ensure_excel_exists():
        return pd.DataFrame(DEFAULT_CATALOG, columns=["Artículo","Precio"])
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=CAT_SHEET)
        df = df.rename(columns={"Articulo":"Artículo","precio":"Precio"})
        if "Artículo" not in df or "Precio" not in df:
            raise ValueError
        df["Artículo"] = df["Artículo"].astype(str)
        df["Precio"] = pd.to_numeric(df["Precio"], errors="coerce").fillna(0.0)
        return df[["Artículo","Precio"]]
    except Exception:
        df = pd.DataFrame(DEFAULT_CATALOG, columns=["Artículo","Precio"])
        try: write_sheet_replace(df, CAT_SHEET)
        except Exception: pass
        return df

def save_catalog_df(df: pd.DataFrame):
    clean = df.copy()
    clean["Artículo"] = clean["Artículo"].astype(str).str.strip()
    clean["Precio"] = pd.to_numeric(clean["Precio"], errors="coerce").fillna(0.0)
    clean = clean[clean["Artículo"] != ""].drop_duplicates(subset=["Artículo"], keep="last")
    write_sheet_replace(clean[["Artículo","Precio"]], CAT_SHEET)

def find_header_row_and_map(ws):
    """Detect the header row using tolerant matching (accents/case/extra spaces)."""
    max_rows = min(ws.max_row, 250)
    max_cols = min(ws.max_column, 50)
    for r in range(1, max_rows+1):
        raw_vals = [ws.cell(r,c).value for c in range(1, max_cols+1)]
        canon = [strip_accents_lower(v) for v in raw_vals]
        if {"fecha","cantidad","nombre del articulo"}.issubset(set(canon)):
            col_map = {}
            for c, raw in enumerate(raw_vals, start=1):
                std = HEADER_SYNONYMS.get(strip_accents_lower(raw))
                if std in EXPECTED:
                    col_map[std] = c
            return r, col_map
    return None, {}

def find_next_empty_data_row(ws, header_row, key_cols):
    """First empty row after header based on key columns (Fecha/Cantidad/Nombre...)."""
    r = header_row + 1
    while r <= ws.max_row:
        if all(ws.cell(r, c).value in (None, "") for c in key_cols.values()):
            return r
        r += 1
    return ws.max_row + 1

def append_sale_to_sheet(sheet_name: str, row_dict: dict) -> dict:
    """Append row to the detected table; returns debug info."""
    wb = open_wb()
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"No se encontró la hoja '{sheet_name}'.")
    ws = wb[sheet_name]

    header_row, col_map = find_header_row_and_map(ws)
    if not header_row:
        raise RuntimeError("No se encontraron cabeceras (Fecha/Cantidad/Nombre del Artículo).")

    # auto-calc Venta Total if missing
    if "Venta Total" in col_map and (("Venta Total" not in row_dict) or row_dict["Venta Total"] in (None, "")):
        try:
            row_dict["Venta Total"] = float(row_dict.get("Cantidad", 0)) * float(row_dict.get("Precio Unitario", 0))
        except Exception:
            row_dict["Venta Total"] = None

    key_cols = {k: col_map[k] for k in col_map if k in ["Fecha","Cantidad","Nombre del Artículo"]}
    if not key_cols: key_cols = col_map
    next_row = find_next_empty_data_row(ws, header_row, key_cols)

    # write values only into known columns
    for h, c in col_map.items():
        val = row_dict.get(h, None)
        if h == "Fecha" and isinstance(val, (date, datetime)):
            ws.cell(next_row, c).value = datetime.combine(val, datetime.min.time())
        else:
            ws.cell(next_row, c).value = val

    wb.save(EXCEL_FILE)
    return {"sheet": sheet_name, "header_row": header_row, "written_row": next_row, "columns_used": list(col_map.keys())}

@st.cache_data(show_spinner=False)
def read_current_table(sheet_name: str) -> pd.DataFrame:
    if not ensure_excel_exists(): return pd.DataFrame(columns=EXPECTED)
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = wb[sheet_name]
        hr, cmap = find_header_row_and_map(ws)
        if not hr: return pd.DataFrame(columns=EXPECTED)
        rows = []
        r = hr + 1
        while r <= ws.max_row:
            if all(ws.cell(r, c).value in (None, "") for c in cmap.values()): break
            rows.append({h: ws.cell(r, c).value for h, c in cmap.items()})
            r += 1
        df = pd.DataFrame(rows)
        if "Fecha" in df: df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.date
        for c in ["Cantidad","Precio Unitario","Venta Total"]:
            if c in df: df[c] = pd.to_numeric(df[c], errors="coerce")
        for c in EXPECTED:
            if c not in df: df[c] = None
        return df[EXPECTED]
    except Exception:
        return pd.DataFrame(columns=EXPECTED)

# ================== Upload / sheet select ==================
with st.container():
    st.subheader("📂 Tu archivo Excel")
    up = st.file_uploader("Sube tu Excel (.xlsx). Las ventas se escribirán en tu tabla principal.", type=["xlsx"])
    if up is not None:
        with open(EXCEL_FILE, "wb") as f: f.write(up.getbuffer())
        st.success("Excel guardado.")
        st.cache_data.clear()

    # Let you PICK the target sheet (default to Sheet1 or first)
    if ensure_excel_exists():
        wb = open_wb()
        sheets = wb.sheetnames
        default = DEFAULT_SHEET if DEFAULT_SHEET in sheets else sheets[0]
        sheet_selected = st.selectbox("Hoja de destino para guardar ventas", sheets, index=sheets.index(default))
        st.session_state["sheet_selected"] = sheet_selected

        with open(EXCEL_FILE, "rb") as f:
            st.download_button("⬇️ Descargar Excel actualizado", f, file_name=EXCEL_FILE,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Aún no has subido el archivo.")
        st.session_state["sheet_selected"] = DEFAULT_SHEET

# ================== Catálogo (CRUD) ==================
st.divider()
st.subheader("🗂️ Catálogo (añadir/editar/borrar)")
catalog_df = load_catalog_df()
if "Eliminar" not in catalog_df.columns: catalog_df["Eliminar"] = False

edited = st.data_editor(
    catalog_df, num_rows="dynamic", hide_index=True, use_container_width=True,
    column_config={
        "Artículo": st.column_config.TextColumn(required=True),
        "Precio": st.column_config.NumberColumn(min_value=0.0, step=1.0, format="%.2f"),
        "Eliminar": st.column_config.CheckboxColumn(help="Marca para borrar esta fila"),
    },
    key="catalog_editor"
)

c1, c2 = st.columns([1,1])
with c1:
    if st.button("💾 Guardar catálogo", use_container_width=True):
        to_save = edited.copy()
        if "Eliminar" in to_save: to_save = to_save[to_save["Eliminar"] == False].drop(columns=["Eliminar"])
        save_catalog_df(to_save)
        st.success("Catálogo guardado en la hoja 'Catalogo'.")
        st.cache_data.clear(); st.rerun()
with c2:
    if st.button("↩️ Deshacer cambios", use_container_width=True):
        st.cache_data.clear(); st.rerun()

# ================== Tiles ==================
st.divider()
st.subheader("🧱 ELIGE UN ARTÍCULO")
tiles = load_catalog_df().sort_values("Artículo").reset_index(drop=True)
search = st.text_input("Buscar artículo", placeholder="escribe para filtrar…")
if search: tiles = tiles[tiles["Artículo"].str.contains(search, case=False, na=False)]

if "articulo_sel" not in st.session_state: st.session_state.articulo_sel = ""
if "precio_sel" not in st.session_state: st.session_state.precio_sel = 0.0

per_row = 4
for i in range(0, len(tiles), per_row):
    cols = st.columns(per_row)
    for col, row in zip(cols, tiles.iloc[i:i+per_row].itertuples(index=False)):
        name, price = row
        with col:
            if st.button(f"{name}\n${float(price):.2f}", key=f"tile_{name}", use_container_width=True):
                st.session_state.articulo_sel = name
                st.session_state.precio_sel = float(price)

# ================== Sales form (writes to Excel) ==================
st.divider()
st.subheader("🧾 Guardar venta en tu Excel (tabla principal)")
left, right = st.columns(2)
with left:
    fecha = st.date_input("Fecha", value=date.today())
    cantidad = st.number_input("Cantidad", min_value=1, step=1, value=1)
    articulo = st.text_input("Nombre del Artículo", value=st.session_state.articulo_sel)
with right:
    metodo = st.radio("Método de Pago", ["E", "T"], horizontal=True, help="E=Efectivo, T=Tarjeta")
    precio_unit = st.number_input("Precio Unitario", min_value=0.0, step=1.0, value=float(st.session_state.precio_sel), format="%.2f")
    venta_total = st.number_input("Venta Total (auto)", min_value=0.0, step=1.0, value=float(cantidad)*float(precio_unit), format="%.2f")

comentarios = st.text_area("Comentarios (opcional)", value="")
disabled = (not ensure_excel_exists()) or (not articulo) or (precio_unit <= 0)

debug_box = st.empty()

if st.button("💾 Guardar venta", type="primary", use_container_width=True, disabled=disabled):
    try:
        sheet_name = st.session_state.get("sheet_selected", DEFAULT_SHEET)
        info = append_sale_to_sheet(sheet_name, {
            "Fecha": fecha,
            "Cantidad": int(cantidad),
            "Nombre del Artículo": articulo,
            "Método de Pago": metodo,
            "Precio Unitario": float(precio_unit),
            "Venta Total": float(venta_total),
            "Comentarios": (comentarios or "").strip() or None,
        })
        st.success("✅ Venta guardada en tu Excel.")
        debug_box.info(f"Guardado en hoja **{info['sheet']}**, fila **{info['written_row']}** (cabeceras detectadas en fila {info['header_row']}).")
        st.balloons()
        st.session_state.articulo_sel = ""
        st.session_state.precio_sel = 0.0
        st.cache_data.clear()
    except Exception as e:
        st.error(f"No se pudo escribir: {e}")

with st.expander("📊 Vista rápida de ventas (lectura de tu tabla)"):
    sheet_name = st.session_state.get("sheet_selected", DEFAULT_SHEET)
    st.dataframe(read_current_table(sheet_name), use_container_width=True, hide_index=True)
