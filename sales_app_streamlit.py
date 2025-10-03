# sales_app_streamlit.py
import streamlit as st
import pandas as pd
import os
from datetime import date, datetime
from openpyxl import load_workbook

# =========================
# Config
# =========================
EXCEL_FILE = "mimamuni sales datta+.xlsx"  # main workbook
RAW_SHEET  = "Sheet1"                      # where your main table lives
CAT_SHEET  = "Catalogo"                    # persistent catalog sheet (Artículo, Precio)

EXPECTED = ["Fecha","Cantidad","Nombre del Artículo","Método de Pago","Precio Unitario","Venta Total","Comentarios"]

DEFAULT_CATALOG = [
    {"Artículo": "bolsa",   "Precio": 120.0},
    {"Artículo": "jeans",   "Precio": 50.0},
    {"Artículo": "t-shirt", "Precio": 25.0},
    {"Artículo": "jacket",  "Precio": 120.0},
    {"Artículo": "cinturón","Precio": 20.0},
]

st.set_page_config(page_title="Ventas - Tienda de Ropa", page_icon="🛍️", layout="wide")

# ---------- light styling ----------
st.markdown("""
<style>
/* nicer buttons for the tiles */
button[kind="secondary"] {
  border-radius: 16px !important;
  padding: 18px 14px !important;
  min-height: 80px !important;
  white-space: pre-line !important;
  font-weight: 600 !important;
}
/* cardy sections */
.block-container { padding-top: 1.5rem; }
</style>
""", unsafe_allow_html=True)

st.title("🛍️ Registro de Ventas")
st.caption("• Crea/edita artículos • Elige con un clic • Guarda la venta en tu Excel (Sheet1).")

# =======================================================
# Helpers
# =======================================================
def ensure_excel_exists() -> bool:
    return os.path.exists(EXCEL_FILE)

def open_wb():
    if not ensure_excel_exists():
        raise FileNotFoundError("Excel no encontrado. Sube tu archivo primero.")
    return load_workbook(EXCEL_FILE)

def write_sheet_replace(df: pd.DataFrame, sheet_name: str):
    """Hard-replace a sheet with df content (create if missing)."""
    from openpyxl.utils.dataframe import dataframe_to_rows
    wb = open_wb()
    if sheet_name in wb.sheetnames:
        ws_old = wb[sheet_name]; wb.remove(ws_old)
    ws = wb.create_sheet(sheet_name)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    wb.save(EXCEL_FILE)

@st.cache_data(show_spinner=False)
def load_catalog_df() -> pd.DataFrame:
    """Load persistent catalog from Excel. If missing, seed defaults."""
    if not ensure_excel_exists():
        return pd.DataFrame(DEFAULT_CATALOG, columns=["Artículo","Precio"])
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=CAT_SHEET)
        df = df.rename(columns={"Articulo":"Artículo","precio":"Precio"})
        if "Artículo" not in df.columns or "Precio" not in df.columns:
            raise ValueError("Wrong columns in catalog.")
        df["Artículo"] = df["Artículo"].astype(str)
        df["Precio"]   = pd.to_numeric(df["Precio"], errors="coerce").fillna(0.0)
        return df[["Artículo","Precio"]]
    except Exception:
        df = pd.DataFrame(DEFAULT_CATALOG, columns=["Artículo","Precio"])
        try:
            write_sheet_replace(df, CAT_SHEET)
        except Exception:
            pass
        return df

def save_catalog_df(df: pd.DataFrame):
    clean = df.copy()
    clean["Artículo"] = clean["Artículo"].astype(str).str.strip()
    clean["Precio"]   = pd.to_numeric(clean["Precio"], errors="coerce").fillna(0.0)
    clean = clean[clean["Artículo"] != ""]
    clean = clean.drop_duplicates(subset=["Artículo"], keep="last")
    write_sheet_replace(clean[["Artículo","Precio"]], CAT_SHEET)

def find_header_row_and_map(ws):
    """Find the header row on Sheet1 and return (row_idx, {header: col_idx})."""
    max_rows = min(ws.max_row, 200)
    max_cols = min(ws.max_column, 40)
    for r in range(1, max_rows+1):
        vals = [str(ws.cell(r,c).value).strip() if ws.cell(r,c).value is not None else "" for c in range(1, max_cols+1)]
        if {"Fecha","Cantidad","Nombre del Artículo"}.issubset(set(vals)):
            col_map = {}
            for c in range(1, max_cols+1):
                v = ws.cell(r,c).value
                if v is not None and str(v).strip() in EXPECTED:
                    col_map[str(v).strip()] = c
            return r, col_map
    return None, {}

def find_next_empty_data_row(ws, header_row, key_cols):
    """First empty row after header based on key columns."""
    r = header_row + 1
    while r <= ws.max_row:
        if all(ws.cell(r, c).value in (None, "") for c in key_cols.values()):
            return r
        r += 1
    return ws.max_row + 1

def append_sale_to_sheet1(row_dict):
    """Append a sale to the main table in Sheet1."""
    wb = open_wb()
    if RAW_SHEET not in wb.sheetnames:
        raise ValueError(f"No se encontró la hoja {RAW_SHEET}.")
    ws = wb[RAW_SHEET]

    header_row, col_map = find_header_row_and_map(ws)
    if not header_row:
        raise RuntimeError("No se encontraron las cabeceras (Fecha/Cantidad/Nombre del Artículo) en Sheet1.")

    # auto Venta Total
    if "Venta Total" in col_map and (("Venta Total" not in row_dict) or row_dict["Venta Total"] in (None, "")):
        try:
            row_dict["Venta Total"] = float(row_dict.get("Cantidad", 0)) * float(row_dict.get("Precio Unitario", 0))
        except Exception:
            row_dict["Venta Total"] = None

    key_cols = {k: col_map[k] for k in col_map if k in ["Fecha","Cantidad","Nombre del Artículo"]}
    if not key_cols:
        key_cols = col_map
    next_row = find_next_empty_data_row(ws, header_row, key_cols)

    # write
    for h, c in col_map.items():
        val = row_dict.get(h, None)
        if h == "Fecha" and isinstance(val, (date, datetime)):
            ws.cell(next_row, c).value = datetime.combine(val, datetime.min.time())
        else:
            ws.cell(next_row, c).value = val

    wb.save(EXCEL_FILE)

@st.cache_data(show_spinner=False)
def read_current_table() -> pd.DataFrame:
    if not ensure_excel_exists():
        return pd.DataFrame(EXPECTED, columns=EXPECTED).iloc[:0]
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = wb[RAW_SHEET]
        header_row, col_map = find_header_row_and_map(ws)
        if not header_row:
            return pd.DataFrame(columns=EXPECTED)
        rows = []
        r = header_row + 1
        while r <= ws.max_row:
            if all(ws.cell(r, c).value in (None, "") for c in col_map.values()):
                break
            rec = {h: ws.cell(r, c).value for h, c in col_map.items()}
            rows.append(rec); r += 1
        df = pd.DataFrame(rows)
        if "Fecha" in df: df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.date
        for c in ["Cantidad","Precio Unitario","Venta Total"]:
            if c in df: df[c] = pd.to_numeric(df[c], errors="coerce")
        for c in EXPECTED:
            if c not in df: df[c] = None
        return df[EXPECTED]
    except Exception:
        return pd.DataFrame(columns=EXPECTED)

# =======================================================
# Upload / Download
# =======================================================
with st.container():
    st.subheader("📂 Tu archivo Excel")
    up = st.file_uploader("Sube tu Excel (.xlsx). El app escribirá en la tabla principal de Sheet1.", type=["xlsx"])
    if up is not None:
        with open(EXCEL_FILE, "wb") as f: f.write(up.getbuffer())
        st.success("Excel guardado.")
        st.cache_data.clear()
    if ensure_excel_exists():
        with open(EXCEL_FILE, "rb") as f:
            st.download_button("⬇️ Descargar Excel actualizado", f, file_name=EXCEL_FILE,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Aún no has subido el archivo.")

# =======================================================
# Tabs: Ventas | Catálogo
# =======================================================
tab_sales, tab_catalog = st.tabs(["🧾 Registrar venta", "🗂️ Catálogo (añadir/editar/borrar)"])

# ---------- CATALOGO (CRUD persistente) ----------
with tab_catalog:
    st.markdown("Administra tu catálogo. Cambia nombres y precios, añade filas o marca para borrar y **Guarda**.")
    cat_df = load_catalog_df()
    if "Eliminar" not in cat_df.columns:
        cat_df["Eliminar"] = False

    edited = st.data_editor(
        cat_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Artículo": st.column_config.TextColumn(required=True),
            "Precio":   st.column_config.NumberColumn(min_value=0.0, step=1.0, format="%.2f"),
            "Eliminar": st.column_config.CheckboxColumn(help="Marca para borrar esta fila"),
        },
        key="catalog_editor"
    )

    c1, c2, c3 = st.columns([1,1,3])
    with c1:
        if st.button("💾 Guardar catálogo", use_container_width=True):
            to_save = edited.copy()
            if "Eliminar" in to_save.columns:
                to_save = to_save[to_save["Eliminar"] == False].drop(columns=["Eliminar"])
            save_catalog_df(to_save)
            st.success("Catálogo guardado en la hoja 'Catalogo'.")
            st.cache_data.clear()   # refresh tiles immediately
            st.rerun()
    with c2:
        if st.button("↩️ Deshacer cambios", use_container_width=True):
            st.cache_data.clear(); st.rerun()

# ---------- REGISTRAR VENTA ----------
with tab_sales:
    # quick add form (adds to catalog and refreshes tiles instantly)
    with st.expander("➕ Añadir artículo rápido al catálogo", expanded=False):
        with st.form("quick_add_form"):
            q1, q2, q3 = st.columns([2,1,1])
            with q1:
                qa_name = st.text_input("Nombre del artículo", placeholder="bolsa")
            with q2:
                qa_price = st.number_input("Precio", min_value=0.0, value=0.0, step=1.0, format="%.2f")
            with q3:
                save_q = st.form_submit_button("Agregar/Actualizar")
            if save_q and qa_name.strip():
                dfc = load_catalog_df()
                dfc = dfc[["Artículo","Precio"]]
                # update or append
                mask = dfc["Artículo"].str.lower().eq(qa_name.strip().lower())
                if mask.any():
                    dfc.loc[mask, "Precio"] = float(qa_price)
                else:
                    dfc = pd.concat([dfc, pd.DataFrame([{"Artículo": qa_name.strip(), "Precio": float(qa_price)}])], ignore_index=True)
                save_catalog_df(dfc)
                st.success(f"Guardado en catálogo: {qa_name.strip()} → {float(qa_price):.2f}")
                st.cache_data.clear(); st.rerun()

    st.subheader("🧱 ELIGE UN ARTÍCULO")
    tiles = load_catalog_df().sort_values("Artículo").reset_index(drop=True)
    # search
    search = st.text_input("Buscar artículo", placeholder="escribe para filtrar…")
    if search:
        tiles = tiles[tiles["Artículo"].str.contains(search, case=False, na=False)]

    # prepare session state
    if "articulo_sel" not in st.session_state: st.session_state.articulo_sel = ""
    if "precio_sel"   not in st.session_state: st.session_state.precio_sel   = 0.0

    # show tiles (nice grid)
    per_row = 4
    items = list(tiles.itertuples(index=False, name=None))  # (Artículo, Precio)
    for i in range(0, len(items), per_row):
        cols = st.columns(per_row)
        for col, (name, price) in zip(cols, items[i:i+per_row]):
            with col:
                if st.button(f"{name}\n${float(price):.2f}", key=f"tile_{name}", use_container_width=True):
                    st.session_state.articulo_sel = name
                    st.session_state.precio_sel   = float(price)

    st.divider()
    st.subheader("🧾 Formulario de venta (se guarda en Sheet1)")

    left, right = st.columns(2)
    with left:
        fecha    = st.date_input("Fecha", value=date.today())
        cantidad = st.number_input("Cantidad", min_value=1, step=1, value=1)
        articulo = st.text_input("Nombre del Artículo", value=st.session_state.articulo_sel)
    with right:
        metodo      = st.radio("Método de Pago", ["E", "T"], horizontal=True, help="E=Efectivo, T=Tarjeta")
        precio_unit = st.number_input("Precio Unitario", min_value=0.0, step=1.0,
                                      value=float(st.session_state.precio_sel), format="%.2f")
        venta_total = st.number_input("Venta Total (auto)", min_value=0.0, step=1.0,
                                      value=float(cantidad)*float(precio_unit), format="%.2f")

    comentarios = st.text_area("Comentarios (opcional)", value="")

    disabled = (not ensure_excel_exists()) or (not articulo) or (precio_unit <= 0)
    if st.button("💾 Guardar venta en Excel", type="primary", use_container_width=True, disabled=disabled):
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
            st.success("✅ Venta guardada en tu Excel (Sheet1).")
            st.balloons()
            st.session_state.articulo_sel = ""
            st.session_state.precio_sel   = 0.0
            st.cache_data.clear()
        except Exception as e:
            st.error(f"No se pudo escribir en Sheet1: {e}")

    st.divider()
    with st.expander("📊 Vista rápida de ventas (Sheet1)", expanded=False):
        st.dataframe(read_current_table(), use_container_width=True, hide_index=True)
