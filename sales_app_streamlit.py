# sales_app_streamlit.py — Catalog CRUD + write into Sheet1 table (VENTAS DIARIAS)
import streamlit as st
import pandas as pd
import os, unicodedata
from datetime import date, datetime
from openpyxl import load_workbook

EXCEL_FILE   = "mimamuni sales datta+.xlsx"   # your workbook name
TARGET_SHEET = "Sheet1"                       # where VENTAS DIARIAS lives
CAT_SHEET    = "Catalogo"                     # persistent catalog (Artículo, Precio)

EXPECTED = ["Fecha","Cantidad","Nombre del Artículo","Método de Pago","Precio Unitario","Venta Total","Comentarios"]
DEFAULT_CATALOG = [
    {"Artículo":"bolsa","Precio":120.0},
    {"Artículo":"jeans","Precio":50.0},
    {"Artículo":"t-shirt","Precio":25.0},
    {"Artículo":"jacket","Precio":120.0},
    {"Artículo":"cinturón","Precio":20.0},
]

# ---------- polish ----------
st.set_page_config(page_title="Ventas - Tienda de Ropa", page_icon="🛍️", layout="wide")
st.markdown("""
<style>
button[kind="secondary"]{border-radius:16px;padding:16px 12px;min-height:80px;white-space:pre-line;font-weight:600}
.block-container{padding-top:1.2rem}
</style>
""", unsafe_allow_html=True)

st.title("🛍️ Registro de Ventas")
st.caption("Crea/edita artículos → elige con un clic → guarda la venta en la tabla principal de Sheet1.")

# ---------- helpers ----------
def ensure_excel_exists() -> bool:
    return os.path.exists(EXCEL_FILE)

def open_wb():
    if not ensure_excel_exists():
        raise FileNotFoundError("Excel no encontrado. Sube tu archivo primero.")
    return load_workbook(EXCEL_FILE)

def write_sheet_replace(df: pd.DataFrame, sheet_name: str):
    """Replace a sheet with df (create if missing)."""
    from openpyxl.utils.dataframe import dataframe_to_rows
    wb = open_wb()
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    ws = wb.create_sheet(sheet_name)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    wb.save(EXCEL_FILE)

def strip_accents_lower(s:str) -> str:
    if s is None: return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode("ascii")
    return s.strip().lower()

HEADER_SYNONYMS = {
    "fecha":"Fecha",
    "cantidad":"Cantidad",
    "nombre del articulo":"Nombre del Artículo",
    "nombre del artículo":"Nombre del Artículo",
    "articulo":"Nombre del Artículo",
    "artículo":"Nombre del Artículo",
    "metodo de pago":"Método de Pago",
    "metodo pago":"Método de Pago",
    "precio unitario":"Precio Unitario",
    "venta total":"Venta Total",
    "comentarios":"Comentarios",
    "comentario":"Comentarios",
}

@st.cache_data(show_spinner=False)
def load_catalog_df() -> pd.DataFrame:
    """Load catalog from Excel or seed defaults."""
    if not ensure_excel_exists():
        return pd.DataFrame(DEFAULT_CATALOG, columns=["Artículo","Precio"])
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=CAT_SHEET)
        df = df.rename(columns={"Articulo":"Artículo","precio":"Precio"})
        df["Artículo"] = df["Artículo"].astype(str)
        df["Precio"]   = pd.to_numeric(df["Precio"], errors="coerce").fillna(0.0)
        return df[["Artículo","Precio"]]
    except Exception:
        df = pd.DataFrame(DEFAULT_CATALOG, columns=["Artículo","Precio"])
        try: write_sheet_replace(df, CAT_SHEET)
        except Exception: pass
        return df

def save_catalog_df(df: pd.DataFrame):
    """Persist catalog → Catalogo sheet."""
    clean = df.copy()
    clean["Artículo"] = clean["Artículo"].astype(str).str.strip()
    clean["Precio"]   = pd.to_numeric(clean["Precio"], errors="coerce").fillna(0.0)
    clean = clean[clean["Artículo"]!=""].drop_duplicates(subset=["Artículo"], keep="last")
    write_sheet_replace(clean[["Artículo","Precio"]], CAT_SHEET)

def find_header_row_and_map(ws):
    """Find header row in Sheet1, tolerant to accents/case."""
    max_rows = min(ws.max_row, 250)
    max_cols = min(ws.max_column, 50)
    for r in range(1, max_rows+1):
        raw = [ws.cell(r,c).value for c in range(1, max_cols+1)]
        canon = [strip_accents_lower(v) for v in raw]
        if {"fecha","cantidad","nombre del articulo"}.issubset(set(canon)):
            cmap = {}
            for c, v in enumerate(raw, start=1):
                std = HEADER_SYNONYMS.get(strip_accents_lower(v))
                if std in EXPECTED:
                    cmap[std] = c
            return r, cmap
    return None, {}

def find_next_row_by_fecha(ws, header_row, fecha_col):
    """
    Append directly under the last row where Fecha has a value.
    This ignores pre-filled zeros/formulas in other columns.
    """
    r = header_row + 1
    last = header_row
    while r <= ws.max_row:
        val = ws.cell(r, fecha_col).value
        if val not in (None, ""):  # anything in Fecha counts as filled
            last = r
            r += 1
        else:
            break
    return last + 1

def append_sale_to_sheet(row: dict) -> dict:
    wb = open_wb()
    if TARGET_SHEET not in wb.sheetnames:
        raise ValueError(f"No se encontró la hoja '{TARGET_SHEET}'.")
    ws = wb[TARGET_SHEET]

    header_row, col_map = find_header_row_and_map(ws)
    if not header_row:
        raise RuntimeError("No se detectaron cabeceras (Fecha/Cantidad/Nombre del Artículo).")

    # where to write? → first empty 'Fecha' row
    if "Fecha" not in col_map:
        raise RuntimeError("No se encontró la columna 'Fecha' en la tabla.")
    next_row = find_next_row_by_fecha(ws, header_row, col_map["Fecha"])

    # auto Venta Total if missing
    if "Venta Total" in col_map and (("Venta Total" not in row) or row["Venta Total"] in (None, "")):
        try: row["Venta Total"] = float(row.get("Cantidad",0))*float(row.get("Precio Unitario",0))
        except Exception: row["Venta Total"] = None

    # write only mapped columns
    for h, c in col_map.items():
        v = row.get(h, None)
        if h == "Fecha" and isinstance(v, (date, datetime)):
            ws.cell(next_row, c).value = datetime.combine(v, datetime.min.time())
        else:
            ws.cell(next_row, c).value = v

    wb.save(EXCEL_FILE)
    return {"header_row": header_row, "written_row": next_row, "columns": list(col_map.keys())}

@st.cache_data(show_spinner=False)
def read_current_table() -> pd.DataFrame:
    if not ensure_excel_exists(): return pd.DataFrame(columns=EXPECTED)
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = wb[TARGET_SHEET]
        hr, cmap = find_header_row_and_map(ws)
        if not hr: return pd.DataFrame(columns=EXPECTED)
        rows = []
        r = hr + 1
        while r <= ws.max_row:
            if ws.cell(r, cmap["Fecha"]).value in (None, ""): break
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

# ================= Upload / Download =================
with st.container():
    st.subheader("📂 Tu archivo Excel")
    up = st.file_uploader("Sube tu Excel (.xlsx). Las ventas se guardarán en la tabla de Sheet1 (VENTAS DIARIAS).", type=["xlsx"])
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

# ================= Catalog: CRUD (persists to Excel) =================
st.divider()
st.subheader("🗂️ Catálogo (añadir / editar / borrar)")
cat_df = load_catalog_df()
if "Eliminar" not in cat_df.columns:
    cat_df["Eliminar"] = False

edited = st.data_editor(
    cat_df, num_rows="dynamic", hide_index=True, use_container_width=True,
    column_config={
        "Artículo": st.column_config.TextColumn(required=True),
        "Precio":   st.column_config.NumberColumn(min_value=0.0, step=1.0, format="%.2f"),
        "Eliminar": st.column_config.CheckboxColumn(help="Marca para borrar esta fila"),
    },
    key="catalog_editor"
)

left, right = st.columns([1,1])
with left:
    if st.button("💾 Guardar catálogo", use_container_width=True):
        to_save = edited.copy()
        if "Eliminar" in to_save:
            to_save = to_save[to_save["Eliminar"]==False].drop(columns=["Eliminar"])
        save_catalog_df(to_save)
        st.success("Catálogo guardado en la hoja 'Catalogo'.")
        st.cache_data.clear()     # ensures tiles refresh immediately
        st.rerun()
with right:
    if st.button("↩️ Deshacer cambios", use_container_width=True):
        st.cache_data.clear(); st.rerun()

# ================= Tiles from saved catalog =================
st.divider()
st.subheader("🧱 ELIGE UN ARTÍCULO")
tiles = load_catalog_df().sort_values("Artículo").reset_index(drop=True)
search = st.text_input("Buscar artículo", placeholder="escribe para filtrar…")
if search:
    tiles = tiles[tiles["Artículo"].str.contains(search, case=False, na=False)]

if "articulo_sel" not in st.session_state: st.session_state.articulo_sel = ""
if "precio_sel"   not in st.session_state: st.session_state.precio_sel   = 0.0

per_row = 4
for i in range(0, len(tiles), per_row):
    cols = st.columns(per_row)
    for col, (name, price) in zip(cols, tiles.itertuples(index=False, name=None)[i:i+per_row]):
        with col:
            if st.button(f"{name}\n${float(price):.2f}", key=f"tile_{name}", use_container_width=True):
                st.session_state.articulo_sel = name
                st.session_state.precio_sel   = float(price)

# quick add on the same page (guaranteed to save & refresh)
with st.expander("➕ Añadir artículo rápido al catálogo", expanded=False):
    q1, q2, q3 = st.columns([2,1,1])
    with q1:
        qa_name = st.text_input("Nombre del artículo", placeholder="bolsa", key="qa_name")
    with q2:
        qa_price = st.number_input("Precio", min_value=0.0, value=0.0, step=1.0, format="%.2f", key="qa_price")
    with q3:
        if st.button("Agregar/Actualizar", use_container_width=True, key="qa_btn"):
            dfc = load_catalog_df()
            mask = dfc["Artículo"].str.lower().eq(qa_name.strip().lower())
            if mask.any():
                dfc.loc[mask, "Precio"] = float(qa_price)
            else:
                dfc = pd.concat([dfc, pd.DataFrame([{"Artículo": qa_name.strip(), "Precio": float(qa_price)}])], ignore_index=True)
            save_catalog_df(dfc)
            st.success(f"Guardado en catálogo: {qa_name.strip()} → {float(qa_price):.2f}")
            st.cache_data.clear(); st.rerun()

# ================= Sales form (writes into Sheet1 table) =================
st.divider()
st.subheader("🧾 Guardar venta en la tabla de Sheet1 (VENTAS DIARIAS)")
c1, c2 = st.columns(2)
with c1:
    fecha    = st.date_input("Fecha", value=date.today())
    cantidad = st.number_input("Cantidad", min_value=1, step=1, value=1)
    articulo = st.text_input("Nombre del Artículo", value=st.session_state.articulo_sel)
with c2:
    metodo      = st.radio("Método de Pago", ["E","T"], horizontal=True, help="E=Efectivo, T=Tarjeta")
    precio_unit = st.number_input("Precio Unitario", min_value=0.0, step=1.0, value=float(st.session_state.precio_sel), format="%.2f")
    venta_total = st.number_input("Venta Total (auto)", min_value=0.0, step=1.0, value=float(cantidad)*float(precio_unit), format="%.2f")
comentarios = st.text_area("Comentarios (opcional)", value="")

disabled = (not ensure_excel_exists()) or (not articulo) or (precio_unit <= 0)
debug_box = st.empty()

if st.button("💾 Guardar venta", type="primary", use_container_width=True, disabled=disabled):
    try:
        info = append_sale_to_sheet({
            "Fecha": fecha,
            "Cantidad": int(cantidad),
            "Nombre del Artículo": articulo,
            "Método de Pago": metodo,
            "Precio Unitario": float(precio_unit),
            "Venta Total": float(venta_total),
            "Comentarios": (comentarios or "").strip() or None,
        })
        st.success("✅ Venta guardada en la tabla.")
        debug_box.info(f"CABECERA fila {info['header_row']} → escrita en fila **{info['written_row']}**.")
        st.balloons()
        st.session_state.articulo_sel = ""
        st.session_state.precio_sel = 0.0
        st.cache_data.clear()
    except Exception as e:
        st.error(f"No se pudo escribir: {e}")

with st.expander("📊 Vista rápida de ventas (Sheet1)"):
    st.dataframe(read_current_table(), use_container_width=True, hide_index=True)

