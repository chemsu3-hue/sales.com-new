# sales_app_streamlit.py — robusto: catálogo + escritura en Sheet1 con override de cabeceras
import streamlit as st
import pandas as pd
import os, unicodedata, re
from datetime import date, datetime
from openpyxl import load_workbook

# ==============================
# Config
# ==============================
DEFAULT_EXCEL  = "mimamuni sales datta+.xlsx"
DEFAULT_SHEET  = "Sheet1"
CAT_SHEET      = "Catalogo"

EXPECTED = ["Fecha","Cantidad","Nombre del Artículo","Método de Pago","Precio Unitario","Venta Total","Comentarios"]

DEFAULT_CATALOG = [
    {"Artículo":"bolsa","Precio":120.0},
    {"Artículo":"jeans","Precio":50.0},
    {"Artículo":"t-shirt","Precio":25.0},
    {"Artículo":"jacket","Precio":120.0},
    {"Artículo":"cinturón","Precio":20.0},
]

st.set_page_config(page_title="Ventas - Tienda de Ropa", page_icon="🛍️", layout="wide")
st.markdown("""
<style>
button[kind="secondary"]{border-radius:16px;padding:16px 12px;min-height:80px;white-space:pre-line;font-weight:600}
.block-container{padding-top:1.2rem}
</style>
""", unsafe_allow_html=True)

st.title("🛍️ Registro de Ventas")
st.caption("Catálogo editable • Tiles rápidos • Guarda en la tabla de Sheet1 (VENTAS DIARIAS).")

# ==============================
# Utilidades
# ==============================
def canon(s: str) -> str:
    """Normaliza: sin acentos, colapsa espacios (incluye NBSP), minúsculas."""
    if s is None: return ""
    s = str(s)
    s = s.replace("\u00A0", " ")            # NBSP → espacio normal
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

HEADER_SYNONYMS = {
    "fecha": "Fecha",
    "cantidad": "Cantidad",
    "nombre del articulo": "Nombre del Artículo",
    "nombre del artículo": "Nombre del Artículo",
    "articulo": "Nombre del Artículo",
    "artículo": "Nombre del Artículo",
    "producto": "Nombre del Artículo",
    "descripcion": "Nombre del Artículo",
    "descripción": "Nombre del Artículo",
    "metodo de pago": "Método de Pago",
    "metodo pago": "Método de Pago",
    "medio de pago": "Método de Pago",
    "precio unitario": "Precio Unitario",
    "venta total": "Venta Total",
    "comentarios": "Comentarios",
    "comentario": "Comentarios",
}

def ensure_excel(path: str) -> bool:
    return os.path.exists(path)

def open_wb(path: str):
    if not ensure_excel(path):
        raise FileNotFoundError("Excel no encontrado. Sube tu archivo primero.")
    return load_workbook(path)

def sheet_headers(path: str, sheet_name: str, max_cols: int = 60, header_row_hint: int | None = None):
    """Devuelve (header_row, lista_de_encabezados_raw) detectando la fila de cabeceras."""
    wb = open_wb(path)
    ws = wb[sheet_name]
    max_rows = min(ws.max_row, 300)

    def row_vals(r: int):
        return [ws.cell(r, c).value for c in range(1, max_cols+1)]

    # Si el usuario forzó fila, úsala
    if header_row_hint and 1 <= header_row_hint <= ws.max_row:
        return header_row_hint, row_vals(header_row_hint)

    # Autodetección: buscamos una fila que contenga al menos fecha, cantidad y nombre del artículo (canónicos)
    for r in range(1, max_rows+1):
        vals = row_vals(r)
        can = [canon(v) for v in vals]
        if {"fecha","cantidad","nombre del articulo"} <= set(can):
            return r, vals
    return None, []

def build_col_map(headers_raw):
    """Construye el mapa estándar -> índice_columna usando sinónimos tolerantes."""
    col_map = {}
    for idx, h in enumerate(headers_raw, start=1):
        std = HEADER_SYNONYMS.get(canon(h))
        if std in EXPECTED:
            col_map[std] = idx
    return col_map

def next_row_by_fecha(path: str, sheet_name: str, header_row: int, col_map: dict):
    wb = open_wb(path)
    ws = wb[sheet_name]
    if "Fecha" not in col_map:
        raise RuntimeError("No se encontró la columna 'Fecha' en la tabla.")
    r = header_row + 1
    last = header_row
    while r <= ws.max_row:
        if ws.cell(r, col_map["Fecha"]).value not in (None, ""):
            last = r
            r += 1
        else:
            break
    return last + 1

def append_row(path: str, sheet_name: str, header_row: int, col_map: dict, row_dict: dict):
    wb = open_wb(path)
    ws = wb[sheet_name]

    # calcular Venta Total si falta
    if "Venta Total" in col_map and (row_dict.get("Venta Total") in (None, "")):
        try:
            row_dict["Venta Total"] = float(row_dict.get("Cantidad", 0)) * float(row_dict.get("Precio Unitario", 0))
        except Exception:
            row_dict["Venta Total"] = None

    r = next_row_by_fecha(path, sheet_name, header_row, col_map)

    for std, col in col_map.items():
        v = row_dict.get(std, None)
        if std == "Nombre del Artículo" and v is not None:
            v = str(v)  # aseguramos string
        if std == "Fecha" and isinstance(v, (date, datetime)):
            ws.cell(r, col).value = datetime.combine(v, datetime.min.time())
        else:
            ws.cell(r, col).value = v

    wb.save(path)
    return r

def write_sheet_replace(path: str, df: pd.DataFrame, sheet_name: str):
    from openpyxl.utils.dataframe import dataframe_to_rows
    wb = open_wb(path)
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    ws = wb.create_sheet(sheet_name)
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)
    wb.save(path)

# ==============================
# Cargar/seleccionar Excel
# ==============================
if "excel_path" not in st.session_state:
    st.session_state.excel_path = DEFAULT_EXCEL

uploaded = st.file_uploader("📂 Sube tu Excel (.xlsx). Se guardará en esa tabla de Sheet1.", type=["xlsx"])
if uploaded is not None:
    st.session_state.excel_path = uploaded.name
    with open(st.session_state.excel_path, "wb") as f:
        f.write(uploaded.getbuffer())
    st.success(f"Excel guardado como: {st.session_state.excel_path}")
    st.cache_data.clear()

excel_path = st.session_state.excel_path

if ensure_excel(excel_path):
    wb = open_wb(excel_path)
    sheets = wb.sheetnames
    target_sheet = st.selectbox("Hoja de destino (donde está VENTAS DIARIAS)", sheets,
                                index=sheets.index(DEFAULT_SHEET) if DEFAULT_SHEET in sheets else 0)
else:
    st.info("Aún no has subido el archivo; se usará el nombre por defecto si existe.")
    target_sheet = DEFAULT_SHEET

# ==============================
# Diagnóstico de cabeceras
# ==============================
st.divider()
st.subheader("🛠️ Diagnóstico de cabeceras (por si 'Nombre del Artículo' no aparece en el Excel)")
header_row_hint = st.number_input("Fila de cabeceras (0 = detectar automático)", min_value=0, value=0, step=1)
forced = header_row_hint if header_row_hint > 0 else None

if ensure_excel(excel_path) and target_sheet in open_wb(excel_path).sheetnames:
    hr, headers_raw = sheet_headers(excel_path, target_sheet, header_row_hint=forced)
    col_map = build_col_map(headers_raw)
    st.caption(f"Fila detectada: **{hr}** | Encabezados: { [h for h in headers_raw if h not in (None,'')] }")
    # Si falta 'Nombre del Artículo', permitimos override manual
    if "Nombre del Artículo" not in col_map and headers_raw:
        opciones = [h for h in headers_raw if h not in (None, "")]
        override = st.selectbox("Selecciona columna para 'Nombre del Artículo' (si no se detectó):", opciones)
        if override:
            # inyectamos override
            idx = headers_raw.index(override) + 1
            col_map["Nombre del Artículo"] = idx
            st.info(f"Usando override: 'Nombre del Artículo' → columna {idx}")
else:
    hr, col_map = None, {}

# ==============================
# Catálogo (CRUD persistente)
# ==============================
st.divider()
st.subheader("🗂️ Catálogo (añadir / editar / borrar)")

def load_catalog_df(path: str) -> pd.DataFrame:
    if not ensure_excel(path):
        return pd.DataFrame(DEFAULT_CATALOG, columns=["Artículo","Precio"])
    try:
        df = pd.read_excel(path, sheet_name=CAT_SHEET)
        df = df.rename(columns={"Articulo":"Artículo","precio":"Precio"})
        df["Artículo"] = df["Artículo"].astype(str)
        df["Precio"]   = pd.to_numeric(df["Precio"], errors="coerce").fillna(0.0)
        return df[["Artículo","Precio"]]
    except Exception:
        df = pd.DataFrame(DEFAULT_CATALOG, columns=["Artículo","Precio"])
        try: write_sheet_replace(path, df, CAT_SHEET)
        except Exception: pass
        return df

def save_catalog_df(path: str, df: pd.DataFrame):
    clean = df.copy()
    clean["Artículo"] = clean["Artículo"].astype(str).strip()
    clean["Precio"]   = pd.to_numeric(clean["Precio"], errors="coerce").fillna(0.0)
    clean = clean[clean["Artículo"]!=""].drop_duplicates(subset=["Artículo"], keep="last")
    write_sheet_replace(path, clean[["Artículo","Precio"]], CAT_SHEET)

cat = load_catalog_df(excel_path)
if "Eliminar" not in cat.columns: cat["Eliminar"] = False

edited = st.data_editor(
    cat, num_rows="dynamic", hide_index=True, use_container_width=True,
    column_config={
        "Artículo": st.column_config.TextColumn(required=True),
        "Precio":   st.column_config.NumberColumn(min_value=0.0, step=1.0, format="%.2f"),
        "Eliminar": st.column_config.CheckboxColumn(help="Marca para borrar esta fila")
    },
    key="catalog_editor"
)

c1, c2 = st.columns([1,1])
with c1:
    if st.button("💾 Guardar catálogo", use_container_width=True, disabled=not ensure_excel(excel_path)):
        to_save = edited.copy()
        if "Eliminar" in to_save:
            to_save = to_save[to_save["Eliminar"]==False].drop(columns=["Eliminar"])
        save_catalog_df(excel_path, to_save)
        st.success("Catálogo guardado en 'Catalogo'.")
        st.cache_data.clear(); st.rerun()
with c2:
    if st.button("↩️ Deshacer cambios", use_container_width=True):
        st.cache_data.clear(); st.rerun()

# ==============================
# Tiles (ELIGE UN ARTÍCULO)
# ==============================
st.divider()
st.subheader("🧱 ELIGE UN ARTÍCULO")
tiles = load_catalog_df(excel_path).sort_values("Artículo").reset_index(drop=True)
busca = st.text_input("Buscar artículo", placeholder="escribe para filtrar…")
if busca: tiles = tiles[tiles["Artículo"].str.contains(busca, case=False, na=False)]

if "articulo_sel" not in st.session_state: st.session_state.articulo_sel = ""
if "precio_sel"   not in st.session_state: st.session_state.precio_sel   = 0.0

per_row = 4
items = list(tiles.itertuples(index=False, name=None))
for i in range(0, len(items), per_row):
    cols = st.columns(per_row)
    for col, (name, price) in zip(cols, items[i:i+per_row]):
        with col:
            if st.button(f"{name}\n${float(price):.2f}", key=f"tile_{name}", use_container_width=True):
                st.session_state.articulo_sel = name
                st.session_state.precio_sel   = float(price)

with st.expander("➕ Añadir artículo rápido", expanded=False):
    q1, q2, q3 = st.columns([2,1,1])
    with q1: qa_name = st.text_input("Nombre", key="qa_name")
    with q2: qa_price = st.number_input("Precio", min_value=0.0, step=1.0, value=0.0, format="%.2f", key="qa_price")
    with q3:
        if st.button("Guardar", use_container_width=True, key="qa_btn", disabled=not ensure_excel(excel_path)):
            dfc = load_catalog_df(excel_path)
            mask = dfc["Artículo"].str.lower().eq((qa_name or "").strip().lower())
            if mask.any():
                dfc.loc[mask, "Precio"] = float(qa_price)
            else:
                dfc = pd.concat([dfc, pd.DataFrame([{"Artículo": (qa_name or "").strip(), "Precio": float(qa_price)}])], ignore_index=True)
            save_catalog_df(excel_path, dfc)
            st.success(f"Guardado: {qa_name} → {float(qa_price):.2f}")
            st.cache_data.clear(); st.rerun()

# ==============================
# Formulario de venta
# ==============================
st.divider()
st.subheader("🧾 Guardar venta en tu tabla de Sheet1")

left, right = st.columns(2)
with left:
    fecha    = st.date_input("Fecha", value=date.today())
    cantidad = st.number_input("Cantidad", min_value=1, step=1, value=1)
    articulo = st.text_input("Nombre del Artículo", value=st.session_state.articulo_sel)
with right:
    metodo      = st.radio("Método de Pago", ["E","T"], horizontal=True)
    precio_unit = st.number_input("Precio Unitario", min_value=0.0, step=1.0, value=float(st.session_state.precio_sel), format="%.2f")
    venta_total = st.number_input("Venta Total (auto)", min_value=0.0, step=1.0, value=float(cantidad)*float(precio_unit), format="%.2f")
comentarios = st.text_area("Comentarios (opcional)")

disabled = (not ensure_excel(excel_path)) or (not articulo) or (precio_unit <= 0)
debug_box = st.empty()

# Botón de prueba (muestra dónde escribiría y aplica el override si lo definiste)
if st.button("🧪 Probar ubicación (no modifica nada)", use_container_width=True, disabled=not ensure_excel(excel_path)):
    try:
        hr, headers_raw = sheet_headers(excel_path, target_sheet, header_row_hint=forced)
        if not hr: raise RuntimeError("No se detectó la fila de cabeceras.")
        cmap = build_col_map(headers_raw)
        # override manual si el usuario lo definió
        if "Nombre del Artículo" not in cmap and headers_raw:
            opciones = [h for h in headers_raw if h not in (None, "")]
            # si el usuario ya eligió en el selectbox, estará en la sesión
            chosen = st.session_state.get("Nombre_Override")
        else:
            chosen = None
        if chosen and chosen in headers_raw:
            cmap["Nombre del Artículo"] = headers_raw.index(chosen) + 1
        r = next_row_by_fecha(excel_path, target_sheet, hr, cmap)
        debug_box.info(f"Cabeceras en fila {hr}. Escribiría en fila **{r}**. Columnas usadas: {list(cmap.keys())}")
    except Exception as e:
        st.error(f"Prueba falló: {e}")

if st.button("💾 Guardar venta", type="primary", use_container_width=True, disabled=disabled):
    try:
        hr, headers_raw = sheet_headers(excel_path, target_sheet, header_row_hint=forced)
        if not hr: raise RuntimeError("No se detectó la fila de cabeceras.")
        cmap = build_col_map(headers_raw)
        # override manual: guarda la elección del usuario (si existe selectbox arriba)
        # Para simplificar, volvemos a calcular aquí:
        if "Nombre del Artículo" not in cmap and headers_raw:
            # intenta heurística: busca la primera cabecera que contenga "nombre" y "art"
            for idx, h in enumerate(headers_raw, start=1):
                c = canon(h)
                if "nombre" in c and ("art" in c or "prod" in c or "descr" in c):
                    cmap["Nombre del Artículo"] = idx
                    break
        new_row = {
            "Fecha": fecha,
            "Cantidad": int(cantidad),
            "Nombre del Artículo": (articulo or "").strip(),
            "Método de Pago": metodo,
            "Precio Unitario": float(precio_unit),
            "Venta Total": float(venta_total),
            "Comentarios": (comentarios or "").strip() or None,
        }
        written_row = append_row(excel_path, target_sheet, hr, cmap, new_row)
        st.success(f"✅ Venta guardada en fila {written_row}.")
        st.balloons()
        st.session_state.articulo_sel = ""
        st.session_state.precio_sel   = 0.0
        st.cache_data.clear()
    except Exception as e:
        st.error(f"No se pudo escribir: {e}")

# Vista rápida
with st.expander("📊 Vista rápida (lectura de tu tabla)"):
    try:
        df_view = pd.read_excel(excel_path, sheet_name=target_sheet, header=None)
        st.dataframe(df_view, use_container_width=True, hide_index=True)
    except Exception:
        pass

# Descargar
if ensure_excel(excel_path):
    with open(excel_path, "rb") as f:
        st.download_button("⬇️ Descargar Excel actualizado", f, file_name=os.path.basename(excel_path),
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
