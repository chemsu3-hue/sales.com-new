import streamlit as st
import pandas as pd
import os
from datetime import date
from openpyxl import load_workbook

EXCEL_FILE = "mimamuni sales datta+.xlsx"
VENTAS_SHEET = "Ventas"
RAW_SHEET = "Sheet1"

st.set_page_config(page_title="Ventas - Tienda de Ropa", page_icon="🛍️", layout="centered")
st.title("🛍️ Registro de Ventas")
st.caption("Sube tu archivo Excel una vez. El formulario añadirá ventas a la hoja “Ventas”.")

# -------- Excel helpers --------
def ensure_excel_exists() -> bool:
    """Return True if target Excel file exists in working dir."""
    return os.path.exists(EXCEL_FILE)

@st.cache_data(show_spinner=False)
def load_table() -> pd.DataFrame:
    """Load existing Ventas table or create an empty schema."""
    if not ensure_excel_exists():
        return pd.DataFrame(columns=["Fecha","Cantidad","Nombre del Artículo","Método de Pago","Precio Unitario","Venta Total","Comentarios"])
    # Try dedicated sheet first
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=VENTAS_SHEET)
        df.columns = [str(c) for c in df.columns]
        return df
    except Exception:
        # Fallback: try to parse from a raw dashboard sheet
        try:
            raw = pd.read_excel(EXCEL_FILE, sheet_name=RAW_SHEET, header=None)
            header_row_idx = None
            for i in range(len(raw)):
                vals = raw.iloc[i].astype(str).tolist()
                if {"Fecha","Cantidad","Nombre del Artículo"}.issubset(set(vals)):
                    header_row_idx = i; break
            if header_row_idx is None:
                raise RuntimeError("No se encontraron cabeceras en la hoja RAW.")
            headers = raw.iloc[header_row_idx].tolist()
            table = raw.iloc[header_row_idx+1:].copy()
            table.columns = headers
            table = table.dropna(axis=1, how="all")
            key = [c for c in ["Fecha","Cantidad","Nombre del Artículo"] if c in table.columns]
            table = table.dropna(subset=key, how="all")
            cols = ["Fecha","Cantidad","Nombre del Artículo","Método de Pago","Precio Unitario","Venta Total","Comentarios"]
            df = table[[c for c in cols if c in table.columns]].copy()
            df = df.loc[:, ~df.columns.duplicated()]
            # Coerce
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
    """Append one row into the Ventas sheet (create/replace sheet safely)."""
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
    # If the file didn't exist (e.g., just uploaded), create it now
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        pd.DataFrame([row]).to_excel(writer, sheet_name=VENTAS_SHEET, index=False)

# -------- Upload / download controls --------
st.subheader("📂 Tu archivo Excel")
uploaded = st.file_uploader("Sube tu Excel (.xlsx) — se guardará como “mimamuni sales datta+.xlsx”", type=["xlsx"])
if uploaded is not None:
    with open(EXCEL_FILE, "wb") as f:
        f.write(uploaded.getbuffer())
    st.success("Excel guardado. Ya puedes usar el formulario de ventas.")
    st.cache_data.clear()

if ensure_excel_exists():
    with open(EXCEL_FILE, "rb") as f:
        st.download_button("⬇️ Descargar Excel actualizado", f, file_name=EXCEL_FILE, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Aún no has subido el archivo. Sube tu Excel arriba para empezar.")

# -------- Data preview --------
df = load_table()
with st.expander("📊 Ver datos actuales", expanded=False):
    st.dataframe(df, use_container_width=True)

# -------- Entry form --------
st.subheader("➕ Añadir una venta")
col1, col2 = st.columns(2)
with col1:
    fecha = st.date_input("Fecha", value=date.today())
    cantidad = st.number_input("Cantidad", min_value=1, step=1, value=1)
    articulo_opts = df["Nombre del Artículo"].dropna().astype(str).unique().tolist() if "Nombre del Artículo" in df.columns else []
    articulo = st.selectbox("Nombre del Artículo", ["(Nuevo)…"] + articulo_opts, index=0)
    if articulo == "(Nuevo)…":
        articulo = st.text_input("Escribe el nombre del artículo", value="")
with col2:
    metodo = st.radio("Método de Pago", ["E", "T"], horizontal=True, help="E=Efectivo, T=Tarjeta")
    precio_unit = st.number_input("Precio Unitario", min_value=0.0, step=1.0, value=0.0, format="%.2f")
    venta_total = st.number_input("Venta Total (auto)", min_value=0.0, step=1.0, value=float(cantidad)*float(precio_unit), format="%.2f")

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
    st.cache_data.clear()
    st.rerun()

st.caption("Nota: en Streamlit Cloud, el archivo se guarda en el servidor y puede reiniciarse al re-desplegar. Descárgalo si quieres conservar los cambios.")
