
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

@st.cache_data(show_spinner=False)
def load_or_seed_data(excel_path: str) -> pd.DataFrame:
    if not os.path.exists(excel_path):
        return pd.DataFrame(columns=["Fecha","Cantidad","Nombre del Artículo","Método de Pago","Precio Unitario","Venta Total","Comentarios"])
    try:
        df = pd.read_excel(excel_path, sheet_name=VENTAS_SHEET)
        df.columns = [str(c) for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame(columns=["Fecha","Cantidad","Nombre del Artículo","Método de Pago","Precio Unitario","Venta Total","Comentarios"])

def save_append_row(excel_path: str, row: dict):
    if os.path.exists(excel_path):
        try:
            from openpyxl import load_workbook
            book = load_workbook(excel_path)
            with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                writer.book = book
                if VENTAS_SHEET in writer.book.sheetnames:
                    existing = pd.read_excel(excel_path, sheet_name=VENTAS_SHEET)
                    existing.columns = [str(c) for c in existing.columns]
                    new_df = pd.concat([existing, pd.DataFrame([row])], ignore_index=True)
                    idx = writer.book.sheetnames.index(VENTAS_SHEET)
                    std = writer.book.worksheets[idx]
                    writer.book.remove(std)
                    writer.book.create_sheet(VENTAS_SHEET, idx)
                    new_df.to_excel(writer, sheet_name=VENTAS_SHEET, index=False)
                else:
                    pd.DataFrame([row]).to_excel(writer, sheet_name=VENTAS_SHEET, index=False)
            return
        except Exception as e:
            st.error(f"Error guardando en Excel: {e}")
            return
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        pd.DataFrame([row]).to_excel(writer, sheet_name=VENTAS_SHEET, index=False)

df = load_or_seed_data(EXCEL_FILE)

with st.expander("📊 Ver datos actuales", expanded=False):
    st.dataframe(df, use_container_width=True)

st.subheader("➕ Añadir una venta")
col1, col2 = st.columns(2)
with col1:
    fecha = st.date_input("Fecha", value=date.today())
    cantidad = st.number_input("Cantidad", min_value=1, step=1, value=1)
    articulo_default = df["Nombre del Artículo"].dropna().astype(str).unique().tolist() if "Nombre del Artículo" in df.columns else []
    articulo = st.selectbox("Nombre del Artículo", options=["(Nuevo)…"] + articulo_default, index=0)
    if articulo == "(Nuevo)…":
        articulo = st.text_input("Escribe el nombre del artículo", value="")
with col2:
    metodo = st.radio("Método de Pago", options=["E", "T"], horizontal=True, help="E=Efectivo, T=Tarjeta")
    precio_unit = st.number_input("Precio Unitario", min_value=0.0, step=1.0, value=0.0, format="%.2f")
    total_calc = float(cantidad) * float(precio_unit)
    venta_total = st.number_input("Venta Total (auto)", min_value=0.0, step=1.0, value=total_calc, format="%.2f", help="Se calcula automáticamente; puedes ajustarlo si es necesario.")
comentarios = st.text_area("Comentarios (opcional)", value="")

valid = bool(articulo) and precio_unit > 0

if st.button("Guardar en Excel", type="primary", use_container_width=True, disabled=not valid):
    new_row = {
        "Fecha": fecha,
        "Cantidad": int(cantidad),
        "Nombre del Artículo": articulo,
        "Método de Pago": metodo,
        "Precio Unitario": float(precio_unit),
        "Venta Total": float(venta_total),
        "Comentarios": comentarios if comentarios.strip() else None,
    }
    save_append_row(EXCEL_FILE, new_row)
    st.success("✅ Venta guardada.")
    st.cache_data.clear()
    st.rerun()
