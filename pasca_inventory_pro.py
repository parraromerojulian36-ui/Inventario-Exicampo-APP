import streamlit as st
import pandas as pd
import openpyxl
import os
import tempfile
from datetime import datetime

# ==========================================
# CONFIGURACIÓN DE INTERFAZ
# ==========================================
st.set_page_config(page_title="PASCA Inventory System", layout="wide")

st.markdown("""
<style>
.stNumberInput label { font-size: 18px !important; font-weight: bold !important; }
.big-font {
    font-size: 36px !important;
    font-weight: bold;
    text-align: center;
    color: #ffffff;
    background-color: #2E7D32;
    padding: 20px;
    border-radius: 15px;
}
.product-header {
    background-color: #ffffff;
    padding: 25px;
    border-radius: 20px;
    border-left: 12px solid #4CAF50;
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# FUNCIONES AUXILIARES
# ==========================================
def clean_code(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    return s[:-2] if s.endswith(".0") else s


def load_pasca_data(uploaded_file):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name

    wb = openpyxl.load_workbook(tmp_path)

    df_sistema = pd.read_excel(tmp_path, sheet_name='SISTEMA')
    df_sistema.iloc[:, 0] = df_sistema.iloc[:, 0].apply(clean_code)

    df_conteo = pd.read_excel(tmp_path, sheet_name='CONTEO_F')

    # Detectar encabezado
    for i, row in df_conteo.iterrows():
        if "CODIGO" in str(row.values).upper():
            df_conteo.columns = df_conteo.iloc[i]
            df_conteo = df_conteo.iloc[i+1:].reset_index(drop=True)
            break

    df_conteo.iloc[:, 0] = df_conteo.iloc[:, 0].apply(clean_code)

    st.session_state.temp_file_path = tmp_path
    return df_conteo, df_sistema, wb


def save_full_inventory(df_conteo, df_sistema, wb):
    # --- Guardar conteo ---
    sheet_conteo = wb['CONTEO_F']

    for i, row in df_conteo.iterrows():
        for j, value in enumerate(row.values, 1):
            sheet_conteo.cell(row=i+2, column=j).value = value

    # --- Hoja resultado ---
    sheet_res = wb['RESULTADO']

    # Limpiar
    for row in sheet_res.iter_rows(min_row=5):
        for cell in row:
            cell.value = None

    row_res = 5

    for _, row_c in df_conteo.iterrows():
        code = clean_code(row_c.iloc[0])
        name = row_c.iloc[1]
        total_fisico = row_c.iloc[11] if pd.notnull(row_c.iloc[11]) else 0

        mask_s = df_sistema.iloc[:, 0].astype(str) == code
        res_s = df_sistema[mask_s]

        if not res_s.empty:
            total_sistema = res_s.iloc[0, 2] if pd.notnull(res_s.iloc[0, 2]) else 0
            diferencia = total_fisico - total_sistema

            faltante = abs(diferencia) if diferencia < 0 else "-"
            sobrante = diferencia if diferencia > 0 else "-"

            sheet_res.cell(row=row_res, column=1).value = code
            sheet_res.cell(row=row_res, column=2).value = name
            sheet_res.cell(row=row_res, column=3).value = total_fisico
            sheet_res.cell(row=row_res, column=4).value = total_sistema
            sheet_res.cell(row=row_res, column=5).value = diferencia
            sheet_res.cell(row=row_res, column=6).value = faltante
            sheet_res.cell(row=row_res, column=7).value = sobrante

            row_res += 1

    # Guardar archivo
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        final_path = tmp.name

    wb.save(final_path)

    with open(final_path, "rb") as f:
        data = f.read()

    try:
        os.remove(st.session_state.temp_file_path)
        os.remove(final_path)
    except:
        pass

    return data

# ==========================================
# INTERFAZ
# ==========================================
st.title("📦 PASCA Inventory System")

with st.sidebar:
    sucursal = st.selectbox("Sucursal", ["PASCA", "SUBIA", "SIBATE", "GRANADA"])
    fecha_actual = datetime.now().strftime("%d-%m-%Y")

uploaded_file = st.file_uploader("Sube el Excel", type=["xlsx"])

if uploaded_file:
    if 'df_inv' not in st.session_state:
        df_c, df_s, wb = load_pasca_data(uploaded_file)
        st.session_state.df_inv = df_c
        st.session_state.df_sistema = df_s
        st.session_state.wb = wb

    df_conteo = st.session_state.df_inv
    df_sistema = st.session_state.df_sistema

    search = st.text_input("Buscar producto").upper()

    if search:
        mask = df_sistema.iloc[:, 0].astype(str) == search
        res = df_sistema[mask]

        if not res.empty:
            code = clean_code(res.iloc[0, 0])
            name = res.iloc[0, 1]

            st.write(name)

            idx = df_conteo[df_conteo.iloc[:, 0] == code].index

            if not idx.empty:
                idx = idx[0]

                inputs = []
                for i in range(8):
                    val = df_conteo.iloc[idx, i+3]
                    val = int(val) if pd.notnull(val) else 0
                    inputs.append(st.number_input(f"Campo {i+1}", value=val))

                total = sum(inputs)
                st.write("TOTAL:", total)

                if st.button("Guardar"):
                    for i in range(8):
                        df_conteo.iloc[idx, i+3] = inputs[i]

                    df_conteo.iloc[idx, 11] = total
                    st.success("Guardado")

    if st.button("Exportar"):
        data = save_full_inventory(df_conteo, df_sistema, st.session_state.wb)

        st.download_button(
            "Descargar Excel",
            data=data,
            file_name=f"INVENTARIO_{sucursal}_{fecha_actual}.xlsx"
        )