import streamlit as st
import pandas as pd
import openpyxl
import os
import tempfile

# ==========================================
# CONFIGURACIÓN (SIEMPRE ARRIBA)
# ==========================================
st.set_page_config(page_title="PASCA Inventory Pro", layout="wide")

st.markdown("""
<style>
.stNumberInput label { font-size: 18px !important; font-weight: bold !important; }
.big-font {
    font-size: 36px;
    font-weight: bold;
    text-align: center;
    color: white;
    background-color: #2E7D32;
    padding: 20px;
    border-radius: 15px;
}
.product-header {
    background-color: white;
    padding: 25px;
    border-radius: 20px;
    border-left: 12px solid #4CAF50;
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# FUNCIONES
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
    df_sistema.columns = df_sistema.columns.str.strip()
    df_sistema.iloc[:, 0] = df_sistema.iloc[:, 0].apply(clean_code)

    df_conteo = pd.read_excel(tmp_path, sheet_name='CONTEO_F')

    header_row_index = 0
    for i, row in df_conteo.iterrows():
        if "CODIGO" in str(row.values).upper():
            header_row_index = i
            break

    df_conteo.columns = df_conteo.iloc[header_row_index].str.strip()
    df_conteo = df_conteo.iloc[header_row_index + 1:].reset_index(drop=True)
    df_conteo = df_conteo.astype(object)
    df_conteo.iloc[:, 0] = df_conteo.iloc[:, 0].apply(clean_code)

    st.session_state.temp_file_path = tmp_path
    return df_conteo, df_sistema, wb


def save_to_excel(df_conteo, wb):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        final_path = tmp.name

    sheet = wb['CONTEO_F']

    start_row = 1
    for row in sheet.iter_rows(max_row=10):
        for cell in row:
            if cell.value and "CODIGO" in str(cell.value).upper():
                start_row = cell.row + 1
                break
        else:
            continue
        break

    for i, row in df_conteo.iterrows():
        row_num = start_row + i
        for col_num, value in enumerate(row.values, 1):
            sheet.cell(row=row_num, column=col_num).value = value

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
st.title("📦 PASCA Inventory Pro")

uploaded_file = st.file_uploader("Sube el Excel del sistema", type=["xlsx"])

if uploaded_file:

    if 'df_inv' not in st.session_state:
        df_c, df_s, wb = load_pasca_data(uploaded_file)
        st.session_state.df_inv = df_c
        st.session_state.df_sistema = df_s
        st.session_state.wb = wb

    df_conteo = st.session_state.df_inv
    df_sistema = st.session_state.df_sistema
    wb = st.session_state.wb

    # 🔍 BUSCAR
    search_term = st.text_input("Buscar producto").strip().upper()

    if search_term:
        mask = (
            (df_sistema.iloc[:, 0].astype(str) == search_term) |
            (df_sistema.iloc[:, 1].astype(str).str.contains(search_term, case=False))
        )

        res = df_sistema[mask]

        if not res.empty:
            code = clean_code(res.iloc[0, 0])
            name = res.iloc[0, 1]

            st.write(f"**{name}**")

            idx_list = df_conteo[df_conteo.iloc[:, 0] == code].index

            if not idx_list.empty:
                idx = idx_list[0]

                inputs = {}
                col_names = ["BO1","BO2","BO3","AL1","AL2","AL3","VALES","VENCIDOS"]

                for i, col in enumerate(col_names):
                    val = df_conteo.iloc[idx, i+3]
                    val = int(val) if pd.notnull(val) else 0
                    inputs[col] = st.number_input(col, value=val)

                total = sum(inputs.values())
                st.markdown(f"<div class='big-font'>TOTAL: {total}</div>", unsafe_allow_html=True)

                if st.button("Guardar", type="primary"):
                    for i, col in enumerate(col_names):
                        df_conteo.iloc[idx, i+3] = inputs[col]

                    df_conteo.iloc[idx, 11] = total

                    st.balloons()
                    st.success("Guardado correctamente")
        else:
            st.error("Producto no encontrado")

    st.divider()

    if st.button("💾 Descargar Excel"):
        data = save_to_excel(df_conteo, wb)

        st.download_button(
            "Descargar archivo",
            data=data,
            file_name="INVENTARIO_FINAL.xlsx"
        )