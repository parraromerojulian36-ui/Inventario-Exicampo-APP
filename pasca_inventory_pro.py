import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# ==========================================
# CONFIGURACIÓN UI
# ==========================================
st.set_page_config(page_title="PASCA Inventory Pro", layout="wide")

st.markdown("""
<style>
.main { background-color: #f5f7f9; }

.stNumberInput label {
    font-size: 18px !important;
    font-weight: bold !important;
}

.big-font {
    font-size: 36px !important;
    font-weight: bold;
    text-align: center;
    color: white;
    background-color: #2E7D32;
    padding: 20px;
    border-radius: 15px;
    margin: 20px 0;
}

.product-header {
    background-color: white;
    padding: 25px;
    border-radius: 20px;
    border-left: 12px solid #4CAF50;
    margin-bottom: 25px;
}

div.stButton > button {
    width: 100%;
    height: 60px;
    font-size: 18px;
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# UTILIDADES
# ==========================================
def clean_code(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    return s[:-2] if s.endswith(".0") else s

# ==========================================
# DATA
# ==========================================
def load_pasca_data(file):
    wb = openpyxl.load_workbook(file)

    df_sistema = pd.read_excel(file, sheet_name='SISTEMA')
    df_sistema.columns = df_sistema.columns.str.strip()
    df_sistema.iloc[:, 0] = df_sistema.iloc[:, 0].apply(clean_code)

    df_conteo = pd.read_excel(file, sheet_name='CONTEO_F')

    header_row_index = 0
    for i, row in df_conteo.iterrows():
        if "CODIGO" in str(row.values).upper():
            header_row_index = i
            break

    df_conteo.columns = df_conteo.iloc[header_row_index].str.strip()
    df_conteo = df_conteo.iloc[header_row_index + 1:].reset_index(drop=True)

    df_conteo = df_conteo.astype(object)
    df_conteo.iloc[:, 0] = df_conteo.iloc[:, 0].apply(clean_code)

    return df_conteo, df_sistema, wb


def save_to_excel(df_conteo, wb):
    sheet = wb['CONTEO_F']

    start_row = 1
    for row in sheet.iter_rows(max_row=10):
        for cell in row:
            if cell.value and "CODIGO" in str(cell.value).upper():
                start_row = cell.row + 1
                break

    for i, row in df_conteo.iterrows():
        for j, val in enumerate(row.values, 1):
            sheet.cell(row=start_row + i, column=j).value = val

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# APP
# ==========================================
st.title("📦 PASCA Inventory Pro")
st.markdown("Gestión de conteo físico inteligente")

uploaded_file = st.file_uploader("Sube el Excel", type=["xlsx"])

if uploaded_file:

    if "data_loaded" not in st.session_state:
        df_c, df_s, wb = load_pasca_data(uploaded_file)
        st.session_state.df_conteo = df_c
        st.session_state.df_sistema = df_s
        st.session_state.wb = wb
        st.session_state.data_loaded = True

    df_conteo = st.session_state.df_conteo
    df_sistema = st.session_state.df_sistema
    wb = st.session_state.wb

    # ======================================
    # BUSCADOR
    # ======================================
    st.subheader("🔍 Buscar Producto")

    search = st.text_input("Código o nombre").upper().strip()

    if search:
        mask = (
            (df_sistema.iloc[:, 0].astype(str) == search) |
            (df_sistema.iloc[:, 1].astype(str).str.contains(search, case=False))
        )

        results = df_sistema[mask]

        if not results.empty:
            if len(results) > 1:
                st.warning("Selecciona una opción")
                for _, r in results.iterrows():
                    code = clean_code(r.iloc[0])
                    name = r.iloc[1]

                    if st.button(f"{name} ({code})"):
                        st.session_state.selected = (code, name)
            else:
                r = results.iloc[0]
                st.session_state.selected = (clean_code(r[0]), r[1])
        else:
            st.error("No encontrado")

    # ======================================
    # EDICIÓN
    # ======================================
    if "selected" in st.session_state:
        code, name = st.session_state.selected

        row_idx = df_conteo[df_conteo.iloc[:, 0] == code].index

        if not row_idx.empty:
            idx = row_idx[0]

            st.markdown(f"""
            <div class="product-header">
                <b>{name}</b><br>
                Código: {code}
            </div>
            """, unsafe_allow_html=True)

            cols = ["BO1","BO2","BO3","AL1","AL2","AL3","VALES","VENCIDOS"]

            vals = df_conteo.iloc[idx, 3:11].values
            vals = [int(v) if pd.notnull(v) else 0 for v in vals]

            inputs = {}

            c1 = st.columns(4)
            c2 = st.columns(4)

            for i, col in enumerate(cols):
                container = c1 if i < 4 else c2
                with container[i % 4]:
                    inputs[col] = st.number_input(col, value=vals[i], min_value=0)

            total = sum(inputs.values())

            st.markdown(f"<div class='big-font'>TOTAL: {total}</div>", unsafe_allow_html=True)

            if st.button("Guardar", type="primary"):
                map_cols = {
                    "BO1":3,"BO2":4,"BO3":5,
                    "AL1":6,"AL2":7,"AL3":8,
                    "VALES":9,"VENCIDOS":10
                }

                for k,v in inputs.items():
                    df_conteo.iloc[idx, map_cols[k]] = v

                df_conteo.iloc[idx, 11] = total

                st.success("Guardado")
                st.balloons()

        else:
            st.error("No está en CONTEO_F")

    # ======================================
    # EXPORTAR
    # ======================================
    st.divider()

    if st.button("Exportar Excel"):
        data = save_to_excel(df_conteo, wb)
        st.download_button("Descargar", data, "inventario.xlsx")