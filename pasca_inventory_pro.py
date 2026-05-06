import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# ==========================================
# CONFIGURACIÓN DE INTERFAZ
# ==========================================
st.set_page_config(page_title="PASCA Inventory Smart", layout="wide")

st.markdown("""
<style>
.stButton>button {
    width: 100%;
    height: 70px;
    font-size: 20px !important;
    font-weight: bold !important;
    border-radius: 12px !important;
    border: 2px solid #4CAF50 !important;
    margin-bottom: 10px;
}
.big-font {
    font-size: 28px !important;
    font-weight: bold;
    text-align: center;
    color: #2E7D32;
}
.product-card {
    background-color: #f0f2f6;
    padding: 15px;
    border-radius: 15px;
    border-left: 8px solid #4CAF50;
    margin-bottom: 15px;
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# LÓGICA DE DATOS
# ==========================================
def load_pasca_data(file):
    wb = openpyxl.load_workbook(file)

    df_pres = pd.read_excel(file, sheet_name='PRESENTACIÓN')
    df_pres.columns = df_pres.columns.str.strip()

    mapping_pres = {}
    for _, row in df_pres.iterrows():
        name = str(row['DESCRIPCION']).strip().upper()
        code = str(row['CODIGO']).strip()
        factor = row['PRESENTACION'] if pd.notnull(row['PRESENTACION']) else 1

        mapping_pres[name] = {'factor': factor, 'code': code}
        mapping_pres[code] = {'factor': factor, 'code': code}

    df_conteo = pd.read_excel(file, sheet_name='CONTEO_F')

    header_row_index = 0
    for i, row in df_conteo.iterrows():
        if "CODIGO" in str(row.values).upper():
            header_row_index = i
            break

    df_conteo.columns = df_conteo.iloc[header_row_index].str.strip()
    df_conteo = df_conteo.iloc[header_row_index + 1:].reset_index(drop=True)

    return df_conteo, wb, mapping_pres


def save_to_excel(df_conteo, wb):
    sheet = wb['CONTEO_F']

    start_row = 1
    for row in sheet.iter_rows(max_row=10):
        for cell in row:
            if cell.value and "CODIGO" in str(cell.value).upper():
                start_row = cell.row + 1
                break

    for i, row in df_conteo.iterrows():
        row_num = start_row + i
        for col_num, value in enumerate(row.values, 1):
            sheet.cell(row=row_num, column=col_num).value = value

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# ==========================================
# APP
# ==========================================
st.title("📦 PASCA Inventory Smart")

# SIDEBAR
with st.sidebar:
    st.header("📍 Ubicación")
    bodegas = ["BO1", "BO2", "BO3", "AL1", "AL2", "AL3"]
    st.session_state.selected_bodega = st.selectbox("Bodega Actual", bodegas)
    st.info("Modo: Búsqueda Rápida y Manual")

# CARGA ARCHIVO
uploaded_file = st.file_uploader("Cargar Plantilla de Sistema", type=["xlsx"])

if uploaded_file:

    if 'df_inv' not in st.session_state:
        df_c, wb, mapping = load_pasca_data(uploaded_file)
        st.session_state.df_inv = df_c
        st.session_state.wb_inv = wb
        st.session_state.mapping_pres = mapping

    df = st.session_state.df_inv
    wb = st.session_state.wb_inv
    mapping = st.session_state.mapping_pres

    col_search, col_data = st.columns([1, 1])

    # ==========================================
    # BUSCADOR
    # ==========================================
    with col_search:
        st.subheader("🔍 Buscar Producto")

        search_term = st.text_input("Escribe el nombre o código...", "").upper()

        if search_term:
            matches = [k for k in mapping.keys() if search_term in k]

            if matches:
                st.write(f"Encontrados {len(matches)} productos:")

                for m in matches[:10]:
                    if st.button(f"👉 {m}"):
                        st.session_state.selected_prod = m
            else:
                st.error("No se encontró ningún producto.")

    # ==========================================
    # CONTEO
    # ==========================================
    with col_data:
        st.subheader("📝 Conteo")

        if 'selected_prod' in st.session_state:
            prod_key = st.session_state.selected_prod
            prod_info = mapping[prod_key]

            factor = prod_info['factor']
            code = prod_info['code']

            mask = df.iloc[:, 0].astype(str).str.strip() == str(code).strip()
            prod_row_idx = df[mask].index

            if not prod_row_idx.empty:
                idx = prod_row_idx[0]
                prod_name = df.iloc[idx, 1]

                st.markdown(f"""
                <div class="product-card">
                    <div class="big-font">{prod_name}</div>
                    <p><b>Código:</b> {code} | <b>Factor:</b> {factor}</p>
                </div>
                """, unsafe_allow_html=True)

                c1, c2 = st.columns(2)

                with c1:
                    cajas = st.number_input("📦 Cajas", min_value=0, step=1)

                with c2:
                    sueltos = st.number_input("Unidades", min_value=0, step=1)

                total = (cajas * factor) + sueltos

                st.markdown(f"<div class='big-font'>Total: {total}</div>", unsafe_allow_html=True)

                if st.button("✅ Confirmar y Guardar"):

                    bodega_map = {
                        "BO1": 3, "BO2": 4, "BO3": 5,
                        "AL1": 6, "AL2": 7, "AL3": 8
                    }

                    col_idx = bodega_map[st.session_state.selected_bodega]
                    df.iloc[idx, col_idx] = total

                    st.balloons()
                    st.success(f"Guardado {total} en {st.session_state.selected_bodega}")

            else:
                st.error("El producto no está en la hoja de conteo.")

    # ==========================================
    # EXPORTAR
    # ==========================================
    st.divider()

    if st.button("💾 EXPORTAR INVENTARIO FINAL"):
        final_bytes = save_to_excel(df, wb)

        st.download_button(
            label="Descargar Excel para Sistema",
            data=final_bytes,
            file_name="INVENTARIO_PASCA_FINAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )