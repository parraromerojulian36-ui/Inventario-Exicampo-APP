import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# ==========================================
# CONFIGURACIÓN DE INTERFAZ
# ==========================================
st.set_page_config(page_title="PASCA Inventory Editor", layout="wide")

st.markdown("""
<style>
.stNumberInput label {
    font-size: 16px !important;
    font-weight: bold !important;
}
.big-font {
    font-size: 32px !important;
    font-weight: bold;
    text-align: center;
    color: #1B5E20;
    background-color: #C8E6C9;
    padding: 10px;
    border-radius: 10px;
    border: 2px solid #4CAF50;
}
.product-header {
    background-color: #E8F5E9;
    padding: 20px;
    border-radius: 15px;
    border-left: 10px solid #2E7D32;
    margin-bottom: 20px;
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# UTILIDAD
# ==========================================
def clean_code(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

# ==========================================
# LÓGICA DE DATOS
# ==========================================
def load_pasca_data(file):
    wb = openpyxl.load_workbook(file)

    # SISTEMA
    df_sistema = pd.read_excel(file, sheet_name='SISTEMA')
    df_sistema.columns = df_sistema.columns.str.strip()
    df_sistema.iloc[:, 0] = df_sistema.iloc[:, 0].apply(clean_code)

    # CONTEO_F
    df_conteo = pd.read_excel(file, sheet_name='CONTEO_F')

    header_row_index = 0
    for i, row in df_conteo.iterrows():
        if "CODIGO" in str(row.values).upper():
            header_row_index = i
            break

    df_conteo.columns = df_conteo.iloc[header_row_index].str.strip()
    df_conteo = df_conteo.iloc[header_row_index + 1:].reset_index(drop=True)

    # Permitir tipos mixtos (evita errores al escribir)
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
        row_num = start_row + i
        for col_num, value in enumerate(row.values, 1):
            sheet.cell(row=row_num, column=col_num).value = value

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# APP
# ==========================================
st.title("📦 PASCA Inventory Editor")
st.markdown("Busque un producto → Edite → Guarde")

uploaded_file = st.file_uploader("Cargar Plantilla de Sistema", type=["xlsx"])

if uploaded_file:

    if 'df_inv' not in st.session_state:
        df_c, df_s, wb = load_pasca_data(uploaded_file)
        st.session_state.df_inv = df_c
        st.session_state.df_sistema = df_s
        st.session_state.wb_inv = wb

    df_conteo = st.session_state.df_inv
    df_sistema = st.session_state.df_sistema
    wb = st.session_state.wb_inv

    # ==========================================
    # BUSCAR
    # ==========================================
    st.subheader("🔍 Buscar Producto")

    search_term = st.text_input("Ingrese Código o Nombre...").strip().upper()

    if search_term:

        mask_s = (
            (df_sistema.iloc[:, 0].astype(str) == search_term) |
            (df_sistema.iloc[:, 1].astype(str).str.contains(search_term, case=False))
        )

        res_sistema = df_sistema[mask_s]

        if not res_sistema.empty:

            real_code = clean_code(res_sistema.iloc[0, 0])
            prod_name = res_sistema.iloc[0, 1]
            stock_sistema = res_sistema.iloc[0, 2]

            mask_c = df_conteo.iloc[:, 0].astype(str) == real_code
            prod_row_idx = df_conteo[mask_c].index

            if not prod_row_idx.empty:
                idx = prod_row_idx[0]

                st.markdown(f"""
                <div class="product-header">
                    <div style="font-size: 24px; font-weight: bold;">{prod_name}</div>
                    <div style="font-size: 18px;">
                        Código: {real_code} | <b>Stock Sistema: {stock_sistema}</b>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                # ==========================================
                # INPUTS
                # ==========================================
                st.write("### 📝 Ingreso de Cantidades")

                col_names = ["BO1", "BO2", "BO3", "AL1", "AL2", "AL3", "VALES", "VENCIDOS"]

                raw_vals = df_conteo.iloc[idx, 3:11].values
                current_values = [
                    int(v) if pd.notnull(v) and str(v).replace('.', '').isdigit() else 0
                    for v in raw_vals
                ]

                inputs = {}
                row1 = st.columns(4)
                row2 = st.columns(4)

                for i, col_name in enumerate(col_names):
                    target = row1 if i < 4 else row2
                    with target[i % 4]:
                        inputs[col_name] = st.number_input(
                            col_name,
                            min_value=0,
                            value=current_values[i]
                        )

                total_fisico = sum(inputs.values())

                st.markdown(
                    f"<div class='big-font'>TOTAL FÍSICO: {total_fisico}</div>",
                    unsafe_allow_html=True
                )

                # ==========================================
                # GUARDAR
                # ==========================================
                if st.button("✅ GUARDAR CAMBIOS EN EXCEL"):

                    bodega_map = {
                        "BO1": 3, "BO2": 4, "BO3": 5,
                        "AL1": 6, "AL2": 7, "AL3": 8,
                        "VALES": 9, "VENCIDOS": 10
                    }

                    for col_name, value in inputs.items():
                        df_conteo.iloc[idx, bodega_map[col_name]] = value

                    df_conteo.iloc[idx, 11] = total_fisico

                    st.balloons()
                    st.success(f"{prod_name} actualizado correctamente")

            else:
                st.error("Está en SISTEMA pero no en CONTEO_F")

        else:
            st.error("No existe en SISTEMA")

    # ==========================================
    # EXPORTAR
    # ==========================================
    st.divider()

    if st.button("💾 EXPORTAR ARCHIVO FINAL"):
        final_bytes = save_to_excel(df_conteo, wb)

        st.download_button(
            label="Descargar Excel",
            data=final_bytes,
            file_name="INVENTARIO_PASCA_FINAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )