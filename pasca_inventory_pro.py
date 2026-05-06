import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# ==========================================
# CONFIGURACIÓN DE INTERFAZ
# ==========================================
st.set_page_config(page_title="PASCA Inventory Pro", layout="wide")

st.markdown("""
<style>
.stNumberInput label {
    font-size: 18px !important;
    font-weight: bold !important;
}

.big-font {
    font-size: 36px !important;
    font-weight: bold;
    text-align: center;
    color: #ffffff;
    background-color: #2E7D32;
    padding: 20px;
    border-radius: 15px;
    border: 3px solid #1B5E20;
    box-shadow: 0px 4px 10px rgba(0,0,0,0.2);
    margin: 20px 0;
}

.product-header {
    background-color: #ffffff;
    padding: 25px;
    border-radius: 20px;
    border-left: 12px solid #4CAF50;
    box-shadow: 0px 2px 15px rgba(0,0,0,0.1);
    margin-bottom: 25px;
}

div.stButton > button {
    width: 100%;
    text-align: left !important;
    height: 60px !important;
    font-size: 18px !important;
    border-radius: 10px !important;
    border: 1px solid #ddd !important;
    background-color: white !important;
    color: #333 !important;
}

.stButton > button[kind="primary"] {
    height: 80px !important;
    font-size: 24px !important;
    background-color: #4CAF50 !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)


# ==========================================
# FUNCIONES AUXILIARES
# ==========================================
def clean_code(val):
    """Limpia códigos (quita .0 y espacios)"""
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
    """Carga y limpia datos del Excel"""
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
    df_conteo = df_conteo.astype(object)
    df_conteo.iloc[:, 0] = df_conteo.iloc[:, 0].apply(clean_code)

    return df_conteo, df_sistema, wb


def save_to_excel(df_conteo, wb):
    """Guarda cambios en el archivo Excel"""
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
# INTERFAZ PRINCIPAL
# ==========================================
st.title("📦 PASCA Inventory Pro")

uploaded_file = st.file_uploader("Sube el Excel del sistema", type=["xlsx"])


# ==========================================
# CARGA INICIAL
# ==========================================
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
    # BUSCADOR
    # ==========================================
    st.subheader("🔍 Buscar Producto")

    search_term = st.text_input("Ingrese Código o Nombre").strip().upper()

    if search_term:
        mask_s = (
            (df_sistema.iloc[:, 0].astype(str) == search_term) |
            (df_sistema.iloc[:, 1].astype(str).str.contains(search_term, case=False))
        )

        res_sistema = df_sistema[mask_s]

        if not res_sistema.empty:

            if len(res_sistema) > 1:
                st.warning("⚠️ Múltiples resultados:")
                for idx_s in res_sistema.index:
                    p_name = res_sistema.loc[idx_s, df_sistema.columns[1]]
                    p_code = clean_code(res_sistema.loc[idx_s, df_sistema.columns[0]])

                    if st.button(f"👉 {p_name} (Cód: {p_code})", key=f"btn_{p_code}"):
                        st.session_state.selected_code = p_code
                        st.session_state.selected_name = p_name
            else:
                st.session_state.selected_code = clean_code(res_sistema.iloc[0, 0])
                st.session_state.selected_name = res_sistema.iloc[0, 1]

        else:
            st.error("❌ Producto no encontrado")
            st.session_state.pop("selected_code", None)


    # ==========================================
    # PANEL DE EDICIÓN
    # ==========================================
    if 'selected_code' in st.session_state:

        real_code = st.session_state.selected_code
        prod_name = st.session_state.selected_name

        # Info sistema
        res_s_info = df_sistema[df_sistema.iloc[:, 0].astype(str) == real_code]
        stock_sys = res_s_info.iloc[0, 2] if not res_s_info.empty else "N/A"

        # Buscar en conteo
        mask_c = df_conteo.iloc[:, 0].astype(str) == real_code
        prod_row_idx = df_conteo[mask_c].index

        # Si no existe → crear fila
        if prod_row_idx.empty:
            st.info(f"✨ Agregando {prod_name} al conteo...")

            new_row = [real_code, prod_name] + [0] * 10
            df_conteo.loc[len(df_conteo)] = new_row

            prod_row_idx = [len(df_conteo) - 1]
            st.session_state.df_inv = df_conteo

        idx = prod_row_idx[0]

        # Header producto
        st.markdown(f"""
        <div class="product-header">
            <div style="font-size: 26px; font-weight: bold;">
                {prod_name}
            </div>
            <div>
                Código: {real_code} | Stock: {stock_sys}
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Inputs
        st.write("### 📝 Cantidades")

        col_names = ["BO1", "BO2", "BO3", "AL1", "AL2", "AL3", "VALES", "VENCIDOS"]

        current_vals = df_conteo.iloc[idx, 3:11].values
        current_values = [
            int(v) if pd.notnull(v) and str(v).replace('.', '').isdigit() else 0
            for v in current_vals
        ]

        inputs = {}

        row1 = st.columns(4)
        row2 = st.columns(4)

        for i, col_name in enumerate(col_names):
            container = row1 if i < 4 else row2

            with container[i % 4]:
                inputs[col_name] = st.number_input(
                    col_name,
                    min_value=0,
                    value=current_values[i]
                )

        total_fisico = sum(inputs.values())

        st.markdown(
            f"<div class='big-font'>TOTAL: {total_fisico}</div>",
            unsafe_allow_html=True
        )

        # Guardar
        if st.button("✅ GUARDAR", type="primary"):

            bodega_map = {
                "BO1": 3, "BO2": 4, "BO3": 5,
                "AL1": 6, "AL2": 7, "AL3": 8,
                "VALES": 9, "VENCIDOS": 10
            }

            for col_name, value in inputs.items():
                df_conteo.iloc[idx, bodega_map[col_name]] = value

            df_conteo.iloc[idx, 11] = total_fisico

            st.success("Guardado correctamente")
            st.balloons()


    # ==========================================
    # DESCARGA
    # ==========================================
    st.divider()

    if st.button("💾 DESCARGAR EXCEL"):
        final_bytes = save_to_excel(df_conteo, wb)

        st.download_button(
            label="Descargar archivo",
            data=final_bytes,
            file_name="INVENTARIO_FINAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )