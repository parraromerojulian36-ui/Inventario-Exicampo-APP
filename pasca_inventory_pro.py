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
.stNumberInput label {
    font-size: 18px !important;
    font-weight: bold !important;
}
.big-font {
    font-size: 36px;
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
</style>
""", unsafe_allow_html=True)

# ==========================================
# UTILIDAD
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

uploaded_file = st.file_uploader("Sube el Excel del sistema", type=["xlsx"])

if uploaded_file:

    if 'loaded' not in st.session_state:
        df_c, df_s, wb = load_pasca_data(uploaded_file)
        st.session_state.df_conteo = df_c
        st.session_state.df_sistema = df_s
        st.session_state.wb = wb
        st.session_state.loaded = True

    df_conteo = st.session_state.df_conteo
    df_sistema = st.session_state.df_sistema
    wb = st.session_state.wb

    # ======================================
    # BUSCAR
    # ======================================
    st.subheader("🔍 Buscar Producto")

    search_term = st.text_input("Código o nombre").strip().upper()

    if search_term:

        mask = (
            (df_sistema.iloc[:, 0].astype(str) == search_term) |
            (df_sistema.iloc[:, 1].astype(str).str.contains(search_term, case=False))
        )

        res = df_sistema[mask]

        if not res.empty:

            if len(res) > 1:
                st.warning("⚠️ Múltiples resultados")

                for i in range(len(res)):
                    name = res.iloc[i, 1]
                    code = clean_code(res.iloc[i, 0])

                    if st.button(f"{name} ({code})", key=f"btn_{i}"):
                        st.session_state.selected = (code, name)

            else:
                st.session_state.selected = (
                    clean_code(res.iloc[0, 0]),
                    res.iloc[0, 1]
                )

        else:
            st.error("Producto no encontrado")
            if 'selected' in st.session_state:
                del st.session_state.selected

    # ======================================
    # EDICIÓN
    # ======================================
    if 'selected' in st.session_state:

        code, name = st.session_state.selected

        stock_row = df_sistema[df_sistema.iloc[:, 0].astype(str) == code]
        stock_sys = stock_row.iloc[0, 2] if not stock_row.empty else "N/A"

        row_idx = df_conteo[df_conteo.iloc[:, 0].astype(str) == code].index

        # 🔥 CREAR PRODUCTO SI NO EXISTE
        if row_idx.empty:
            st.info(f"Agregando producto nuevo: {name}")

            new_row = [code, name] + [0]*10
            df_conteo.loc[len(df_conteo)] = new_row

            row_idx = [len(df_conteo)-1]
            st.session_state.df_conteo = df_conteo

        idx = row_idx[0]

        st.markdown(f"""
        <div class="product-header">
            <b>{name}</b><br>
            Código: {code} | Stock Sistema: {stock_sys}
        </div>
        """, unsafe_allow_html=True)

        cols = ["BO1","BO2","BO3","AL1","AL2","AL3","VALES","VENCIDOS"]

        raw = df_conteo.iloc[idx, 3:11].values
        vals = [int(v) if pd.notnull(v) and str(v).replace('.', '').isdigit() else 0 for v in raw]

        inputs = {}

        r1 = st.columns(4)
        r2 = st.columns(4)

        for i, col in enumerate(cols):
            container = r1 if i < 4 else r2
            with container[i % 4]:
                inputs[col] = st.number_input(col, value=vals[i], min_value=0)

        total = sum(inputs.values())

        st.markdown(f"<div class='big-font'>TOTAL FÍSICO: {total}</div>", unsafe_allow_html=True)

        if st.button("Guardar", type="primary"):

            map_cols = {
                "BO1":3,"BO2":4,"BO3":5,
                "AL1":6,"AL2":7,"AL3":8,
                "VALES":9,"VENCIDOS":10
            }

            for k, v in inputs.items():
                df_conteo.iloc[idx, map_cols[k]] = v

            df_conteo.iloc[idx, 11] = total

            st.success("Guardado correctamente")
            st.balloons()

    # ======================================
    # EXPORTAR
    # ======================================
    st.divider()

    if st.button("💾 Descargar Excel Final"):
        file = save_to_excel(df_conteo, wb)
        st.download_button("Descargar", file, "inventario_final.xlsx")