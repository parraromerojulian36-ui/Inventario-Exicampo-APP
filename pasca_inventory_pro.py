import streamlit as st
import pandas as pd
import google.generativeai as genai
from PIL import Image
import openpyxl
from io import BytesIO

# ==========================================
# CONFIGURACIÓN DE INTERFAZ Y ESTILO (UX TABLET)
# ==========================================
st.set_page_config(page_title="PASCA Inventory Pro", layout="wide")

st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
        height: 80px;
        font-size: 24px !important;
        font-weight: bold !important;
        border-radius: 15px !important;
        border: 2px solid #4CAF50 !important;
    }
    .stSelectbox label, .stNumberInput label {
        font-size: 20px !important;
        font-weight: bold !important;
    }
    .big-font {
        font-size: 30px !important;
        font-weight: bold;
        text-align: center;
        color: #2E7D32;
    }
    .product-card {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 15px;
        border-left: 10px solid #4CAF50;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# LÓGICA DE DATOS
# ==========================================
def load_pasca_data(file):
    """Carga el Excel y crea el mapeo de presentaciones y conteo."""
    wb = openpyxl.load_workbook(file)

    # PRESENTACIONES
    df_pres = pd.read_excel(file, sheet_name='PRESENTACIÓN')
    mapping_pres = {}

    for _, row in df_pres.iterrows():
        name = str(row['DESCRIPCION']).strip().upper()
        code = str(row['CODIGO']).strip()
        factor = row['PRESENTACION'] if pd.notnull(row['PRESENTACION']) else 1

        mapping_pres[name] = {'factor': factor, 'code': code}
        mapping_pres[code] = {'factor': factor, 'code': code}

    # CONTEO
    df_conteo = pd.read_excel(file, sheet_name='CONTEO_F', skiprows=3)
    df_conteo.columns = df_conteo.iloc[0]
    df_conteo = df_conteo[1:].reset_index(drop=True)

    return df_conteo, wb, mapping_pres


def save_to_excel(df_conteo, wb):
    """Guarda los datos actualizados en el libro original."""
    sheet = wb['CONTEO_F']

    for i, row in df_conteo.iterrows():
        row_num = i + 5
        for col_num, value in enumerate(row.values, 1):
            sheet.cell(row=row_num, column=col_num).value = value

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# ==========================================
# LÓGICA DE IA
# ==========================================
def identify_with_gemini(image, api_key):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash')

    prompt = (
        "Analiza la etiqueta del producto agroquímico. "
        "Extrae el nombre comercial exacto o el código. "
        "Devuelve solo el texto, sin comentarios."
    )

    response = model.generate_content([prompt, image])
    return response.text.strip().upper()


# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
st.title("📦 PASCA Inventory Pro")

# SIDEBAR
with st.sidebar:
    st.header("🔑 Acceso")
    api_key = st.text_input("Gemini API Key", type="password")

    st.divider()

    bodegas = ["BO1", "BO2", "BO3", "AL1", "AL2", "AL3"]
    st.session_state.selected_bodega = st.selectbox("📍 Bodega Actual", bodegas)


# CARGA DE ARCHIVO
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

    col_cam, col_data = st.columns([1, 1])

    # ==========================================
    # CÁMARA
    # ==========================================
    with col_cam:
        st.subheader("📷 Identificación")
        img_file = st.camera_input("Capturar Etiqueta")

        if img_file:
            img = Image.open(img_file)

            with st.spinner("Analizando producto..."):
                detected_text = identify_with_gemini(img, api_key)
                st.session_state.detected_text = detected_text

            st.success(f"Detectado: {detected_text}")

    # ==========================================
    # CONTEO
    # ==========================================
    with col_data:
        st.subheader("📝 Conteo")

        if 'detected_text' in st.session_state:
            term = st.session_state.detected_text

            matches = [k for k in mapping.keys() if term in k or k in term]

            if matches:

                if len(matches) > 1:
                    st.warning("⚠️ Varias coincidencias encontradas")
                    selected_match = st.selectbox("Seleccionar presentación", matches)
                else:
                    selected_match = matches[0]

                prod_info = mapping[selected_match]
                factor = prod_info['factor']
                code = prod_info['code']

                prod_row_idx = df[df.iloc[:, 0].astype(str) == code].index

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
                    st.error("Producto no encontrado en CONTEO_F")

            else:
                st.error("No se encontró el producto en presentaciones")

    # ==========================================
    # EXPORTAR
    # ==========================================
    st.divider()

    if st.button("💾 EXPORTAR INVENTARIO FINAL"):
        final_bytes = save_to_excel(df, wb)

        st.download_button(
            label="Descargar Excel",
            data=final_bytes,
            file_name="INVENTARIO_PASCA_FINAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )