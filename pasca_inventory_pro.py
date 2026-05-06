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

# CSS Personalizado para botones gigantes y aspecto de App Nativa
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
# LÓGICA DE DATOS (EL MOTOR)
# ==========================================
def load_pasca_data(file):
    """Carga el Excel y crea el mapeo de presentaciones y conteo."""
    wb = openpyxl.load_workbook(file)
    
    # 1. Mapeo de Presentaciones (Factor de Caja)
    df_pres = pd.read_excel(file, sheet_name='PRESENTACIÓN')
    # Creamos un diccionario { 'Nombre/Codigo': Factor }
    # Usamos el nombre para facilitar la búsqueda de la IA
    mapping_pres = {}
    for _, row in df_pres.iterrows():
        name = str(row['DESCRIPCION']).strip().upper()
        code = str(row['CODIGO']).strip()
        factor = row['PRESENTACION'] if pd.notnull(row['PRESENTACION']) else 1
        mapping_pres[name] = {'factor': factor, 'code': code}
        mapping_pres[code] = {'factor': factor, 'code': code}

    # 2. Carga de Conteo Físico
    df_conteo = pd.read_excel(file, sheet_name='CONTEO_F', skiprows=3)
    # Ajustamos columnas para que coincidan con el archivo (CODIGO, DESCRIPCION, etc)
    df_conteo.columns = df_conteo.iloc[0]
    df_conteo = df_conteo[1:].reset_index(drop=True)
    
    return df_conteo, wb, mapping_pres

def save_to_excel(df_conteo, wb):
    """Guarda los datos actualizados en el libro original."""
    sheet = wb['CONTEO_F']
    # Escribimos los datos del dataframe en la hoja a partir de la fila 4
    for i, row in df_conteo.iterrows():
        row_num = i + 5 # Offset según estructura
        for col_num, value in enumerate(row.values, 1):
            sheet.cell(row=row_num, column=col_num).value = value
    
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# LÓGICA DE IA (VISION)
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
# FLUJO DE LA APP
# ==========================================
st.title("📦 PASCA Inventory Pro")

with st.sidebar:
    st.header("🔑 Acceso")
    api_key = st.text_input("Gemini API Key", type="password")
    st.divider()
    # Selector Global de Bodega
    bodegas = ["BO1", "BO2", "BO3", "AL1", "AL2", "AL3"]
    st.session_state.selected_bodega = st.selectbox("📍 Bodega Actual", bodegas)

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

    with col_cam:
        st.subheader("📷 Identificación")
        img_file = st.camera_input("Capturar Etiqueta")
        
        if img_file:
            img = Image.open(img_file)
            with st.spinner("Analizando producto..."):
                detected_text = identify_with_gemini(img, api_key)
                st.session_state.detected_text = detected_text
                st.success(f"Detectado: {detected_text}")

    with col_data:
        st.subheader("📝 Conteo")
        if 'detected_text' in st.session_state:
            term = st.session_state.detected_text
            
            # Búsqueda de coincidencias en el mapeo de presentaciones
            matches = [k for k in mapping.keys() if term in k or k in term]
            
            if matches:
                # Si hay más de un match, mostrar menú de selección
                if len(matches) > 1:
                    st.warning("⚠️ Se encontraron varias presentaciones. Seleccione la correcta:")
                    selected_match = st.selectbox("Presentación", matches)
                else:
                    selected_match = matches[0]
                
                # Obtener datos del producto seleccionado
                prod_info = mapping[selected_match]
                factor = prod_info['factor']
                code = prod_info['code']
                
                # Buscar la fila exacta en el DataFrame de conteo
                prod_row_idx = df[df.iloc[:, 0].astype(str) == code].index
                
                if not prod_row_idx.empty:
                    idx = prod_row_idx[0]
                    prod_name = df.iloc[idx, 1]
                    
                    st.markdown(f"""<div class="product-card">
                        <div class="big-font">{prod_name}</div>
                        <p><b>Código:</b> {code} | <b>Factor Caja:</b> {factor} unds.</p>
                    </div>""", unsafe_allow_html=True)
                    
                    # CALCULADORA DE CAJAS
                    c1, c2 = st.columns(2)
                    with c1:
                        cajas = st.number_input("📦 Cajas", min_value=0, step=1)
                    with c2:
                        sueltos = st.number_input("낱 Unidades Sueltas", min_value=0, step=1)
                    
                    total = (cajas * factor) + sueltos
                    st.markdown(f"<div class='big-font'>Total: {total}</div>", unsafe_allow_html=True)
                    
                    if st.button("✅ Confirmar y Guardar"):
                        # Ubicar columna de bodega
                        # En la plantilla, BO1 es col 3, BO2 col 4...
                        bodega_map = {"BO1": 3, "BO2": 4, "BO3": 5, "AL1": 6, "AL2": 7, "AL3": 8}
                        col_idx = bodega_map[st.session_state.selected_bodega]
                        
                        df.iloc[idx, col_idx] = total
                        st.balloons()
                        st.success(f"Guardado {total} unidades en {st.session_state.selected_bodega}")
                else:
                    st.error("El producto fue identificado pero no existe en la hoja de CONTEO_F.")
            else:
                st.error("No se pudo encontrar el producto en el catálogo de presentaciones.")

    # Descarga final
    st.divider()
    if st.button("💾 EXPORTAR INVENTARIO FINAL"):
        final_bytes = save_to_excel(df, wb)
        st.download_button(
            label="Descargar Excel para Sistema",
            data=final_bytes,
            file_name="INVENTARIO_PASCA_FINAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )