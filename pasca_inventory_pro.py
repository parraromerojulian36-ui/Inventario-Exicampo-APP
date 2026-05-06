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
# LÓGICA DE DATOS (EL MOTOR CORREGIDO)
# ==========================================
def load_pasca_data(file):
    """Carga el Excel y crea el mapeo de presentaciones y conteo con limpieza de columnas."""
    wb = openpyxl.load_workbook(file)
    
    # 1. Mapeo de Presentaciones (Factor de Caja)
    df_pres = pd.read_excel(file, sheet_name='PRESENTACIÓN')
    # LIMPIEZA CRÍTICA: Eliminamos espacios en blanco de los nombres de las columnas
    df_pres.columns = df_pres.columns.str.strip()
    
    mapping_pres = {}
    # Buscamos las columnas correctas sin importar los espacios
    col_desc = 'DESCRIPCION'
    col_code = 'CODIGO'
    col_pres = 'PRESENTACION'

    for _, row in df_pres.iterrows():
        name = str(row[col_desc]).strip().upper()
        code = str(row[col_code]).strip()
        factor = row[col_pres] if pd.notnull(row[col_pres]) else 1
        mapping_pres[name] = {'factor': factor, 'code': code}
        mapping_pres[code] = {'factor': factor, 'code': code}

    # 2. Carga de Conteo Físico
    # Cargamos la hoja saltando las filas vacías hasta llegar a los datos
    df_conteo = pd.read_excel(file, sheet_name='CONTEO_F')
    
    # Buscamos la fila donde está la palabra "CODIGO" para establecer el encabezado
    header_row_index = 0
    for i, row in df_conteo.iterrows():
        if "CODIGO" in str(row.values).upper():
            header_row_index = i
            break
            
    # Reestructuramos el DataFrame desde la fila de encabezados
    df_conteo.columns = df_conteo.iloc[header_row_index].str.strip()
    df_conteo = df_conteo.iloc[header_row_index + 1:].reset_index(drop=True)
    
    return df_conteo, wb, mapping_pres

def save_to_excel(df_conteo, wb):
    """Guarda los datos actualizados en el libro original."""
    sheet = wb['CONTEO_F']
    
    # Encontramos la fila de inicio buscando "CODIGO" en la hoja
    start_row = 1
    for row in sheet.iter_rows(max_row=10):
        for cell in row:
            if cell.value and "CODIGO" in str(cell.value).upper():
                start_row = cell.row + 1
                break
            # Escribimos los valores del DataFrame en las celdas
    for i, row in df_conteo.iterrows():
        row_num = start_row + i
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
    bodegas = ["BO1", "BO2", "BO3", "AL1", "AL2", "AL3"]
    st.session_state.selected_bodega = st.selectbox("📍 Bodega Actual", bodegas)

uploaded_file = st.file_uploader("Cargar Plantilla de Sistema", type=["xlsx"])

if uploaded_file:
    if 'df_inv' not in st.session_state:
        try:
            df_c, wb, mapping = load_pasca_data(uploaded_file)
            st.session_state.df_inv = df_c
            st.session_state.wb_inv = wb
            st.session_state.mapping_pres = mapping
        except Exception as e:
            st.error(f"Error al cargar el Excel: {e}")
            st.stop()

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
            matches = [k for k in mapping.keys() if term in k or k in term]
            
            if matches:
                if len(matches) > 1:
                    st.warning("⚠️ Se encontraron varias presentaciones. Seleccione la correcta:")
                    selected_match = st.selectbox("Presentación", matches)
                else:
                    selected_match = matches[0]
                
                prod_info = mapping[selected_match]
                factor = prod_info['factor']
                code = prod_info['code']
                
                # Búsqueda robusta en la columna de código (Columna 0)
                mask = df.iloc[:, 0].astype(str).str.strip() == str(code).strip()
                prod_row_idx = df[mask].index
                
                if not prod_row_idx.empty:
                    idx = prod_row_idx[0]
                    prod_name = df.iloc[idx, 1]
                    
                    st.markdown(f"""<div class="product-card">
                        <div class="big-font">{prod_name}</div>
                        <p><b>Código:</b> {code} | <b>Factor Caja:</b> {factor} unds.</p>
                    </div>""", unsafe_allow_html=True)
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        cajas = st.number_input("📦 Cajas", min_value=0, step=1)
                    with c2:
                        sueltos = st.number_input("낱 Unidades Sueltas", min_value=0, step=1)
                    
                    total = (cajas * factor) + sueltos
                    st.markdown(f"<div class='big-font'>Total: {total}</div>", unsafe_allow_html=True)
                    
                    if st.button("✅ Confirmar y Guardar"):
                        bodega_map = {"BO1": 3, "BO2": 4, "BO3": 5, "AL1": 6, "AL2": 7, "AL3": 8}
                        col_idx = bodega_map[st.session_state.selected_bodega]
                        
                        df.iloc[idx, col_idx] = total
                        st.balloons()
                        st.success(f"Guardado {total} unidades en {st.session_state.selected_bodega}")
                else:
                    st.error("El producto fue identificado pero no existe en la hoja de CONTEO_F.")
            else:
                st.error("No se pudo encontrar el producto en el catálogo de presentaciones.")

    st.divider()
    if st.button("💾 EXPORTAR INVENTARIO FINAL"):
        final_bytes = save_to_excel(df, wb)
        st.download_button(
            label="Descargar Excel para Sistema",
            data=final_bytes,
            file_name="INVENTARIO_PASCA_FINAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )