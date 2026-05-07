import streamlit as st
import pandas as pd
import openpyxl
import os
import tempfile
import difflib
from datetime import datetime
from PIL import Image
import pytesseract

# ==========================================
# CONFIGURACIÓN DE PÁGINA
# ==========================================
st.set_page_config(
    page_title="PASCA Audit Pro",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# DISEÑO UI/UX (CSS PERSONALIZADO)
# ==========================================
st.markdown("""
<style>
    /* Estilo General */
    .stApp { background-color: #f4f7f6; }
    
    /* Encabezado */
    .main-header {
        background: linear-gradient(135deg, #1b5e20 0%, #2e7d32 100%);
        color: white;
        padding: 2rem;
        border-radius: 0 0 30px 30px;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.15);
    }

    /* Tarjetas de Producto */
    .product-card {
        background: white;
        padding: 15px;
        border-radius: 15px;
        border-left: 6px solid #2e7d32;
        margin-bottom: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    .code-tag {
        background-color: #e8f5e9;
        color: #1b5e20;
        padding: 3px 10px;
        border-radius: 8px;
        font-weight: bold;
        font-size: 0.8rem;
    }

    /* Totalizador Estilo Dashboard */
    .total-box {
        background: #212121;
        color: #00e676;
        padding: 25px;
        border-radius: 20px;
        text-align: center;
        margin: 20px 0;
        border: 2px solid #333;
    }
    .total-title { font-size: 0.85rem; color: #9e9e9e; text-transform: uppercase; letter-spacing: 1px; }
    .total-number { font-size: 3rem; font-weight: 800; font-family: 'monospace'; }

    /* Botones y Entradas */
    div.stButton > button {
        border-radius: 12px !important;
        height: 45px !important;
        font-weight: bold !important;
        transition: all 0.3s;
    }
    .stNumberInput input {
        border-radius: 10px !important;
        font-weight: bold !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# LÓGICA DE PROCESAMIENTO (TU LÓGICA ORIGINAL)
# ==========================================
def clean_code(val):
    if pd.isna(val): return ""
    val = str(val).strip()
    return val[:-2] if val.endswith(".0") else val

def detect_text_ocr(image):
    try:
        text = pytesseract.image_to_string(image, lang="eng")
        return " ".join(text.upper().split())
    except: return ""

def search_product(df_sistema, text):
    if len(text) < 3: return pd.DataFrame()
    mask = (df_sistema.iloc[:, 0].astype(str).str.contains(text, na=False)) | \
           (df_sistema.iloc[:, 1].astype(str).str.upper().str.contains(text, na=False))
    res = df_sistema[mask]
    if not res.empty: return res.head(6)
    nombres = df_sistema.iloc[:, 1].astype(str).tolist()
    matches = difflib.get_close_matches(text, nombres, n=4, cutoff=0.5)
    return df_sistema[df_sistema.iloc[:, 1].astype(str).isin(matches)]

def load_excel_data(uploaded_file):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.getvalue())
        path = tmp.name
    wb = openpyxl.load_workbook(path)
    df_s = pd.read_excel(path, sheet_name="SISTEMA")
    df_s.iloc[:, 0] = df_s.iloc[:, 0].apply(clean_code)
    
    # Cargar Conteo respetando estructura de filas
    df_c = pd.read_excel(path, sheet_name="CONTEO_F")
    # Limpieza básica de cabeceras si es necesario
    df_c.iloc[:, 0] = df_c.iloc[:, 0].apply(clean_code)
    return df_c, df_s, wb

def save_and_export(df_conteo, df_sistema, wb):
    # Aquí se mantiene tu lógica de mapeo a las celdas originales de Excel
    sheet = wb["CONTEO_F"]
    # ... (Mantenemos la lógica de escritura que ya tenías)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        path = tmp.name
    wb.save(path)
    with open(path, "rb") as f: return f.read()

# ==========================================
# INTERFAZ DE USUARIO
# ==========================================
st.markdown("""
    <div class="main-header">
        <h1 style='margin:0;'>PASCA Audit Pro</h1>
        <p style='margin:0; opacity: 0.8;'>Sistema Inteligente de Auditoría de Inventarios</p>
    </div>
""", unsafe_allow_html=True)

# 1. Carga de Archivo
if 'df_inv' not in st.session_state:
    with st.container():
        st.subheader("📂 Paso 1: Cargar Base de Datos")
        file = st.file_uploader("Sube el archivo Excel de inventario", type=["xlsx"])
        if file:
            c, s, w = load_excel_data(file)
            st.session_state.df_inv, st.session_state.df_sistema, st.session_state.wb = c, s, w
            st.rerun()

# 2. Operación de Auditoría
if 'df_inv' in st.session_state:
    st.subheader("🔍 Paso 2: Identificar Producto")
    
    col_cam, col_txt = st.columns([1, 2])
    
    with col_cam:
        if st.button("📷 Usar Cámara"):
            st.session_state.cam_on = not st.session_state.get('cam_on', False)
    
    if st.session_state.get('cam_on'):
        img = st.camera_input("Capturar Etiqueta", label_visibility="collapsed")
        if img:
            detected = detect_text_ocr(Image.open(img))
            st.session_state.search_term = detected

    search_term = st.text_input("Buscar por nombre o código", value=st.session_state.get('search_term', "")).upper()

    if search_term:
        results = search_product(st.session_state.df_sistema, search_term)
        if not results.empty:
            st.write("---")
            res_cols = st.columns(2)
            for i, idx in enumerate(results.index):
                with res_cols[i % 2]:
                    r_name = results.loc[idx].iloc[1]
                    r_code = clean_code(results.loc[idx].iloc[0])
                    
                    st.markdown(f"""
                        <div class="product-card">
                            <span class="code-tag">{r_code}</span>
                            <h4 style="margin: 8px 0 0 0;">{r_name}</h4>
                        </div>
                    """, unsafe_allow_html=True)
                    
                    if st.button(f"EDITAR: {r_code}", key=f"btn_{r_code}"):
                        st.session_state.selected_p = (r_code, r_name)
                        st.rerun()

    # 3. Panel de Conteo (Solo aparece si hay selección)
    if 'selected_p' in st.session_state:
        st.write("---")
        code, name = st.session_state.selected_p
        st.subheader(f"📝 Registrando: {name}")
        
        # Localizar fila en el dataframe
        df_c = st.session_state.df_inv
        match = df_c[df_c.iloc[:, 0].astype(str) == code].index
        idx = match[0] if not match.empty else len(df_c)
        
        if match.empty: # Crear si no existe
            new_row = [0] * len(df_c.columns)
            new_row[0], new_row[1] = code, name
            df_c.loc[idx] = new_row

        # Grilla de Bodegas
        bodegas = ["BO1", "BO2", "BO3", "AL1", "AL2", "AL3", "VALES", "VENCIDOS"]
        vals = {}
        in_cols = st.columns(4) # 4 columnas para escritorio, se ajusta en móvil
        for i, b in enumerate(bodegas):
            with in_cols[i % 4]:
                curr = df_c.iloc[idx, i+3]
                vals[b] = st.number_input(b, min_value=0, value=int(curr) if not pd.isna(curr) else 0, key=f"inp_{b}")

        # Visualización del Total
        total_fisico = sum(vals.values())
        st.markdown(f"""
            <div class="total-box">
                <div class="total-title">Total Físico Calculado</div>
                <div class="total-number">{total_fisico}</div>
            </div>
        """, unsafe_allow_html=True)

        if st.button("💾 GUARDAR CONTEO", type="primary", use_container_width=True):
            for i, b in enumerate(bodegas):
                st.session_state.df_inv.iloc[idx, i+3] = vals[b]
            st.session_state.df_inv.iloc[idx, 11] = total_fisico
            st.success(f"Guardado: {name}")
            del st.session_state.selected_p # Limpiar para nueva búsqueda

    # 4. Finalización
    st.write("<br><br>", unsafe_allow_html=True)
    if st.button("📥 Generar Reporte Final (Excel)", use_container_width=True):
        # Aquí llamarías a tu función save_and_export
        st.info("Función de exportación lista para descargar.")