import streamlit as st
import pandas as pd
import openpyxl
import os
import tempfile
import difflib
import html
from datetime import datetime
from PIL import Image
import pytesseract

# ==========================================
# CONFIGURACIÓN DE PÁGINA
# ==========================================
st.set_page_config(
    page_title="PASCA Inventory Audit Pro",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# CSS MEJORADO (RESPONSIVE)
# ==========================================
st.markdown("""
<style>
    /* Ajustes generales de contenedor */
    .block-container {
        padding-top: 1rem;
        padding-bottom: 5rem;
    }

    /* Tarjetas de producto */
    .product-card {
        background: #ffffff;
        padding: 16px;
        border-radius: 12px;
        border-left: 6px solid #2E7D32;
        margin-bottom: 15px;
        box-shadow: 0px 2px 10px rgba(0,0,0,0.08);
    }

    .product-card h3 {
        margin: 0 0 8px 0;
        color: #2E7D32;
        font-size: 1.1rem;
        line-height: 1.3;
    }

    .product-card p {
        margin: 4px 0;
        font-size: 0.9rem;
        color: #333;
    }

    /* Indicador de Total */
    .big-total {
        background: #2E7D32;
        color: white;
        text-align: center;
        padding: 20px;
        border-radius: 12px;
        font-size: 2rem;
        font-weight: bold;
        margin: 20px 0;
        box-shadow: 0px 4px 12px rgba(46, 125, 50, 0.3);
    }

    /* Estilo global para botones de Streamlit */
    div.stButton > button {
        border-radius: 10px !important;
        height: auto !important;
        min-height: 48px !important;
        font-weight: bold !important;
        font-size: 16px !important;
        margin-top: 5px;
    }

    /* Ajustes específicos para móviles */
    @media (max-width: 768px) {
        .big-total {
            font-size: 1.5rem;
            padding: 15px;
        }
        .product-card h3 {
            font-size: 1rem;
        }
        div.stButton > button {
            font-size: 14px !important;
            min-height: 44px !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# UTILIDADES
# ==========================================
def clean_code(val):
    if pd.isna(val):
        return ""
    val = str(val).strip()
    if val.endswith(".0"):
        val = val[:-2]
    return val

def detect_text_ocr(image):
    try:
        text = pytesseract.image_to_string(image, lang="eng")
        text = text.upper().replace("\n", " ").replace("  ", " ")
        return text.strip()
    except Exception as e:
        return f"ERROR OCR: {str(e)}"

def search_product(df_sistema, detected_text):
    detected_text = detected_text.upper()
    # Búsqueda Exacta o Parcial
    exact = df_sistema[
        (df_sistema.iloc[:, 0].astype(str).str.upper().str.contains(detected_text, na=False)) |
        (df_sistema.iloc[:, 1].astype(str).str.upper().str.contains(detected_text, na=False))
    ]
    if not exact.empty:
        return exact
    # Búsqueda por similitud
    nombres = df_sistema.iloc[:, 1].astype(str).tolist()
    similars = difflib.get_close_matches(detected_text, nombres, n=5, cutoff=0.3)
    if similars:
        return df_sistema[df_sistema.iloc[:, 1].astype(str).isin(similars)]
    return pd.DataFrame()

def load_excel(uploaded_file):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.getvalue())
        temp_path = tmp.name
    
    wb = openpyxl.load_workbook(temp_path)
    df_sistema = pd.read_excel(temp_path, sheet_name="SISTEMA")
    df_sistema.columns = df_sistema.columns.str.strip()
    df_sistema.iloc[:, 0] = df_sistema.iloc[:, 0].apply(clean_code)

    df_conteo = pd.read_excel(temp_path, sheet_name="CONTEO_F")
    header_row = 0
    for i, row in df_conteo.iterrows():
        if "CODIGO" in str(row.values).upper():
            header_row = i
            break
    df_conteo.columns = df_conteo.iloc[header_row].astype(str).str.strip()
    df_conteo = df_conteo.iloc[header_row + 1:].reset_index(drop=True)
    df_conteo = df_conteo.astype(object)
    df_conteo.iloc[:, 0] = df_conteo.iloc[:, 0].apply(clean_code)

    st.session_state.temp_file = temp_path
    return df_conteo, df_sistema, wb

def save_full_audit(df_conteo, df_sistema, wb):
    sheet = wb["CONTEO_F"]
    start_row = 1
    for row in sheet.iter_rows(max_row=15):
        for cell in row:
            if cell.value and "CODIGO" in str(cell.value).upper():
                start_row = cell.row + 1
                break

    for i, row in df_conteo.iterrows():
        row_num = start_row + i
        for col_num, value in enumerate(row.values, 1):
            sheet.cell(row=row_num, column=col_num).value = value

    result_sheet = wb["RESULTADO"]
    for row in result_sheet.iter_rows(min_row=5):
        for cell in row: cell.value = None

    row_res = 5
    for _, row_c in df_conteo.iterrows():
        code = clean_code(row_c.iloc[0])
        name = row_c.iloc[1]
        total_fisico = row_c.iloc[11] if not pd.isna(row_c.iloc[11]) else 0
        match = df_sistema[df_sistema.iloc[:, 0].astype(str) == code]

        if not match.empty:
            total_sistema = match.iloc[0, 2] if not pd.isna(match.iloc[0, 2]) else 0
            diferencia = total_fisico - total_sistema
            faltantes = abs(diferencia) if diferencia < 0 else 0
            sobrantes = diferencia if diferencia > 0 else 0

            result_sheet.cell(row=row_res, column=1).value = code
            result_sheet.cell(row=row_res, column=2).value = name
            result_sheet.cell(row=row_res, column=3).value = total_fisico
            result_sheet.cell(row=row_res, column=4).value = total_sistema
            result_sheet.cell(row=row_res, column=5).value = diferencia
            result_sheet.cell(row=row_res, column=6).value = faltantes
            result_sheet.cell(row=row_res, column=7).value = sobrantes
            row_res += 1

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        path = tmp.name
    wb.save(path)
    with open(path, "rb") as f:
        data = f.read()
    return data

# ==========================================
# INTERFAZ DE USUARIO (UI)
# ==========================================
st.title("📦 PASCA Audit Pro")

with st.sidebar:
    st.header("Configuración")
    sucursal = st.selectbox("Sucursal", ["PASCA", "SUBIA", "SIBATE", "GRANADA"])
    fecha = datetime.now().strftime("%d-%m-%Y")

uploaded_file = st.file_uploader("Sube el archivo Excel de Inventario", type=["xlsx"])

if uploaded_file:
    if "df_inv" not in st.session_state:
        df_c, df_s, wb = load_excel(uploaded_file)
        st.session_state.df_inv = df_c
        st.session_state.df_sistema = df_s
        st.session_state.wb = wb

    df_conteo = st.session_state.df_inv
    df_sistema = st.session_state.df_sistema
    wb = st.session_state.wb

    # --- SECCIÓN OCR ---
    st.subheader("📷 Escanear Producto")
    img_file = st.camera_input("Tomar foto")

    if img_file:
        image = Image.open(img_file)
        with st.spinner("Procesando imagen..."):
            detected_text = detect_text_ocr(image)
        st.info(f"Texto detectado: {detected_text}")
        results = search_product(df_sistema, detected_text)

        if results.empty:
            st.warning("No se encontraron coincidencias.")
        else:
            for idx in results.index:
                p_name = str(results.loc[idx].iloc[1])
                p_code = clean_code(results.loc[idx].iloc[0])
                p_stock = results.loc[idx].iloc[2]
                
                st.markdown(f"""
                <div class="product-card">
                    <h3>{html.escape(p_name)}</h3>
                    <p><b>Código:</b> {html.escape(p_code)}</p>
                    <p><b>Stock Sistema:</b> {p_stock}</p>
                </div>
                """, unsafe_allow_html=True)
                
                if st.button(f"SELECCIONAR {p_code}", key=f"ocr_{p_code}", use_container_width=True):
                    st.session_state.selected_code = p_code
                    st.session_state.selected_name = p_name
                    st.rerun()

    # --- BÚSQUEDA MANUAL ---
    st.divider()
    search = st.text_input("🔍 Búsqueda manual (Nombre o Código)").upper()
    if search:
        mask = (df_sistema.iloc[:, 0].astype(str).str.contains(search, na=False)) | \
               (df_sistema.iloc[:, 1].astype(str).str.upper().str.contains(search, na=False))
        results = df_sistema[mask].head(10)

        for idx in results.index:
            m_name = results.loc[idx].iloc[1]
            m_code = clean_code(results.loc[idx].iloc[0])
            st.markdown(f"""
            <div class="product-card">
                <h3>{m_name}</h3>
                <p><b>Código:</b> {m_code}</p>
            </div>
            """, unsafe_allow_html=True)
            if st.button(f"SELECCIONAR {m_code}", key=f"man_{m_code}", use_container_width=True):
                st.session_state.selected_code = m_code
                st.session_state.selected_name = m_name
                st.rerun()

    # --- EDITOR DE CANTIDADES ---
    if "selected_code" in st.session_state:
        st.divider()
        code = st.session_state.selected_code
        name = st.session_state.selected_name
        
        st.markdown(f"""
        <div style="background:#e8f5e9; padding:15px; border-radius:10px; border:1px solid #2e7d32">
            <h2 style="margin:0; color:#1b5e20;">📝 Editando: {name}</h2>
            <p style="margin:0; color:#1b5e20;">Código: {code}</p>
        </div>
        """, unsafe_allow_html=True)

        idxs = df_conteo[df_conteo.iloc[:, 0].astype(str) == code].index
        if idxs.empty:
            new_row = [0] * len(df_conteo.columns)
            new_row[0], new_row[1] = code, name
            df_conteo.loc[len(df_conteo)] = new_row
            idx = len(df_conteo) - 1
        else:
            idx = idxs[0]

        bodegas = ["BO1", "BO2", "BO3", "AL1", "AL2", "AL3", "VALES", "VENCIDOS"]
        inputs = {}
        cols = st.columns(2)
        for i, bodega in enumerate(bodegas):
            with cols[i % 2]:
                val_orig = df_conteo.iloc[idx, i+3]
                inputs[bodega] = st.number_input(bodega, min_value=0, value=int(val_orig) if not pd.isna(val_orig) else 0, key=f"inp_{bodega}")

        total = sum(inputs.values())
        st.markdown(f'<div class="big-total">TOTAL: {total}</div>', unsafe_allow_html=True)

        if st.button("💾 GUARDAR CAMBIOS", type="primary", use_container_width=True):
            for i, bodega in enumerate(bodegas):
                df_conteo.iloc[idx, i+3] = inputs[bodega]
            df_conteo.iloc[idx, 11] = total
            st.success("¡Datos actualizados!")
            st.balloons()

    # --- EXPORTACIÓN ---
    st.divider()
    if st.button("📥 GENERAR EXCEL FINAL", use_container_width=True):
        final_data = save_full_audit(df_conteo, df_sistema, wb)
        st.download_button(
            label="⬇️ DESCARGAR AHORA",
            data=final_data,
            file_name=f"AUDITORIA_{sucursal}_{fecha}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )