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
# CSS MEJORADO (COMPACTO Y RESPONSIVE)
# ==========================================
st.markdown("""
<style>
    .block-container {
        padding-top: 1rem;
        padding-bottom: 5rem;
    }

    /* Tarjetas mucho más compactas */
    .product-card {
        background: #ffffff;
        padding: 10px;
        border-radius: 8px;
        border-left: 4px solid #2E7D32;
        margin-bottom: 10px;
        box-shadow: 0px 2px 5px rgba(0,0,0,0.05);
    }

    .product-card h3 {
        margin: 0 0 5px 0;
        color: #2E7D32;
        font-size: 0.95rem;
        line-height: 1.2;
    }

    .product-card p {
        margin: 2px 0;
        font-size: 0.8rem;
        color: #555;
    }

    .big-total {
        background: #2E7D32;
        color: white;
        text-align: center;
        padding: 15px;
        border-radius: 10px;
        font-size: 1.6rem;
        font-weight: bold;
        margin: 15px 0;
    }

    /* Botones más pequeños para ahorrar scroll vertical */
    div.stButton > button {
        border-radius: 8px !important;
        height: auto !important;
        min-height: 38px !important;
        font-weight: bold !important;
        font-size: 13px !important;
        padding: 4px 8px !important;
    }

    @media (max-width: 768px) {
        .big-total { font-size: 1.3rem; }
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# UTILIDADES
# ==========================================
def clean_code(val):
    if pd.isna(val): return ""
    val = str(val).strip()
    if val.endswith(".0"): val = val[:-2]
    return val

def detect_text_ocr(image):
    try:
        text = pytesseract.image_to_string(image, lang="eng")
        text = text.upper().replace("\n", " ").strip()
        return " ".join(text.split())
    except Exception: return ""

def search_product(df_sistema, detected_text):
    if len(detected_text) < 3: return pd.DataFrame()
    detected_text = detected_text.upper()
    
    # Búsqueda Exacta o Parcial Directa
    exact = df_sistema[
        (df_sistema.iloc[:, 0].astype(str).str.upper().str.contains(detected_text, na=False)) |
        (df_sistema.iloc[:, 1].astype(str).str.upper().str.contains(detected_text, na=False))
    ]
    if not exact.empty:
        return exact.head(6)
    
    # Búsqueda por similitud MÁS ESTRICTA (0.5) para evitar basura
    nombres = df_sistema.iloc[:, 1].astype(str).tolist()
    similars = difflib.get_close_matches(detected_text, nombres, n=4, cutoff=0.5)
    if similars:
        return df_sistema[df_sistema.iloc[:, 1].astype(str).isin(similars)]
    return pd.DataFrame()

def load_excel(uploaded_file):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.getvalue())
        temp_path = tmp.name
    
    wb = openpyxl.load_workbook(temp_path)
    df_sistema = pd.read_excel(temp_path, sheet_name="SISTEMA")
    df_sistema.iloc[:, 0] = df_sistema.iloc[:, 0].apply(clean_code)

    df_conteo = pd.read_excel(temp_path, sheet_name="CONTEO_F")
    header_row = 0
    for i, row in df_conteo.iterrows():
        if "CODIGO" in str(row.values).upper():
            header_row = i
            break
    df_conteo.columns = df_conteo.iloc[header_row].astype(str).str.strip()
    df_conteo = df_conteo.iloc[header_row + 1:].reset_index(drop=True)
    df_conteo.iloc[:, 0] = df_conteo.iloc[:, 0].apply(clean_code)

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

    # Lógica de Hoja RESULTADO (se mantiene igual para integridad de datos)
    result_sheet = wb["RESULTADO"]
    for row in result_sheet.iter_rows(min_row=5):
        for cell in row: cell.value = None
    row_res = 5
    for _, row_c in df_conteo.iterrows():
        code = clean_code(row_c.iloc[0])
        total_fisico = row_c.iloc[11] if not pd.isna(row_c.iloc[11]) else 0
        match = df_sistema[df_sistema.iloc[:, 0].astype(str) == code]
        if not match.empty:
            total_sistema = match.iloc[0, 2] if not pd.isna(match.iloc[0, 2]) else 0
            diferencia = total_fisico - total_sistema
            result_sheet.cell(row=row_res, column=1).value = code
            result_sheet.cell(row=row_res, column=2).value = row_c.iloc[1]
            result_sheet.cell(row=row_res, column=3).value = total_fisico
            result_sheet.cell(row=row_res, column=4).value = total_sistema
            result_sheet.cell(row=row_res, column=5).value = diferencia
            result_sheet.cell(row=row_res, column=6).value = abs(diferencia) if diferencia < 0 else 0
            result_sheet.cell(row=row_res, column=7).value = diferencia if diferencia > 0 else 0
            row_res += 1

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        path = tmp.name
    wb.save(path)
    with open(path, "rb") as f: return f.read()

# ==========================================
# INTERFAZ DE USUARIO (UI)
# ==========================================
st.title("📦 PASCA Audit Pro")

uploaded_file = st.file_uploader("Subir Inventario (Excel)", type=["xlsx"])

if uploaded_file:
    if "df_inv" not in st.session_state:
        df_c, df_s, wb = load_excel(uploaded_file)
        st.session_state.df_inv, st.session_state.df_sistema, st.session_state.wb = df_c, df_s, wb

    # --- SECCIÓN OCR EN GRILLA ---
    st.subheader("📷 Escanear")
    img_file = st.camera_input("Foto etiqueta", label_visibility="collapsed")

    if img_file:
        with st.spinner("Buscando..."):
            detected_text = detect_text_ocr(Image.open(img_file))
            results = search_product(st.session_state.df_sistema, detected_text)

        if results.empty:
            st.warning("Sin coincidencias claras. Prueba búsqueda manual.")
        else:
            st.caption(f"Resultados para: {detected_text}")
            # MOSTRAR EN 2 COLUMNAS PARA AHORRAR ESPACIO
            grid_cols = st.columns(2)
            for i, idx in enumerate(results.index):
                with grid_cols[i % 2]:
                    p_name = str(results.loc[idx].iloc[1])
                    p_code = clean_code(results.loc[idx].iloc[0])
                    
                    st.markdown(f"""
                    <div class="product-card">
                        <h3>{p_name[:35]}...</h3>
                        <p><b>Cod:</b> {p_code}</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    if st.button(f"EDITAR {p_code}", key=f"ocr_{p_code}", use_container_width=True):
                        st.session_state.selected_code = p_code
                        st.session_state.selected_name = p_name
                        st.rerun()

    # --- BÚSQUEDA MANUAL ---
    st.divider()
    search = st.text_input("🔍 Búsqueda manual").upper()
    if search:
        mask = (st.session_state.df_sistema.iloc[:, 0].astype(str).str.contains(search, na=False)) | \
               (st.session_state.df_sistema.iloc[:, 1].astype(str).str.upper().str.contains(search, na=False))
        results_m = st.session_state.df_sistema[mask].head(6)
        
        grid_cols_m = st.columns(2)
        for i, idx in enumerate(results_m.index):
            with grid_cols_m[i % 2]:
                m_name = results_m.loc[idx].iloc[1]
                m_code = clean_code(results_m.loc[idx].iloc[0])
                st.markdown(f'<div class="product-card"><h3>{m_name[:35]}</h3><p>Cod: {m_code}</p></div>', unsafe_allow_html=True)
                if st.button(f"SEL. {m_code}", key=f"man_{m_code}", use_container_width=True):
                    st.session_state.selected_code, st.session_state.selected_name = m_code, m_name
                    st.rerun()

    # --- EDITOR DE CANTIDADES ---
    if "selected_code" in st.session_state:
        st.divider()
        code, name = st.session_state.selected_code, st.session_state.selected_name
        
        st.markdown(f"**Editando:** {name} (`{code}`)")

        # Buscar índice en el dataframe de conteo
        df_c = st.session_state.df_inv
        idxs = df_c[df_c.iloc[:, 0].astype(str) == code].index
        if idxs.empty:
            new_row = [0] * len(df_c.columns)
            new_row[0], new_row[1] = code, name
            df_c.loc[len(df_c)] = new_row
            idx = len(df_c) - 1
        else: idx = idxs[0]

        bodegas = ["BO1", "BO2", "BO3", "AL1", "AL2", "AL3", "VALES", "VENCIDOS"]
        inputs = {}
        # Entradas en 2 columnas también
        ed_cols = st.columns(2)
        for i, bodega in enumerate(bodegas):
            with ed_cols[i % 2]:
                val = df_c.iloc[idx, i+3]
                inputs[bodega] = st.number_input(bodega, min_value=0, value=int(val) if not pd.isna(val) else 0, key=f"inp_{bodega}")

        total = sum(inputs.values())
        st.markdown(f'<div class="big-total">TOTAL: {total}</div>', unsafe_allow_html=True)

        if st.button("💾 GUARDAR CAMBIOS", type="primary", use_container_width=True):
            for i, bodega in enumerate(bodegas):
                st.session_state.df_inv.iloc[idx, i+3] = inputs[bodega]
            st.session_state.df_inv.iloc[idx, 11] = total
            st.success("Guardado")

    # --- EXPORTACIÓN ---
    st.divider()
    if st.button("📥 EXPORTAR EXCEL FINAL", use_container_width=True):
        final_data = save_full_audit(st.session_state.df_inv, st.session_state.df_sistema, st.session_state.wb)
        st.download_button("⬇️ DESCARGAR ARCHIVO", data=final_data, file_name=f"AUDITORIA_{datetime.now().strftime('%d-%m')}.xlsx", use_container_width=True)