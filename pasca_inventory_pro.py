import streamlit as st
import pandas as pd
import openpyxl
import os
import tempfile
from datetime import datetime
import google.generativeai as genai
from PIL import Image
from io import BytesIO

# OCR
import pytesseract
import cv2
import numpy as np
import re

# ==========================================
# CONFIG UI
# ==========================================
st.set_page_config(page_title="PASCA Inventory Audit Pro", layout="wide")

st.markdown("""
<style>
.stNumberInput label { font-size: 16px !important; font-weight: bold !important; }
.big-font {
    font-size: 28px;
    font-weight: bold;
    text-align: center;
    color: white;
    background-color: #2E7D32;
    padding: 12px;
    border-radius: 10px;
}
.product-header {
    background: white;
    padding: 15px;
    border-radius: 12px;
    border-left: 8px solid #4CAF50;
    margin-bottom: 15px;
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# UTILIDADES
# ==========================================
def clean_code(val):
    if pd.isna(val): return ""
    val = str(val).strip()
    return val[:-2] if val.endswith(".0") else val

# ==========================================
# PREPROCESAMIENTO OCR
# ==========================================
def preprocess_image(image):
    img = np.array(image)
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
    return thresh

# ==========================================
# IA HÍBRIDA (OCR + GEMINI)
# ==========================================
def identify_product_hybrid(image, api_key, model_name):
    try:
        # 1) OCR GRATIS
        processed = preprocess_image(image)
        text = pytesseract.image_to_string(processed)
        text_clean = text.upper().strip()

        # busca códigos tipo ABC123 o numéricos largos
        match = re.findall(r"\b[A-Z0-9]{6,}\b", text_clean)
        if match:
            best_match = sorted(match, key=len, reverse=True)[0]
            return best_match

        # 2) GEMINI (FALLBACK)
        genai.configure(api_key=api_key)

        modelos = [
            "models/gemini-2.5-flash",
            "models/gemini-2.0-flash",
            "models/gemini-2.0-flash-lite"
        ]
        if model_name not in modelos:
            modelos.insert(0, model_name)

        buffer = BytesIO()
        image.convert("RGB").save(buffer, format="JPEG")
        img_bytes = buffer.getvalue()

        prompt = (
            "Lee la etiqueta del producto agroquímico. "
            "Extrae el nombre o código exacto. "
            "Devuelve SOLO el resultado."
        )

        for m in modelos:
            try:
                model_ai = genai.GenerativeModel(m)
                response = model_ai.generate_content([
                    prompt,
                    {"mime_type": "image/jpeg", "data": img_bytes}
                ])
                if response and hasattr(response, "text"):
                    return response.text.strip().upper()
            except Exception as e:
                if "429" in str(e):
                    continue
                else:
                    return f"ERROR: {str(e)}"

        return "ERROR: Sin cuota disponible"

    except Exception as e:
        return f"ERROR: {str(e)}"

# ==========================================
# AGREGAR PRODUCTO A CONTEO
# ==========================================
def add_product_to_conteo(df_conteo, code, name):
    exists = df_conteo[df_conteo.iloc[:, 0].astype(str) == code]
    if exists.empty:
        new_row = [code, name] + [0] * (len(df_conteo.columns) - 2)
        df_conteo.loc[len(df_conteo)] = new_row
    return df_conteo

# ==========================================
# DATA
# ==========================================
def load_pasca_data(uploaded_file):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name

    wb = openpyxl.load_workbook(tmp_path)

    # SISTEMA
    df_sistema = pd.read_excel(tmp_path, sheet_name='SISTEMA')
    df_sistema.columns = df_sistema.columns.str.strip()
    df_sistema.iloc[:, 0] = df_sistema.iloc[:, 0].apply(clean_code)

    # CONTEO_F
    df_conteo = pd.read_excel(tmp_path, sheet_name='CONTEO_F')

    header_row = 0
    for i, row in df_conteo.iterrows():
        if "CODIGO" in str(row.values).upper():
            header_row = i
            break

    df_conteo.columns = df_conteo.iloc[header_row].astype(str).str.strip()
    df_conteo = df_conteo.iloc[header_row + 1:].reset_index(drop=True)
    df_conteo = df_conteo.astype(object)
    df_conteo.iloc[:, 0] = df_conteo.iloc[:, 0].apply(clean_code)

    st.session_state.temp_file = tmp_path
    return df_conteo, df_sistema, wb

def save_full_audit(df_conteo, df_sistema, wb):
    sheet = wb['CONTEO_F']

    # ubicar inicio de datos
    start_row = 1
    for row in sheet.iter_rows(max_row=10):
        for cell in row:
            if cell.value and "CODIGO" in str(cell.value).upper():
                start_row = cell.row + 1
                break

    # escribir CONTEO_F
    for i, row in df_conteo.iterrows():
        row_num = start_row + i
        for col_num, value in enumerate(row.values, 1):
            sheet.cell(row=row_num, column=col_num).value = value

    # limpiar RESULTADO
    sheet_res = wb['RESULTADO']
    for row in sheet_res.iter_rows(min_row=5):
        for cell in row:
            cell.value = None

    # llenar RESULTADO
    row_res = 5
    for _, row_c in df_conteo.iterrows():
        code = clean_code(row_c.iloc[0])
        name = row_c.iloc[1]
        total_fisico = row_c.iloc[11] if pd.notnull(row_c.iloc[11]) else 0

        match = df_sistema[df_sistema.iloc[:, 0].astype(str) == code]

        if not match.empty:
            total_sistema = match.iloc[0, 2] if pd.notnull(match.iloc[0, 2]) else 0
            diff = total_fisico - total_sistema

            sheet_res.cell(row=row_res, column=1).value = code
            sheet_res.cell(row=row_res, column=2).value = name
            sheet_res.cell(row=row_res, column=3).value = total_fisico
            sheet_res.cell(row=row_res, column=4).value = total_sistema
            sheet_res.cell(row=row_res, column=5).value = diff
            sheet_res.cell(row=row_res, column=6).value = abs(diff) if diff < 0 else "-"
            sheet_res.cell(row=row_res, column=7).value = diff if diff > 0 else "-"
            row_res += 1

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        path = tmp.name

    wb.save(path)

    with open(path, "rb") as f:
        data = f.read()

    try:
        os.remove(st.session_state.temp_file)
        os.remove(path)
    except:
        pass

    return data

# ==========================================
# UI
# ==========================================
st.title("📦 PASCA Inventory Audit Pro")

api_key = st.secrets.get("GEMINI_API_KEY", "")

with st.sidebar:
    if not api_key:
        api_key = st.text_input("API Key", type="password")
    else:
        st.success("API Key cargada")

    model_choice = st.selectbox("Modelo IA", [
        "models/gemini-2.5-flash",
        "models/gemini-2.0-flash",
        "models/gemini-2.0-flash-lite"
    ])

    sucursal = st.selectbox("Sucursal", ["PASCA", "SUBIA", "SIBATE", "GRANADA"])
    fecha = datetime.now().strftime("%d-%m-%Y")

uploaded_file = st.file_uploader("Sube Excel", type=["xlsx"])

# ==========================================
# APP
# ==========================================
if uploaded_file:

    if "df_inv" not in st.session_state:
        df_c, df_s, wb = load_pasca_data(uploaded_file)
        st.session_state.df_inv = df_c
        st.session_state.df_sistema = df_s
        st.session_state.wb = wb

    df_conteo = st.session_state.df_inv
    df_sistema = st.session_state.df_sistema
    wb = st.session_state.wb

    # -------------------------
    # CAMARA
    # -------------------------
    st.subheader("📷 Identificación")
    img_file = st.camera_input("Foto del producto")

    if img_file:
        if not api_key:
            st.error("Falta API Key")
        else:
            img = Image.open(img_file)

            with st.spinner("Analizando..."):
                detected = identify_product_hybrid(img, api_key, model_choice)

            if detected.startswith("ERROR"):
                st.error(detected)
            else:
                st.success(f"Detectado: {detected}")

                mask = (
                    (df_sistema.iloc[:, 0].astype(str) == detected) |
                    (df_sistema.iloc[:, 1].astype(str).str.contains(detected, case=False))
                )

                res = df_sistema[mask]

                for idx in res.index:
                    name = res.iloc[res.index.get_loc(idx), 1]
                    code = clean_code(res.iloc[res.index.get_loc(idx), 0])

                    if st.button(f"{name} ({code})", key=f"cam_{code}"):
                        st.session_state.df_inv = add_product_to_conteo(df_conteo, code, name)
                        st.session_state.selected_code = code
                        st.session_state.selected_name = name
                        st.rerun()

    # -------------------------
    # BUSCADOR
    # -------------------------
    st.subheader("🔍 Búsqueda")
    search = st.text_input("Código o nombre").upper()

    if search:
        mask = (
            (df_sistema.iloc[:, 0].astype(str) == search) |
            (df_sistema.iloc[:, 1].astype(str).str.contains(search, case=False))
        )

        res = df_sistema[mask]

        for idx in res.index:
            name = res.iloc[res.index.get_loc(idx), 1]
            code = clean_code(res.iloc[res.index.get_loc(idx), 0])

            if st.button(f"{name} ({code})", key=f"bus_{code}"):
                st.session_state.df_inv = add_product_to_conteo(df_conteo, code, name)
                st.session_state.selected_code = code
                st.session_state.selected_name = name
                st.rerun()

    # -------------------------
    # EDITOR
    # -------------------------
    if "selected_code" in st.session_state:

        code = st.session_state.selected_code
        name = st.session_state.selected_name

        match = df_sistema[df_sistema.iloc[:, 0].astype(str) == code]
        stock = match.iloc[0, 2] if not match.empty else "N/A"

        idx = df_conteo[df_conteo.iloc[:, 0].astype(str) == code].index[0]

        st.markdown(f"""
        <div class="product-header">
        <b>{name}</b><br>
        Código: {code} | Stock sistema: {stock}
        </div>
        """, unsafe_allow_html=True)

        cols = ["BO1","BO2","BO3","AL1","AL2","AL3","VALES","VENCIDOS"]

        values = (
            df_conteo.iloc[idx, 3:11]
            .fillna(0)
            .astype(int)
            .tolist()
        )

        while len(values) < 8:
            values.append(0)

        inputs = {}

        for i, col in enumerate(cols):
            inputs[col] = st.number_input(col, 0, value=int(values[i]))

        total = sum(inputs.values())
        st.markdown(f"<div class='big-font'>TOTAL: {total}</div>", unsafe_allow_html=True)

        if st.button("GUARDAR"):
            map_cols = {"BO1":3,"BO2":4,"BO3":5,"AL1":6,"AL2":7,"AL3":8,"VALES":9,"VENCIDOS":10}
            for k, v in inputs.items():
                df_conteo.iloc[idx, map_cols[k]] = v

            df_conteo.iloc[idx, 11] = total
            st.success("Guardado")

    # -------------------------
    # TABLA EN VIVO
    # -------------------------
    st.subheader("📊 Conteo actual")
    st.dataframe(df_conteo)

    # -------------------------
    # EXPORTAR
    # -------------------------
    if st.button("EXPORTAR"):
        data = save_full_audit(df_conteo, df_sistema, wb)

        st.download_button(
            "Descargar Excel",
            data,
            file_name=f"INVENTARIO_{sucursal}_{fecha}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )