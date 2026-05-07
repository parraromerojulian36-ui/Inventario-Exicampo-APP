import streamlit as st
import pandas as pd
import openpyxl
import os
import tempfile
from datetime import datetime
from PIL import Image
import pytesseract
import cv2
import numpy as np
from difflib import get_close_matches

# ==========================================
# CONFIG
# ==========================================
st.set_page_config(
    page_title="PASCA Inventory Audit Pro",
    layout="wide"
)

# ==========================================
# CSS RESPONSIVE
# ==========================================
st.markdown("""
<style>

html, body, [class*="css"]  {
    font-size: 16px;
}

.big-font {
    font-size: 24px !important;
    font-weight: bold !important;
    text-align: center !important;
    color: white !important;
    background-color: #2E7D32 !important;
    padding: 15px !important;
    border-radius: 12px !important;
    margin-top: 20px !important;
}

.product-card {
    background: white;
    padding: 18px;
    border-radius: 15px;
    border-left: 8px solid #4CAF50;
    margin-top: 15px;
    margin-bottom: 20px;
    box-shadow: 0px 2px 8px rgba(0,0,0,0.1);
}

.stButton button {
    width: 100%;
    border-radius: 10px;
    font-weight: bold;
}

@media (max-width: 768px) {

    .big-font {
        font-size: 20px !important;
    }

    h1 {
        font-size: 28px !important;
    }

}

</style>
""", unsafe_allow_html=True)

# ==========================================
# LIMPIAR CODIGO
# ==========================================
def clean_code(val):

    if pd.isna(val):
        return ""

    val = str(val).strip()

    if val.endswith(".0"):
        val = val[:-2]

    return val

# ==========================================
# OCR + MATCH INTELIGENTE
# ==========================================
def identify_product_ocr(image, df_sistema):

    try:

        img = np.array(image)
        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

        # mejorar calidad
        img = cv2.resize(img, None, fx=3, fy=3)

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        gray = cv2.GaussianBlur(gray, (3,3), 0)

        _, thresh = cv2.threshold(
            gray,
            0,
            255,
            cv2.THRESH_BINARY + cv2.THRESH_OTSU
        )

        text = pytesseract.image_to_string(
            thresh,
            config='--psm 6'
        )

        text = text.upper().strip()

        if not text:
            return None

        # ==================================
        # BUSQUEDA INTELIGENTE
        # ==================================
        nombres = (
            df_sistema.iloc[:,1]
            .astype(str)
            .str.upper()
            .tolist()
        )

        coincidencia = get_close_matches(
            text,
            nombres,
            n=1,
            cutoff=0.3
        )

        if coincidencia:

            producto = coincidencia[0]

            fila = df_sistema[
                df_sistema.iloc[:,1]
                .astype(str)
                .str.upper() == producto
            ]

            if not fila.empty:

                codigo = clean_code(fila.iloc[0,0])
                stock = fila.iloc[0,2]

                return {
                    "texto": text,
                    "producto": producto,
                    "codigo": codigo,
                    "stock": stock
                }

        return {
            "texto": text,
            "producto": None
        }

    except Exception as e:

        return {
            "error": str(e)
        }

# ==========================================
# CARGAR EXCEL
# ==========================================
def load_pasca_data(uploaded_file):

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:

        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name

    wb = openpyxl.load_workbook(tmp_path)

    # SISTEMA
    df_sistema = pd.read_excel(
        tmp_path,
        sheet_name='SISTEMA'
    )

    df_sistema.columns = df_sistema.columns.str.strip()

    df_sistema.iloc[:,0] = (
        df_sistema.iloc[:,0]
        .apply(clean_code)
    )

    # CONTEO
    df_conteo = pd.read_excel(
        tmp_path,
        sheet_name='CONTEO_F'
    )

    header_row = 0

    for i, row in df_conteo.iterrows():

        if "CODIGO" in str(row.values).upper():

            header_row = i
            break

    df_conteo.columns = (
        df_conteo.iloc[header_row]
        .astype(str)
        .str.strip()
    )

    df_conteo = (
        df_conteo.iloc[header_row + 1:]
        .reset_index(drop=True)
    )

    df_conteo = df_conteo.astype(object)

    df_conteo.iloc[:,0] = (
        df_conteo.iloc[:,0]
        .apply(clean_code)
    )

    st.session_state.temp_file = tmp_path

    return df_conteo, df_sistema, wb

# ==========================================
# GUARDAR EXCEL
# ==========================================
def save_full_audit(df_conteo, df_sistema, wb):

    sheet = wb['CONTEO_F']

    start_row = 1

    for row in sheet.iter_rows(max_row=10):

        for cell in row:

            if cell.value and "CODIGO" in str(cell.value).upper():

                start_row = cell.row + 1
                break

    # guardar conteo
    for i, row in df_conteo.iterrows():

        row_num = start_row + i

        for col_num, value in enumerate(row.values, 1):

            sheet.cell(
                row=row_num,
                column=col_num
            ).value = value

    # RESULTADO
    sheet_res = wb['RESULTADO']

    for row in sheet_res.iter_rows(min_row=5):

        for cell in row:

            cell.value = None

    row_res = 5

    for _, row_c in df_conteo.iterrows():

        code = clean_code(row_c.iloc[0])
        name = row_c.iloc[1]

        total_fisico = (
            row_c.iloc[11]
            if pd.notnull(row_c.iloc[11])
            else 0
        )

        match = df_sistema[
            df_sistema.iloc[:,0].astype(str) == code
        ]

        if not match.empty:

            total_sistema = (
                match.iloc[0,2]
                if pd.notnull(match.iloc[0,2])
                else 0
            )

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

    return data

# ==========================================
# UI
# ==========================================
st.title("📦 PASCA Inventory Audit Pro")

with st.sidebar:

    sucursal = st.selectbox(
        "Sucursal",
        ["PASCA","SUBIA","SIBATE","GRANADA"]
    )

    fecha = datetime.now().strftime("%d-%m-%Y")

uploaded_file = st.file_uploader(
    "Sube el Excel",
    type=["xlsx"]
)

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

    # ======================================
    # CAMARA
    # ======================================
    st.subheader("📷 Escanear Producto")

    img_file = st.camera_input(
        "Tomar foto"
    )

    if img_file:

        img = Image.open(img_file)

        with st.spinner("Analizando producto..."):

            detected = identify_product_ocr(
                img,
                df_sistema
            )

        if not detected:

            st.error("No se pudo detectar.")

        elif "error" in detected:

            st.error(detected["error"])

        elif detected["producto"]:

            st.success("✅ Producto encontrado")

            st.markdown(f"""
            <div class="product-card">

            <h3>{detected['producto']}</h3>

            <b>Código:</b> {detected['codigo']}<br>
            <b>Stock Sistema:</b> {detected['stock']}<br>
            <b>OCR Detectó:</b> {detected['texto']}

            </div>
            """, unsafe_allow_html=True)

            st.session_state.selected_code = detected["codigo"]
            st.session_state.selected_name = detected["producto"]

        else:

            st.warning(
                f"OCR detectó: {detected['texto']}"
            )

    # ======================================
    # BUSQUEDA MANUAL
    # ======================================
    st.subheader("🔍 Buscar Manual")

    search = st.text_input(
        "Código o nombre"
    ).upper()

    if search:

        mask = (
            (df_sistema.iloc[:,0].astype(str) == search)
            |
            (
                df_sistema.iloc[:,1]
                .astype(str)
                .str.contains(search, case=False)
            )
        )

        resultados = df_sistema[mask]

        for idx in resultados.index:

            name = resultados.iloc[
                resultados.index.get_loc(idx),
                1
            ]

            code = clean_code(
                resultados.iloc[
                    resultados.index.get_loc(idx),
                    0
                ]
            )

            if st.button(
                f"{name} ({code})",
                key=f"bus_{code}"
            ):

                st.session_state.selected_code = code
                st.session_state.selected_name = name

                st.rerun()

    # ======================================
    # EDITOR
    # ======================================
    if "selected_code" in st.session_state:

        code = st.session_state.selected_code
        name = st.session_state.selected_name

        match = df_sistema[
            df_sistema.iloc[:,0].astype(str) == code
        ]

        stock = (
            match.iloc[0,2]
            if not match.empty
            else 0
        )

        idxs = df_conteo[
            df_conteo.iloc[:,0].astype(str) == code
        ].index

        if idxs.empty:

            nueva_fila = [0] * len(df_conteo.columns)

            nueva_fila[0] = code
            nueva_fila[1] = name

            df_conteo.loc[len(df_conteo)] = nueva_fila

            idx = len(df_conteo) - 1

        else:

            idx = idxs[0]

        st.markdown(f"""
        <div class="product-card">

        <h3>{name}</h3>

        <b>Código:</b> {code}<br>
        <b>Stock Sistema:</b> {stock}

        </div>
        """, unsafe_allow_html=True)

        cols_names = [
            "BO1",
            "BO2",
            "BO3",
            "AL1",
            "AL2",
            "AL3",
            "VALES",
            "VENCIDOS"
        ]

        valores = (
            df_conteo.iloc[idx, 3:11]
            .fillna(0)
            .astype(int)
            .tolist()
        )

        inputs = {}

        row1 = st.columns(2)
        row2 = st.columns(2)
        row3 = st.columns(2)
        row4 = st.columns(2)

        filas = [row1, row2, row3, row4]

        for i, col in enumerate(cols_names):

            fila_actual = filas[i // 2]

            with fila_actual[i % 2]:

                inputs[col] = st.number_input(
                    col,
                    min_value=0,
                    value=int(valores[i])
                )

        total = sum(inputs.values())

        st.markdown(f"""
        <div class="big-font">
        TOTAL: {total}
        </div>
        """, unsafe_allow_html=True)

        if st.button(
            "💾 GUARDAR",
            type="primary"
        ):

            map_cols = {
                "BO1":3,
                "BO2":4,
                "BO3":5,
                "AL1":6,
                "AL2":7,
                "AL3":8,
                "VALES":9,
                "VENCIDOS":10
            }

            for k, v in inputs.items():

                df_conteo.iloc[idx, map_cols[k]] = v

            df_conteo.iloc[idx,11] = total

            st.success("Guardado correctamente")

    # ======================================
    # EXPORTAR
    # ======================================
    st.divider()

    if st.button("📥 EXPORTAR EXCEL FINAL"):

        data = save_full_audit(
            df_conteo,
            df_sistema,
            wb
        )

        filename = (
            f"INVENTARIO_{sucursal}_{fecha}.xlsx"
        )

        st.download_button(
            "Descargar Excel",
            data=data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )