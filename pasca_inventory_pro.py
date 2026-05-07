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
# UTILIDADES (asegúrate de tenerlas definidas)
# ==========================================
def clean_code(x):
    return str(x).strip()

def detect_text_ocr(image):
    return pytesseract.image_to_string(image)

def search_product(df, text):
    text = str(text).upper()
    mask = (
        df.iloc[:, 0].astype(str).str.upper().str.contains(text, na=False) |
        df.iloc[:, 1].astype(str).str.upper().str.contains(text, na=False)
    )
    return df[mask]

# ==========================================
# CARGA EXCEL
# ==========================================
def load_excel(uploaded_file):
    wb = openpyxl.load_workbook(uploaded_file)
    df_conteo = pd.read_excel(uploaded_file, sheet_name=0)
    df_sistema = pd.read_excel(uploaded_file, sheet_name=1)

    header_row = 0
    for i, row in df_conteo.iterrows():
        if "CODIGO" in str(row.values).upper():
            header_row = i
            break

    df_conteo.columns = df_conteo.iloc[header_row].astype(str).str.strip()
    df_conteo = df_conteo.iloc[header_row + 1:].reset_index(drop=True)

    df_conteo = df_conteo.astype(object)
    df_conteo.iloc[:, 0] = df_conteo.iloc[:, 0].apply(clean_code)

    return df_conteo, df_sistema, wb

# ==========================================
# EXPORTAR
# ==========================================
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
        for cell in row:
            cell.value = None

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
# UI
# ==========================================
st.title("📦 Control Inventarios Exicampo")

with st.sidebar:
    st.header("Configuración")
    sucursal = st.selectbox("Sucursal", ["PASCA", "SUBIA", "SIBATE", "GRANADA"])
    fecha = datetime.now().strftime("%d-%m-%Y")

uploaded_file = st.file_uploader("Sube Excel", type=["xlsx"])

if uploaded_file:

    if "df_inv" not in st.session_state:
        df_c, df_s, wb = load_excel(uploaded_file)
        st.session_state.df_inv = df_c
        st.session_state.df_sistema = df_s
        st.session_state.wb = wb

    df_conteo = st.session_state.df_inv
    df_sistema = st.session_state.df_sistema
    wb = st.session_state.wb

    # ==========================
    # VISTA EDICIÓN
    # ==========================
    if "selected_code" in st.session_state:

        code = st.session_state.selected_code
        name = st.session_state.selected_name

        if st.button("⬅️ Volver"):
            del st.session_state.selected_code
            del st.session_state.selected_name
            st.rerun()

        match = df_sistema[df_sistema.iloc[:, 0].astype(str) == code]
        stock = match.iloc[0, 2] if not match.empty else 0

        idxs = df_conteo[df_conteo.iloc[:, 0].astype(str) == code].index

        if idxs.empty:
            new_row = [0] * len(df_conteo.columns)
            new_row[0] = code
            new_row[1] = name

            df_conteo.loc[len(df_conteo)] = new_row
            idx = len(df_conteo) - 1
        else:
            idx = idxs[0]

        st.markdown(f"""
        <div class="product-card">
            <h3>{name}</h3>
            <p><b>Código:</b> {code}</p>
            <p><b>Stock Sistema:</b> {stock}</p>
        </div>
        """, unsafe_allow_html=True)

        bodegas = ["BO1", "BO2", "BO3", "AL1", "AL2", "AL3", "VALES", "VENCIDOS"]

        values = []
        for i in range(3, 11):
            try:
                val = df_conteo.iloc[idx, i]
                values.append(int(val) if not pd.isna(val) else 0)
            except:
                values.append(0)

        inputs = {}
        cols = st.columns(2)

        for i, bodega in enumerate(bodegas):
            with cols[i % 2]:
                inputs[bodega] = st.number_input(
                    bodega,
                    min_value=0,
                    value=values[i],
                    key=f"{code}_{bodega}"
                )

        total = sum(inputs.values())
        st.markdown(f"### TOTAL: {total}")

        if st.button("💾 GUARDAR", type="primary"):
            mapa = {"BO1":3, "BO2":4, "BO3":5, "AL1":6, "AL2":7, "AL3":8, "VALES":9, "VENCIDOS":10}

            for k, v in inputs.items():
                df_conteo.iloc[idx, mapa[k]] = v

            df_conteo.iloc[idx, 11] = total
            st.success("Guardado correctamente")

    # ==========================
    # VISTA BÚSQUEDA
    # ==========================
    else:

        st.subheader("📷 OCR Cámara")
        img_file = st.camera_input("Tomar foto")

        if img_file:
            image = Image.open(img_file)

            with st.spinner("Leyendo..."):
                detected = detect_text_ocr(image)

            st.info(f"OCR: {detected}")

            results = search_product(df_sistema, detected)

            if results.empty:
                st.error("Sin coincidencias")
            else:
                for idx in results.index:
                    name_p = str(results.loc[idx].iloc[1])
                    code_p = clean_code(results.loc[idx].iloc[0])
                    stock_p = results.loc[idx].iloc[2]

                    st.markdown(f"""
                    <div class="product-card">
                        <h3>{html.escape(name_p)}</h3>
                        <p><b>Código:</b> {code_p}</p>
                        <p><b>Stock:</b> {stock_p}</p>
                    </div>
                    """, unsafe_allow_html=True)

                    if st.button(f"Seleccionar {code_p}", key=f"ocr_{code_p}"):
                        st.session_state.selected_code = code_p
                        st.session_state.selected_name = name_p
                        st.rerun()

        st.divider()
        st.subheader("🔍 Buscar manual")

        search = st.text_input("Código o nombre").upper()

        if search and len(search) >= 2:
            mask = (
                df_sistema.iloc[:, 0].astype(str).str.upper().str.startswith(search, na=False) |
                df_sistema.iloc[:, 1].astype(str).str.upper().str.startswith(search, na=False)
            )

            results = df_sistema[mask].head(5)

            for idx in results.index:
                name = results.loc[idx].iloc[1]
                code = clean_code(results.loc[idx].iloc[0])
                stock = results.loc[idx].iloc[2]

                st.markdown(f"""
                <div class="product-card">
                    <h3>{name}</h3>
                    <p><b>Código:</b> {code}</p>
                    <p><b>Stock:</b> {stock}</p>
                </div>
                """, unsafe_allow_html=True)

                if st.button(f"Seleccionar {code}", key=f"manual_{code}"):
                    st.session_state.selected_code = code
                    st.session_state.selected_name = name
                    st.rerun()

    # ==========================
    # EXPORTAR
    # ==========================
    st.divider()

    if st.button("📥 EXPORTAR EXCEL FINAL"):
        data = save_full_audit(df_conteo, df_sistema, wb)

        filename = f"INVENTARIO_{sucursal}_{fecha}.xlsx"

        st.download_button(
            "Descargar Excel",
            data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )