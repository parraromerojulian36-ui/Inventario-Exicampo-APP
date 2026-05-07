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
# CONFIG
# ==========================================
st.set_page_config(
    page_title="PASCA Inventory Audit Pro",
    layout="wide"
)

# ==========================================
# CSS
# ==========================================
st.markdown("""
<style>

.block-container{
    padding-top: 1rem;
}

.product-card{
    background: white;
    padding: 18px;
    border-radius: 16px;
    border-left: 8px solid #2E7D32;
    margin-bottom: 15px;
    box-shadow: 0px 2px 10px rgba(0,0,0,0.08);
}

.product-card h3{
    margin:0;
    color:#2E7D32;
    font-size:22px;
}

.big-total{
    background:#2E7D32;
    color:white;
    text-align:center;
    padding:18px;
    border-radius:16px;
    font-size:34px;
    font-weight:bold;
    margin-top:20px;
}

.stButton button{
    width:100%;
    border-radius:12px;
    height:50px;
    font-weight:bold;
}

@media (max-width: 768px){

    .product-card{
        padding:14px;
    }

    .product-card h3{
        font-size:18px;
    }

    .big-total{
        font-size:26px;
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


# ==========================================
# OCR
# ==========================================
def detect_text_ocr(image):
    try:
        text = pytesseract.image_to_string(image, lang="eng")

        text = text.upper()
        text = text.replace("\n", " ")
        text = text.replace("  ", " ")

        return text.strip()

    except Exception as e:
        return f"ERROR OCR: {str(e)}"


# ==========================================
# BUSCADOR INTELIGENTE
# ==========================================
def search_product(df_sistema, detected_text):

    detected_text = detected_text.upper()

    # EXACTO
    exact = df_sistema[
        (
            df_sistema.iloc[:, 0]
            .astype(str)
            .str.upper()
            .str.contains(detected_text, na=False)
        )
        |
        (
            df_sistema.iloc[:, 1]
            .astype(str)
            .str.upper()
            .str.contains(detected_text, na=False)
        )
    ]

    if not exact.empty:
        return exact

    # SIMILAR
    nombres = df_sistema.iloc[:, 1].astype(str).tolist()

    similars = difflib.get_close_matches(
        detected_text,
        nombres,
        n=10,
        cutoff=0.2
    )

    if similars:
        return df_sistema[
            df_sistema.iloc[:, 1].astype(str).isin(similars)
        ]

    return pd.DataFrame()


# ==========================================
# CARGAR EXCEL
# ==========================================
def load_excel(uploaded_file):

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:

        tmp.write(uploaded_file.getvalue())

        temp_path = tmp.name

    wb = openpyxl.load_workbook(temp_path)

    # SISTEMA
    df_sistema = pd.read_excel(
        temp_path,
        sheet_name="SISTEMA"
    )

    df_sistema.columns = df_sistema.columns.str.strip()

    df_sistema.iloc[:, 0] = (
        df_sistema.iloc[:, 0]
        .apply(clean_code)
    )

    # CONTEO
    df_conteo = pd.read_excel(
        temp_path,
        sheet_name="CONTEO_F"
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

    df_conteo.iloc[:, 0] = (
        df_conteo.iloc[:, 0]
        .apply(clean_code)
    )

    st.session_state.temp_file = temp_path

    return df_conteo, df_sistema, wb


# ==========================================
# EXPORTAR
# ==========================================
def save_full_audit(df_conteo, df_sistema, wb):

    # ======================================
    # CONTEO_F
    # ======================================
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

            sheet.cell(
                row=row_num,
                column=col_num
            ).value = value

    # ======================================
    # RESULTADO
    # ======================================
    result_sheet = wb["RESULTADO"]

    for row in result_sheet.iter_rows(min_row=5):

        for cell in row:

            cell.value = None

    row_res = 5

    for _, row_c in df_conteo.iterrows():

        code = clean_code(row_c.iloc[0])

        name = row_c.iloc[1]

        total_fisico = row_c.iloc[11]

        if pd.isna(total_fisico):
            total_fisico = 0

        match = df_sistema[
            df_sistema.iloc[:, 0]
            .astype(str)
            == code
        ]

        if not match.empty:

            total_sistema = match.iloc[0, 2]

            if pd.isna(total_sistema):
                total_sistema = 0

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

    # ======================================
    # GUARDAR
    # ======================================
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

    st.header("Configuración")

    sucursal = st.selectbox(
        "Sucursal",
        ["PASCA", "SUBIA", "SIBATE", "GRANADA"]
    )

    fecha = datetime.now().strftime("%d-%m-%Y")

uploaded_file = st.file_uploader(
    "Sube Excel",
    type=["xlsx"]
)

# ==========================================
# APP
# ==========================================
if uploaded_file:

    # CARGAR
    if "df_inv" not in st.session_state:

        df_c, df_s, wb = load_excel(uploaded_file)

        st.session_state.df_inv = df_c
        st.session_state.df_sistema = df_s
        st.session_state.wb = wb

    df_conteo = st.session_state.df_inv
    df_sistema = st.session_state.df_sistema
    wb = st.session_state.wb

    # ======================================
    # OCR
    # ======================================
    st.subheader("📷 Cámara OCR")

    img_file = st.camera_input(
        "Tomar foto del producto"
    )

    if img_file:

        image = Image.open(img_file)

        with st.spinner("Leyendo etiqueta..."):

            detected_text = detect_text_ocr(image)

        st.success(f"OCR Detectó: {detected_text}")

        results = search_product(
            df_sistema,
            detected_text
        )

        if results.empty:

            st.error("No se encontraron productos.")

        else:

            st.write("### Productos encontrados")

            for idx in results.index:

                product_name = str(results.loc[idx].iloc[1])

                product_code = clean_code(
                    results.loc[idx].iloc[0]
                )

                stock = results.loc[idx].iloc[2]

                safe_name = html.escape(product_name)
                safe_code = html.escape(product_code)
                safe_stock = html.escape(str(stock))
                safe_detected = html.escape(str(detected_text))

                st.markdown(f"""
                <div class="product-card">

                    <h3>{safe_name}</h3>

                    <p><b>Código:</b> {safe_code}</p>

                    <p><b>Stock Sistema:</b> {safe_stock}</p>

                    <p><b>OCR Detectó:</b> {safe_detected}</p>

                </div>
                """, unsafe_allow_html=True)

                if st.button(
                    f"Seleccionar {product_code}",
                    key=f"ocr_{product_code}"
                ):

                    st.session_state.selected_code = product_code
                    st.session_state.selected_name = product_name

                    st.rerun()

    # ======================================
    # BUSQUEDA MANUAL
    # ======================================
    st.subheader("🔍 Buscar manualmente")

    search = st.text_input(
        "Código o nombre"
    ).upper()

    if search:

        mask = (
            (
                df_sistema.iloc[:, 0]
                .astype(str)
                .str.contains(search, na=False)
            )
            |
            (
                df_sistema.iloc[:, 1]
                .astype(str)
                .str.upper()
                .str.contains(search, na=False)
            )
        )

        results = df_sistema[mask]

        for idx in results.index:

            name = results.loc[idx].iloc[1]

            code = clean_code(
                results.loc[idx].iloc[0]
            )

            stock = results.loc[idx].iloc[2]

            st.markdown(f"""
            <div class="product-card">

            <h3>{name}</h3>

            <p><b>Código:</b> {code}</p>

            <p><b>Stock Sistema:</b> {stock}</p>

            </div>
            """, unsafe_allow_html=True)

            if st.button(
                f"Seleccionar {code}",
                key=f"manual_{code}"
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
            df_sistema.iloc[:, 0]
            .astype(str)
            == code
        ]

        stock = (
            match.iloc[0, 2]
            if not match.empty
            else 0
        )

        idxs = df_conteo[
            df_conteo.iloc[:, 0]
            .astype(str)
            == code
        ].index

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

        bodegas = [
            "BO1",
            "BO2",
            "BO3",
            "AL1",
            "AL2",
            "AL3",
            "VALES",
            "VENCIDOS"
        ]

        values = []

        for i in range(3, 11):

            try:

                val = df_conteo.iloc[idx, i]

                if pd.isna(val):
                    val = 0

                values.append(int(val))

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

        st.markdown(f"""
        <div class="big-total">
        TOTAL: {total}
        </div>
        """, unsafe_allow_html=True)

        if st.button(
            "💾 GUARDAR",
            type="primary"
        ):

            mapa = {
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

                df_conteo.iloc[idx, mapa[k]] = v

            df_conteo.iloc[idx, 11] = total

            st.success("Producto guardado")

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

        filename = f"INVENTARIO_{sucursal}_{fecha}.xlsx"

        st.download_button(
            "Descargar Excel",
            data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )