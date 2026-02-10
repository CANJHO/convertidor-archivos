import streamlit as st
import pandas as pd
import pdfplumber
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import io
import re

st.set_page_config(page_title="Convertidor Inteligente", layout="wide")

st.title("Convertidor de Archivos Inteligente")
st.write("Convierte PDF, Excel, Word y CSV automáticamente")

# =====================================================
# FUNCION LIMPIAR ENCABEZADOS REPETIDOS
# =====================================================

def limpiar_dataframe(df):

    if df is None or df.empty:
        return df

    palabras_encabezado = [
        "ORDEN","CÓDIGO","CODIGO","CURSO","REQ",
        "H TEO","H PRA","CRÉD","CRED","TIP CUR",
        "TIP_CUR","ÁREA","AREA"
    ]

    filas_limpias = []

    for _, fila in df.iterrows():

        valores = [str(v).upper().strip() for v in fila.values]

        coincidencias = sum(
            1 for v in valores if v in palabras_encabezado
        )

        if coincidencias < 3:
            filas_limpias.append(fila)

    df_limpio = pd.DataFrame(filas_limpias)

    df_limpio.columns = df.columns
    df_limpio.reset_index(drop=True, inplace=True)

    return df_limpio


# =====================================================
# METODO 1 TABLAS REALES
# =====================================================

def extraer_tablas_pdf(file):

    filas = []

    with pdfplumber.open(file) as pdf:

        for pagina in pdf.pages:

            tablas = pagina.extract_tables()

            for tabla in tablas:

                for fila in tabla:

                    if fila and any(celda is not None for celda in fila):
                        filas.append(fila)

    if filas:

        df = pd.DataFrame(filas)

        df.columns = df.iloc[0]

        df = df[1:]

        df = limpiar_dataframe(df)

        return df

    return None


# =====================================================
# METODO 2 TEXTO
# =====================================================

def extraer_texto_pdf(file):

    filas = []

    texto = ""

    with pdfplumber.open(file) as pdf:

        for pagina in pdf.pages:

            contenido = pagina.extract_text()

            if contenido:
                texto += contenido + "\n"

    lineas = texto.split("\n")

    patron = re.compile(
        r'^(\d+)\s+'
        r'(P\d+A\d+)\s+'
        r'(.+?)\s+'
        r'(P\d+A\d+)?\s*'
        r'(\d+)\s+(\d+)\s+(\d+)\s+'
        r'([OE])\s+'
        r'(EC|EF|GE)'
    )

    for linea in lineas:

        match = patron.match(linea.strip())

        if match:
            filas.append(match.groups())

    if filas:

        columnas = [
            "ORDEN","CODIGO","CURSO","REQ",
            "H_TEO","H_PRA","CRED","TIP_CUR","AREA"
        ]

        df = pd.DataFrame(filas, columns=columnas)

        df = limpiar_dataframe(df)

        return df

    return None


# =====================================================
# FUNCION INTELIGENTE
# =====================================================

def procesar_pdf(file, nombre_archivo=""):

    df = extraer_tablas_pdf(file)

    if df is None or df.empty:

        df = extraer_texto_pdf(file)

    if df is None:
        return None

    df.insert(0, "ARCHIVO", nombre_archivo)

    return df


# =====================================================
# CONVERSION INDIVIDUAL
# =====================================================

st.header("Conversión Individual")

archivo = st.file_uploader(
    "Sube archivo",
    type=["pdf","xlsx","csv","docx"]
)

conversion = st.selectbox(
    "Convertir a",
    ["Excel (.xlsx)", "CSV (.csv)", "PDF (.pdf)"]
)


def convertir_individual(archivo):

    extension = archivo.name.split(".")[-1].lower()

    if extension == "pdf":

        df = procesar_pdf(archivo, archivo.name)

        if df is None:
            st.error("No se pudo procesar")
            return

        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            "Descargar Excel",
            output,
            "convertido.xlsx"
        )


# ejecutar individual
if archivo:
    convertir_individual(archivo)


# =====================================================
# CONVERSION MASIVA STREAMLIT CLOUD
# =====================================================

st.header("Conversión Masiva (Streamlit Cloud Compatible)")

archivos_masivos = st.file_uploader(
    "Sube múltiples PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

if st.button("Convertir TODOS"):

    if archivos_masivos:

        progreso = st.progress(0)

        total = len(archivos_masivos)

        dfs = []

        for i, archivo in enumerate(archivos_masivos):

            df = procesar_pdf(
                archivo,
                archivo.name
            )

            if df is not None:
                dfs.append(df)

            progreso.progress((i+1)/total)

        if dfs:

            final = pd.concat(dfs, ignore_index=True)

            output = io.BytesIO()

            final.to_excel(output, index=False)

            output.seek(0)

            st.success("Conversión completada")

            st.download_button(
                "Descargar Excel Consolidado",
                output,
                "TODOS_LOS_PLANES.xlsx"
            )

        else:
            st.error("No se pudo convertir")
