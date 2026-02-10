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

# LECTOR CSV INTELIGENTE NORMAL

# =====================================================

def leer_csv_seguro(file):


    codificaciones = ["utf-8", "latin1", "cp1252"]

for enc in codificaciones:

    try:

        file.seek(0)

        df = pd.read_csv(
            file,
            encoding=enc,
            sep=None,
            engine="python"
        )

        if df.shape[1] >= 1:
            return df

    except:
        continue


try:

    file.seek(0)
    df = pd.read_excel(file)
    return df

except:
    pass


raise Exception("No se pudo leer el archivo CSV")


    # =====================================================

    # CSV → EXCEL PARA ARCHIVOS MUY GRANDES (1GB+)

    # =====================================================

def csv_a_excel_grande(file):


    output = io.BytesIO()

    writer = pd.ExcelWriter(
        output,
        engine="openpyxl"
    )

file.seek(0)

chunk_iter = pd.read_csv(
    file,
    sep=None,
    engine="python",
    encoding_errors="ignore",
    chunksize=50000
)

fila_inicio = 0

encabezado_escrito = False

progreso = st.progress(0)
total_chunks = 0

# contar chunks aproximados
file.seek(0)
for _ in pd.read_csv(file, sep=None, engine="python", encoding_errors="ignore", chunksize=50000):
    total_chunks += 1

file.seek(0)

chunk_iter = pd.read_csv(
    file,
    sep=None,
    engine="python",
    encoding_errors="ignore",
    chunksize=50000
)

for i, chunk in enumerate(chunk_iter):

    chunk = limpiar_dataframe(chunk)

    chunk.to_excel(
        writer,
        index=False,
        startrow=fila_inicio,
        header=not encabezado_escrito
    )

    encabezado_escrito = True

    fila_inicio += len(chunk)

    progreso.progress((i+1)/total_chunks)

writer.close()

output.seek(0)

return output


# =====================================================

# METODO 1 TABLAS REALES PDF

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

# METODO 2 TEXTO PDF

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

# FUNCION INTELIGENTE PDF

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

```
extension = archivo.name.split(".")[-1].lower()

output = io.BytesIO()


# PDF → EXCEL
if extension == "pdf" and conversion == "Excel (.xlsx)":

    df = procesar_pdf(archivo, archivo.name)

    if df is None:
        st.error("No se pudo procesar el PDF")
        return

    df.to_excel(output, index=False)

    nombre = "convertido.xlsx"


# CSV → EXCEL (OPTIMIZADO PARA ARCHIVOS GIGANTES)
elif extension == "csv" and conversion == "Excel (.xlsx)":

    try:

        output = csv_a_excel_grande(archivo)

        nombre = "convertido.xlsx"

    except Exception as e:

        st.error(f"Error procesando archivo grande: {str(e)}")
        return


# EXCEL → CSV
elif extension == "xlsx" and conversion == "CSV (.csv)":

    df = pd.read_excel(archivo)

    df.to_csv(output, index=False)

    nombre = "convertido.csv"


# WORD → PDF
elif extension == "docx" and conversion == "PDF (.pdf)":

    doc = Document(archivo)

    c = canvas.Canvas(output, pagesize=letter)

    y = 750

    for para in doc.paragraphs:

        c.drawString(30, y, para.text)

        y -= 20

        if y < 50:
            c.showPage()
            y = 750

    c.save()

    nombre = "convertido.pdf"


else:

    st.warning("Conversión no soportada")
    return


output.seek(0)

st.success("Archivo convertido correctamente")

st.download_button(
    "Descargar archivo",
    output,
    nombre
)
```

if archivo:

```
if st.button("Convertir archivo"):

    convertir_individual(archivo)
```

# =====================================================

# CONVERSION MASIVA

# =====================================================

st.header("Conversión Masiva (Streamlit Cloud Compatible)")

archivos_masivos = st.file_uploader(
"Sube múltiples PDFs",
type=["pdf"],
accept_multiple_files=True
)

if st.button("Convertir TODOS"):

```
if archivos_masivos:

    progreso = st.progress(0)

    total = len(archivos_masivos)

    dfs = []

    for i, archivo_pdf in enumerate(archivos_masivos):

        df = procesar_pdf(
            archivo_pdf,
            archivo_pdf.name
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
        
```
