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
            filas_limpias.append(fila.tolist())  # ← importante

    # 🔥 SI SE ELIMINÓ TODO, DEVOLVER DF ORIGINAL
    if not filas_limpias:
        return df

    # 🔥 CREAR CON MISMAS COLUMNAS
    df_limpio = pd.DataFrame(filas_limpias, columns=df.columns)

    df_limpio.reset_index(drop=True, inplace=True)

    return df_limpio




# =====================================================
# CSV GRANDE → EXCEL (HASTA 1GB+)
# =====================================================

def csv_a_excel_grande(file):

    # copiar archivo en memoria (SOLUCION STREAMLIT CLOUD)
    contenido = file.read()

    buffer = io.BytesIO(contenido)

    output = io.BytesIO()

    writer = pd.ExcelWriter(
        output,
        engine="openpyxl"
    )

    fila_inicio = 0
    encabezado_escrito = False

    progreso = st.progress(0)

    chunk_iter = pd.read_csv(
        buffer,
        sep=None,
        engine="python",
        encoding_errors="ignore",
        chunksize=50000
    )

    total_chunks = 0

    buffer.seek(0)

    for _ in pd.read_csv(
        buffer,
        sep=None,
        engine="python",
        encoding_errors="ignore",
        chunksize=50000
    ):
        total_chunks += 1

    buffer.seek(0)

    chunk_iter = pd.read_csv(
        buffer,
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
# EXTRAER TABLAS PDF
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

    if not filas:
        return None

    df = pd.DataFrame(filas)

    # 🔥 Validar que haya al menos 2 filas
    if len(df) < 2:
        return None

    # Intentar usar primera fila como encabezado
    posible_header = df.iloc[0]

    # Validar que realmente parezca encabezado
    if any(str(x).upper() in ["ORDEN","CÓDIGO","CODIGO","CURSO"] for x in posible_header):
        df.columns = posible_header
        df = df[1:]
    else:
        # si no parece encabezado, no lo usamos
        df.columns = [f"COL_{i}" for i in range(len(df.columns))]

    df = df.reset_index(drop=True)

    df = limpiar_dataframe(df)

    if df.empty:
        return None

    return df


    filas = []

    with pdfplumber.open(file) as pdf:

        for pagina in pdf.pages:

            tablas = pagina.extract_tables()

            for tabla in tablas:

                for fila in tabla:

                    if fila and any(celda is not None for celda in fila):
                        filas.append(fila)

    if not filas:
        return None

    df = pd.DataFrame(filas)

    # usar primera fila como encabezado
    df.columns = df.iloc[0]
    df = df[1:]

    df = df.reset_index(drop=True)

    # 🔥 ELIMINAR FILAS QUE SEAN SOLO ENCABEZADO REPETIDO
    df = df[
        ~(
            df.iloc[:,0].astype(str).str.upper().str.contains("ORDEN|CÓDIGO|CODIGO")
        )
    ]

    df = df[
        ~(
            df.iloc[:,1].astype(str).str.upper().str.contains("CURSO")
        )
    ]

    df = limpiar_dataframe(df)

    if df.empty:
        return None

    return df


# =====================================================
# EXTRAER TEXTO PDF
# =====================================================

def extraer_texto_pdf(file):
    """
    Extrae Planes AKADEMIC tipo:
      P09-20242-
      MATEMÁTICA I 00 3.0 2.0 5.0 4.0 Ningun Requisito
      P09A1101

    Devuelve DF con:
      PLAN, SEMESTRE, CODIGO (completo concatenado), CURSO, HT, HP, TH, CRED, REQ
    """
    import re
    import pandas as pd
    import pdfplumber

    def norm_ws(s: str) -> str:
        return " ".join((s or "").split())

    def is_prefix_line(line: str) -> bool:
        s = (line or "").strip().replace(" ", "")
        return re.fullmatch(r"P\d{2}-\d+-", s) is not None

    def is_code_line(line: str) -> bool:
        # Ej: P09A1101 o 11001 o P09A2155
        s = norm_ws(line)
        return re.fullmatch(r"[A-Z0-9]{4,15}", s) is not None

    def parse_course_line_without_code(line: str):
        """
        MATEMÁTICA I 00 3.0 2.0 5.0 4.0 Ningun Requisito
        """
        s = norm_ws(line)
        toks = s.split()
        if len(toks) < 6:
            return None
        try:
            esp_idx = toks.index("00")   # columna Esp
        except ValueError:
            return None

        if esp_idx < 1:
            return None

        curso = " ".join(toks[:esp_idx]).strip()
        after = toks[esp_idx + 1:]
        if len(after) < 4:
            return None

        ht, hp, th, cred = after[0], after[1], after[2], after[3]
        req = " ".join(after[4:]).strip() if len(after) > 4 else ""
        return curso, ht, hp, th, cred, req

    # --- leer líneas de todo el PDF
    file.seek(0)
    lines = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            lines.extend(txt.split("\n"))

    # --- parseo
    rows = []
    semestre = ""

    i = 0
    while i < len(lines):
        line = (lines[i] or "").strip()
        if not line:
            i += 1
            continue

        # Semestre
        if line.upper().startswith("SEMESTRE"):
            semestre = line.split(":")[-1].strip().upper()
            i += 1
            continue

        # Prefijo de código (P09-20242-)
        if is_prefix_line(line):
            prefix = line.strip().replace(" ", "")

            # Patrón esperado: prefix, curso_line, code_line
            if i + 2 < len(lines):
                curso_line = (lines[i + 1] or "").strip()
                code_line = (lines[i + 2] or "").strip()

                parsed = parse_course_line_without_code(curso_line)
                if parsed and is_code_line(code_line):
                    curso, ht, hp, th, cred, req = parsed
                    codigo_completo = prefix + norm_ws(code_line)

                    # PLAN lo saco del prefix: P09-20242-
                    plan = prefix.split("-")[1] if "-" in prefix else ""

                    rows.append([
                        plan, semestre, codigo_completo, curso,
                        ht, hp, th, cred, req
                    ])

                    i += 3
                    continue

            i += 1
            continue

        i += 1

    if not rows:
        return None

    df = pd.DataFrame(rows, columns=[
        "PLAN", "SEMESTRE", "CODIGO", "CURSO",
        "HT", "HP", "TH", "CRED", "REQ"
    ])

    df.reset_index(drop=True, inplace=True)
    return df


# =====================================================
# PROCESAR PDF
# =====================================================
def procesar_pdf(file, nombre_archivo=""):
    file.seek(0)

    df = extraer_texto_pdf(file)

    if df is None or df.empty:
        return None

    # Mantener tu columna ARCHIVO
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

    output = io.BytesIO()

    if extension == "pdf" and conversion == "Excel (.xlsx)":

        df = procesar_pdf(archivo, archivo.name)

        if df is None:
            st.error("No se pudo procesar el PDF")
            return

        df.to_excel(output, index=False)

        nombre = "convertido.xlsx"


    elif extension == "csv" and conversion == "Excel (.xlsx)":

        try:

            output = csv_a_excel_grande(archivo)

            nombre = "convertido.xlsx"

        except Exception as e:

            st.error(str(e))
            return


    elif extension == "xlsx" and conversion == "CSV (.csv)":

        df = pd.read_excel(archivo)

        df.to_csv(output, index=False)

        nombre = "convertido.csv"


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


if archivo:

    if st.button("Convertir archivo"):

        convertir_individual(archivo)


# =====================================================
# CONVERSION MASIVA PDF
# =====================================================

st.header("Conversión Masiva")

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

        for i, archivo_pdf in enumerate(archivos_masivos):

            archivo_pdf.seek(0)  # 🔥 reset obligatorio

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
