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
# LIMPIAR ENCABEZADOS REPETIDOS
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
            filas_limpias.append(fila.tolist())

    if not filas_limpias:
        return df

    df_limpio = pd.DataFrame(filas_limpias, columns=df.columns)
    df_limpio.reset_index(drop=True, inplace=True)

    return df_limpio


# =====================================================
# CSV GRANDE → EXCEL (HASTA 1GB+)
# =====================================================
def csv_a_excel_grande(file):

    contenido = file.read()
    buffer = io.BytesIO(contenido)

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")

    fila_inicio = 0
    encabezado_escrito = False

    progreso = st.progress(0)

    # contar chunks
    total_chunks = 0
    buffer.seek(0)
    for _ in pd.read_csv(
        buffer, sep=None, engine="python",
        encoding_errors="ignore", chunksize=50000
    ):
        total_chunks += 1

    buffer.seek(0)
    chunk_iter = pd.read_csv(
        buffer, sep=None, engine="python",
        encoding_errors="ignore", chunksize=50000
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
        progreso.progress((i + 1) / max(total_chunks, 1))

    writer.close()
    output.seek(0)
    return output


# =====================================================
# EXTRAER TABLAS PDF (NO USADO PARA PLANES, PERO LO DEJAMOS)
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

    if len(df) < 2:
        return None

    posible_header = df.iloc[0]

    if any(str(x).upper() in ["ORDEN", "CÓDIGO", "CODIGO", "CURSO"] for x in posible_header):
        df.columns = posible_header
        df = df[1:]
    else:
        df.columns = [f"COL_{i}" for i in range(len(df.columns))]

    df = df.reset_index(drop=True)
    df = limpiar_dataframe(df)

    if df.empty:
        return None

    return df


# =====================================================
# EXTRAER TEXTO PDF (UNIVERSAL AKADEMIC: OB + AF/P01 REAL)
# =====================================================
def extraer_texto_pdf(file):
    """
    Extractor UNIVERSAL para Planes AKADEMIC UAI.

    Soporta:
    A) OB:
        P09-20242-
        MATEMÁTICA I 00 3.0 2.0 5.0 4.0 Ningun Requisito
        P09A1101

    B) AF/P01 REAL (como tus AF-14210 / AF-20201):
        P01-
        14210- MATEMÁTICA I 3.0 2.0 5.0 4.0 O NINGUNO
        10A01
    """
    import io
    import re
    import pandas as pd
    import pdfplumber

    def norm_ws(s: str) -> str:
        return " ".join((s or "").split())

    def is_prefix_ob(line: str) -> bool:
        s = (line or "").strip().replace(" ", "")
        return re.fullmatch(r"P\d{2}-\d+-", s) is not None

    def is_prefix_p01(line: str) -> bool:
        s = (line or "").strip().replace(" ", "")
        return re.fullmatch(r"P\d{2}-", s) is not None

    def is_digits_prefix_token(tok: str) -> bool:
        return re.fullmatch(r"\d+-", tok or "") is not None

    def is_code_line(line: str) -> bool:
        s = norm_ws(line)
        return re.fullmatch(r"[A-Z0-9]{3,15}", s) is not None

    def parse_ob_course_line(line: str):
        # MATEMÁTICA I 00 3.0 2.0 5.0 4.0 Ningun Requisito
        s = norm_ws(line)
        toks = s.split()
        if len(toks) < 6:
            return None
        try:
            esp_idx = toks.index("00")
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

    def parse_p01_course_line(line: str):
        """
        MATEMÁTICA I 3.0 2.0 5.0 4.0 O NINGUNO
        """
        s = norm_ws(line)
        toks = s.split()
        if len(toks) < 6:
            return None

        float_pat = re.compile(r"^\d+(\.\d+)?$")
        idx_num = None
        for i in range(0, len(toks) - 3):
            if (float_pat.match(toks[i]) and float_pat.match(toks[i+1]) and
                float_pat.match(toks[i+2]) and float_pat.match(toks[i+3])):
                idx_num = i
                break

        if idx_num is None:
            return None

        curso = " ".join(toks[:idx_num]).strip()
        ht, hp, th, cred = toks[idx_num], toks[idx_num+1], toks[idx_num+2], toks[idx_num+3]

        tail = toks[idx_num+4:]
        # quitar TC (O/E) si existe
        if tail and re.fullmatch(r"[A-Z]", tail[0], re.IGNORECASE):
            tail = tail[1:]

        req = " ".join(tail).strip() if tail else ""
        return curso, ht, hp, th, cred, req

    # ✅ Streamlit-safe: abrir desde bytes SIEMPRE (especialmente masivo)
    try:
        pdf_bytes = file.getvalue()
    except Exception:
        file.seek(0)
        pdf_bytes = file.read()

    bio = io.BytesIO(pdf_bytes)

    lines = []
    with pdfplumber.open(bio) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            lines.extend(txt.split("\n"))

    rows = []
    semestre = ""
    i = 0

    while i < len(lines):
        line = (lines[i] or "").strip()
        if not line:
            i += 1
            continue

        # Semestre / Ciclo
        if line.upper().startswith("SEMESTRE"):
            semestre = line.split(":")[-1].strip().upper()
            i += 1
            continue

        if line.upper().startswith("CICLO"):
            semestre = line.strip().upper()
            i += 1
            continue

        # ---------- FORMATO OB ----------
        if is_prefix_ob(line):
            prefix = line.strip().replace(" ", "")
            if i + 2 < len(lines):
                curso_line = (lines[i + 1] or "").strip()
                code_line = (lines[i + 2] or "").strip()

                parsed = parse_ob_course_line(curso_line)
                if parsed and is_code_line(code_line):
                    curso, ht, hp, th, cred, req = parsed
                    codigo = prefix + norm_ws(code_line)
                    plan = prefix.split("-")[1] if "-" in prefix else ""
                    rows.append([plan, semestre, codigo, curso, ht, hp, th, cred, req])
                    i += 3
                    continue

            i += 1
            continue

        # ---------- FORMATO P01 / AF REAL ----------
        if is_prefix_p01(line):
            p01 = line.strip().replace(" ", "")  # P01-

            # Esperado real:
            # i+1: "14210- MATEMÁTICA I 3.0 2.0 5.0 4.0 O NINGUNO"  (plan+curso juntos)
            # i+2: "10A01"                                          (codigo)
            if i + 2 < len(lines):
                l_mid = (lines[i + 1] or "").strip()
                toks_mid = norm_ws(l_mid).split()

                if toks_mid and is_digits_prefix_token(toks_mid[0]):
                    mid_token = toks_mid[0]  # 14210-
                    curso_line = " ".join(toks_mid[1:]).strip()  # resto

                    # Buscar código más abajo si el curso se partió en líneas
                    k = i + 2
                    extra = []
                    while k < len(lines) and not is_code_line((lines[k] or "").strip()):
                        extra.append((lines[k] or "").strip())
                        k += 1

                    if extra:
                        curso_line = (curso_line + " " + " ".join(extra)).strip()

                    if k < len(lines):
                        code_line = (lines[k] or "").strip()

                        parsed = parse_p01_course_line(curso_line)
                        if parsed:
                            curso, ht, hp, th, cred, req = parsed
                            codigo = p01 + mid_token + norm_ws(code_line)  # P01-14210-10A01
                            plan = mid_token.replace("-", "").strip()
                            rows.append([plan, semestre, codigo, curso, ht, hp, th, cred, req])
                            i = k + 1
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
    # No dependas de seek() para pdfplumber; pero lo dejamos por compatibilidad
    try:
        file.seek(0)
    except Exception:
        pass

    df = extraer_texto_pdf(file)

    if df is None or df.empty:
        return None

    df.insert(0, "ARCHIVO", nombre_archivo)
    return df


# =====================================================
# CONVERSION INDIVIDUAL
# =====================================================
st.header("Conversión Individual")

archivo = st.file_uploader(
    "Sube archivo",
    type=["pdf", "xlsx", "csv", "docx"]
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
            st.error("No se pudo procesar el PDF (formato no reconocido o texto no extraíble).")
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
            try:
                archivo_pdf.seek(0)
            except Exception:
                pass

            df = procesar_pdf(archivo_pdf, archivo_pdf.name)

            if df is not None and not df.empty:
                dfs.append(df)
            else:
                st.warning(f"No se pudo extraer: {archivo_pdf.name}")

            progreso.progress((i + 1) / total)

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
            st.error("No se pudo convertir ningún PDF.")
