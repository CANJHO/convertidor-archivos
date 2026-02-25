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
    - OB: P09-20242- + (curso con '00') + codigo (P09A1101)
    - AF/P01: P01- + 14210-/20201- + COD + (curso en 1 o varias líneas) + (números en misma línea o línea aparte)

    ✅ Corrección clave:
    - Evita confundir "PERSONAL" / "HUMANO" como código.
      Ahora el código debe tener al menos 1 dígito.
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

    def is_digits_prefix(line: str) -> bool:
        # "14210-" o "20201-"
        s = (line or "").strip().replace(" ", "")
        return re.fullmatch(r"\d+-", s) is not None

    def is_digits_prefix_token(tok: str) -> bool:
        return re.fullmatch(r"\d+-", tok or "") is not None

    def is_code_line(line: str) -> bool:
        """
        ✅ El código debe tener al menos 1 dígito.
        Evita que "PERSONAL" o "HUMANO" se confundan como código.
        """
        s = norm_ws(line).replace(" ", "")
        return re.fullmatch(r"(?=.*\d)[A-Z0-9]{3,15}", s) is not None

    float_pat = re.compile(r"^\d+(\.\d+)?$")

    def has_4_floats_in_row(text: str) -> bool:
        toks = norm_ws(text).split()
        for i in range(0, len(toks) - 3):
            if all(float_pat.match(toks[i + j]) for j in range(4)):
                return True
        return False

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

    def parse_p01_course_line(text: str):
        """
        CURSO ... 3.0 2.0 5.0 4.0 O NINGUNO
        """
        s = norm_ws(text)
        toks = s.split()
        if len(toks) < 6:
            return None

        idx_num = None
        for i in range(0, len(toks) - 3):
            if all(float_pat.match(toks[i + j]) for j in range(4)):
                idx_num = i
                break
        if idx_num is None:
            return None

        curso = " ".join(toks[:idx_num]).strip()
        ht, hp, th, cred = toks[idx_num], toks[idx_num + 1], toks[idx_num + 2], toks[idx_num + 3]

        tail = toks[idx_num + 4:]
        # quitar TC (O/E) si existe
        if tail and re.fullmatch(r"[A-Z]", tail[0], re.IGNORECASE):
            tail = tail[1:]
        req = " ".join(tail).strip() if tail else ""
        return curso, ht, hp, th, cred, req

    # ✅ Streamlit-safe: bytes SIEMPRE
    try:
        pdf_bytes = file.getvalue()
    except Exception:
        try:
            file.seek(0)
        except Exception:
            pass
        pdf_bytes = file.read()

    bio = io.BytesIO(pdf_bytes)

    # leer líneas completas
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

        # ---------- FORMATO P01 / AF ----------
        # Maneja 2 variantes:
        # A) P01- / 14210- / 10A01 / CURSO... (multi línea) / NUMS...
        # B) P01- / "14210- CURSO... NUMS..." / (PERSONAL/HUMANO...) / 25A29
        if is_prefix_p01(line):
            p01 = line.strip().replace(" ", "")  # P01-

            # Necesitamos mirar desde i+1 hacia adelante para detectar la variante real
            if i + 1 >= len(lines):
                i += 1
                continue

            # Caso B: la línea i+1 trae "14210- CURSO..."
            l1 = (lines[i + 1] or "").strip()
            toks1 = norm_ws(l1).split()

            # Detectar "14210-" como primer token
            if toks1 and is_digits_prefix_token(toks1[0]):
                plan_token = toks1[0]  # 14210-
                rest_after_plan = " ".join(toks1[1:]).strip()

                # Buscamos el código real hacia abajo (saltando PERSONAL/HUMANO/etc.)
                k = i + 2
                extra_parts = []

                while k < len(lines):
                    s = (lines[k] or "").strip()
                    if not s:
                        k += 1
                        continue

                    if is_code_line(s):
                        break

                    # cortar si empieza otro bloque
                    if is_prefix_p01(s) or is_prefix_ob(s) or s.upper().startswith("CICLO") or s.upper().startswith("SEMESTRE"):
                        break

                    # acumular como parte del curso (PERSONAL/HUMANO, etc.)
                    extra_parts.append(s)
                    k += 1

                if k < len(lines) and is_code_line((lines[k] or "").strip()):
                    code_line = (lines[k] or "").strip()

                    candidate = " ".join([rest_after_plan, *extra_parts]).strip()

                    # Si candidate no trae números aún, intenta sumar 1 línea más (números abajo)
                    if not has_4_floats_in_row(candidate) and (k + 1) < len(lines):
                        candidate2 = (candidate + " " + (lines[k + 1] or "").strip()).strip()
                        if has_4_floats_in_row(candidate2):
                            candidate = candidate2

                    parsed = parse_p01_course_line(candidate)
                    if parsed:
                        curso, ht, hp, th, cred, req = parsed
                        codigo_full = f"{p01}{plan_token}{norm_ws(code_line)}"
                        plan = plan_token.replace("-", "").strip()
                        rows.append([plan, semestre, codigo_full, curso, ht, hp, th, cred, req])
                        i = k + 1
                        continue

            # Caso A: P01- / 14210- / 10A01 / curso...
            if i + 2 < len(lines):
                mid = (lines[i + 1] or "").strip()   # 14210-
                code = (lines[i + 2] or "").strip()  # 10A01 / 25A27 / 20A40

                if is_digits_prefix(mid) and is_code_line(code):
                    j = i + 3
                    course_parts = []
                    nums_line = None

                    while j < len(lines):
                        s = (lines[j] or "").strip()
                        if not s:
                            j += 1
                            continue

                        if is_prefix_p01(s) or is_prefix_ob(s) or s.upper().startswith("CICLO") or s.upper().startswith("SEMESTRE") or s.upper().startswith("CODIGO"):
                            break

                        if has_4_floats_in_row(s):
                            nums_line = s
                            j += 1
                            break

                        course_parts.append(s)
                        j += 1

                    if nums_line is None:
                        candidate = " ".join(course_parts).strip()
                    else:
                        candidate = (" ".join(course_parts) + " " + nums_line).strip()

                    parsed = parse_p01_course_line(candidate)
                    if parsed:
                        curso, ht, hp, th, cred, req = parsed
                        codigo_full = f"{p01}{mid.replace(' ', '')}{norm_ws(code)}"
                        plan = mid.replace("-", "").strip()
                        rows.append([plan, semestre, codigo_full, curso, ht, hp, th, cred, req])
                        i = j
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