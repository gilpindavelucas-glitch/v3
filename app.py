# app.py
import streamlit as st
from io import BytesIO
import os
import re
import shutil
import datetime
import pandas as pd
from pathlib import Path

# Lectura de .docx y .pdf
import docx2txt
import PyPDF2
import dateparser

# Excel
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ---------- Config ----------
BRAND_BLUE = "#003399"
BRAND_WHITE = "#FFFFFF"
APP_TITLE = "Ford Fiorasi – Procesador de Antecedentes Disciplinarios"
LOGO_PATH = "logo_ford_fiorasi.png"  # poner el logo en el repo
OUTPUT_FOLDER = "procesados_output"
# ----------------------------

st.set_page_config(page_title=APP_TITLE, layout="wide", page_icon=":shield:")

# Branding header
st.markdown(
    f"""
    <div style="display:flex;align-items:center;background:{BRAND_BLUE};padding:10px;border-radius:8px;">
      <img src="data:image/png;base64,{'' if not Path(LOGO_PATH).exists() else ''}" 
           style="height:64px;margin-right:16px;border-radius:8px;"/>
      <h1 style="color:{BRAND_WHITE};margin:0">{APP_TITLE}</h1>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("---")

st.write("**Instrucciones:** arrastra y suelta los archivos .docx y .pdf a continuación (puedes seleccionar varios). Luego presiona **Procesar antecedentes**.")

uploaded_files = st.file_uploader("Archivos (.docx, .pdf) — múltiple", accept_multiple_files=True, type=["docx", "pdf"])

# Opcional: carpeta de salida en servidor
output_folder_input = st.text_input("Carpeta de salida en el servidor (opcional)", value=OUTPUT_FOLDER)

# Keywords y heurísticas
TYPE_KEYWORDS = {
    "Llamado de atención": ["llamado de atención", "llamado de atencion", "llamado de atención:"],
    "Apercibimiento": ["apercebimiento", "aperibimiento"],
    "Solicitud de descargo": ["solicitud de descargo", "solicitar descargo", "requerimos descargo"],
    "Contestación": ["contestación", "contestacion", "contestó", "contestó:"],
    "Acta": ["acta", "acta de", "acta de incumplimiento"],
}

RESPONSE_KEYWORDS = ["contestó", "contestación", "contestacion", "presentó descargo", "presento descargo", "se recibió descargo", "adjunta contestación", "adjunto descargo", "se presentó descargo"]

def extract_text_from_docx_bytes(bts):
    try:
        tmp = "tmp_docx.docx"
        with open(tmp, "wb") as f:
            f.write(bts)
        text = docx2txt.process(tmp) or ""
        os.remove(tmp)
        return text
    except Exception as e:
        return ""

def extract_text_from_pdf_bytes(bts):
    try:
        reader = PyPDF2.PdfReader(BytesIO(bts))
        text = []
        for page in reader.pages:
            page_text = page.extract_text() or ""
            text.append(page_text)
        return "\n".join(text)
    except Exception as e:
        return ""

def extract_date(text):
    # Buscar fecha por regex (varios formatos) y parsear con dateparser
    patterns = [
        r"\b(\d{1,2}\s+de\s+[A-Za-z]+(?:\s+de\s+\d{4})?)\b",   # e.g., 5 de marzo de 2024
        r"\b(\d{1,2}/\d{1,2}/\d{2,4})\b",
        r"\b(\d{4}-\d{1,2}-\d{1,2})\b",
        r"\b(\d{1,2}\s+[A-Za-z]+\s+\d{4})\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            dt = dateparser.parse(m.group(1), languages=['es'])
            if dt:
                return dt.date().isoformat()
    return ""

def detect_type(text_lower):
    found = []
    for t, keys in TYPE_KEYWORDS.items():
        for k in keys:
            if k in text_lower:
                found.append(t)
                break
    return ", ".join(found) if found else "No determinado"

def detect_response(text_lower):
    for k in RESPONSE_KEYWORDS:
        if k in text_lower:
            return "Sí"
    return "No"

def first_relevant_paragraph(text):
    # dividir por dobles saltos y tomar primer párrafo con >30 caracteres
    parts = re.split(r"\n\s*\n", text)
    for p in parts:
        p_clean = p.strip()
        if len(p_clean) > 30:
            return p_clean.replace("\n", " ")
    # fallback: primeras 200 chars
    return text.strip().replace("\n", " ")[:500]

def guess_name(text):
    # Heurística: buscar líneas con "Señor", "Sra.", o mayúsculas seguidas o "Apellido y Nombre"
    # Buscar "Apellido y Nombre" literal
    m = re.search(r"(Apellido[s]?\s*y\s*Nombre[s]?[:\s]+)([A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚñáéíóú\s,.-]+)", text, flags=re.IGNORECASE)
    if m:
        return m.group(2).strip()
    # buscar línea con "Sr.", "Sra.", "Señor", "Sra" seguido de nombre
    m = re.search(r"\b(Sr\.|Sra\.|Señor|Señora)\s+([A-Z][a-zA-ZáéíóúñÁÉÍÓÚÑ\s]+)", text)
    if m:
        return m.group(2).strip()
    # buscar mayúsculas sostenidas (APELLIDO, NOMBRE)
    m = re.search(r"\n([A-ZÁÉÍÓÚÑ]{2,}[,\sA-ZÁÉÍÓÚÑ\-\.]{5,})\n", text)
    if m:
        cand = m.group(1).strip()
        cand = re.sub(r"\s{2,}", " ", cand)
        return cand.title()
    # fallback vacío
    return ""

# Procesar botón
if st.button("Procesar antecedentes"):
    if not uploaded_files:
        st.error("No se han subido archivos. Seleccioná los documentos .docx y .pdf o subí una carpeta/comprimido con ellos.")
    else:
        out_folder = Path(output_folder_input)
        out_folder.mkdir(parents=True, exist_ok=True)

        records = []
        per_employee_files = {}

        for up in uploaded_files:
            filename = up.name
            content = up.read()
            text = ""
            if filename.lower().endswith(".docx"):
                text = extract_text_from_docx_bytes(content)
            elif filename.lower().endswith(".pdf"):
                text = extract_text_from_pdf_bytes(content)
            text = text or ""

            text_lower = text.lower()
            name = guess_name(text) or "SIN_NOMBRE_DETECTADO"
            date_iso = extract_date(text) or ""
            tipo = detect_type(text_lower)
            desc = first_relevant_paragraph(text)
            contestacion = detect_response(text_lower)

            # Registro
            records.append({
                "Apellido y Nombre": name,
                "Fecha emisión": date_iso,
                "Tipo de antecedente": tipo,
                "Descripción breve": desc,
                "Contestación/Descargo": contestacion,
                "Archivo original": filename
            })

            # Guardar archivo en subcarpeta por empleado
            emp_folder_name = re.sub(r"[^\w\-_\. ]", "_", name) or "SIN_NOMBRE"
            emp_folder = out_folder / emp_folder_name
            emp_folder.mkdir(parents=True, exist_ok=True)
            # Guardar el archivo original en la subcarpeta
            target_path = emp_folder / filename
            with open(target_path, "wb") as f:
                f.write(content)

            # Track files por empleado
            per_employee_files.setdefault(name, []).append({
                "filename": filename,
                "path": str(target_path),
                "fecha": date_iso,
                "tipo": tipo,
                "descripcion": desc,
                "contestacion": contestacion
            })

        # DataFrame base
        df = pd.DataFrame(records)
        # ordenar alfabéticamente por Apellido y Nombre
        df_sorted = df.sort_values(by=["Apellido y Nombre"]).reset_index(drop=True)

        # Generar hoja resumen por empleado
        resumen_rows = []
        for emp, files in per_employee_files.items():
            cant = len(files)
            tipos = sorted({f["tipo"] for f in files if f["tipo"]})
            fechas = [f["fecha"] for f in files if f["fecha"]]
            ultima_fecha = max(fechas) if fechas else ""
            # sintetizar hechos: concatenar primeros 200 caracteres de cada descripción
            sintet = " | ".join([f["descripcion"][:180] for f in files])
            resumen_rows.append({
                "Apellido y Nombre": emp,
                "Cantidad de antecedentes": cant,
                "Tipos recibidos": ", ".join(tipos) if tipos else "",
                "Última fecha": ultima_fecha,
                "Síntesis de hechos": sintet
            })
        df_resumen = pd.DataFrame(resumen_rows)
        df_resumen = df_resumen.sort_values(by=["Apellido y Nombre"]).reset_index(drop=True)

        # Nombre archivo Excel con año actual
        year = datetime.date.today().year
        excel_name = f"FordFiorasi_Antecedentes_Base_{year}.xlsx"
        excel_path = out_folder / excel_name

        # Guardar a Excel con dos hojas
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df_sorted.to_excel(writer, sheet_name="Base completa", index=False)
            df_resumen.to_excel(writer, sheet_name="Resumen por empleado", index=False)

        st.success(f"Procesamiento completado. Archivo Excel generado: `{excel_path}`")
        with open(excel_path, "rb") as f:
            st.download_button("Descargar Excel", f, file_name=excel_name)

        st.markdown("### Archivos movidos por empleado")
        for emp, files in per_employee_files.items():
            st.write(f"**{emp}** — {len(files)} archivo(s)")
            for fi in files:
                st.write(f"- {fi['filename']}  ({fi['tipo']}) -> {fi['path']}")

        st.info("Si querés que el sistema busque coincidencias más exactas de nombre (DNI, legajo), hay que incluir una base maestra de empleados para emparejar por similitud. Puedo añadir esa función si la necesitás.")
