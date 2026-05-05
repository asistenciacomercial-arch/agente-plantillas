from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from docx import Document
import shutil
import os
import re
from datetime import datetime

app = FastAPI()

# =========================
# 🧠 UTILIDADES
# =========================

def obtener_fecha():
    meses = {
        "January": "enero","February": "febrero","March": "marzo",
        "April": "abril","May": "mayo","June": "junio",
        "July": "julio","August": "agosto","September": "septiembre",
        "October": "octubre","November": "noviembre","December": "diciembre"
    }
    now = datetime.now()
    return f"{now.day} de {meses[now.strftime('%B')]} de {now.year}"

def consecutivo():
    return datetime.now().strftime("%Y%m%d%H%M")

# =========================
# 🧠 EXTRAER TEXTO COMPLETO
# =========================

def extraer_texto(doc):
    texto = []

    for p in doc.paragraphs:
        texto.append(p.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texto.append(cell.text)

    return "\n".join(texto).lower()

# =========================
# 🧠 EXTRACCIÓN INTELIGENTE
# =========================

def extraer_datos(texto):
    datos = {}

    # 📌 correo
    correo = re.findall(r'[\w\.-]+@[\w\.-]+', texto)
    if correo:
        datos["correo"] = correo[0]

    # 📌 teléfono
    tel = re.findall(r'\b3\d{9}\b', texto)
    if tel:
        datos["telefono"] = tel[0]

    # 📌 nombre (busca después de "contacto")
    match = re.search(r'contacto\s*[:\-]?\s*([a-z\s]+)', texto)
    if match:
        datos["nombre"] = match.group(1).strip().upper()

    # 📌 cargo
    match = re.search(r'cargo\s*[:\-]?\s*([a-z\s]+)', texto)
    if match:
        datos["cargo"] = match.group(1).strip()

    # 📌 compañía
    match = re.search(r'compa[ñn]ia\s*[:\-]?\s*([a-z0-9\s\.]+)', texto)
    if match:
        datos["compania"] = match.group(1).strip().upper()

    # 📌 ciudad
    if "bogotá" in texto:
        datos["ciudad"] = "Bogotá"
    elif "medellin" in texto:
        datos["ciudad"] = "Medellín"
    else:
        datos["ciudad"] = "Bogotá"

    return datos

# =========================
# 🧠 DETECTAR SERVICIO
# =========================

def detectar_servicio(texto):
    if "vigilancia" in texto:
        if "sin arma" in texto:
            if "mensual" in texto:
                return "plantillas/vigilancia_sin_arma_m.docx"
            return "plantillas/vigilancia_sin_arma_e_12h.docx"
        else:
            if "mensual" in texto:
                return "plantillas/vigilancia_armada_m.docx"
            return "plantillas/vigilancia_armada_e.docx"

    if "escolta" in texto:
        if "motorizado" in texto:
            return "plantillas/escolta_motorizado.docx"
        if "conductor" in texto:
            return "plantillas/escolta_conductor_ev.docx"
        return "plantillas/escolta_a_pie.docx"

    if "confiabilidad" in texto:
        return "plantillas/confiabilidad.docx"

    if "electronica" in texto:
        return "plantillas/seguridad_electronica.docx"

    if "evento" in texto:
        return "plantillas/seguridad_en_eventos.docx"

    return None

# =========================
# 🧠 TRATAMIENTO
# =========================

def tratamiento(nombre, cargo):
    if "gerente" in (cargo or "").lower():
        return "Doctor"
    return "Señor"

# =========================
# 🧠 REEMPLAZO
# =========================

def reemplazar(doc, contexto):
    def reemplazar_en_runs(paragraph):
        for run in paragraph.runs:
            texto = run.text
            for k, v in contexto.items():
                texto = texto.replace(f"{{{{{k}}}}}", str(v))
            run.text = texto

    # Párrafos normales
    for p in doc.paragraphs:
        reemplazar_en_runs(p)

    # Tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    reemplazar_en_runs(p)
# =========================
# 🚀 ENDPOINT
# =========================

@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):
    try:
        temp = "temp.docx"
        out = "resultado.docx"

        with open(temp, "wb") as f:
            shutil.copyfileobj(file.file, f)

        doc = Document(temp)

        texto = extraer_texto(doc)
        datos = extraer_datos(texto)

        plantilla = detectar_servicio(texto)

        if not plantilla or not os.path.exists(plantilla):
            return {"error": "No se detectó plantilla"}

        nombre = datos.get("nombre", "CLIENTE")
        contexto = {
            "consecutivo": consecutivo(),
            "fecha": obtener_fecha(),
            "tratamiento": tratamiento(nombre, datos.get("cargo")),
            "nombre": nombre,
            "cargo": datos.get("cargo", ""),
            "compania": datos.get("compania", ""),
            "correo": datos.get("correo", ""),
            "telefono": datos.get("telefono", ""),
            "ciudad": datos.get("ciudad", ""),
            "alcance": datos.get("ciudad", "")
        }

        doc_final = Document(plantilla)
        reemplazar(doc_final, contexto)
        doc_final.save(out)

        return FileResponse(out, filename="resultado.docx")

    except Exception as e:
        return {"error": str(e)}
