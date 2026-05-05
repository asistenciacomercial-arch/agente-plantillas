from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from docx import Document
from docxtpl import DocxTemplate
from io import BytesIO
import shutil
from datetime import datetime

app = FastAPI()

# =========================
# LIMPIEZA
# =========================
def limpiar_nombre(nombre):
    return (
        nombre.upper()
        .replace("SR.", "")
        .replace("SRA.", "")
        .replace("DR.", "")
        .replace("DRA.", "")
        .replace("E-MAIL", "")
        .replace("EMAIL", "")
        .strip()
    )

def primer_nombre(nombre):
    return nombre.split()[0].capitalize() if nombre else ""

def obtener_tratamiento(cargo):
    cargo = cargo.lower()
    if any(x in cargo for x in ["gerente", "director", "presidente"]):
        return "Doctor"
    return "Señor"

# =========================
# FECHA
# =========================
def fecha_es():
    meses = [
        "enero","febrero","marzo","abril","mayo","junio",
        "julio","agosto","septiembre","octubre","noviembre","diciembre"
    ]
    hoy = datetime.now()
    return f"{hoy.day} de {meses[hoy.month-1]} de {hoy.year}"

# =========================
# EXTRAER DATOS TABLA
# =========================
import re

def extraer_datos(doc):
    texto = ""

    # 🔥 leer párrafos
    for p in doc.paragraphs:
        texto += p.text + "\n"

    # 🔥 leer tablas (CLAVE)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texto += cell.text + "\n"

    lineas = [l.strip() for l in texto.split("\n") if l.strip()]

    datos = {
        "nombre": "",
        "cargo": "",
        "compania": "",
        "correo": "",
        "telefono": "",
        "ciudad": "",
    }

    # =====================
    # NOMBRE (línea con SR o DOCTOR)
    # =====================
    for l in lineas:
        if any(x in l.upper() for x in ["SR.", "SEÑOR", "DOCTOR"]):
            datos["nombre"] = limpiar_nombre(l)
            break

    # =====================
    # TELÉFONO
    # =====================
    for l in lineas:
        limpio = l.replace(" ", "")
        if re.match(r"3\d{9}", limpio):
            datos["telefono"] = l
            break

    # =====================
    # CIUDAD
    # =====================
    for l in lineas:
        if "COLOMBIA" in l.upper():
            datos["ciudad"] = l.split(",")[0]
            break

    # =====================
    # CORREO
    # =====================
    match = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", texto)
    if match:
        datos["correo"] = match.group()

    # =====================
    # COMPAÑÍA
    # =====================
    for l in lineas:
        if "EDIFICIO" in l.upper():
            datos["compania"] = l
            break

    # =====================
    # CARGO (si existe)
    # =====================
    for l in lineas:
        if any(x in l.lower() for x in ["gerente", "director", "presidente"]):
            datos["cargo"] = l
            break

    return datos

    # =====================
    # TELÉFONO
    # =====================
    for l in lineas:
        if re.search(r"\b3\d{9}\b", l.replace(" ", "")):
            datos["telefono"] = l
            break

    # =====================
    # CIUDAD
    # =====================
    for l in lineas:
        if "COLOMBIA" in l.upper():
            datos["ciudad"] = l.split(",")[0]
            break

    # =====================
    # CORREO
    # =====================
    match = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", texto)
    if match:
        datos["correo"] = match.group()

    # =====================
    # CARGO (heurística)
    # =====================
    for l in lineas:
        if any(p in l.lower() for p in ["gerente", "director", "presidente", "jefe"]):
            datos["cargo"] = l
            break

    # =====================
    # COMPAÑÍA
    # =====================
    for l in lineas:
        if "EDIFICIO" in l.upper():
            datos["compania"] = l
            break

    return datos
    
# =========================
# DETECTAR SERVICIO (CLAVE)
# =========================
def detectar_servicio(doc):
    for table in doc.tables:
        for row in table.rows:
            textos = [c.text.strip().lower() for c in row.cells]

            if any("vigilancia" in t for t in textos) and "x" in textos:
                return "vigilancia"

            if any("escolta" in t for t in textos) and "x" in textos:
                return "escolta"

            if any("confiabilidad" in t for t in textos) and "x" in textos:
                return "confiabilidad"

            if any("seguridad electronica" in t for t in textos) and "x" in textos:
                return "electronica"

            if any("monitoreo" in t for t in textos) and "x" in textos:
                return "monitoreo"

    return "vigilancia"  # fallback
# DETECTAR TIPO
# =========================
def detectar_tipo(doc):
    texto = "\n".join([p.text.lower() for p in doc.paragraphs])

    tipo = ""

    if "sin arma" in texto:
        tipo += "_sin_arma"
    elif "armada" in texto:
        tipo += "_armada"

    if "mensual" in texto:
        tipo += "_m"
    elif "evento" in texto or "día" in texto:
        tipo += "_e"

    if "fortalecimiento" in texto:
        tipo += "_f"

    return tipo
    
def detectar_detalle(doc):
    texto = "\n".join([p.text.lower() for p in doc.paragraphs])

    if "armada" in texto:
        return "armada"
    if "sin arma" in texto:
        return "sin_arma"
    if "escolta" in texto:
        return "escolta"

    return "sin_arma"  # fallback seguro
        
def detectar_modalidad(doc):
    texto = "\n".join([p.text.lower() for p in doc.paragraphs])

    if "mensual" in texto:
        return "m"
    if "festivo" in texto or "fin de semana" in texto:
        return "e"
    if "fortalecimiento" in texto:
        return "f"

    return "m"
# =========================
# SELECCIONAR PLANTILLA
# =========================
def seleccionar_plantilla(servicio, detalle, modalidad):
    if servicio == "vigilancia":
        if detalle == "armada":
            if modalidad == "m":
                return "plantillas/vigilancia_armada_m.docx"
            if modalidad == "f":
                return "plantillas/vigilancia_armada_m_f.docx"
            return "plantillas/vigilancia_armada_e.docx"

        if detalle == "sin_arma":
            if modalidad == "m":
                return "plantillas/vigilancia_sin_arma_m.docx"
            if modalidad == "f":
                return "plantillas/vigilancia_sin_arma_f_m.docx"
            return "plantillas/vigilancia_sin_arma_e_12h.docx"

    if servicio == "escolta":
        if detalle == "motorizado":
            return "plantillas/escolta_motorizado.docx"
        return "plantillas/escolta_a_pie.docx"

    if servicio == "confiabilidad":
        return "plantillas/confiabilidad.docx"

    if servicio == "electronica":
        return "plantillas/seguridad_electronica.docx"

    if servicio == "monitoreo":
        return "plantillas/monitoreo.docx"

    return "plantillas/vigilancia_sin_arma_m.docx"
    
# =========================
# REEMPLAZO SEGURO
# =========================

# =========================
# API
# =========================
@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):
    try:
        temp = "entrada.docx"

        # guardar archivo
        with open(temp, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        file.file.close()

        # leer levantamiento
        doc = Document(temp)

        datos = extraer_datos(doc)

        print("DATOS EXTRAIDOS:", datos)

        servicio = detectar_servicio(doc)
        detalle = detectar_detalle(doc)
        modalidad = detectar_modalidad(doc)

        print("SERVICIO:", servicio)
        print("DETALLE:", detalle)
        print("MODALIDAD:", modalidad)

        # seleccionar plantilla
        plantilla = seleccionar_plantilla(servicio, detalle, modalidad)

        # 🔥 crear reemplazos (ESTO TE FALTABA)
        reemplazos = {
            "consecutivo": datetime.now().strftime("%Y%m%d%H%M"),
            "fecha": fecha_es(),
            "nombre": datos.get("nombre", ""),
            "cargo": datos.get("cargo", ""),
            "compania": datos.get("compania", ""),
            "correo": datos.get("correo", ""),
            "telefono": datos.get("telefono", ""),
            "ciudad": datos.get("ciudad", ""),
            "alcance": datos.get("ciudad", ""),
            "tratamiento": obtener_tratamiento(datos.get("cargo", "")),
            "saludo": "Estimado",
            "nombre_corto": datos.get("nombre", "").split(" ")[0].capitalize() if datos.get("nombre") else ""
        }

        # generar documento
        doc_tpl = DocxTemplate(plantilla)
        doc_tpl.render(reemplazos)

        buffer = BytesIO()
        doc_tpl.save(buffer)
        buffer.seek(0)

        return StreamingResponse(
            buffer,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": "attachment; filename=resultado.docx"
            }
        )

    except Exception as e:
        return {"error": str(e)}
