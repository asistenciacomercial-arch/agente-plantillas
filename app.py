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
        .strip()
    )

def obtener_tratamiento(cargo):
    cargo = (cargo or "").lower()
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
# EXTRAER DATOS (FORMATO REAL)
# =========================
def extraer_datos(doc):
    datos = {
        "nombre": "",
        "cargo": "",
        "compania": "",
        "correo": "",
        "telefono": "",
        "ciudad": "",
    }

    for table in doc.tables:
        for row in table.rows:
            textos = [c.text.strip() for c in row.cells]

            for i, texto in enumerate(textos):
                t = texto.lower()

                # 🔥 NOMBRE
                if "contacto" in t and i + 1 < len(textos):
                    datos["nombre"] = limpiar_nombre(textos[i + 1])

                # 🔥 EMPRESA
                elif "compañ" in t or "edificio" in t:
                    if i + 1 < len(textos):
                        datos["compania"] = textos[i + 1].upper()

                # 🔥 CORREO (muchas variantes)
                elif any(x in t for x in ["mail", "correo"]):
                    if i + 1 < len(textos):
                        datos["correo"] = textos[i + 1]

                # 🔥 TELÉFONO (muchas variantes)
                elif any(x in t for x in ["tel", "cel"]):
                    if i + 1 < len(textos):
                        datos["telefono"] = textos[i + 1].replace(" ", "")

                # 🔥 CARGO (variantes)
                elif any(x in t for x in ["cargo", "funcion"]):
                    if i + 1 < len(textos):
                        datos["cargo"] = textos[i + 1]

                # 🔥 CIUDAD
                elif "ciudad" in t:
                    if i + 1 < len(textos):
                        datos["ciudad"] = textos[i + 1]

    # fallback ciudad
    if not datos["ciudad"]:
        for p in doc.paragraphs:
            if "colombia" in p.text.lower():
                datos["ciudad"] = p.text.replace(", Colombia", "").strip()

    print("DATOS FINALES:", datos)  # 👈 IMPORTANTE DEBUG

    return datos
# =========================
# DETECCIÓN SERVICIO
# =========================
def detectar_servicio(doc):
    for table in doc.tables:
        for row in table.rows:
            textos = [c.text.strip().lower() for c in row.cells]

            if "vigilancia" in textos and "x" in textos:
                return "vigilancia"

    return "vigilancia"

def detectar_detalle(doc):
    texto = "\n".join([p.text.lower() for p in doc.paragraphs])

    if "armada" in texto:
        return "armada"
    if "sin arma" in texto:
        return "sin_arma"

    return "sin_arma"

def detectar_modalidad(doc):
    texto = "\n".join([p.text.lower() for p in doc.paragraphs])

    if "mensual" in texto:
        return "m"
    return "m"

# =========================
# PLANTILLA
# =========================
def seleccionar_plantilla(servicio, detalle, modalidad):
    return "plantillas/vigilancia_sin_arma_m.docx"

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

        # leer documento
        doc = Document(temp)

        datos = extraer_datos(doc)

        print("DATOS:", datos)

        # validar nombre
        nombre = datos.get("nombre")
        if not nombre:
            nombre = "CLIENTE"

        tratamiento = obtener_tratamiento(datos.get("cargo", ""))

        servicio = detectar_servicio(doc)
        detalle = detectar_detalle(doc)
        modalidad = detectar_modalidad(doc)

        plantilla = seleccionar_plantilla(servicio, detalle, modalidad)

        # =========================
        # REEMPLAZOS CORRECTOS
        # =========================
        reemplazos = {
            "consecutivo": datetime.now().strftime("%Y%m%d%H%M"),
            "fecha": fecha_es(),

            "tratamiento": tratamiento,
            "nombre": nombre,

            "cargo": datos.get("cargo", ""),
            "compania": datos.get("compania", ""),
            "correo": datos.get("correo", ""),
            "telefono": datos.get("telefono", ""),
            "ciudad": datos.get("ciudad", ""),

            "saludo": f"Estimado {tratamiento} {nombre}",
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
