from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from docx import Document
from docxtpl import DocxTemplate
import shutil
import os
from datetime import datetime

app = FastAPI()

# ----------------------------
# FORMATEOS
# ----------------------------

def fecha_espanol():
    meses = [
        "enero","febrero","marzo","abril","mayo","junio",
        "julio","agosto","septiembre","octubre","noviembre","diciembre"
    ]
    now = datetime.now()
    return f"{now.day} de {meses[now.month-1]} de {now.year}"

def limpiar_nombre(nombre):
    nombre = nombre.replace("Sr.", "").replace("Sra.", "").strip()
    return nombre

def primer_nombre(nombre):
    return nombre.split()[0].capitalize()

def determinar_tratamiento(nombre, cargo):
    cargo = cargo.lower()
    if "gerente" in cargo or "director" in cargo or "presidente" in cargo:
        return "Doctor"
    return "Señor"

# ----------------------------
# EXTRACCIÓN DE DATOS
# ----------------------------

def extraer_datos(doc):
    datos = {}

    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]

            if len(cells) >= 2:
                key = cells[0].lower()
                value = cells[1]

                if "contacto" in key:
                    datos["nombre"] = limpiar_nombre(value)

                elif "cargo" in key:
                    datos["cargo"] = value

                elif "compañía" in key:
                    datos["compania"] = value.upper()

                elif "teléfono" in key:
                    datos["telefono"] = value

                elif "ciudad" in key:
                    datos["ciudad"] = value

                elif "e- mail" in key:
                    datos["correo"] = value

                elif "tipo de servicio" in key:
                    datos["servicio"] = value.lower()

                elif "tiempo de servicio" in key:
                    datos["modalidad"] = value.lower()

    return datos

# ----------------------------
# DETECCIÓN AVANZADA
# ----------------------------

def detectar_subtipo_y_flags(doc):
    texto = "\n".join([p.text.lower() for p in doc.paragraphs])

    resultado = {
        "arma": "sin_arma",
        "fortalecimiento": False
    }

    if "arma" in texto:
        resultado["arma"] = "armada"

    if "fortalecimiento" in texto:
        resultado["fortalecimiento"] = True

    return resultado

# ----------------------------
# MAPEO DE PLANTILLAS REAL
# ----------------------------

def seleccionar_plantilla(datos, extra):
    servicio = datos.get("servicio", "")

    if "vigilancia" in servicio:
        arma = extra["arma"]
        modalidad = "m" if "mensual" in datos.get("modalidad","") else "e"
        f = "_f" if extra["fortalecimiento"] else ""

        return f"plantillas/vigilancia_{arma}_{modalidad}{f}.docx"

    if "escolta" in servicio:
        return "plantillas/escolta_mensual.docx"

    if "confiabilidad" in servicio:
        return "plantillas/confiabilidad.docx"

    if "monitoreo" in servicio:
        return "plantillas/monitoreo.docx"

    if "electronica" in servicio:
        return "plantillas/seguridad_electronica.docx"

    if "eventos" in servicio:
        return "plantillas/seguridad_en_eventos.docx"

    return None

# ----------------------------
# ENDPOINT
# ----------------------------

@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):
    try:
        temp_path = "temp.docx"
        output_path = "resultado.docx"

        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        doc = Document(temp_path)

        datos = extraer_datos(doc)
        extra = detectar_subtipo_y_flags(doc)

        plantilla = seleccionar_plantilla(datos, extra)

        if not plantilla or not os.path.exists(plantilla):
            return {"error": f"No existe plantilla: {plantilla}"}

        tratamiento = determinar_tratamiento(datos["nombre"], datos["cargo"])

        contexto = {
            "consecutivo": datetime.now().strftime("%Y%m%d%H%M"),
            "fecha": fecha_espanol(),
            "tratamiento": tratamiento,
            "nombre": datos["nombre"].upper(),
            "nombre_simple": primer_nombre(datos["nombre"]),
            "cargo": datos["cargo"],
            "compania": datos["compania"],
            "correo": datos.get("correo",""),
            "telefono": datos.get("telefono",""),
            "ciudad": datos.get("ciudad","")
        }

        doc = DocxTemplate(plantilla)
        doc.render(contexto)
        doc.save(output_path)

        return FileResponse(
            output_path,
            filename="resultado.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        return {"error": str(e)}
