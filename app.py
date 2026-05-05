from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from docx import Document
import shutil
import os
from datetime import datetime

app = FastAPI()


@app.get("/")
def home():
    return {"status": "ok"}


# -------------------------
# EXTRAER DATOS DE TABLAS
# -------------------------
def extraer_datos(doc):
    datos = {
        "nombre": "",
        "cargo": "",
        "compania": "",
        "correo": "",
        "telefono": "",
        "ciudad": "",
        "servicio": ""
    }

    for table in doc.tables:
        for row in table.rows:
            celdas = [c.text.strip() for c in row.cells]

            texto = " ".join(celdas).lower()

            if "contacto" in texto:
                datos["nombre"] = celdas[-1]

            elif "cargo" in texto:
                datos["cargo"] = celdas[-1]

            elif "compañía" in texto or "compania" in texto:
                datos["compania"] = celdas[-1]

            elif "e-mail" in texto or "correo" in texto:
                datos["correo"] = celdas[-1]

            elif "teléfono" in texto or "telefono" in texto:
                datos["telefono"] = celdas[-1]

            elif "ciudad" in texto:
                datos["ciudad"] = celdas[-1]

    return datos


# -------------------------
# DETECTAR SERVICIO (X)
# -------------------------
def detectar_servicio(doc):
    for table in doc.tables:
        for row in table.rows:
            celdas = [c.text.strip().lower() for c in row.cells]

            if len(celdas) < 2:
                continue

            # Detectar si hay X en la fila
            tiene_x = any("x" in c for c in celdas)

            if not tiene_x:
                continue

            texto_fila = " ".join(celdas)

            # 🔥 MAPEO COMPLETO
            if "vigilancia" in texto_fila:
                return "vigilancia"

            if "seguridad electronica" in texto_fila:
                return "electronica"

            if "confiabilidad" in texto_fila:
                return "confiabilidad"

            if "escolta" in texto_fila:
                return "escolta"

            if "monitoreo" in texto_fila:
                return "monitoreo"

            if "eventos" in texto_fila:
                return "eventos"

            if "logisticos" in texto_fila:
                return "logistico"

    return None
    
# -------------------------
# SELECCIONAR PLANTILLA
# -------------------------
def elegir_plantilla(servicio, texto):
    texto = texto.lower()

    if servicio == "vigilancia":
        if "sin arma" in texto:
            return "plantillas/vigilancia_sin_arma_m.docx"
        else:
            return "plantillas/vigilancia_armada_m.docx"

    elif servicio == "escolta":
        if "motorizado" in texto:
            return "plantillas/escolta_motorizado.docx"
        elif "conductor" in texto:
            return "plantillas/escolta_conductor_ev.docx"
        else:
            return "plantillas/escolta_a_pie.docx"

    elif servicio == "confiabilidad":
        return "plantillas/confiabilidad.docx"

    elif servicio == "electronica":
        return "plantillas/seguridad_electronica.docx"

    return "plantillas/monitoreo.docx"


# -------------------------
# REEMPLAZO SEGURO
# -------------------------
def reemplazar(doc, contexto):
    for p in doc.paragraphs:
        for k, v in contexto.items():
            if f"{{{{{k}}}}}" in p.text:
                p.text = p.text.replace(f"{{{{{k}}}}}", str(v))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for k, v in contexto.items():
                        if f"{{{{{k}}}}}" in p.text:
                            p.text = p.text.replace(f"{{{{{k}}}}}", str(v))


# -------------------------
# FORMATEOS
# -------------------------
def formatear_nombre(nombre):
    return nombre.upper()


def primer_nombre(nombre):
    return nombre.split()[0].capitalize()


def obtener_tratamiento(cargo):
    cargo = cargo.lower()
    if "gerente" in cargo or "director" in cargo:
        return "Doctor"
    return "Señor"


def fecha_es():
    meses = [
        "enero","febrero","marzo","abril","mayo","junio",
        "julio","agosto","septiembre","octubre","noviembre","diciembre"
    ]
    hoy = datetime.now()
    return f"{hoy.day} de {meses[hoy.month-1]} de {hoy.year}"


# -------------------------
# ENDPOINT
# -------------------------
@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):
    try:
        temp = "temp.docx"
        salida = "resultado.docx"

        with open(temp, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        doc_input = Document(temp)

        texto = "\n".join([p.text for p in doc_input.paragraphs])

        datos = extraer_datos(doc_input)
        servicio = detectar_servicio(doc_input)

        plantilla = elegir_plantilla(servicio, texto)

        if not os.path.exists(plantilla):
            return {"error": f"No existe plantilla: {plantilla}"}

        doc = Document(plantilla)

        tratamiento = obtener_tratamiento(datos["cargo"])

        contexto = {
            "nombre": formatear_nombre(datos["nombre"]),
            "nombre_simple": primer_nombre(datos["nombre"]),
            "cargo": datos["cargo"],
            "compania": datos["compania"].upper(),
            "correo": datos["correo"],
            "telefono": datos["telefono"],
            "ciudad": datos["ciudad"],
            "fecha": fecha_es(),
            "alcance": datos["ciudad"],
            "tratamiento": tratamiento,
            "saludo": f"Estimado {primer_nombre(datos['nombre'])}"
        }

        reemplazar(doc, contexto)

        doc.save(salida)

        return FileResponse(
            salida,
            filename="resultado.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        return {"error": str(e)}
