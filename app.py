from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from docxtpl import DocxTemplate
from docx import Document
import shutil
import os
from datetime import datetime

app = FastAPI()

# ------------------------
# UTILIDADES
# ------------------------

def obtener_fecha():
    meses = {
        "January": "enero", "February": "febrero", "March": "marzo",
        "April": "abril", "May": "mayo", "June": "junio",
        "July": "julio", "August": "agosto", "September": "septiembre",
        "October": "octubre", "November": "noviembre", "December": "diciembre"
    }
    now = datetime.now()
    mes = meses[now.strftime("%B")]
    return f"{now.day} de {mes} de {now.year}"


def primer_nombre(nombre):
    return nombre.split()[0].capitalize()


def obtener_tratamiento(nombre, cargo):
    cargo = cargo.lower()

    if "gerente" in cargo or "director" in cargo:
        return "Doctor"

    if nombre.strip().endswith("a"):
        return "Señora"

    return "Señor"


def obtener_consecutivo():
    return datetime.now().strftime("%Y%m%d%H%M")


# ------------------------
# EXTRAER DATOS
# ------------------------

def extraer_datos_tabla(doc):
    data = {
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
            cells = [c.text.strip() for c in row.cells]

            if len(cells) < 2:
                continue

            key = cells[0].lower()
            value = cells[1]

            if "contacto" in key:
                data["nombre"] = value

            elif "cargo" in key:
                data["cargo"] = value

            elif "compañ" in key:
                data["compania"] = value

            elif "mail" in key or "correo" in key:
                data["correo"] = value

            elif "tel" in key:
                data["telefono"] = value

            elif "ciudad" in key:
                data["ciudad"] = value

            elif "servicio" in key or "tipo" in key:
                data["servicio"] = value

    return data


# ------------------------
# DETECTAR SERVICIO COMPLETO
# ------------------------

def detectar_servicio_completo(data):
    texto = data["servicio"].lower()

    if "vigilancia" in texto:
        if "sin arma" in texto:
            return "vigilancia_sin_arma"
        elif "arma" in texto:
            return "vigilancia_armada"
        else:
            return "vigilancia"

    elif "escolta" in texto:
        if "motorizado" in texto:
            return "escolta_motorizado"
        elif "conductor" in texto:
            return "escolta_conductor"
        else:
            return "escolta_a_pie"

    elif "monitoreo" in texto:
        return "monitoreo"

    elif "ciber" in texto:
        return "ciberseguridad"

    return "general"


# ------------------------
# PLANTILLA BASE
# ------------------------

def seleccionar_plantilla_por_tipo(data):
    texto = data["servicio"].lower()

    if "mensual" in texto:
        return "plantillas/base_m.docx"

    elif "eventual" in texto or "día" in texto:
        return "plantillas/base_e.docx"

    elif "fortalecimiento" in texto:
        return "plantillas/base_f.docx"

    elif "valor agregado" in texto:
        return "plantillas/base_av.docx"

    return "plantillas/base_e.docx"


# ------------------------
# ALCANCE
# ------------------------

def generar_alcance(servicio):

    if servicio == "vigilancia_armada":
        return "Servicio de vigilancia armada con control de accesos, rondas de seguridad y reacción ante incidentes."

    elif servicio == "vigilancia_sin_arma":
        return "Servicio de vigilancia sin arma enfocado en control de accesos y prevención de riesgos."

    elif servicio == "escolta_motorizado":
        return "Servicio de escolta motorizado con reacción inmediata y cobertura en desplazamientos."

    elif servicio == "escolta_conductor":
        return "Servicio de escolta conductor con protección y conducción segura."

    elif servicio == "escolta_a_pie":
        return "Servicio de escolta a pie para protección directa del cliente."

    elif servicio == "monitoreo":
        return "Servicio de monitoreo electrónico con respuesta ante eventos."

    elif servicio == "ciberseguridad":
        return "Capacitación en ciberseguridad preventiva."

    return "Servicio ajustado a las necesidades del cliente."


# ------------------------
# API
# ------------------------

@app.get("/")
def home():
    return {"status": "ok"}


@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):
    try:
        temp_path = "temp.docx"
        output_path = "resultado.docx"

        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        doc = Document(temp_path)

        data = extraer_datos_tabla(doc)

        servicio = detectar_servicio_completo(data)
        plantilla = seleccionar_plantilla_por_tipo(data)
        alcance = generar_alcance(servicio)

        nombre = data["nombre"].upper()
        nombre_simple_val = primer_nombre(data["nombre"])

        context = {
            "consecutivo": obtener_consecutivo(),
            "fecha": obtener_fecha(),

            "tratamiento": obtener_tratamiento(data["nombre"], data["cargo"]),

            "nombre": nombre,
            "nombre_simple": nombre_simple_val,

            "cargo": data["cargo"],
            "compania": data["compania"].upper(),

            "correo": data["correo"],
            "telefono": data["telefono"],
            "ciudad": data["ciudad"],

            "alcance": alcance
        }

        doc = DocxTemplate(plantilla)
        doc.render(context)
        doc.save(output_path)

        return FileResponse(
            path=output_path,
            filename="resultado.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        return {"error": str(e)}
