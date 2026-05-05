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
        "servicio": "",
        "texto_completo": ""
    }

    texto_total = []

    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]

            texto_total.extend(cells)

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

            elif "servicio" in key:
                data["servicio"] = value

    data["texto_completo"] = " ".join(texto_total).lower()

    return data


# ------------------------
# DETECTAR SERVICIO POR TABLA (X)
# ------------------------

def detectar_servicio_desde_tabla(doc):
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip().lower() for c in row.cells]

            if len(cells) < 2:
                continue

            servicio = cells[0]
            marca = cells[1]

            if "x" in marca:

                if "vigilancia" in servicio:
                    return "vigilancia"

                elif "escolta" in servicio:
                    return "escolta"

                elif "monitoreo" in servicio:
                    return "monitoreo"

                elif "electronica" in servicio:
                    return "seguridad_electronica"

                elif "confiabilidad" in servicio:
                    return "confiabilidad"

                elif "eventos" in servicio:
                    return "seguridad_eventos"

    return None


# ------------------------
# DETECTAR SUBTIPO (TEXTO)
# ------------------------

def detectar_subtipo(data):
    texto = data["texto_completo"]

    subtipo = ""

    if "sin arma" in texto:
        subtipo = "sin_arma"

    elif "arma" in texto:
        subtipo = "armada"

    if "motorizado" in texto:
        subtipo = "motorizado"

    elif "conductor" in texto:
        subtipo = "conductor"

    elif "a pie" in texto:
        subtipo = "a_pie"

    if "12" in texto:
        subtipo += "_12h"

    return subtipo


# ------------------------
# SELECCIONAR PLANTILLA
# ------------------------

def seleccionar_plantilla(servicio_base, subtipo, data):
    texto = data["texto_completo"]

    # VIGILANCIA
    if servicio_base == "vigilancia":

        if "sin_arma" in subtipo:

            if "mensual" in texto and "fortalecimiento" in texto:
                return "plantillas/vigilancia_sin_arma_f_m.docx"

            elif "mensual" in texto:
                return "plantillas/vigilancia_sin_arma_m.docx"

            else:
                return "plantillas/vigilancia_sin_arma_e_12h.docx"

        else:

            if "mensual" in texto and "fortalecimiento" in texto:
                return "plantillas/vigilancia_armada_m_f.docx"

            elif "mensual" in texto:
                return "plantillas/vigilancia_armada_m.docx"

            else:
                return "plantillas/vigilancia_armada_e.docx"

    # ESCOLTA
    elif servicio_base == "escolta":

        if "motorizado" in subtipo:
            return "plantillas/escolta_motorizado.docx"

        elif "conductor" in subtipo:
            return "plantillas/escolta_conductor_ev.docx"

        elif "mensual" in texto:
            return "plantillas/escolta_mensual.docx"

        else:
            return "plantillas/escolta_a_pie.docx"

    # OTROS
    elif servicio_base == "monitoreo":
        return "plantillas/monitoreo.docx"

    elif servicio_base == "seguridad_eventos":
        return "plantillas/seguridad_en_eventos.docx"

    elif servicio_base == "seguridad_electronica":
        return "plantillas/seguridad_electronica.docx"

    elif servicio_base == "confiabilidad":
        return "plantillas/confiabilidad.docx"

    return "plantillas/monitoreo.docx"


# ------------------------
# ALCANCE DINÁMICO
# ------------------------

def generar_alcance(servicio_base, subtipo):

    if servicio_base == "vigilancia":

        if "sin_arma" in subtipo:
            return "Servicio de vigilancia sin arma con control de accesos y prevención de riesgos."

        else:
            return "Servicio de vigilancia armada con control de accesos y reacción ante incidentes."

    elif servicio_base == "escolta":

        if "motorizado" in subtipo:
            return "Servicio de escolta motorizado con reacción inmediata."

        elif "conductor" in subtipo:
            return "Servicio de escolta conductor con protección en desplazamientos."

        else:
            return "Servicio de escolta a pie."

    elif servicio_base == "monitoreo":
        return "Servicio de monitoreo electrónico."

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

        servicio_base = detectar_servicio_desde_tabla(doc)

        if not servicio_base:
            return {"error": "No se detectó servicio en la tabla (X)"}

        subtipo = detectar_subtipo(data)

        plantilla = seleccionar_plantilla(servicio_base, subtipo, data)

        if not os.path.exists(plantilla):
            return {"error": f"No existe la plantilla: {plantilla}"}

        alcance = generar_alcance(servicio_base, subtipo)

        context = {
            "consecutivo": obtener_consecutivo(),
            "fecha": obtener_fecha(),

            "tratamiento": obtener_tratamiento(data["nombre"], data["cargo"]),

            "nombre": data["nombre"].upper(),
            "nombre_simple": primer_nombre(data["nombre"]),

            "cargo": data["cargo"],
            "compania": data["compania"].upper(),

            "correo": data["correo"],
            "telefono": data["telefono"],
            "ciudad": data["ciudad"],

            "alcance": alcance
        }

        doc_tpl = DocxTemplate(plantilla)
        doc_tpl.render(context)
        doc_tpl.save(output_path)

        return FileResponse(
            path=output_path,
            filename="resultado.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        return {"error": str(e)}
