from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from docx import Document
import shutil
import os
from datetime import datetime

app = FastAPI()


# =========================
# 🧠 UTILIDADES
# =========================

def limpiar(texto):
    return texto.strip().replace("\n", " ")


def limpiar_nombre(nombre):
    return nombre.replace("Sr.", "").replace("Sra.", "").replace("Dr.", "").strip()


def primer_nombre(nombre):
    return nombre.split()[0].capitalize()


def obtener_tratamiento(nombre, cargo):
    nombre = nombre.lower()
    cargo = (cargo or "").lower()

    if "gerente" in cargo or "director" in cargo or "doctor" in cargo:
        return "Doctor"

    if nombre.endswith("a"):
        return "Señora"

    return "Señor"


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


def obtener_consecutivo():
    return datetime.now().strftime("%Y%m%d%H%M")


# =========================
# 🧠 EXTRACCIÓN INTELIGENTE
# =========================

def extraer_datos(doc):
    datos = {}

    for table in doc.tables:
        for row in table.rows:
            cells = [limpiar(c.text) for c in row.cells if limpiar(c.text)]

            for i in range(len(cells)):
                texto = cells[i].lower()

                if "contacto" in texto and i+1 < len(cells):
                    datos["nombre"] = limpiar_nombre(cells[i+1])

                if "cargo" in texto and i+1 < len(cells):
                    datos["cargo"] = cells[i+1]

                if "compañ" in texto and i+1 < len(cells):
                    datos["compania"] = cells[i+1].upper()

                if "mail" in texto or "correo" in texto:
                    if i+1 < len(cells):
                        datos["correo"] = cells[i+1]

                if "tel" in texto and i+1 < len(cells):
                    datos["telefono"] = cells[i+1]

                if "ciudad" in texto and i+1 < len(cells):
                    datos["ciudad"] = cells[i+1]

                if "tipo de servicio" in texto and i+1 < len(cells):
                    datos["servicio"] = cells[i+1].lower()

                if "tiempo de servicio" in texto and i+1 < len(cells):
                    datos["modalidad"] = cells[i+1].lower()

    return datos


# =========================
# 🧠 DETECTAR PLANTILLA
# =========================

def detectar_plantilla(datos):
    servicio = datos.get("servicio", "")
    modalidad = datos.get("modalidad", "")

    if "vigilancia" in servicio:
        if "sin arma" in servicio:
            if "mensual" in modalidad:
                return "plantillas/vigilancia_sin_arma_m.docx"
            else:
                return "plantillas/vigilancia_sin_arma_e_12h.docx"
        else:
            if "mensual" in modalidad:
                return "plantillas/vigilancia_armada_m.docx"
            else:
                return "plantillas/vigilancia_armada_e.docx"

    if "escolta" in servicio:
        if "motorizado" in servicio:
            return "plantillas/escolta_motorizado.docx"
        if "conductor" in servicio:
            return "plantillas/escolta_conductor_ev.docx"
        return "plantillas/escolta_a_pie.docx"

    if "confiabilidad" in servicio:
        return "plantillas/confiabilidad.docx"

    if "electronica" in servicio:
        return "plantillas/seguridad_electronica.docx"

    if "evento" in servicio:
        return "plantillas/seguridad_en_eventos.docx"

    return None


# =========================
# 🧠 REEMPLAZAR VARIABLES
# =========================

def reemplazar(doc, contexto):
    for p in doc.paragraphs:
        for k, v in contexto.items():
            if f"{{{{{k}}}}}" in p.text:
                p.text = p.text.replace(f"{{{{{k}}}}}", str(v))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for k, v in contexto.items():
                    if f"{{{{{k}}}}}" in cell.text:
                        cell.text = cell.text.replace(f"{{{{{k}}}}}", str(v))


# =========================
# 🚀 ENDPOINT
# =========================

@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):
    try:
        temp_path = "temp.docx"
        output_path = "resultado.docx"

        # Guardar archivo
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        doc = Document(temp_path)

        # EXTRAER DATOS
        datos = extraer_datos(doc)

        # VALIDAR
        if not datos.get("nombre"):
            return {"error": "No se pudo extraer el nombre"}

        # PLANTILLA
        plantilla = detectar_plantilla(datos)
        if not plantilla:
            return {"error": "No se detectó plantilla"}

        if not os.path.exists(plantilla):
            return {"error": f"No existe la plantilla: {plantilla}"}

        # FORMATEO
        nombre = datos.get("nombre", "CLIENTE")
        cargo = datos.get("cargo", "")

        contexto = {
            "consecutivo": obtener_consecutivo(),
            "fecha completa actual": obtener_fecha(),

            "tratamiento": obtener_tratamiento(nombre, cargo),
            "nombre": nombre.upper(),
            "nombre_corto": primer_nombre(nombre),

            "cargo": cargo,
            "compania": datos.get("compania", "").upper(),
            "correo": datos.get("correo", ""),
            "telefono": datos.get("telefono", ""),
            "ciudad": datos.get("ciudad", ""),
            "alcance": datos.get("ciudad", "")
        }

        # GENERAR DOCUMENTO
        doc_out = Document(plantilla)
        reemplazar(doc_out, contexto)
        doc_out.save(output_path)

        return FileResponse(
            path=output_path,
            filename="resultado.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        return {"error": str(e)}
