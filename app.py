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
        "direccion": ""
    }

    for table in doc.tables:
        for row in table.rows:

            celdas = [c.text.strip() for c in row.cells]

            # recorrer en pares (campo - valor)
            for i in range(0, len(celdas), 2):

                if i + 1 >= len(celdas):
                    continue

                campo = celdas[i].lower()
                valor = celdas[i + 1].strip()

                if "contacto" in campo:
                    datos["nombre"] = valor.replace("Sr.", "").strip()

                elif "cargo" in campo:
                    datos["cargo"] = valor

                elif "compañía" in campo or "compania" in campo:
                    datos["compania"] = valor

                elif "e- mail" in campo or "correo" in campo:
                    datos["correo"] = valor

                elif "teléfono" in campo or "telefono" in campo:
                    datos["telefono"] = valor

                elif "ciudad" in campo:
                    datos["ciudad"] = valor

                elif "dirección" in campo or "direccion" in campo:
                    datos["direccion"] = valor

    return datos
    
# -------------------------
# DETECTAR SERVICIO (X)
# -------------------------
def detectar_servicio(doc):
    servicios_detectados = []

    for table in doc.tables:
        for row in table.rows:

            celdas = [c.text.strip().lower() for c in row.cells if c.text.strip()]

            # 🔒 Si la fila está vacía → ignorar
            if not celdas:
                continue

            texto_fila = " ".join(celdas)

            # 🔍 detectar marca (X o variantes)
            tiene_x = any(
                c in ["x", "X", "✔", "✓"] or c.strip() == "x"
                for c in celdas
            )

            if not tiene_x:
                continue

            # 🔥 DETECCIÓN REAL
            if "vigilancia" in texto_fila:
                servicios_detectados.append("vigilancia")

            elif "seguridad electronica" in texto_fila:
                servicios_detectados.append("electronica")

            elif "confiabilidad" in texto_fila:
                servicios_detectados.append("confiabilidad")

            elif "escolta" in texto_fila:
                servicios_detectados.append("escolta")

            elif "monitoreo" in texto_fila:
                servicios_detectados.append("monitoreo")

            elif "eventos" in texto_fila:
                servicios_detectados.append("eventos")

    # 🧠 Prioridad (por si hay más de uno marcado)
    prioridad = [
        "vigilancia",
        "escolta",
        "electronica",
        "confiabilidad",
        "monitoreo",
        "eventos"
    ]

    for p in prioridad:
        if p in servicios_detectados:
            return p

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
