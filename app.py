from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import shutil
from docx import Document
from docxtpl import DocxTemplate
from datetime import datetime

app = FastAPI()


# =========================
# 🔹 FUNCIONES AUXILIARES
# =========================
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
            if len(row.cells) < 2:
                continue

            clave = row.cells[0].text.strip().lower()
            valor = row.cells[1].text.strip()

            if "contacto" in clave:
                data["nombre"] = valor

            elif "cargo" in clave:
                data["cargo"] = valor

            elif "compañía" in clave or "compania" in clave:
                data["compania"] = valor

            elif "mail" in clave or "correo" in clave:
                data["correo"] = valor

            elif "teléfono" in clave or "telefono" in clave:
                data["telefono"] = valor

            elif "ciudad" in clave:
                data["ciudad"] = valor

            elif "servicio" in clave:
                data["servicio"] = valor

    return data
    
def obtener_tratamiento(nombre, cargo):
    nombre = nombre.lower()
    cargo = cargo.lower()

    genero = "femenino" if nombre.endswith("a") else "masculino"

    cargos_doctor = ["doctor", "ingeniero", "director", "gerente"]

    if any(c in cargo for c in cargos_doctor):
        return "Doctora" if genero == "femenino" else "Doctor"
    else:
        return "Señora" if genero == "femenino" else "Señor"


def obtener_saludo(nombre):
    return "Distinguida" if nombre.lower().endswith("a") else "Distinguido"


def obtener_fecha():
    return datetime.now().strftime("%d de %B de %Y")


def generar_consecutivo():
    # 🔥 simple (puedes cambiar luego)
    return datetime.now().strftime("%Y%m%d%H%M")


def seleccionar_plantilla(texto):
    texto = texto.lower()

    if "escolta" in texto and "motorizado" in texto:
        return "plantillas/escolta_motorizado.docx"

    elif "escolta" in texto:
        return "plantillas/escolta_a_pie.docx"

    elif "vigilancia" in texto and "armada" in texto:
        return "plantillas/vigilancia_armada_e.docx"

    elif "monitoreo" in texto:
        return "plantillas/monitoreo.docx"

    elif "ciberseguridad" in texto:
        return "plantillas/capacitacion_ciberseguridad.docx"

    else:
        return "plantillas/confiabilidad.docx"


# =========================
# 🔹 ENDPOINTS
# =========================
@app.get("/")
def home():
    return {"status": "ok"}


@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):
    try:
        temp_path = "temp.docx"
        output_path = "resultado.docx"

        # Guardar archivo
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Leer documento
        doc = Document(temp_path)
        texto = "\n".join([p.text for p in doc.paragraphs])

        # Seleccionar plantilla
        plantilla = seleccionar_plantilla(texto)

        # 🔥 EXTRAER DATOS (básico - puedes mejorar)
        data = extraer_datos_tabla(doc)

        # fallback si algo no se detecta
        for k in data:
            if not data[k]:
                data[k] = "No especificado"
        texto_lower = texto.lower()

        if "escolta" in texto_lower:
            data["alcance"] = "Servicio de escolta con protección personal y acompañamiento según requerimiento."
        elif "vigilancia" in texto_lower:
            data["alcance"] = "Servicio de vigilancia física con control de accesos, rondas de seguridad y prevención."
        elif "monitoreo" in texto_lower:
            data["alcance"] = "Servicio de monitoreo remoto de sistemas de seguridad electrónica."
        else:
            data["alcance"] = "Servicio de seguridad adaptado a las necesidades del cliente."
        if "mensual" in texto_lower:
            data["modalidad"] = "Mensual"
        elif "eventual" in texto_lower:
            data["modalidad"] = "Eventual"
        elif "fortalecimiento" in texto_lower:
            data["modalidad"] = "Fortalecimiento"
        elif "valor agregado" in texto_lower:
            data["modalidad"] = "Valor agregado"
        else:
            data["modalidad"] = ""
        
        # Variables dinámicas
        consecutivo = generar_consecutivo()
        tratamiento = obtener_tratamiento(data["nombre"], data["cargo"])
        saludo = obtener_saludo(data["nombre"])
        fecha = obtener_fecha()

        # Generar documento
        doc = DocxTemplate(plantilla)

        context = {
            "consecutivo": consecutivo,
            "fecha": fecha,
            "tratamiento": tratamiento,
            "saludo": saludo,
            "nombre": data["nombre"],
            "cargo": data["cargo"],
            "compania": data["compania"],
            "correo": data["correo"],
            "telefono": data["telefono"],
            "ciudad": data["ciudad"],
            "alcance": data["alcance"]
        }

        doc.render(context)
        doc.save(output_path)

        return FileResponse(
            path=output_path,
            filename="resultado.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        return {"error": str(e)}
