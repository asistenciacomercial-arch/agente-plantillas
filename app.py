from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import shutil
import os
from docx import Document
from docxtpl import DocxTemplate
from datetime import datetime

# Google Sheets
import gspread
from oauth2client.service_account import ServiceAccountCredentials

app = FastAPI()


# =========================
# 🔹 GOOGLE SHEETS
# =========================
def obtener_consecutivo():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name("credenciales.json", scope)
    client = gspread.authorize(creds)

    sheet = client.open("TU_SHEET").sheet1

    # 🔥 SOLO LEE (turnero)
    consecutivo = sheet.acell("A1").value

    return consecutivo if consecutivo else "SIN-NUMERO"


# =========================
# 🔹 SALUDO Y TRATAMIENTO
# =========================
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


# =========================
# 🔹 FECHA
# =========================
def obtener_fecha():
    return datetime.now().strftime("%d de %B de %Y")


# =========================
# 🔹 DETECTAR PLANTILLA
# =========================
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

        # 🔥 SELECCIONAR PLANTILLA AUTOMÁTICA
        plantilla = seleccionar_plantilla(texto)

        # 🔥 EXTRAER DATOS (simple - puedes mejorar luego)
        data = {
            "nombre": "Cliente Ejemplo",
            "cargo": "Director",
            "compania": "Empresa XYZ",
            "correo": "correo@email.com",
            "telefono": "123456789",
            "ciudad": "Bogotá",
            "alcance": texto[:500]  # ejemplo
        }

        # 🔥 VARIABLES DINÁMICAS
        consecutivo = obtener_consecutivo()
        tratamiento = obtener_tratamiento(data["nombre"], data["cargo"])
        saludo = obtener_saludo(data["nombre"])
        fecha = obtener_fecha()

        # 🔥 RENDER DOCX
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
