from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from docx import Document
import shutil
import os
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
            celdas = [c.text.strip() for c in row.cells]

            for i in range(0, len(celdas), 2):
                if i + 1 >= len(celdas):
                    continue

                campo = celdas[i].lower()
                valor = celdas[i + 1].strip()

                if not valor:
                    continue

                if "contacto" in campo:
                    datos["nombre"] = limpiar_nombre(valor)

                elif "cargo" in campo:
                    datos["cargo"] = valor

                elif "compañ" in campo:
                    datos["compania"] = valor.upper()

                elif "mail" in campo or "correo" in campo:
                    datos["correo"] = valor

                elif "tel" in campo:
                    datos["telefono"] = valor

                elif "ciudad" in campo:
                    datos["ciudad"] = valor

    return datos

# =========================
# DETECTAR SERVICIO (CLAVE)
# =========================
def detectar_servicio(doc):
    texto = "\n".join([p.text.lower() for p in doc.paragraphs])

    # prioridad: tabla (X)
    for table in doc.tables:
        for row in table.rows:
            celdas = [c.text.strip().lower() for c in row.cells]

            if "vigilancia" in celdas and "x" in celdas:
                return "vigilancia"

            if "escoltas" in celdas and "x" in celdas:
                return "escolta"

            if "monitoreo" in celdas and "x" in celdas:
                return "monitoreo"

            if "confiabilidad" in celdas and "x" in celdas:
                return "confiabilidad"

    # fallback por texto
    if "vigilancia" in texto:
        return "vigilancia"
    if "escolta" in texto:
        return "escolta"
    if "monitoreo" in texto:
        return "monitoreo"

    return "vigilancia"

# =========================
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

# =========================
# SELECCIONAR PLANTILLA
# =========================
def seleccionar_plantilla(servicio, tipo):
    base = f"plantillas/{servicio}{tipo}.docx"

    if os.path.exists(base):
        return base

    # fallback seguro
    return "plantillas/vigilancia_sin_arma_m.docx"

# =========================
# REEMPLAZO SEGURO
# =========================
def reemplazar(doc, data):
    def reemplazar_texto(parrafo):
        texto = parrafo.text

        for key, val in data.items():
            if key in texto:
                texto = texto.replace(key, val)

        # 🔥 REEMPLAZO SEGURO SIN ROMPER XML
        if parrafo.text != texto:
            parrafo.text = texto

    for p in doc.paragraphs:
        reemplazar_texto(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    reemplazar_texto(p)
                    
# =========================
# API
# =========================
@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):
    try:
        temp = "entrada.docx"
        output = "resultado.docx"

        with open(temp, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        file.file.close()

        doc = Document(temp)

        datos = extraer_datos(doc)
        servicio = detectar_servicio(doc)
        tipo = detectar_tipo(doc)

        print("DATOS:", datos)
        print("SERVICIO:", servicio)
        print("TIPO:", tipo)

        plantilla = seleccionar_plantilla(servicio, tipo)

        doc_final = Document(plantilla)

        tratamiento = obtener_tratamiento(datos["cargo"])

        reemplazos = {
            "{{consecutivo}}": datetime.now().strftime("%Y%m%d%H%M"),
            "{{fecha}}": fecha_es(),
            "{{tratamiento}}": tratamiento,
            "{{nombre}}": datos["nombre"],
            "{{cargo}}": datos["cargo"],
            "{{compania}}": datos["compania"],
            "{{correo}}": datos["correo"],
            "{{telefono}}": datos["telefono"],
            "{{ciudad}}": datos["ciudad"],
            "{{alcance}}": datos["ciudad"],
            "{{saludo}}": "Estimado",
            "{{nombre_corto}}": primer_nombre(datos["nombre"]),
        }

        reemplazar(doc_final, reemplazos)

        if os.path.exists(output):
            os.remove(output)

        doc_final.save(output)

        return FileResponse(output, filename="resultado.docx")

    except Exception as e:
        return {"error": str(e)}
