from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from docx import Document
import shutil
import os
from datetime import datetime

app = FastAPI()


# =========================
# EXTRAER DATOS
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
# REEMPLAZO SEGURO (NO ROMPE DOCX)
# =========================
def reemplazar(doc, data):
    def procesar(p):
        texto = p.text
        nuevo = texto

        for k, v in data.items():
            nuevo = nuevo.replace(k, v)

        if nuevo != texto:
            # 🔥 esto evita corrupción
            p.clear()
            p.add_run(nuevo)

    for p in doc.paragraphs:
        procesar(p)

    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    procesar(p)


# =========================
# API
# =========================
@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):
    try:
        temp = "entrada.docx"
        output = "resultado.docx"

        # 🔥 guardar archivo correctamente (IMPORTANTE)
        with open(temp, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        file.file.close()

        # 🔥 leer documento
        doc = Document(temp)

        datos = extraer_datos(doc)

        # DEBUG (te ayuda a ver si está leyendo bien)
        print("DATOS EXTRAIDOS:", datos)

        plantilla = "plantillas/vigilancia_sin_arma_m.docx"

        doc_final = Document(plantilla)

        reemplazos = {
            "{{consecutivo}}": datetime.now().strftime("%Y%m%d%H%M"),
            "{{fecha}}": fecha_es(),
            "{{nombre}}": datos["nombre"],
            "{{cargo}}": datos["cargo"],
            "{{compania}}": datos["compania"],
            "{{correo}}": datos["correo"],
            "{{telefono}}": datos["telefono"],
            "{{ciudad}}": datos["ciudad"],
            "{{alcance}}": datos["ciudad"],
        }

        reemplazar(doc_final, reemplazos)

        # 🔥 eliminar anterior si existe
        if os.path.exists(output):
            os.remove(output)

        doc_final.save(output)

        return FileResponse(
            output,
            filename="resultado.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        return {"error": str(e)}
        
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
    cargo = cargo.lower()

    if "gerente" in cargo or "director" in cargo or "presidente" in cargo:
        return "Doctor"
    return "Señor"
    tratamiento = obtener_tratamiento(datos["cargo"])

    reemplazos = {
        "{{tratamiento}}": tratamiento,
    }
    def limpiar_saludo(nombre):
    return nombre.split()[0].capitalize()
    "{{saludo}}": "Estimado",
    "{{nombre_corto}}": limpiar_saludo(datos["nombre"]),
