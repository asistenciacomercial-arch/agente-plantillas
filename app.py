from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from docx import Document
import shutil
import os
from datetime import datetime

app = FastAPI()

# =========================
# EXTRAER DATOS BIEN (TABLA)
# =========================
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

                elif "compañía" in campo or "compania" in campo:
                    datos["compania"] = valor.upper()

                elif "mail" in campo or "correo" in campo:
                    datos["correo"] = valor

                elif "teléfono" in campo or "telefono" in campo:
                    datos["telefono"] = valor

                elif "ciudad" in campo:
                    datos["ciudad"] = valor

                elif "dirección" in campo or "direccion" in campo:
                    datos["direccion"] = valor

    return datos


# =========================
# LIMPIAR NOMBRE
# =========================
def limpiar_nombre(nombre):
    return (
        nombre.replace("Sr.", "")
        .replace("Sra.", "")
        .replace("Dr.", "")
        .replace("Dra.", "")
        .strip()
        .upper()
    )


# =========================
# PRIMER NOMBRE
# =========================
def primer_nombre(nombre):
    return nombre.split()[0].capitalize() if nombre else ""


# =========================
# TRATAMIENTO
# =========================
def obtener_tratamiento(cargo):
    cargo = cargo.lower()

    if "presidente" in cargo or "gerente" in cargo or "director" in cargo:
        return "Doctor"
    return "Señor"


# =========================
# SALUDO
# =========================
def obtener_saludo(tratamiento):
    if tratamiento == "Doctor":
        return "Estimado"
    return "Estimado"


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
# REEMPLAZO SIN DAÑAR FORMATO
# =========================
def reemplazar_en_doc(doc, data):
    def reemplazar_texto(parrafo):
        texto = parrafo.text
        for key, val in data.items():
            if key in texto:
                texto = texto.replace(key, val)

        # 🔥 BORRAR RUNS SIN ROMPER DOCX
        if parrafo.text != texto:
            for run in parrafo.runs:
                run.text = ""
            parrafo.runs[0].text = texto

    # párrafos normales
    for p in doc.paragraphs:
        reemplazar_texto(p)

    # tablas
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
        temp = "temp.docx"
        shutil.copyfileobj(file.file, open(temp, "wb"))

        doc = Document(temp)

        datos = extraer_datos(doc)

        # VALIDACIÓN
        if not datos["nombre"]:
            return {"error": "No se pudo extraer el nombre"}
        if not datos["compania"]:
            return {"error": "No se pudo extraer la compañía"}

        tratamiento = obtener_tratamiento(datos["cargo"])
        saludo = obtener_saludo(tratamiento)

        plantilla = "plantillas/vigilancia_sin_arma_m.docx"

        doc_final = Document(plantilla)

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
            "{{direccion}}": datos["direccion"],
            "{{saludo}}": saludo,
            "{{primer_nombre}}": primer_nombre(datos["nombre"]),
            "{{alcance}}": datos["ciudad"],
        }

        reemplazar_en_doc(doc_final, reemplazos)

        output = "resultado.docx"
        doc_final.save(output)

        return FileResponse(output, filename="resultado.docx")

    except Exception as e:
        return {"error": str(e)}
