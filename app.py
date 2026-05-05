from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from docx import Document
from docxtpl import DocxTemplate
import shutil
import os
from datetime import datetime

app = FastAPI()

# ----------------------------
# FORMATEOS
# ----------------------------

def fecha_espanol():
    meses = [
        "enero","febrero","marzo","abril","mayo","junio",
        "julio","agosto","septiembre","octubre","noviembre","diciembre"
    ]
    now = datetime.now()
    return f"{now.day} de {meses[now.month-1]} de {now.year}"

def limpiar_nombre(nombre):
    return nombre.replace("Sr.", "").replace("Sra.", "").strip()

def primer_nombre(nombre):
    return nombre.split()[0].capitalize()

def determinar_tratamiento(nombre, cargo):
    cargo = cargo.lower()
    if any(x in cargo for x in ["gerente", "director", "presidente"]):
        return "Doctor"
    return "Señor"

# ----------------------------
# EXTRACCIÓN ROBUSTA
# ----------------------------

def extraer_datos(doc):
    datos = {}

    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]

            # ⚠️ SI NO HAY SUFICIENTES CELDAS → SALTA
            if len(cells) < 2:
                continue

            texto_fila = " ".join(cells).lower()

            try:
                if "contacto" in texto_fila:
                    datos["nombre"] = limpiar_nombre(cells[-1])

                elif "cargo" in texto_fila:
                    datos["cargo"] = cells[-1]

                elif "compañía" in texto_fila:
                    datos["compania"] = cells[-1].upper()

                elif "teléfono" in texto_fila:
                    datos["telefono"] = cells[-1]

                elif "ciudad" in texto_fila:
                    datos["ciudad"] = cells[-1]

                elif "mail" in texto_fila or "correo" in texto_fila:
                    datos["correo"] = cells[-1]

                elif "tipo de servicio" in texto_fila:
                    datos["servicio"] = cells[-1].lower()

                elif "tiempo de servicio" in texto_fila:
                    datos["modalidad"] = cells[-1].lower()

            except Exception:
                # evita que una fila dañada tumbe todo
                continue

    return datos
# ----------------------------
# DETECCIÓN DE DETALLES
# ----------------------------

def detectar_detalles(doc):
    texto = "\n".join([p.text.lower() for p in doc.paragraphs])

    resultado = {
        "arma": "sin_arma",
        "fortalecimiento": False
    }

    if "arma" in texto:
        resultado["arma"] = "armada"

    if "fortalecimiento" in texto:
        resultado["fortalecimiento"] = True

    return resultado

# ----------------------------
# SELECCIÓN DE PLANTILLA
# ----------------------------

def seleccionar_plantilla(datos, extra):
    servicio = datos.get("servicio", "")

    if "vigilancia" in servicio:
        arma = extra["arma"]
        modalidad = "m" if "mensual" in datos.get("modalidad","") else "e"
        f = "_f" if extra["fortalecimiento"] else ""

        ruta = f"plantillas/vigilancia_{arma}_{modalidad}{f}.docx"
        return ruta

    if "escolta" in servicio:
        return "plantillas/escolta_mensual.docx"

    if "confiabilidad" in servicio:
        return "plantillas/confiabilidad.docx"

    if "monitoreo" in servicio:
        return "plantillas/monitoreo.docx"

    if "electronica" in servicio:
        return "plantillas/seguridad_electronica.docx"

    if "eventos" in servicio:
        return "plantillas/seguridad_en_eventos.docx"

    return None

# ----------------------------
# ENDPOINT
# ----------------------------

@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):
    try:
        temp_path = "temp.docx"
        output_path = "resultado.docx"

        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        doc = Document(temp_path)

        datos = extraer_datos(doc)
        extra = detectar_detalles(doc)

        plantilla = seleccionar_plantilla(datos, extra)

        if not plantilla or not os.path.exists(plantilla):
            return {"error": f"No existe plantilla: {plantilla}"}

        tratamiento = determinar_tratamiento(
            datos.get("nombre", ""),
            datos.get("cargo", "")
        )

        contexto = {
            "consecutivo": datetime.now().strftime("%Y%m%d%H%M"),
            "fecha": fecha_espanol(),

            "tratamiento": tratamiento,

            # Nombre completo en MAYÚSCULA
            "nombre": datos.get("nombre", "").upper(),

            # Solo primer nombre capitalizado
            "nombre_simple": primer_nombre(datos.get("nombre", "")),

            "cargo": datos.get("cargo", ""),
            "compania": datos.get("compania", ""),
            "correo": datos.get("correo", "No especificado"),
            "telefono": datos.get("telefono", "No especificado"),

            # 🔥 CLAVE: alcance correcto
            "alcance": datos.get("ciudad") or "Bogotá"
        }

        doc_tpl = DocxTemplate(plantilla)
        doc_tpl.render(contexto)
        doc_tpl.save(output_path)

        return FileResponse(
            output_path,
            filename="resultado.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        return {"error": str(e)}
