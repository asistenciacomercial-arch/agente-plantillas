from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from docx import Document
from docxtpl import DocxTemplate
from io import BytesIO
import shutil
from datetime import datetime

app = FastAPI()

# =========================
# LIMPIEZA
# =========================
def limpiar_nombre(nombre):

    eliminar = [
        "SR.",
        "SRA.",
        "DR.",
        "DRA."
    ]

    nombre = nombre.upper()

    for e in eliminar:
        nombre = nombre.replace(e, "")

    return " ".join(nombre.split()).strip()
    
def obtener_tratamiento(cargo):
    cargo = (cargo or "").lower()
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
# EXTRAER DATOS (FORMATO REAL)
# =========================
def extraer_datos(doc):

    datos = {
        "nombre": "",
        "primer_nombre": "",
        "cargo": "",
        "compania": "",
        "correo": "",
        "telefono": "",
        "direccion": "",
        "ciudad": "Bogotá"
    }

    for table in doc.tables:

        for row in table.rows:

            cells = [c.text.strip() for c in row.cells]

            if len(cells) < 4:
                continue

            # COMPAÑIA + DIRECCION
            if "Compañía" in cells[0]:

                datos["compania"] = cells[1].upper()
                datos["direccion"] = cells[3]

            # CONTACTO + EMAIL
           if "Contacto" in cells[0]:

                nombre = limpiar_nombre(cells[1])
            
                datos["nombre"] = nombre
                datos["primer_nombre"] = nombre.split()[0].capitalize()
            
            # 🔥 BUSCAR EMAIL EN CUALQUIER CELDA
            for c in cells:
            
                if "@" in c:
                    datos["correo"] = c.strip()

            # CARGO + TELEFONO
            if "Cargo" in cells[0]:

                datos["cargo"] = cells[1]
                datos["telefono"] = cells[3]

    print(datos)

    return datos
    
# =========================
# DETECCIÓN SERVICIO
# =========================
def detectar_servicio(doc):
    for table in doc.tables:
        for row in table.rows:
            textos = [c.text.strip().lower() for c in row.cells]

            if "vigilancia" in textos and "x" in textos:
                return "vigilancia"

    return "vigilancia"

def detectar_detalle(doc):
    texto = "\n".join([p.text.lower() for p in doc.paragraphs])

    if "armada" in texto:
        return "armada"
    if "sin arma" in texto:
        return "sin_arma"

    return "sin_arma"

def detectar_modalidad(doc):
    texto = "\n".join([p.text.lower() for p in doc.paragraphs])

    if "mensual" in texto:
        return "m"
    return "m"

# =========================
# PLANTILLA
# =========================
def seleccionar_plantilla(servicio, detalle, modalidad):
    return "plantillas/vigilancia_sin_arma_m.docx"

# =========================
# API
# =========================
@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):
    try:
        temp = "entrada.docx"

        # guardar archivo
        with open(temp, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        file.file.close()

        # leer documento
        doc = Document(temp)

        datos = extraer_datos(doc)

        print("DATOS:", datos)

        # validar nombre
        nombre = datos.get("nombre")
        if not nombre:
            nombre = "CLIENTE"

        tratamiento = obtener_tratamiento(datos.get("cargo", ""))

        servicio = detectar_servicio(doc)
        detalle = detectar_detalle(doc)
        modalidad = detectar_modalidad(doc)

        plantilla = seleccionar_plantilla(servicio, detalle, modalidad)

        # =========================
        # REEMPLAZOS CORRECTOS
        # =========================
        reemplazos = {
        
            "consecutivo": datetime.now().strftime("%Y%m%d%H%M"),
        
            "fecha": fecha_es(),
        
            "tratamiento": obtener_tratamiento(datos["cargo"]),
        
            "nombre": datos["nombre"],
        
            "primer_nombre": datos["primer_nombre"],
        
            "cargo": datos["cargo"],
        
            "compania": datos["compania"],
        
            "correo": datos["correo"],
        
            "telefono": datos["telefono"],
        
            "direccion": datos["direccion"],
        
            "ciudad": datos["ciudad"],
        
            # saludo abajo
            "saludo": f"Estimado {obtener_tratamiento(datos['cargo']).lower()} {datos['primer_nombre']}"
        }

        # generar documento
        doc_tpl = DocxTemplate(plantilla)
        doc_tpl.render(reemplazos)

        buffer = BytesIO()
        doc_tpl.save(buffer)
        buffer.seek(0)

        return StreamingResponse(
            buffer,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": "attachment; filename=resultado.docx"
            }
        )

    except Exception as e:
        return {"error": str(e)}
