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
    return (
        nombre.upper()
        .replace("SR.", "")
        .replace("SRA.", "")
        .replace("DR.", "")
        .replace("DRA.", "")
        .replace("E-MAIL", "")
        .replace("EMAIL", "")
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
        "direccion": ""
    }

    def limpiar(txt):
        return (txt or "").strip()

    def normalizar_campo(txt):
        return limpiar(txt).lower().replace(":", "").replace("  ", " ")

    for table in doc.tables:
        for row in table.rows:
            celdas = [limpiar(c.text) for c in row.cells]

            # Procesar en pares: (0,1) y (2,3)
            pares = []
            if len(celdas) >= 2:
                pares.append((celdas[0], celdas[1]))
            if len(celdas) >= 4:
                pares.append((celdas[2], celdas[3]))

            for campo_raw, valor_raw in pares:
                campo = normalizar_campo(campo_raw)
                valor = limpiar(valor_raw)

                if not valor:
                    continue

                # 🔥 MAPEO EXACTO A TU FORMATO
                if "contacto" in campo:
                    datos["nombre"] = limpiar_nombre(valor)

                elif campo == "cargo":
                    datos["cargo"] = valor

                elif "compañ" in campo or "compania" in campo:
                    datos["compania"] = valor.upper()

                elif "e-mail" in campo or "email" in campo or "correo" in campo:
                    datos["correo"] = valor

                elif "tel" in campo:
                    datos["telefono"] = valor

                elif "ciudad" in campo:
                    datos["ciudad"] = valor

                elif "dirección" in campo or "direccion" in campo:
                    datos["direccion"] = valor

    return datos
    
# =========================
# DETECTAR SERVICIO (CLAVE)
# =========================
def detectar_servicio(doc):
    for table in doc.tables:
        for row in table.rows:
            textos = [c.text.strip().lower() for c in row.cells]

            if any("vigilancia" in t for t in textos) and "x" in textos:
                return "vigilancia"

            if any("escolta" in t for t in textos) and "x" in textos:
                return "escolta"

            if any("confiabilidad" in t for t in textos) and "x" in textos:
                return "confiabilidad"

            if any("seguridad electronica" in t for t in textos) and "x" in textos:
                return "electronica"

            if any("monitoreo" in t for t in textos) and "x" in textos:
                return "monitoreo"

    return "vigilancia"  # fallback
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
    
def detectar_detalle(doc):
    texto = "\n".join([p.text.lower() for p in doc.paragraphs])

    if "armada" in texto:
        return "armada"
    if "sin arma" in texto:
        return "sin_arma"
    if "escolta" in texto:
        return "escolta"

    return "sin_arma"  # fallback seguro
        
def detectar_modalidad(doc):
    texto = "\n".join([p.text.lower() for p in doc.paragraphs])

    if "mensual" in texto:
        return "m"
    if "festivo" in texto or "fin de semana" in texto:
        return "e"
    if "fortalecimiento" in texto:
        return "f"

    return "m"
# =========================
# SELECCIONAR PLANTILLA
# =========================
def seleccionar_plantilla(servicio, detalle, modalidad):
    if servicio == "vigilancia":
        if detalle == "armada":
            if modalidad == "m":
                return "plantillas/vigilancia_armada_m.docx"
            if modalidad == "f":
                return "plantillas/vigilancia_armada_m_f.docx"
            return "plantillas/vigilancia_armada_e.docx"

        if detalle == "sin_arma":
            if modalidad == "m":
                return "plantillas/vigilancia_sin_arma_m.docx"
            if modalidad == "f":
                return "plantillas/vigilancia_sin_arma_f_m.docx"
            return "plantillas/vigilancia_sin_arma_e_12h.docx"

    if servicio == "escolta":
        if detalle == "motorizado":
            return "plantillas/escolta_motorizado.docx"
        return "plantillas/escolta_a_pie.docx"

    if servicio == "confiabilidad":
        return "plantillas/confiabilidad.docx"

    if servicio == "electronica":
        return "plantillas/seguridad_electronica.docx"

    if servicio == "monitoreo":
        return "plantillas/monitoreo.docx"

    return "plantillas/vigilancia_sin_arma_m.docx"
    
# =========================
# REEMPLAZO SEGURO
# =========================

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

        # leer levantamiento
        doc = Document(temp)

        datos = extraer_datos(doc)

        print("DATOS EXTRAIDOS:", datos)

        servicio = detectar_servicio(doc)
        detalle = detectar_detalle(doc)
        modalidad = detectar_modalidad(doc)

        print("SERVICIO:", servicio)
        print("DETALLE:", detalle)
        print("MODALIDAD:", modalidad)

        # seleccionar plantilla
        plantilla = seleccionar_plantilla(servicio, detalle, modalidad)

        # 🔥 crear reemplazos (ESTO TE FALTABA)
        reemplazos = {
            "consecutivo": datetime.now().strftime("%Y%m%d%H%M"),
            "fecha": fecha_es(),
            "nombre": datos.get("nombre", ""),
            "cargo": datos.get("cargo", ""),
            "compania": datos.get("compania", ""),
            "correo": datos.get("correo", ""),
            "telefono": datos.get("telefono", ""),
            "ciudad": datos.get("ciudad", ""),
            "alcance": datos.get("ciudad", ""),
            "tratamiento": obtener_tratamiento(datos.get("cargo", "")),
            "saludo": "Estimado",
            "nombre_corto": datos.get("nombre", "").split(" ")[0].capitalize() if datos.get("nombre") else ""
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
