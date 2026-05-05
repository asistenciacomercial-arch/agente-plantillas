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
    return (texto or "").strip().replace("\n", " ")

def limpiar_nombre(nombre):
    return limpiar(nombre).replace("Sr.", "").replace("Sra.", "").replace("Dr.", "").strip()

def primer_nombre(nombre):
    return limpiar(nombre).split()[0].capitalize() if nombre else ""

def obtener_tratamiento(nombre, cargo):
    nombre = (nombre or "").lower()
    cargo = (cargo or "").lower()

    if any(x in cargo for x in ["gerente", "director", "doctor"]):
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
# 🧠 EXTRACCIÓN DE DATOS
# =========================

def extraer_datos(doc):
    datos = {}

    for table in doc.tables:
        for row in table.rows:
            cells = [limpiar(c.text) for c in row.cells if limpiar(c.text)]

            for i in range(len(cells)):
                texto = cells[i].lower()

                # CONTACTO → nombre
                if "contacto" in texto and i+1 < len(cells):
                    datos["nombre"] = limpiar_nombre(cells[i+1])

                # CARGO
                if texto.strip() == "cargo" and i+1 < len(cells):
                    datos["cargo"] = cells[i+1]

                # COMPAÑIA
                if "compañ" in texto and i+1 < len(cells):
                    datos["compania"] = cells[i+1].upper()

                # CORREO (evita agarrar texto del párrafo)
                if texto.strip() in ["correo", "e-mail", "email"] and i+1 < len(cells):
                    datos["correo"] = cells[i+1]

                # TELÉFONO
                if texto.strip() in ["telefono", "teléfono", "tel"] and i+1 < len(cells):
                    datos["telefono"] = cells[i+1]

                # CIUDAD
                if "ciudad" in texto and i+1 < len(cells):
                    datos["ciudad"] = cells[i+1]

                # SERVICIO
                if "tipo de servicio" in texto and i+1 < len(cells):
                    datos["servicio"] = cells[i+1].lower()

                # MODALIDAD
                if "tiempo de servicio" in texto and i+1 < len(cells):
                    datos["modalidad"] = cells[i+1].lower()

    # Fallbacks
    datos.setdefault("correo", "")
    datos.setdefault("telefono", "")
    datos.setdefault("cargo", "")
    datos.setdefault("compania", "")
    datos.setdefault("ciudad", "Bogotá")

    return datos

# =========================
# 🧠 DETECTAR SERVICIO POR "X"
# =========================

def detectar_servicio_por_x(doc):
    for table in doc.tables:
        for row in table.rows:
            cells = [limpiar(c.text).lower() for c in row.cells]

            if "x" in cells:
                idx = cells.index("x")
                if idx > 0:
                    servicio = cells[idx-1]

                    if "vigilancia" in servicio:
                        return "vigilancia"
                    if "escolta" in servicio:
                        return "escolta"
                    if "confiabilidad" in servicio:
                        return "confiabilidad"
                    if "electronica" in servicio:
                        return "seguridad_electronica"
                    if "evento" in servicio:
                        return "evento"

    return None

# =========================
# 🧠 DETECTAR PLANTILLA
# =========================

def detectar_plantilla(datos, servicio_x=None):
    servicio = datos.get("servicio", "")
    modalidad = datos.get("modalidad", "")

    # usar detección por X si existe
    if servicio_x:
        servicio = servicio_x

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
# 🧠 REEMPLAZO FLEXIBLE
# =========================

def reemplazar(doc, contexto):
    alias = {
        "fecha": contexto.get("fecha completa actual", ""),
        "fecha completa actual": contexto.get("fecha completa actual", ""),

        "Cargo": contexto.get("cargo", ""),
        "cargo": contexto.get("cargo", ""),

        "nombre": contexto.get("nombre", ""),
        "Nombre": contexto.get("nombre", ""),

        "correo": contexto.get("correo", ""),
        "telefono": contexto.get("telefono", ""),
        "ciudad": contexto.get("ciudad", ""),
        "compania": contexto.get("compania", ""),
        "alcance": contexto.get("alcance", ""),
        "consecutivo": contexto.get("consecutivo", ""),

        "tratamiento": contexto.get("tratamiento", ""),
        "saludo": f"{contexto.get('tratamiento','')} {contexto.get('nombre_corto','')}",
    }

    def reemplazar_texto(texto):
        for k, v in alias.items():
            texto = texto.replace(f"{{{{{k}}}}}", str(v))
        return texto

    # 🔹 PÁRRAFOS (CLAVE)
    for p in doc.paragraphs:
        texto_original = p.text
        texto_nuevo = reemplazar_texto(texto_original)

        if texto_original != texto_nuevo:
            p.clear()  # borra runs pero mantiene formato general
            p.add_run(texto_nuevo)

    # 🔹 TABLAS
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    texto_original = p.text
                    texto_nuevo = reemplazar_texto(texto_original)

                    if texto_original != texto_nuevo:
                        p.clear()
                        p.add_run(texto_nuevo)

# =========================
# 🚀 ENDPOINT
# =========================

@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):
    try:
        temp_path = "temp.docx"
        output_path = "resultado.docx"

        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        doc = Document(temp_path)

        datos = extraer_datos(doc)

        # detectar servicio por X (más confiable)
        servicio_x = detectar_servicio_por_x(doc)

        plantilla = detectar_plantilla(datos, servicio_x)

        if not plantilla:
            return {"error": "No se detectó plantilla"}

        if not os.path.exists(plantilla):
            return {"error": f"No existe la plantilla: {plantilla}"}

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
