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
# VALIDAR NOMBRE
# =========================
def es_nombre_valido(texto):
    palabras = texto.split()

    if len(palabras) < 2:
        return False

    bloqueados = [
        "LEVANTAMIENTO",
        "NECESIDADES",
        "PROPUESTA",
        "SERVICIO",
        "VIGILANCIA",
        "SEGURIDAD"
    ]

    for b in bloqueados:
        if b in texto:
            return False

    if not texto.isupper():
        return False

    return True
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

    # =========================
    # 1. TABLA PRINCIPAL (CONTACTO)
    # =========================
    if doc.tables:
        tabla = doc.tables[0]  # la primera tabla siempre es contacto

        for row in tabla.rows:
            celdas = [c.text.strip() for c in row.cells]

            if len(celdas) < 2:
                continue

            campo = celdas[0].lower()
            valor = celdas[1].strip()

            if not valor:
                continue

            if "cliente" in campo or "contacto" in campo:
                datos["nombre"] = limpiar_nombre(valor)

            elif "cargo" in campo:
                datos["cargo"] = valor

            elif "mail" in campo or "correo" in campo:
                datos["correo"] = valor

            elif "tel" in campo or "cel" in campo:
                datos["telefono"] = valor

            elif "empresa" in campo or "compañ" in campo or "edificio" in campo:
                datos["compania"] = valor.upper()

    # =========================
    # 2. SEGUNDA TABLA (SI EXISTE)
    # =========================
    if len(doc.tables) > 1:
        tabla2 = doc.tables[1]

        for row in tabla2.rows:
            texto = " ".join([c.text.strip() for c in row.cells]).upper()

            if "COLOMBIA" in texto:
                datos["ciudad"] = texto.replace(", COLOMBIA", "").strip()

    # =========================
    # 3. PÁRRAFOS (RESPALDO REAL)
    # =========================
    for p in doc.paragraphs:
        t = p.text.strip()
        t_up = t.upper()

        # ciudad
        if not datos["ciudad"] and "COLOMBIA" in t_up:
            datos["ciudad"] = t.replace(", Colombia", "").strip()

        # teléfono (línea sola con números)
        if not datos["telefono"]:
            if t.isdigit() and len(t) >= 7:
                datos["telefono"] = t

        # compañía (EDIFICIO ...)
        if not datos["compania"] and "EDIFICIO" in t_up:
            datos["compania"] = t_up
    # =====================
    # CARGO (desde texto real)
    # =====================
    for p in doc.paragraphs:
        t = p.text.strip()
        t_low = t.lower()
    
        if not datos["cargo"]:
            if any(x in t_low for x in [
                "gerente",
                "director",
                "presidente",
                "administrador",
                "coordinador"
            ]):
                datos["cargo"] = t
        if not datos["cargo"]:
        datos["cargo"] = ""
    # =========================
    # 4. LIMPIEZA FINAL
    # =========================
    if not datos["nombre"]:
        datos["nombre"] = "CLIENTE"

    if not datos["ciudad"]:
        datos["ciudad"] = "Bogotá"

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
        nombre = datos.get("nombre") or "Cliente"

        reemplazos = {
            "consecutivo": datetime.now().strftime("%Y%m%d%H%M"),
            "fecha": fecha_es(),
            "nombre": nombre,
            "cargo": datos.get("cargo", ""),
            "compania": datos.get("compania", ""),
            "correo": datos.get("correo", ""),
            "telefono": datos.get("telefono", ""),
            "ciudad": datos.get("ciudad", ""),
            "alcance": datos.get("ciudad", ""),
            "tratamiento": obtener_tratamiento(datos.get("cargo", "")),
            "saludo": "Estimado",
            "nombre_corto": nombre.split()[0],
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
