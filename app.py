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

    dia = str(hoy.day).zfill(2)

    return f"{dia} de {meses[hoy.month-1]} de {hoy.year}"
    
# =========================
# EXTRAER DATOS (FORMATO REAL)
# =========================
def extraer_datos(doc):

    datos = {
        "nombre": "",
        "primer_nombre": "",
        "cargo": "",
        "compañia": "",
        "correo": "",
        "telefono": "",
        "direccion": "",
        "ciudad": "Bogotá"
    }

    for table in doc.tables:

        for row in table.rows:

            cells = [c.text.strip() for c in row.cells]

            # recorrer por pares LABEL -> VALOR
            for i, cell in enumerate(cells):

                texto = cell.lower()

                # =====================
                # COMPAÑIA
                # =====================
                if "compañía" in texto:

                    if i + 1 < len(cells):
                        datos["compañia"] = cells[i + 1].strip().upper()

                # =====================
                # DIRECCION
                # =====================
                if "dirección" in texto:

                    if i + 1 < len(cells):
                        datos["direccion"] = cells[i + 1].strip()

                # =====================
                # CONTACTO
                # =====================
                if "contacto" in texto:

                    if i + 1 < len(cells):

                        nombre = limpiar_nombre(cells[i + 1])

                        datos["nombre"] = nombre.upper()

                        datos["primer_nombre"] = (
                            nombre.split()[0].capitalize()
                            if nombre else ""
                        )

                # =====================
                # EMAIL
                # =====================
                for c in cells:

                    if "@" in c:
                        datos["correo"] = c.strip()

                # =====================
                # CARGO
                # =====================
                if "cargo" in texto:

                    if i + 1 < len(cells):
                        datos["cargo"] = cells[i + 1].strip()

                # =====================
                # TELEFONO
                # =====================
                if "teléfono" in texto or "telefono" in texto:

                    if i + 1 < len(cells):
                        datos["telefono"] = cells[i + 1].strip()

                # =====================
                # CIUDAD
                # =====================
                if "ciudad - lugar" in texto:

                    if i + 1 < len(cells):
                        datos["ciudad"] = cells[i + 1].strip()

    print("DATOS EXTRAIDOS:", datos)

    return datos
    
# =========================
# DETECCIÓN SERVICIO
# =========================
def detectar_servicio(doc):

    for table in doc.tables:

        for row in table.rows:

            cells = [c.text.strip().lower() for c in row.cells]

            print("FILA:", cells)

            for i, cell in enumerate(cells):

                # 🔥 BUSCAR SOLO ESTA CELDA
                if "tipo de servicio" in cell:

                    # tomar celda derecha
                    if i + 1 < len(cells):

                        valor = cells[i + 1].strip().lower()

                        print("TIPO SERVICIO:", valor)

                        # =========================
                        # ESCOLTA
                        # =========================
                        if "escolta" in valor:
                            return "escolta"

                        # =========================
                        # VIGILANCIA
                        # =========================
                        if "vigilancia" in valor:
                            return "vigilancia"

                        # =========================
                        # ELECTRONICA
                        # =========================
                        if "electr" in valor:
                            return "electronica"

                        # =========================
                        # CONFIABILIDAD
                        # =========================
                        if "confiabilidad" in valor:
                            return "confiabilidad"

                        # =========================
                        # MONITOREO
                        # =========================
                        if "monitoreo" in valor:
                            return "monitoreo"

    return None
# =========================
# PLANTILLA
# =========================
def seleccionar_plantilla(servicio, detalle, modalidad):

    print("SERVICIO:", servicio)
    print("DETALLE:", detalle)
    print("MODALIDAD:", modalidad)

    # =========================
    # VIGILANCIA
    # =========================
    if servicio == "vigilancia":

    # 🔥 DOS MODALIDADES
    if detalle == "vigilancia_mixta":

        return "plantillas/vigilancia_mixta.docx"

    # =====================================
    # ARMADA
    # =====================================
    if detalle == "armada":

        if modalidad == "m":
            return "plantillas/vigilancia_armada_m.docx"

        if modalidad == "f":
            return "plantillas/vigilancia_armada_m_f.docx"

        return "plantillas/vigilancia_armada_e.docx"

    # =====================================
    # SIN ARMA
    # =====================================
    if detalle == "sin_arma":

        if modalidad == "m":
            return "plantillas/vigilancia_sin_arma_m.docx"

        if modalidad == "f":
            return "plantillas/vigilancia_sin_arma_f_m.docx"

        return "plantillas/vigilancia_sin_arma_e_12h.docx"

    # =========================
    # ESCOLTA
    # =========================
    if servicio == "escolta":

        if detalle == "mensual":
            return "plantillas/escolta_mensual.docx"
    
        if detalle == "motorizado":
            return "plantillas/escolta_motorizado.docx"
    
        if detalle == "conductor":
            return "plantillas/escolta_conductor.docx"
    
        if detalle == "apie":
            return "plantillas/escolta_apie.docx"

    # =========================
    # CONFIABILIDAD
    # =========================
    if servicio == "confiabilidad":
        return "plantillas/confiabilidad.docx"

    # =========================
    # ELECTRONICA
    # =========================
    if servicio == "electronica":
        return "plantillas/seguridad_electronica.docx"

    # =========================
    # MONITOREO
    # =========================
    if servicio == "monitoreo":
        return "plantillas/monitoreo.docx"

    # =========================
    # DEFAULT
    # =========================
    return "plantillas/vigilancia_sin_arma_m.docx"
    
def generar_titulos(servicio, detalle):

    titulo_ref = ""
    titulo_servicio = ""

    # =====================================
    # ESCOLTAS
    # =====================================
    if servicio == "escolta":

        # 🔥 VARIAS MODALIDADES
        if detalle == "mensual":

            titulo_ref = (
                "Propuesta Servicio de Escolta"
            )

            titulo_servicio = (
                "SERVICIOS DE ESCOLTA "
                "- DIFERENTES MODALIDADES"
            )
            
    # =====================================
    # VIGILANCIA
    # =====================================
    elif servicio == "vigilancia":

    # =====================================
    # MIXTA
    # =====================================
    if detalle == "vigilancia_mixta":

        titulo_ref = (
            "Propuesta servicios de vigilancia"
        )

        titulo_servicio = (
            "SERVICIOS DE VIGILANCIA "
            "- ARMADA Y SIN ARMA"
        )
    
def detectar_detalle(doc):

    texto = ""

    # =========================
    # LEER TODO EL TEXTO TABLAS
    # =========================
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texto += " " + cell.text.lower()

    print("TEXTO DETALLE:", texto)

    # =====================================
    # DETECTAR MODALIDADES ESCOLTA
    # =====================================

    tiene_motorizado = "motorizado" in texto

    tiene_conductor = (
        "conductor escolta" in texto
        or "conductor" in texto
    )

    tiene_apie = (
        "a pie" in texto
        or "acompañante" in texto
        or "acompanante" in texto
    )

    # =========================
    # CONTAR MODALIDADES
    # =========================
    total_modalidades = 0

    if tiene_motorizado:
        total_modalidades += 1

    if tiene_conductor:
        total_modalidades += 1

    if tiene_apie:
        total_modalidades += 1

    # =====================================
    # SI HAY DOS O MÁS → MENSUAL
    # =====================================

    if total_modalidades >= 2:

        print("DETALLE: mensual")
        return "mensual"

    # =====================================
    # SOLO MOTORIZADO
    # =====================================

    if tiene_motorizado:

        print("DETALLE: motorizado")
        return "motorizado"

    # =====================================
    # SOLO CONDUCTOR
    # =====================================

    if tiene_conductor:

        print("DETALLE: conductor")
        return "conductor"

    # =====================================
    # SOLO A PIE
    # =====================================

    if tiene_apie:

        print("DETALLE: apie")
        return "apie"

    # =====================================
    # VIGILANCIA
    # =====================================

    tiene_armada = (
        "armado" in texto
        or "armada" in texto
    )
    
    tiene_sin_arma = (
        "sin arma" in texto
        or "medio de comunicación" in texto
    )
    
    # 🔥 DOS MODALIDADES
    if tiene_armada and tiene_sin_arma:
    
        print("DETALLE: vigilancia_mixta")
        return "vigilancia_mixta"
    
    # 🔥 SOLO ARMADA
    if tiene_armada:
    
        print("DETALLE: armada")
        return "armada"
    
    # 🔥 SOLO SIN ARMA
    if tiene_sin_arma:
    
        print("DETALLE: sin_arma")
        return "sin_arma"
        
def detectar_modalidad(doc):

    texto = ""

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texto += " " + cell.text.lower()

    print("MODALIDAD:", texto)

    if "mensual" in texto:
        return "m"

    if "fortalecimiento" in texto:
        return "f"

    if "evento" in texto:
        return "e"

    return "m"

# =========================
# API
# =========================

from fastapi import Form

@app.post("/procesar/")
async def procesar(
    file: UploadFile = File(...),
    servicio_manual: str = Form(None)
):

    try:

        temp = "entrada.docx"

        # =========================
        # GUARDAR ARCHIVO
        # =========================
        with open(temp, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        file.file.close()

        # =========================
        # LEER DOCUMENTO
        # =========================
        doc = Document(temp)

        # =========================
        # EXTRAER DATOS
        # =========================
        datos = extraer_datos(doc)

        print("DATOS EXTRAIDOS:", datos)

        # =========================
        # VALIDAR NOMBRE
        # =========================
        if not datos.get("nombre"):
            datos["nombre"] = "CLIENTE"

        if not datos.get("primer_nombre"):
            datos["primer_nombre"] = "Cliente"

        # =========================
        # DETECTAR SERVICIO
        # =========================
        servicio = detectar_servicio(doc)

        # 🔥 SERVICIO MANUAL
        if servicio_manual and servicio_manual != "string":
            servicio = servicio_manual

        # 🔥 SI NO DETECTA
        if servicio is None:

            return {
                "requiere_seleccion": True,
                "mensaje": "No se pudo detectar el servicio",
                "opciones": [
                    "vigilancia",
                    "escolta",
                    "confiabilidad",
                    "electronica",
                    "monitoreo"
                ]
            }

        print("SERVICIO:", servicio)

        # =========================
        # DETECTAR DETALLE
        # =========================
        detalle = detectar_detalle(doc)

        print("DETALLE:", detalle)

        # =========================
        # DETECTAR MODALIDAD
        # =========================
        modalidad = detectar_modalidad(doc)

        print("MODALIDAD:", modalidad)

        # =========================
        # SELECCIONAR PLANTILLA
        # =========================
        plantilla = seleccionar_plantilla(
            servicio,
            detalle,
            modalidad
        )

        print("PLANTILLA:", plantilla)
        titulo_ref, titulo_servicio = generar_titulos(
            servicio,
            detalle
        )
        
        # =========================
        # REEMPLAZOS
        # =========================
        reemplazos = {
            
            "titulo_ref": titulo_ref,

            "titulo_servicio": titulo_servicio,
            
            "consecutivo": datetime.now().strftime("%Y%m%d%H%M"),

            "fecha": fecha_es(),

            # ENCABEZADO
            "tratamiento": obtener_tratamiento(
                datos.get("cargo", "")
            ),

            "nombre": datos.get("nombre", ""),

            "primer_nombre": datos.get(
                "primer_nombre",
                ""
            ),

            "cargo": datos.get("cargo", ""),

            "compañia": datos.get(
                "compañia",
                ""
            ),

            "correo": datos.get(
                "correo",
                ""
            ),

            "telefono": datos.get(
                "telefono",
                ""
            ),

            "direccion": datos.get(
                "direccion",
                ""
            ),

            "ciudad": datos.get(
                "ciudad",
                "Bogotá"
            ),

            # SALUDO
            "saludo": (
                f"Estimado "
                f"{obtener_tratamiento(datos.get('cargo', '')).lower()} "
                f"{datos.get('primer_nombre', '')}"
            ),
            
            # ALCANCE
            "alcance": datos.get(
                "ciudad",
                "Bogotá"
            )
        }

        print("REEMPLAZOS:", reemplazos)

        # =========================
        # GENERAR DOCUMENTO
        # =========================
        print("PLANTILLA FINAL:", plantilla)
        
        doc_tpl = DocxTemplate(plantilla)

        doc_tpl.render(reemplazos)

        # =========================
        # GUARDAR EN MEMORIA
        # =========================
        buffer = BytesIO()

        doc_tpl.save(buffer)

        buffer.seek(0)

        # =========================
        # RESPUESTA
        # =========================
        return StreamingResponse(

            buffer,

            media_type=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),

            headers={
                "Content-Disposition":
                "attachment; filename=resultado.docx"
            }
        )

    except Exception as e:

        print("ERROR:", str(e))

        return {
            "error": str(e)
        }
