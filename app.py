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
        "ciudad": "",
    }

    tabla = None

    # =====================================
    # BUSCAR TABLA DEL LEVANTAMIENTO
    # =====================================
    for t in doc.tables:

        texto_tabla = ""
    
        for row in t.rows:
            for cell in row.cells:
                texto_tabla += cell.text.lower() + " "
    
        if (
            "contacto" in texto_tabla
            and "cargo" in texto_tabla
            and (
                "compañía" in texto_tabla
                or "compania" in texto_tabla
            )
        ):
    
            tabla = t
            break

    if tabla is None:
        return datos

    try:

        # =====================================
        # COMPAÑIA
        # fila 4
        # =====================================
        datos["compañia"] = (
            tabla.rows[3].cells[1].text.strip().upper()
        )
        
        # =====================================
        # NOMBRE
        # fila 5
        # =====================================
        nombre = tabla.rows[4].cells[1].text.strip()
        
        nombre = (
            nombre.upper()
            .replace("SR.", "")
            .replace("SRA.", "")
            .replace("DR.", "")
            .replace("DRA.", "")
            .strip()
        )
        
        datos["nombre"] = nombre
        
        if nombre:
            datos["primer_nombre"] = (
                nombre.split()[0].title()
            )
        
        # =====================================
        # CORREO
        # fila 5 col 3
        # =====================================
        correo = tabla.rows[4].cells[3].text.strip()
        
        if "@" in correo:
            datos["correo"] = correo
        
        # =====================================
        # CARGO
        # fila 6
        # =====================================
        datos["cargo"] = (
            tabla.rows[5].cells[1].text.strip()
        )
        
        # =====================================
        # TELEFONO
        # fila 6 col 3
        # =====================================
        telefono = (
            tabla.rows[5].cells[3].text.strip()
        )
        
        datos["telefono"] = telefono
        
        # =====================================
        # DIRECCION
        # fila 4 col 3
        # =====================================
        datos["direccion"] = (
            tabla.rows[3].cells[3].text.strip()
        )

        # =====================================
        # CIUDAD
        # =====================================
        try:

            ciudad = tabla.rows[10].cells[1].text.strip()

            datos["ciudad"] = ciudad

        except:
            datos["ciudad"] = ""

    except Exception as e:

        print("ERROR EXTRACCION:", e)

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

    # =====================================
    # ESCOLTAS
    # =====================================
    if servicio == "escolta":

        # múltiples modalidades
        if detalle == "mensual":
            return "plantillas/escolta_mensual.docx"

        # individuales
        elif modalidad == "motorizado":
            return "plantillas/escolta_motorizado.docx"

        elif modalidad == "conductor":
            return "plantillas/escolta_conductor.docx"

        elif modalidad == "a_pie":
            return "plantillas/escolta_apie.docx"

        else:
            return "plantillas/escolta_mensual.docx"

    # =====================================
    # VIGILANCIA
    # =====================================
    elif servicio == "vigilancia":

        # mezcla armada + sin arma
        if detalle == "vigilancia_mixta":
            return "plantillas/vigilancia_mixta.docx"

        # armada
        elif detalle == "armada":

            if modalidad == "mensual":
                return "plantillas/vigilancia_armada_m.docx"

            elif modalidad == "evento":
                return "plantillas/vigilancia_armada_e.docx"

            else:
                return "plantillas/vigilancia_armada_m.docx"

        # sin arma
        elif detalle == "sin_arma":

            if modalidad == "mensual":
                return "plantillas/vigilancia_sin_arma_m.docx"

            elif modalidad == "evento":
                return "plantillas/vigilancia_sin_arma_e.docx"

            else:
                return "plantillas/vigilancia_sin_arma_m.docx"

        else:
            return "plantillas/vigilancia_sin_arma_m.docx"

    # =====================================
    # CONFIABILIDAD
    # =====================================
    elif servicio == "confiabilidad":

        return "plantillas/confiabilidad.docx"

    # =====================================
    # SEGURIDAD ELECTRONICA
    # =====================================
    elif servicio == "electronica":

        return "plantillas/seguridad_electronica.docx"

    # =====================================
    # MONITOREO
    # =====================================
    elif servicio == "monitoreo":

        return "plantillas/monitoreo.docx"

    # =====================================
    # CONSULTORIA
    # =====================================
    elif servicio == "consultoria":

        return "plantillas/consultoria.docx"

    # =====================================
    # ESTUDIOS DE SEGURIDAD
    # =====================================
    elif servicio == "estudio_seguridad":

        return "plantillas/estudio_seguridad.docx"

    # =====================================
    # CAPACITACIONES
    # =====================================
    elif servicio == "capacitacion":

        return "plantillas/capacitacion.docx"

    # =====================================
    # POLIGRAFIA
    # =====================================
    elif servicio == "poligrafia":

        return "plantillas/poligrafia.docx"

    # =====================================
    # VISITAS DOMICILIARIAS
    # =====================================
    elif servicio == "visita_domiciliaria":

        return "plantillas/visita_domiciliaria.docx"

    # =====================================
    # DEFAULT
    # =====================================
    else:

        return "plantillas/default.docx"
    
def generar_titulos(servicio, detalle):

    titulo_ref = ""
    titulo_servicio = ""

    # =====================================
    # ESCOLTAS
    # =====================================
    if servicio == "escolta":

        # =====================================
        # VARIAS MODALIDADES
        # =====================================
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
                "Propuesta Servicios de Vigilancia"
            )

            titulo_servicio = (
                "SERVICIOS DE VIGILANCIA "
                "- ARMADA Y SIN ARMA"
            )

        # =====================================
        # ARMADA
        # =====================================
        elif detalle == "armada":

            titulo_ref = (
                "Propuesta Servicio de Vigilancia Armada"
            )

            titulo_servicio = (
                "SERVICIO DE VIGILANCIA ARMADA"
            )

        # =====================================
        # SIN ARMA
        # =====================================
        else:

            titulo_ref = (
                "Propuesta Servicio de Vigilancia "
                "con Medio de Comunicación, Sin Arma"
            )

            titulo_servicio = (
                "SERVICIO DE VIGILANCIA "
                "CON MEDIO DE COMUNICACIÓN SIN ARMA"
            )

    return titulo_ref, titulo_servicio
    
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

            "ciudad": datos["ciudad"],
            "alcance": datos["ciudad"],
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
