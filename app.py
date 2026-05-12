import unicodedata

def normalizar(txt):

    if not txt:
        return ""

    txt = txt.lower()

    txt = unicodedata.normalize(
        'NFKD',
        txt
    ).encode(
        'ascii',
        'ignore'
    ).decode('utf-8')

    txt = " ".join(txt.split())

    return txt
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
# =====================================
# DETECTAR GENERO
# =====================================
def detectar_genero(nombre):

    nombre = nombre.lower()

    femeninos = [
        "maria","ana","laura","paula","andrea","carolina",
        "diana","luisa","patricia","camila","valentina",
        "anverly","solange","nancy","lina","adriana",
        "monica","claudia","johana","tatiana","veronica",
        "dayana","yohana","juliana","angela","sandra"
    ]

    masculinos = [
        "juan","carlos","manuel","ivan","andres","luis",
        "felipe","daniel","miguel","jose","eduardo",
        "sebastian","david","alejandro","ricardo",
        "hector","oscar","diego","jorge","kevin"
    ]

    palabras = nombre.split()

    for p in palabras:

        pl = p.lower().strip()

        if pl in femeninos:
            return "f"

        if pl in masculinos:
            return "m"

    return "m"


# =====================================
# PRIMER NOMBRE REAL
# =====================================
def obtener_primer_nombre_real(nombre):

    femeninos = [
        "maria","ana","laura","paula","andrea","carolina",
        "diana","luisa","patricia","camila","valentina",
        "anverly","solange","nancy","lina","adriana"
    ]

    masculinos = [
        "juan","carlos","manuel","ivan","andres","luis",
        "felipe","daniel","miguel","jose","eduardo"
    ]

    palabras = nombre.split()

    for p in palabras:

        pl = p.lower().strip()

        if pl in femeninos or pl in masculinos:
            return p.capitalize()

    return palabras[0].capitalize()


# =====================================
# TRATAMIENTO
# =====================================
def obtener_tratamiento(nombre, cargo):

    genero = detectar_genero(nombre)

    cargo = cargo.lower()

    es_directivo = any(x in cargo for x in [
        "director",
        "directora",
        "gerente",
        "presidente",
        "presidenta",
        "coordinador",
        "coordinadora",
        "jefe",
        "jefa"
    ])

    if es_directivo:

        if genero == "f":
            return "Doctora"

        return "Doctor"

    if genero == "f":
        return "Señora"

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
        if len(tabla.rows) > 4:
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
        for t in doc.tables:
        
            for row in t.rows:
        
                try:
        
                    titulo = row.cells[0].text.lower()
        
                    if "ciudad - lugar" in titulo:
        
                        datos["ciudad"] = (
                            row.cells[1].text.strip()
                        )
        
                        break
        
                except:
                    pass
    except Exception as e:

        print("ERROR EXTRACCION:", e)

    return datos
# =========================
# DETECTAR SERVICIO DESDE TABLA X
# =========================
def detectar_servicio(doc):

    for table in doc.tables:

        for row in table.rows:

            celdas = [c.text.strip().lower() for c in row.cells]

            fila = normalizar(
                " | ".join(celdas)
            )

            print("FILA:", fila)

            # =====================================
            # VIGILANCIA
            # =====================================
            if "vigilancia" in fila and "x" in fila:
                return "vigilancia"

            # =====================================
            # ESCOLTA
            # =====================================
            if "escolta" in fila and "x" in fila:
                return "escolta"

            # =====================================
            # CONFIABILIDAD
            # =====================================
            if "confiabilidad" in fila and "x" in fila:
                return "confiabilidad"

            # =====================================
            # ELECTRONICA
            # =====================================
            if (
                "seguridad electronica" in fila
                or "electrónica" in fila
            ) and "x" in fila:
                return "electronica"

            # =====================================
            # MONITOREO
            # =====================================
            if "monitoreo" in fila and "x" in fila:
                return "monitoreo"
            # =====================================
            # FALLBACK POR TEXTO LIBRE
            # =====================================
        
            texto_total = ""
        
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        texto_total += " " + normalizar(cell.text)
        
            # ESCOLTA
            if (
                "escolta" in texto_total
                or "acompanamiento" in texto_total
                or "acompañamiento" in texto_total
            ):
                return "escolta"
        
            # VIGILANCIA
            if "vigilancia" in texto_total:
                return "vigilancia"
        
            # CONFIABILIDAD
            if "confiabilidad" in texto_total:
                return "confiabilidad"
        
            # ELECTRONICA
            if (
                "seguridad electronica" in texto_total
                or "seguridad electrónica" in texto_total
            ):
                return "electronica"
        
            # MONITOREO
            if "monitoreo" in texto_total:
                return "monitoreo"
        
    return None
# =========================
# PLANTILLA
# =========================
def seleccionar_plantilla(
    servicio,
    detalle,
    modalidad
):

    print("==== SELECCIONAR PLANTILLA ====")
    print("SERVICIO:", servicio)
    print("DETALLE:", detalle)
    print("MODALIDAD:", modalidad)

    # =====================================
    # ESCOLTAS
    # =====================================
    if servicio == "escolta":

        # MULTIPLE
        if detalle == "multiple":

            return (
                "plantillas/"
                "escolta_multiple.docx"
            )

        # CONDUCTOR
        if detalle == "conductor":

            return (
                "plantillas/"
                "escolta_conductor.docx"
            )

        # MOTORIZADO
        if detalle == "motorizado":

            return (
                "plantillas/"
                "escolta_motorizado.docx"
            )

        # A PIE
        if detalle == "a_pie":

            return (
                "plantillas/"
                "escolta_apie.docx"
            )

        # GENERAL
        return (
            "plantillas/"
            "escolta_general.docx"
        )

    # =====================================
    # VIGILANCIA
    # =====================================
    if servicio == "vigilancia":

        if detalle == "vigilancia_mixta":

            return (
                "plantillas/"
                "vigilancia_mixta.docx"
            )

        if detalle == "armada":

            return (
                "plantillas/"
                "vigilancia_armada.docx"
            )

        if detalle == "sin_arma":

            return (
                "plantillas/"
                "vigilancia_sin_arma.docx"
            )

        return (
            "plantillas/"
            "vigilancia_general.docx"
        )

    # =====================================
    # CONFIABILIDAD
    # =====================================
    if servicio == "confiabilidad":

        return (
            "plantillas/"
            "confiabilidad.docx"
        )

    # =====================================
    # ELECTRONICA
    # =====================================
    if servicio == "electronica":

        return (
            "plantillas/"
            "electronica.docx"
        )

    # =====================================
    # MONITOREO
    # =====================================
    if servicio == "monitoreo":

        return (
            "plantillas/"
            "monitoreo.docx"
        )

    # =====================================
    # DEFAULT
    # =====================================
    return (
        "plantillas/"
        "default.docx"
    )
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
        if detalle == "multiple":

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
# =========================
# OBTENER TEXTO COMPLETO
# =========================
def obtener_texto_completo(doc):

    texto = ""

    # =========================
    # PÁRRAFOS
    # =========================
    for p in doc.paragraphs:

        texto += " " + normalizar(p.text)

    # =========================
    # TABLAS
    # =========================
    for table in doc.tables:

        for row in table.rows:

            for cell in row.cells:

                texto += " " + normalizar(cell.text)

    return texto
    
def detectar_detalle(doc):

    texto = obtener_texto_completo(doc)

    print(texto)

    # =====================================
    # VARIABLES
    # =====================================

    tiene_motorizado = (

        "motorizado" in texto
        or "moto" in texto
        or "rodamiento" in texto
    )

    tiene_conductor = (

        "conductor" in texto
        or "chofer" in texto
        or "driver" in texto
        or "vehiculo" in texto
        or "vehículo" in texto
    )

    tiene_apie = (

        "a pie" in texto
        or "acompanante" in texto
        or "acompañante" in texto
    )

    print("MOTORIZADO:", tiene_motorizado)
    print("CONDUCTOR:", tiene_conductor)
    print("A PIE:", tiene_apie)

    modalidades_detectadas = []

    # =====================================
    # MODALIDADES
    # =====================================

    if tiene_motorizado:
        modalidades_detectadas.append("motorizado")

    if tiene_conductor:
        modalidades_detectadas.append("conductor")

    if tiene_apie:
        modalidades_detectadas.append("a_pie")

    # =====================================
    # MULTIPLE
    # =====================================

    if len(modalidades_detectadas) >= 2:

        print("DETALLE: multiple")

        return "multiple"

    # =====================================
    # SOLO UNA
    # =====================================

    if len(modalidades_detectadas) == 1:

        print(
            "DETALLE:",
            modalidades_detectadas[0]
        )

        return modalidades_detectadas[0]

    # =====================================
    # ESCOLTA GENERAL
    # =====================================

    if "escolta" in texto:

        print("DETALLE: general")

        return "general"

    # =====================================
    # VIGILANCIA
    # =====================================

    tiene_armada = (
        "armado" in texto
        or "armada" in texto
    )

    tiene_sin_arma = (
        "sin arma" in texto
        or "medio de comunicacion" in texto
    )

    if tiene_armada and tiene_sin_arma:

        return "vigilancia_mixta"

    if tiene_armada:

        return "armada"

    if tiene_sin_arma:

        return "sin_arma"

    return None
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
        
        modalidad = detectar_modalidad(doc)
        
        print("SERVICIO FINAL:", servicio)
        print("DETALLE FINAL:", detalle)
        print("MODALIDAD FINAL:", modalidad)

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
        print("PLANTILLA FINAL:", plantilla)
        ## =========================
        # REEMPLAZOS
        # =========================
        titulo_ref, titulo_servicio = generar_titulos(
            servicio,
            detalle
        )
        
        # =========================
        # GENERO Y SALUDO
        # =========================
        
        tratamiento = obtener_tratamiento(
            datos["nombre"],
            datos["cargo"]
        )
        
        primer_nombre = obtener_primer_nombre_real(
            datos["nombre"]
        )
        
        genero = detectar_genero(
            datos["nombre"]
        )
        
        saludo_base = "Estimado"
        
        if genero == "f":
            saludo_base = "Estimada"
        
        # =========================
        # REEMPLAZOS
        # =========================

        
        reemplazos = {

            "consecutivo": datetime.now().strftime("%Y%m%d%H%M"),
        
            "fecha": fecha_es(),
        
            "tratamiento": tratamiento,
        
            "nombre": datos["nombre"],
        
            "primer_nombre": primer_nombre,
        
            "cargo": datos["cargo"],
        
            "compañia": datos["compañia"],
        
            "correo": datos["correo"],
        
            "telefono": datos["telefono"],
        
            "direccion": datos["direccion"],
        
            "ciudad": datos["ciudad"],
        
            "alcance": datos["ciudad"],
        
            # SALUDO
            "saludo": (
                f"{saludo_base} "
                f"{tratamiento.lower()} "
                f"{primer_nombre}"
            ),
        
            # TITULOS
            "titulo_ref": titulo_ref,
        
            "titulo_servicio": titulo_servicio
        
        }
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
