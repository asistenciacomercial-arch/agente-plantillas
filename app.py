from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from docx import Document
from docxtpl import DocxTemplate
import shutil

app = FastAPI()

def leer_docx(path):
doc = Document(path)
return "\n".join([p.text for p in doc.paragraphs])

def elegir_plantilla(texto):
texto = texto.lower()

```
# VIGILANCIA
if "vigilancia" in texto:
    if "armada" in texto:
        if "mensual" in texto:
            return "plantillas/vigilancia_armada_m.docx"
        elif "eventual" in texto:
            return "plantillas/vigilancia_armada_e.docx"
    
    if "sin arma" in texto:
        if "mensual" in texto:
            return "plantillas/vigilancia_sin_arma_m.docx"
        elif "eventual" in texto:
            return "plantillas/vigilancia_sin_arma_e_12h.docx"

# ESCOLTA
if "escolta" in texto:
    if "motorizado" in texto:
        return "plantillas/escolta_motorizado.docx"
    elif "conductor" in texto:
        return "plantillas/escolta_conductor_ev.docx"
    else:
        return "plantillas/escolta_mensual.docx"

# OTROS
if "ciberseguridad" in texto:
    return "plantillas/capacitacion_ciberseguridad.docx"

if "monitoreo" in texto:
    return "plantillas/monitoreo.docx"

return "plantillas/confiabilidad.docx"
```

@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):

```
with open("temp.docx", "wb") as buffer:
    shutil.copyfileobj(file.file, buffer)

texto = leer_docx("temp.docx")

plantilla = elegir_plantilla(texto)

doc = DocxTemplate(plantilla)

doc.render({
    "cliente": "Cliente demo",
    "servicio": "Servicio detectado"
})

salida = "resultado.docx"
doc.save(salida)

return FileResponse(salida, filename="resultado.docx")
```
