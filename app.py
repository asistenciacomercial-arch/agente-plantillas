from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from docx import Document
from docxtpl import DocxTemplate
import shutil
import re

app = FastAPI()

def leer_docx(path):
doc = Document(path)
return "\n".join([p.text for p in doc.paragraphs])

def extraer_datos(texto):
datos = {}

```
# Cliente
match = re.search(r'Compañía\s*(.*)', texto)
if match:
    datos["cliente"] = match.group(1).strip()
else:
    datos["cliente"] = "Cliente no detectado"

# Ciudad
if "cali" in texto.lower():
    datos["ciudad"] = "Cali"

# Servicio
if "vigilancia" in texto.lower():
    datos["servicio"] = "vigilancia"
else:
    datos["servicio"] = "general"

return datos
```

def elegir_plantilla(servicio):
if servicio == "vigilancia":
return "plantillas/vigilancia.docx"
return "plantillas/general.docx"

@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):

```
with open("temp.docx", "wb") as buffer:
    shutil.copyfileobj(file.file, buffer)

texto = leer_docx("temp.docx")
datos = extraer_datos(texto)

plantilla = elegir_plantilla(datos["servicio"])

doc = DocxTemplate(plantilla)

doc.render({
    "cliente": datos.get("cliente"),
    "servicio": datos.get("servicio"),
    "ciudad": datos.get("ciudad", "")
})

salida = "resultado.docx"
doc.save(salida)

return FileResponse(salida, filename="resultado.docx")
```
