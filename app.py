from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import shutil
import os

app = FastAPI()

@app.get("/")
def home():
return {"status": "ok"}

@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):
try:
from docx import Document
from docxtpl import DocxTemplate

```
    temp_path = "temp.docx"

    with open(temp_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    doc = Document(temp_path)
    texto = "\n".join([p.text for p in doc.paragraphs])

    plantilla = "plantillas/confiabilidad.docx"

    if not os.path.exists(plantilla):
        return {"error": f"No existe plantilla: {plantilla}"}

    tpl = DocxTemplate(plantilla)

    tpl.render({
        "cliente": "Demo",
        "servicio": "Test"
    })

    salida = "resultado.docx"
    tpl.save(salida)

    os.remove(temp_path)

    return FileResponse(salida, filename="resultado.docx")

except Exception as e:
    return {"error": str(e)}
```
