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
        import shutil

        temp_path = "temp.docx"
        output_path = "resultado.docx"

        # Guardar archivo
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        from docx import Document

        doc = Document(temp_path)
        
        texto = "\n".join([p.text for p in doc.paragraphs])
        
        # ejemplo simple de detección
        if "levantamiento" in texto.lower():
            contenido = "Plantilla de levantamiento detectada"
        else:
            contenido = "Otro tipo de documento"
        
        # crear nuevo documento
        nuevo = Document()
        nuevo.add_heading("Documento generado", 0)
        nuevo.add_paragraph(contenido)
        
        nuevo.save(output_path)

        return FileResponse(
            path=output_path,
            filename="resultado.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        return {"error": str(e)}
        return {"error": str(e)}
