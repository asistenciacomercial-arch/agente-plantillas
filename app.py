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

        # (por ahora solo copiamos)
        shutil.copy(temp_path, output_path)

        return FileResponse(
            path=output_path,
            filename="resultado.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        return {"error": str(e)}
        return {"error": str(e)}
