from fastapi import FastAPI

app = FastAPI()

@app.get("/")
def home():
return {"status": "ok"}

@app.post("/procesar/")
async def procesar():
try:
from docxtpl import DocxTemplate
from docx import Document
return {"mensaje": "imports OK"}
except Exception as e:
return {"error": str(e)}
