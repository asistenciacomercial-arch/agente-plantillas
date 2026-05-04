@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):
    print("🔥 Entró al endpoint")

    try:
        temp_path = "temp.docx"

        print("📁 Guardando archivo...")
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        print("✅ Archivo guardado")

        from docx import Document
        from docxtpl import DocxTemplate

        print("📦 Librerías cargadas")

        return {"status": "procesado"}

    except Exception as e:
        print("❌ ERROR:", str(e))
        return {"error": str(e)}
