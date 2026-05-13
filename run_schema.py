"""Lancement du Schema Designer N1 : python run_schema.py → http://localhost:8001"""
import uvicorn

if __name__ == "__main__":
    uvicorn.run("web.schema_app.main:app", host="127.0.0.1", port=8001, reload=True)
