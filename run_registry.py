"""Lancement du Registry Populator N2 : python run_registry.py → http://localhost:8000"""
import uvicorn

if __name__ == "__main__":
    uvicorn.run("web.registry_app.main:app", host="127.0.0.1", port=8000, reload=True)
