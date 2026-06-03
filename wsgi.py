"""
Punto de entrada para Gunicorn en producción.
Render ejecuta: gunicorn wsgi:app
"""
from app import app

if __name__ == "__main__":
    app.run()
