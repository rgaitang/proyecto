"""Punto de entrada para servidores WSGI (gunicorn en Render/otro hosting).

Uso:  gunicorn wsgi:app
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app

if __name__ == '__main__':
    with app.app_context():
        from models import db, Sucursal
        db.create_all()
        if Sucursal.query.count() == 0:
            print('NO HAY SUCURSALES: ejecuta seed.py localmente o crea un comando de seed.')
    app.run()
