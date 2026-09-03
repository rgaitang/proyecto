"""Punto de entrada para ejecutar la aplicación.

Uso:
    python run.py
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app

if __name__ == '__main__':
    with app.app_context():
        from models import db, Sucursal
        db.create_all()
        if Sucursal.query.count() == 0:
            print('No hay sucursales. Ejecuta primero:  python seed.py')
    print('Sistema de Horarios en:  http://127.0.0.1:5000')
    app.run(debug=True, host='0.0.0.0', port=5000)
