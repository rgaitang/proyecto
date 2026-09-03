#!/usr/bin/env bash
# Script de arranque para Render
set -e

# Ejecutar seed (crea sucursales, turnos, usuario admin y carga empleados si la BD esta vacia)
python seed.py

# Iniciar servidor
gunicorn wsgi:app --workers 2 --timeout 120 --bind 0.0.0.0:$PORT
