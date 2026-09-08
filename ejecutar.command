#!/bin/bash
# Doble clic en este archivo para levantar la aplicación web.
cd "$(dirname "$0")" || exit 1
echo "Instalando dependencias (sólo la primera vez)..."
python3 -m pip install -q -r requirements.txt
echo "Abriendo la aplicación en el navegador..."
python3 -m streamlit run app.py
