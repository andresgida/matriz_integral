@echo off
REM Inicia la interfaz grafica del Monitor de Tickets en el navegador.
cd /d "%~dp0"
py -m streamlit run app.py
pause
