@echo off
title Terminal Financiera — INEHRM Ojo de Dios
cd /d %~dp0

echo ========================================================
echo   CONFIGURANDO TERMINAL EJECUTIVA INEHRM...
echo ========================================================

:: Verifica e instala las librerias minimas de manera silenciosa si no estan
python -m pip install --quiet streamlit pandas openpyxl

echo.
echo [OK] Entorno listo. Abriendo Centro de Mando Financiero...
echo No cierre esta ventana mientras Isaac este usando el sistema.
echo.

:: Lanza streamlit y abre el navegador automaticamente
start http://localhost:8501
streamlit run app.py

pause