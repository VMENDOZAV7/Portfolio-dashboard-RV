@echo off
:: ════════════════════════════════════════════════════
::  ACTUALIZAR PORTFOLIO DASHBOARD
::  Doble-clic para actualizar y publicar el dashboard
:: ════════════════════════════════════════════════════
title Actualizando Portfolio Dashboard...
color 0A

echo.
echo  ╔══════════════════════════════════════╗
echo  ║   Portfolio Dashboard  ·  Updater   ║
echo  ╚══════════════════════════════════════╝
echo.

:: Ir a la carpeta del script
cd /d "%~dp0"

:: Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python no encontrado.
    echo  Descargalo en: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Instalar dependencias si faltan
echo  Verificando dependencias...
pip install openpyxl yfinance --quiet --upgrade

echo.
echo  Ejecutando actualizacion...
echo  ─────────────────────────────────────
python actualizar.py

if errorlevel 1 (
    echo.
    echo  ERROR en la actualizacion. Revisa el mensaje de arriba.
    pause
    exit /b 1
)

echo.
echo  ─────────────────────────────────────
echo  Dashboard actualizado y publicado exitosamente.
echo.
timeout /t 4
