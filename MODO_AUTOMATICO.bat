@echo off
:: ════════════════════════════════════════════════════
::  MODO AUTOMÁTICO  ·  Vigila el Excel continuamente
::  Deja esta ventana abierta mientras trabajas.
::  Cada vez que guardas el Excel, el dashboard
::  se actualiza solo en ~15 segundos.
:: ════════════════════════════════════════════════════
title Auto-Watcher Portfolio Dashboard
color 0B

echo.
echo  ╔══════════════════════════════════════════════╗
echo  ║   Portfolio Dashboard  ·  Auto-Watcher      ║
echo  ║   Dejá esta ventana abierta mientras         ║
echo  ║   trabajas. Se actualiza solo.               ║
echo  ╚══════════════════════════════════════════════╝
echo.

cd /d "%~dp0"

pip install openpyxl yfinance --quiet

echo  Iniciando vigilancia del Excel...
echo  (Ctrl+C para detener)
echo.

python actualizar.py --watch
pause
