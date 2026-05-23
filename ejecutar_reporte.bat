@echo off
cd /d "%~dp0"

echo Ejecutando reporte de ventas...
py reporte_ventas.py

echo.
echo Abriendo reporte final...
start "" "output\reporte_final.xlsx"

echo.
echo Proceso terminado.
pause