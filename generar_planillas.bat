@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt -q
python src\crear_plantilla.py
python src\generar_planilla.py --anio 2026 --mes 6
echo.
echo Planillas en: data\planillas\
pause
