@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt -q
if "%~1"=="" (
  echo Uso: arrastrar el Excel de Legajos Online sobre este .bat
  echo   o: comparar_legajos_online.bat "ruta\al\archivo.xlsx"
  pause
  exit /b 1
)
python src\comparar_legajos_online.py "%~1"
echo.
echo Comparacion en: data\comparaciones\
pause
