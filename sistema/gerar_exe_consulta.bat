@echo off
chcp 65001 >nul
cd /d "%~dp0.."
set "RAIZ=%CD%"
echo Gerando "Consulta Vendas.exe" (leva ~2 minutos)...
python -m PyInstaller --noconfirm --onefile --windowed --name "Consulta Vendas" ^
  --paths "%RAIZ%\sistema" --splash "%RAIZ%\sistema\splash_consulta.png" ^
  --distpath "%RAIZ%" --workpath "%TEMP%\consulta_build" --specpath "%TEMP%\consulta_build" ^
  --exclude-module matplotlib --exclude-module scipy --exclude-module IPython ^
  --exclude-module pytest --exclude-module PyQt5 --exclude-module PySide6 ^
  --exclude-module pyarrow --exclude-module sqlalchemy --exclude-module jinja2 ^
  "%RAIZ%\sistema\consulta_vendas.py"
echo.
echo Pronto: "Consulta Vendas.exe" na pasta do projeto.
pause
