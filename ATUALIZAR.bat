@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Atualizando previsao de falta de pecas...
python -X utf8 -W ignore sistema\atualizar.py %*
echo.
pause
