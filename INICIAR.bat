@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Previsao de Pecas - aplicativo
echo Iniciando o aplicativo de previsao de pecas...
echo O navegador vai abrir sozinho. Deixe esta janela aberta enquanto usa.
python -X utf8 -W ignore sistema\app.py
pause
