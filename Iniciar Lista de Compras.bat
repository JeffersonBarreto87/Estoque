@echo off
title Lista de Compras Inteligente
cd /d "%~dp0"
echo Iniciando Lista de Compras Inteligente...
"C:\Python314\python.exe" main.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Ocorreu um erro. Pressione qualquer tecla para fechar.
    pause > nul
)
