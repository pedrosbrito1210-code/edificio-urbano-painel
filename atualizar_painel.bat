@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Atualizando o painel do edificio...
python atualizar.py %*
echo.
pause
