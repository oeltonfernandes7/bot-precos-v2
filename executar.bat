@echo off
chcp 65001 >nul
title Bot de Precos em Farmacias

echo ----------------------------------------
echo   BOT DE PESQUISA DE PRECOS EM FARMACIAS
echo ----------------------------------------
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado.
    echo        Instale em https://python.org e tente novamente.
    pause
    exit /b 1
)

echo [OK] Python encontrado.
echo.

echo [1/3] Instalando dependencias...
python -m pip install requests --quiet
python -m pip install pandas --quiet
python -m pip install openpyxl --quiet
python -m pip install playwright --quiet
python -m playwright install chromium
echo [OK] Dependencias verificadas.
echo.

if not exist "produtos.xlsx" (
    echo [ERRO] Arquivo produtos.xlsx nao encontrado.
    echo        Crie o arquivo com as colunas: EAN e Nome
    echo        Exemplo: EAN=7891058013202, Nome=Dipirona 500mg
    echo.
    pause
    exit /b 1
)

echo [OK] produtos.xlsx encontrado.
echo.

if not exist "estrategia.json" (
    echo [2/3] Investigando tecnologia dos sites...
    echo       Aguarde, isso pode levar alguns segundos.
    echo.
    python investigar_sites.py
    echo.
) else (
    echo [2/3] estrategia.json ja existe. Pulando investigacao.
    echo       Apague estrategia.json para reinvestigar os sites.
    echo.
)

echo [3/3] Iniciando busca de precos...
echo.
python bot_precos.py

echo.
echo ----------------------------------------
echo   Concluido! Abra o arquivo resultado.csv
echo ----------------------------------------
echo.
pause
