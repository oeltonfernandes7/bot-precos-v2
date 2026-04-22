@echo off
echo Gerando executavel, aguarde...
echo Isso pode demorar de 3 a 10 minutos.
echo.
python -m PyInstaller launcher.py ^
  --onedir ^
  --noconsole ^
  --name "BotPrecos" ^
  --icon NONE ^
  --add-data "app.py;." ^
  --add-data "bot_precos.py;." ^
  --add-data "config.py;." ^
  --add-data "fontes.json;." ^
  --add-data "investigar_sites.py;." ^
  --add-data ".streamlit/config.toml;.streamlit" ^
  --collect-all streamlit ^
  --collect-all plotly ^
  --collect-all altair ^
  --hidden-import openpyxl ^
  --hidden-import pandas ^
  --hidden-import requests ^
  --hidden-import playwright ^
  --hidden-import psutil ^
  --hidden-import streamlit.web.cli ^
  --hidden-import streamlit.runtime.scriptrunner.magic_funcs ^
  --exclude-module tkinter ^
  --exclude-module matplotlib

echo.
echo Copiando arquivos necessarios...
if not exist dist\BotPrecos\logs mkdir dist\BotPrecos\logs
if not exist dist\BotPrecos\historico mkdir dist\BotPrecos\historico
copy fontes.json dist\BotPrecos\
copy config.py dist\BotPrecos\

echo.
echo Criando README...
echo BotPrecos v2.0 > dist\BotPrecos\LEIA-ME.txt
echo. >> dist\BotPrecos\LEIA-ME.txt
echo Como usar: >> dist\BotPrecos\LEIA-ME.txt
echo 1. Abra a pasta BotPrecos >> dist\BotPrecos\LEIA-ME.txt
echo 2. Clique duas vezes em BotPrecos.exe >> dist\BotPrecos\LEIA-ME.txt
echo 3. O navegador abrira automaticamente >> dist\BotPrecos\LEIA-ME.txt
echo 4. Para encerrar feche o navegador >> dist\BotPrecos\LEIA-ME.txt
echo. >> dist\BotPrecos\LEIA-ME.txt
echo Requisitos: >> dist\BotPrecos\LEIA-ME.txt
echo - Windows 10 ou superior >> dist\BotPrecos\LEIA-ME.txt
echo - Conexao com internet >> dist\BotPrecos\LEIA-ME.txt

echo.
echo ----------------------------------------
echo Executavel gerado em: dist\BotPrecos\
echo Copie esta pasta para qualquer computador
echo ----------------------------------------
pause