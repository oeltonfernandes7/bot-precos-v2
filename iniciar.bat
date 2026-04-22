@echo off
echo ========================================
echo    BOT DE PREÇOS v2.0 - INSTALAÇÃO
echo ========================================
echo.

echo Instalando dependencias...
python -m pip install streamlit pandas openpyxl plotly requests playwright

echo.
echo Instalando Playwright browsers...
python -m playwright install chromium

echo.
echo ========================================
echo         INICIANDO APLICAÇÃO
echo ========================================
echo.

echo Abrindo navegador...
start http://localhost:8501

echo.
echo Iniciando Streamlit...
python -m streamlit run app.py --browser.serverAddress localhost --server.headless true

pause