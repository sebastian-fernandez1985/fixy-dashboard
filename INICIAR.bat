@echo off
echo ============================================
echo  FIXY DASHBOARD - Iniciando servidor...
echo ============================================

:: Verificar si Python esta instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado.
    echo Descargalo desde https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Instalar dependencias si no estan
echo Instalando dependencias...
pip install -r requirements.txt --quiet

:: Iniciar el servidor Flask
echo.
echo  Abriendo dashboard en: http://localhost:5000
echo  Para cerrar el servidor, presiona Ctrl+C
echo.
start "" http://localhost:5000
python app.py

pause
