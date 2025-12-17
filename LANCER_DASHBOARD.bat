@echo off

title Dashboard Retailer - Lancement

color 0A

cls

echo.
echo ================================================
echo DASHBOARD COMPLET RETAILER
echo ================================================
echo.

cd /d "%~dp0"

set PYTHON_EXE=python\WPy64-31700\python\python.exe

echo Lancement du script...
echo.

cd Scripts

REM Lancer avec redirection des erreurs
"..\%PYTHON_EXE%" generer_dashboard.py 2>&1

REM Capturer le code d'erreur
set ERROR_CODE=%ERRORLEVEL%

cd ..

echo.
echo ================================================
if %ERROR_CODE% EQU 0 (
    echo SUCCES! Dashboard genere
    echo Consultez le dossier: Resultats\
) else (
    echo ERREUR lors de l'execution
    echo Code erreur: %ERROR_CODE%
)
echo ================================================
echo.
pause
