@echo off
chcp 65001 >nul
title Dashboard Retailer - Lancement
color 0A
cls

echo.
echo =================================================
echo   DASHBOARD COMPLET RETAILER
echo =================================================
echo.

REM Aller dans le dossier du script
cd /d "%~dp0"

set "PYTHON_EXE=F:\python\WPy64-31700\python\python.exe"

echo Lancement du script...
echo.

"%PYTHON_EXE%" generer_dashboard.py
set "ERROR_CODE=%ERRORLEVEL%"

echo.
echo =================================================
if "%ERROR_CODE%"=="0" (
    echo  SUCCES! Dashboard genere
    echo  Consultez le dossier: F:\02_Analyse_Rotation\Dashboard
) else (
    echo  ERREUR lors de l'execution
    echo  Code erreur: %ERROR_CODE%
)
echo =================================================
echo.
pause
