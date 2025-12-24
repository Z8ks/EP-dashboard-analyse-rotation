@echo off
REM ===========================================
REM  LANCEMENT DASHBOARD RETAILER (clé USB)
REM  Structure respectée :
REM  - F:\02_Analyse_Rotation\generer_dashboard.py
REM  - F:\Data\
REM  - F:\python\WPy64-31700\python\python.exe
REM ===========================================

REM Monter la clé actuelle sur F: (peu importe sa lettre réelle)
subst F: "%~d0\"

REM Se placer dans le dossier du script vu depuis F:
cd /d "F:\02_Analyse_Rotation"

echo Dossier courant : %cd%
echo.

echo Lancement de Python...
"F:\python\WPy64-31700\python\python.exe" "generer_dashboard.py"
echo.
echo Code retour Python : %errorlevel%

echo.
pause

REM (Optionnel) libérer la lettre F: après exécution
REM subst F: /d
