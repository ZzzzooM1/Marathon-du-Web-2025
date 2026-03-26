@echo off
echo ============================================
echo   VERIDEM — Installation du pipeline IA
echo ============================================
echo.

REM Creer un environnement virtuel
echo [1/4] Creation de l'environnement Python...
python -m venv venv
call venv\Scripts\activate

REM Forcer la bonne version de transformers
echo [2/4] Installation de transformers==4.44.0 (obligatoire pour Florence-2)...
pip install transformers==4.44.0

REM Installer le reste des dependances
echo [3/4] Installation des autres dependances (peut prendre 5-10 min)...
pip install -r requirements.txt

echo.
echo [4/4] Installation terminee !
echo.
echo ============================================
echo   Pour lancer le serveur :
echo   Double-cliquez sur lancer.bat
echo ============================================
pause
