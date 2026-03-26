@echo off
echo ============================================
echo   VERIDEM — Demarrage du serveur IA
echo ============================================
echo.
call venv\Scripts\activate
echo Demarrage sur http://localhost:5000 ...
echo (La premiere fois, les modeles se telechargent : ~3 GB)
echo.
python server.py
pause
