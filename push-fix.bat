@echo off
REM --- Push du correctif content.js vers GitHub ---
cd /d "%~dp0"
echo Envoi du correctif content.js vers GitHub...
echo.
git push origin main
echo.
echo Termine. Appuyez sur une touche pour fermer.
pause >nul
