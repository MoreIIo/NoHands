@echo off
REM --- Push des modifications du 6 juillet 2026 vers GitHub ---
cd /d "%~dp0"

echo Fichiers modifies aujourd'hui :
echo   - README.md
echo   - sidepanel.css
echo   - sidepanel.html
echo   - sidepanel.js
echo   - update-check.js (nouveau)
echo.

git add -A
git commit -m "Mise a jour du 6 juillet : systeme de verification de mise a jour (update-check.js) + ameliorations interface sidepanel (html/css/js) et README"
git push origin main

echo.
echo Termine. Appuyez sur une touche pour fermer.
pause >nul
