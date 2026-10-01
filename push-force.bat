@echo off
REM --- Push FORCE : le dossier local ecrase GitHub ---
cd /d "%~dp0"
echo Nettoyage du verrou git (si present)...
del /f /q ".git\index.lock" 2>nul
echo.
echo Envoi FORCE de l etat local vers GitHub (branche main)...
git push --force origin main
if errorlevel 1 goto erreur
echo.
echo Termine avec succes. Appuyez sur une touche pour fermer.
pause ^>nul
goto fin
:erreur
echo.
echo ECHEC du push. Verifiez le message ci-dessus (identifiants GitHub ?).
pause ^>nul
:fin
