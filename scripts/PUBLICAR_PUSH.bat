@echo off
chcp 65001 >nul
setlocal
echo ================================================
echo  Publicando commit ya preparado - Grupo J^&D
echo ================================================

cd /d "%~dp0."

echo.
echo  Este .bat SOLO hace push. No hace "git add -A",
echo  asi que no sube archivos sueltos que no quieras
echo  publicar (backups, modulos en pruebas, etc).
echo.

echo [1/3] Eliminando locks si existen...
if exist ".git\index.lock" (
    del /f /q ".git\index.lock"
    echo      index.lock eliminado.
) else (
    echo      Sin locks, OK.
)
if exist ".git\HEAD.lock" del /f /q ".git\HEAD.lock"

echo.
echo [2/3] Commit que se va a publicar:
git log --oneline -1
echo.
echo      Pendientes sin commitear ^(NO se publican^):
git status --short

echo.
echo [3/3] Enviando a GitHub...
git push origin main
if errorlevel 1 goto :error

echo.
echo ================================================
echo  Push enviado. El despliegue lo hace GitHub.
echo ================================================
echo.
echo  Verifica que termine en verde aqui:
echo    https://github.com/dlopezu0221-byte/dashboard-jyd/actions
echo.
echo  Tarda entre 30 y 60 segundos. Despues abre:
echo    https://dashboard.grupoempresarialjd.com/grupo579780/?v=1
echo    https://dashboard.grupoempresarialjd.com/monitores/?v=1
echo.
echo  El sello debe decir "Datos al: 06/10/2026".
echo ================================================
goto :fin

:error
echo.
echo ================================================
echo  ERROR en git push. Revisa el mensaje de arriba.
echo  Nada se publico.
echo ================================================

:fin
echo.
pause
endlocal
