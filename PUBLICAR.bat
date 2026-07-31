@echo off
echo ================================================
echo  Publicando dashboards dashboard-jyd
echo ================================================

cd /d "%~dp0"

echo.
echo [1/5] Eliminando locks si existen...
if exist ".git\index.lock" (
    del /f /q ".git\index.lock"
    echo      index.lock eliminado.
) else (
    echo      index.lock: Sin lock, OK.
)
if exist ".git\HEAD.lock" (
    del /f /q ".git\HEAD.lock"
    echo      HEAD.lock eliminado.
) else (
    echo      HEAD.lock: Sin lock, OK.
)

echo.
echo [2/5] Añadiendo cambios al staging...
git add -A
echo      Done.

echo.
echo [3/5] (paso reservado para futuros cambios)

echo.
echo [4/5] Haciendo commit...
git commit -m "Actualizacion datos al 29 de julio"

echo.
echo [5/5] Publicando en GitHub...
git push origin main

echo.
echo ================================================
echo  Listo. Revisa los mensajes arriba.
echo ================================================
pause
