@echo off
chcp 65001 >nul
echo ================================================
echo  Publicando dashboards - Grupo J^&D
echo ================================================

cd /d "%~dp0"

:: Fecha actual para el commit
for /f "tokens=2 delims==" %%I in ('wmic os get LocalDateTime /value 2^>nul') do set dt=%%I
set DIA=%dt:~6,2%
set MES=%dt:~4,2%
set ANO=%dt:~0,4%
set FECHA=%DIA%/%MES%/%ANO%

echo.
echo [1/5] Eliminando locks si existen...
if exist ".git\index.lock" (
    del /f /q ".git\index.lock"
    echo      index.lock eliminado.
) else (
    echo      Sin locks, OK.
)
if exist ".git\HEAD.lock" (
    del /f /q ".git\HEAD.lock"
)

echo.
echo [2/5] Añadiendo todos los cambios al staging...
git add -A
echo      Incluye: grupo, erika, fabio, estudios, monitores
echo      Done.

echo.
echo [3/5] Verificando cambios pendientes...
git status --short

echo.
echo [4/5] Haciendo commit...
git commit -m "Actualizacion datos %FECHA%"

echo.
echo [5/5] Publicando en GitHub Pages...
git push origin main

echo.
echo ================================================
echo  Listo. Datos publicados al %FECHA%
echo  URL: https://grupoempresarialjd.com/
echo  Monitores: https://grupoempresarialjd.com/monitores/
echo ================================================
pause
