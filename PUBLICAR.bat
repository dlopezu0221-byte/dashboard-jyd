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
echo [1/7] Eliminando locks si existen...
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
echo [2/7] Construyendo carpeta dist con archivos HTML...
if not exist "dist" mkdir "dist"
robocopy . dist *.html /s /purge /xd dist __pycache__ .git node_modules >nul 2>&1
echo      dist actualizado.

echo.
echo [3/7] Añadiendo todos los cambios al staging...
git add -A
echo      Incluye: grupo, erika, fabio, estudios, monitores
echo      Done.

echo.
echo [4/7] Verificando cambios pendientes...
git status --short

echo.
echo [5/7] Haciendo commit...
git commit -m "Actualizacion datos %FECHA%"

echo.
echo [6/7] Publicando en GitHub Pages...
git push origin main

echo.
echo [7/7] Desplegando en Cloudflare Workers...
call npx wrangler deploy --name dashboard-jyd
if %errorlevel% neq 0 (
    echo      ERROR en Cloudflare deploy. Verifica autenticacion con: npx wrangler login
) else (
    echo      Cloudflare deploy exitoso.
)

echo.
echo ================================================
echo  Listo. Datos publicados al %FECHA%
echo  URL: https://dashboard.grupoempresarialjd.com/
echo  Monitores: https://dashboard.grupoempresarialjd.com/monitores/
echo ================================================
pause
