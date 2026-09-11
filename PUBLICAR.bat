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
echo [1/6] Eliminando locks si existen...
if exist ".git\index.lock" (
    del /f /q ".git\index.lock"
    echo      index.lock eliminado.
) else (
    echo      Sin locks, OK.
)
if exist ".git\HEAD.lock" del /f /q ".git\HEAD.lock"

echo.
echo [2/6] Construyendo carpeta dist con archivos HTML...
if not exist "dist" mkdir "dist"
robocopy . dist *.html /s /purge /xd dist __pycache__ .git node_modules >nul 2>&1
echo      dist actualizado.

echo.
echo [3/6] Anadiendo todos los cambios al staging...
git add -A
if %errorlevel% neq 0 (
    echo      ERROR: git add fallo. Verifica que git este instalado.
    goto :error
)
echo      Done.

echo.
echo [4/6] Verificando cambios pendientes...
git status --short

echo.
echo [5/6] Haciendo commit...
git commit -m "Actualizacion datos %FECHA%"
if %errorlevel% neq 0 (
    echo      AVISO: No hay cambios nuevos para commitear, o ya estaba actualizado.
    echo      Intentando push de todas formas...
)

echo.
echo [6/6] Publicando en GitHub Pages...
git push origin main
if %errorlevel% neq 0 (
    echo.
    echo      ERROR en git push. Posibles causas:
    echo      - Token de GitHub vencido: ve a github.com, Settings^> Developer settings^> Tokens
    echo      - Sin conexion a internet
    echo      - Conflicto en el repo
    goto :cloudflare
)
echo      GitHub actualizado correctamente!

:cloudflare
echo.
echo [Extra] Desplegando en Cloudflare Workers...
call npx wrangler deploy --name dashboard-jyd 2>nul
if %errorlevel% neq 0 (
    echo      Cloudflare no disponible o no autenticado.
    echo      Para re-autenticar: npx wrangler login
    echo      ^(El sitio en GitHub Pages ya quedo actualizado^)
) else (
    echo      Cloudflare deploy exitoso.
)

echo.
echo ================================================
echo  Proceso terminado al %FECHA%
echo  GitHub: https://dlopezu0221-byte.github.io/dashboard-jyd/
echo  Dominio: https://dashboard.grupoempresarialjd.com/
echo ================================================
goto :fin

:error
echo.
echo ================================================
echo  ERROR CRITICO - Revisa los mensajes arriba
echo ================================================

:fin
pause
