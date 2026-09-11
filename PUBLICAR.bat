@echo off
chcp 65001 >nul
setlocal
echo ================================================
echo  Publicando dashboards - Grupo J^&D
echo ================================================

cd /d "%~dp0."

:: Fecha actual para el commit.
:: Antes se usaba WMIC, que Windows ya elimino: el commit salia sin fecha.
set FECHA=%DATE%

echo.
echo [1/5] Eliminando locks si existen...
if exist ".git\index.lock" (
    del /f /q ".git\index.lock"
    echo      index.lock eliminado.
) else (
    echo      Sin locks, OK.
)
if exist ".git\HEAD.lock" del /f /q ".git\HEAD.lock"

echo.
echo [2/5] Anadiendo todos los cambios al staging...
git add -A
if errorlevel 1 goto :error
echo      Incluye: grupo, erika, fabio, estudios, modelos, monitores, elite
echo      Done.

echo.
echo [3/5] Cambios pendientes:
git status --short

echo.
echo [4/5] Haciendo commit...
git commit -m "Actualizacion datos %FECHA%"
if errorlevel 1 (
    echo      Sin cambios que commitear ^(o commit fallido^). Se intenta el push igual.
)

echo.
echo [5/5] Publicando en GitHub...
git push origin main
if errorlevel 1 goto :error

echo.
echo ================================================
echo  Push enviado. El despliegue lo hace GitHub.
echo ================================================
echo.
echo  La carpeta dist/ que publica Cloudflare se construye
echo  SOLA dentro de GitHub Actions. No hace falta node,
echo  ni npx, ni robocopy en este equipo.
echo.
echo  Verifica que el despliegue termine en verde aqui:
echo    https://github.com/dlopezu0221-byte/dashboard-jyd/actions
echo.
echo  Tarda entre 30 y 60 segundos. Despues abre el link
echo  con Ctrl+Shift+R ^(o en ventana de incognito^):
echo    https://dashboard.grupoempresarialjd.com/
echo    https://dashboard.grupoempresarialjd.com/monitores/
echo.
echo  Y confirma que el sello diga "Datos al:" con la fecha
echo  de corte que esperas. Si dice una fecha vieja, es cache:
echo  agrega ?v=1 al final de la URL.
echo ================================================
goto :fin

:error
echo.
echo ================================================
echo  ERROR en git. Revisa el mensaje de arriba.
echo  Nada se publico.
echo ================================================

:fin
echo.
pause
endlocal
