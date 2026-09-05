@echo off
chcp 65001 >nul
title RESTAURAR SISTEMA — VERSION 0.1

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║        GRUPO EMPRESARIAL J&D — RESTAURAR VERSIÓN 0.1        ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo  Este script restaurará el sistema al estado exacto de VERSION 0.1.
echo  El archivo gerencia-general.html será reemplazado.
echo.
echo  ADVERTENCIA: Los cambios realizados después de la VERSION 0.1
echo  serán revertidos. Esta acción NO afecta la base de datos Supabase.
echo.
echo  Archivos involucrados:
echo    ORIGEN:  VERSION_0.1\gerencia-general_v0.1.html
echo    DESTINO: gerencia-general.html
echo.

set /p CONFIRM=Escriba SI para confirmar la restauracion:
if /i not "%CONFIRM%"=="SI" (
    echo.
    echo  Restauracion cancelada.
    pause
    exit /b 0
)

echo.
echo  Verificando archivo de restauracion...

if not exist "%~dp0gerencia-general_v0.1.html" (
    echo.
    echo  ERROR: No se encuentra gerencia-general_v0.1.html en esta carpeta.
    echo  Asegurese de ejecutar este script desde la carpeta VERSION_0.1.
    echo.
    pause
    exit /b 1
)

echo  Archivo VERSION 0.1 encontrado. Verificando integridad...

REM Verificar hash SHA-256
for /f "tokens=*" %%A in ('certutil -hashfile "%~dp0gerencia-general_v0.1.html" SHA256 ^| findstr /v "hash" ^| findstr /v "CertUtil"') do set HASH=%%A
set HASH=%HASH: =%

if "%HASH%"=="fa8a9d5a0902eda7eb3274cf1f57f69bf257b01a6a21515d184628a29a4e6b1d" (
    echo  Hash SHA-256 verificado: OK
) else (
    echo.
    echo  ADVERTENCIA: El hash del archivo no coincide con el esperado.
    echo  Hash encontrado:  %HASH%
    echo  Hash esperado:    fa8a9d5a0902eda7eb3274cf1f57f69bf257b01a6a21515d184628a29a4e6b1d
    echo.
    echo  Es posible que el archivo VERSION 0.1 haya sido modificado.
    set /p FORCECOPY=Continuar de todas formas? (SI/NO):
    if /i not "!FORCECOPY!"=="SI" (
        echo  Restauracion cancelada por seguridad.
        pause
        exit /b 1
    )
)

REM Hacer backup del archivo actual antes de reemplazarlo
set BACKUP_NAME=gerencia-general_ANTES_RESTAURAR_%date:~-4%-%date:~3,2%-%date:~0,2%.html
echo  Creando backup del archivo actual como: %BACKUP_NAME%
copy /Y "%~dp0..\gerencia-general.html" "%~dp0..\%BACKUP_NAME%" >nul 2>&1

REM Copiar la VERSION 0.1 al directorio raíz
echo  Restaurando VERSION 0.1...
copy /Y "%~dp0gerencia-general_v0.1.html" "%~dp0..\gerencia-general.html" >nul

if %ERRORLEVEL% EQU 0 (
    echo.
    echo  ╔══════════════════════════════════════════════════════════════╗
    echo  ║           RESTAURACION COMPLETADA EXITOSAMENTE              ║
    echo  ╚══════════════════════════════════════════════════════════════╝
    echo.
    echo  Sistema restaurado a VERSION 0.1.
    echo  Backup del archivo anterior: %BACKUP_NAME%
    echo.
    echo  PASOS SIGUIENTES:
    echo    1. Abra el dashboard en el navegador
    echo    2. Recargue con Ctrl+F5 (forzar recarga sin cache)
    echo    3. Inicie sesion normalmente
    echo    4. Verifique en pestaña Bancolombia que el saldo es ~$33.042.000 COP
    echo    5. Abra consola (F12) y busque mensajes [RESTORE] para confirmar
    echo.
) else (
    echo.
    echo  ERROR: No se pudo copiar el archivo. Verifique permisos.
    echo.
)

pause
