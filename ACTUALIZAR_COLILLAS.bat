@echo off
chcp 65001 >nul
echo ================================================
echo  Inyectando colillas en dashboards de estudios
echo ================================================
cd /d "%~dp0"
python _inyectar_colillas_temp.py
echo.
echo ================================================
echo  Publicando en GitHub...
echo ================================================
git add estudios/cyv-studios837357/index.html estudios/fornax-studios345929/index.html estudios/goldonline078939/index.html
git commit -m "Actualizar colillas estudios"
git push origin main
echo.
echo Listo. Presiona una tecla para cerrar.
pause
