@echo off
setlocal
cd /d "%~dp0"
echo descarga_resultados_100X - ICFES Saber 11
echo Listado: Calendario B con grados 11 o 26
echo.
python descarga_resultados_100X.py ^
  --entrada "EstablecimientosCalendarioB_Grados_11_26.csv" ^
  --salida "resultados_icfes_2026_calendario_B_100_hilos.xlsx" ^
  --hilos 100
set "ICFES_EXIT=%ERRORLEVEL%"
echo.
if "%ICFES_EXIT%"=="0" (
  echo Proceso terminado correctamente.
) else if "%ICFES_EXIT%"=="4" (
  echo Proceso pausado por proteccion HTTP. Ejecute nuevamente mas tarde.
) else if "%ICFES_EXIT%"=="5" (
  echo Quedaron consultas temporales pendientes. Ejecute nuevamente.
) else (
  echo El proceso termino con codigo %ICFES_EXIT%.
)
pause
exit /b %ICFES_EXIT%


