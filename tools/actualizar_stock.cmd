@echo off
rem Arrastra el archivo de stock (.xlsb / .xlsx) sobre este .cmd, o ejecuta:  actualizar_stock.cmd "C:\ruta\archivo.xlsb"
if "%~1"=="" echo Uso: arrastra el archivo de stock sobre este .cmd  ^|  actualizar_stock.cmd "C:\ruta\archivo.xlsb" & pause & exit /b 1
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0actualizar_stock.ps1" -Path "%~1"
echo.
pause
