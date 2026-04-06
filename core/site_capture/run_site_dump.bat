@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "CONFIG_FILE=%SCRIPT_DIR%config.json"

if not "%~1"=="" set "CONFIG_FILE=%~1"

where py >nul 2>nul
if %errorlevel%==0 (
    py "%SCRIPT_DIR%main.py" --config "%CONFIG_FILE%"
) else (
    python "%SCRIPT_DIR%main.py" --config "%CONFIG_FILE%"
)

if errorlevel 1 (
    echo.
    echo El proceso termino con error.
    echo Revisa el mensaje anterior.
    pause
    exit /b %errorlevel%
)

echo.
echo Proceso finalizado.
pause
