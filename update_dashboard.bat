@echo off
REM Actualiza el dashboard Backlog Repara desde el CSV mas reciente
REM y publica los cambios en GitHub. Pensado para correr a diario
REM (doble clic o Programador de tareas de Windows).

cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0update_dashboard.ps1"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] La actualizacion fallo. Revisa update.log
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo Actualizacion completada correctamente.
timeout /t 5
