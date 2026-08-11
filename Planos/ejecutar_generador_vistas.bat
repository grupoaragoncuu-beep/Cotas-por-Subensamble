@echo off
setlocal

cd /d "%~dp0"

echo ============================================
echo  COTAS ABIGAIL - Generador de vistas
echo ============================================
echo.
echo Carpeta: %CD%
echo.
echo Requisitos antes de continuar:
echo   1. Autodesk Inventor abierto
echo   2. Plano .idw o .dwg activo y guardado
echo   3. Regla iLogic "GenerarVistas" en el documento
echo.

where python >nul 2>&1
if %ERRORLEVEL% equ 0 (
    python generador_vistas.py
    goto :fin
)

where py >nul 2>&1
if %ERRORLEVEL% equ 0 (
    py generador_vistas.py
    goto :fin
)

echo ERROR: No se encontro Python en el PATH.
echo Instala Python o agregalo al PATH del sistema.
echo.

:fin
echo.
pause
endlocal
