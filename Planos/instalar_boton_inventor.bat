@echo off
setlocal

cd /d "%~dp0"

echo ============================================
echo  COTAS ABIGAIL - Instalar boton en Inventor
echo ============================================
echo.
echo IMPORTANTE: cierra Inventor antes de continuar.
echo Si Inventor esta abierto, al cerrarlo borrara este registro.
echo.

tasklist /FI "IMAGENAME eq Inventor.exe" 2>nul | find /I "Inventor.exe" >nul
if %ERRORLEVEL% equ 0 (
    echo ADVERTENCIA: Inventor esta abierto en este momento.
    echo Cierralo y vuelve a ejecutar este instalador.
    echo.
    pause
    goto :fin
)

where python >nul 2>&1
if %ERRORLEVEL% equ 0 (
    python instalar_boton_inventor.py
    goto :fin
)

where py >nul 2>&1
if %ERRORLEVEL% equ 0 (
    py instalar_boton_inventor.py
    goto :fin
)

echo ERROR: No se encontro Python en el PATH.
echo.

:fin
echo.
pause
endlocal
