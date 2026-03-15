@echo off
chcp 65001 >nul
title Remates Judiciales - Pipeline Automatizado
color 0A

:: Detectar carpeta donde esta este .bat (funciona en cualquier disco/ruta)
set "BASE_DIR=%~dp0"
if "%BASE_DIR:~-1%"=="\" set "BASE_DIR=%BASE_DIR:~0,-1%"

echo.
echo  ============================================================
echo    SISTEMA DE ANALISIS DE REMATES JUDICIALES
echo  ============================================================
echo    Carpeta: %BASE_DIR%
echo.

:: Buscar el DOCX mas reciente en Diarios/
set "DOCX_DIR=%BASE_DIR%\Diarios"
set "DOCX_FILE="

if not exist "%DOCX_DIR%" (
    echo  [ERROR] No existe la carpeta Diarios\
    echo          Creala en: %DOCX_DIR%
    echo.
    pause
    exit /b 1
)

for /f "delims=" %%F in ('dir /b /o-d "%DOCX_DIR%\*.docx" 2^>nul') do (
    if not defined DOCX_FILE set "DOCX_FILE=%%F"
)

if not defined DOCX_FILE (
    echo  [ERROR] No se encontro ningun archivo .docx en:
    echo          %DOCX_DIR%
    echo.
    echo  Coloca el DOCX semanal en esa carpeta e intenta de nuevo.
    echo.
    pause
    exit /b 1
)

echo  Archivo encontrado: %DOCX_FILE%
echo.
echo  Presiona cualquier tecla para iniciar el pipeline...
echo  (Ctrl+C para cancelar)
echo.
pause >nul

:: Ejecutar pipeline
cd /d "%BASE_DIR%"
python main.py --docx "%DOCX_DIR%\%DOCX_FILE%"

echo.
echo  ============================================================
echo    PIPELINE COMPLETADO
echo  ============================================================
echo.

:: Abrir el reporte mas reciente
for /f "delims=" %%R in ('dir /b /o-d "%BASE_DIR%\Reportes\Reporte_*.xlsx" 2^>nul') do (
    echo  Abriendo reporte: %%R
    start "" "%BASE_DIR%\Reportes\%%R"
    goto :done
)

echo  No se encontro reporte generado.

:done
echo.
pause
