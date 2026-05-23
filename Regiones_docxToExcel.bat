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

:: Verificar que haya DOCX en Diarios/ (se procesan TODOS)
set "DOCX_DIR=%BASE_DIR%\Diarios"

if not exist "%DOCX_DIR%" (
    echo  [ERROR] No existe la carpeta Diarios\
    echo          Creala en: %DOCX_DIR%
    echo.
    pause
    exit /b 1
)

set "DOCX_COUNT=0"
for /f "delims=" %%F in ('dir /b "%DOCX_DIR%\*.docx" 2^>nul') do (
    set /a DOCX_COUNT+=1
)

if %DOCX_COUNT%==0 (
    echo  [ERROR] No se encontro ningun archivo .docx en:
    echo          %DOCX_DIR%
    echo.
    echo  Coloca los DOCX semanales en esa carpeta e intenta de nuevo.
    echo.
    pause
    exit /b 1
)

echo  DOCX encontrados: %DOCX_COUNT%
echo.
echo  Presiona cualquier tecla para iniciar el pipeline...
echo  (Ctrl+C para cancelar)
echo.
pause >nul

:: Ejecutar pipeline (modo carpeta: procesa TODOS los DOCX y los elimina al terminar)
cd /d "%BASE_DIR%"
python main.py --docx-dir "%DOCX_DIR%"

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
