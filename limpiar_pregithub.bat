@echo off
echo ============================================================
echo   LIMPIEZA PRE-GITHUB — D:\Remates
echo ============================================================
echo.
echo Este script ELIMINA archivos y carpetas que no son parte del
echo proyecto en produccion. REVISAR antes de ejecutar.
echo.
pause

echo.
echo [1/8] Eliminando archivos obsoletos de raiz...
del /q "D:\Remates\modulo1_mercurio.py" 2>nul
del /q "D:\Remates\analizar_log.py" 2>nul
del /q "D:\Remates\causas_antes.json" 2>nul
del /q "D:\Remates\causas_despues.json" 2>nul
del /q "D:\Remates\tasacion_cache.db" 2>nul
echo    Hecho.

echo.
echo [2/8] Eliminando reportes viejos de raiz (ya estan en Reportes/)...
del /q "D:\Remates\Reporte_2026-02-28.xlsx" 2>nul
del /q "D:\Remates\Reporte_2026-03-01.xlsx" 2>nul
echo    Hecho.

echo.
echo [3/8] Eliminando archivos de test sandbox...
del /q "D:\Remates\TEST_RESUMEN_16_CAUSAS.docx" 2>nul
del /q "D:\Remates\TEST_SANDBOX_24_CAUSAS.docx" 2>nul
echo    Hecho.

echo.
echo [4/8] Eliminando backup viejo...
rmdir /s /q "D:\Remates\backup_20260301" 2>nul
echo    Hecho.

echo.
echo [5/8] Eliminando carpetas experimentales...
rmdir /s /q "D:\Remates\la API" 2>nul
rmdir /s /q "D:\Remates\htmls de PDJ" 2>nul
rmdir /s /q "D:\Remates\Diarios test" 2>nul
rmdir /s /q "D:\Remates\Diarios_Procesados" 2>nul
echo    Hecho.

echo.
echo [6/8] Eliminando .tessdata (OCR ya no necesario con DOCX)...
rmdir /s /q "D:\Remates\.tessdata" 2>nul
echo    Hecho.

echo.
echo [7/8] Eliminando __pycache__...
rmdir /s /q "D:\Remates\__pycache__" 2>nul
echo    Hecho.

echo.
echo [8/8] Verificando estructura final...
echo.
echo === ARCHIVOS QUE DEBEN QUEDAR ===
dir /b "D:\Remates\*.py" "D:\Remates\*.md" "D:\Remates\*.xlsx" 2>nul
echo.
echo === CARPETAS QUE DEBEN QUEDAR ===
echo    .claude\         (CC settings)
echo    .playwright_profile\  (browser OJV)
echo    Descargas\       (PDFs descargados)
echo    Diarios\         (PDFs diarios fallback)
echo    logs\            (logs de ejecucion)
echo    Reportes\        (reportes generados)
echo.
echo ============================================================
echo   LIMPIEZA COMPLETADA
echo   Ahora: reemplazar PROGRESO.md y CLAUDE_CODE_PROMPT.md
echo   con las versiones actualizadas, agregar .gitignore y
echo   config_template.py, luego: git init, git add, git commit
echo ============================================================
pause
