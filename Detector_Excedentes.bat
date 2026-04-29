@echo off
REM ============================================================
REM   FILTRADOR DE SALDOS - Ejecucion completa
REM   v2 (2026-04-29): incluye reaudit automatico al final.
REM
REM   El reaudit recupera causas con error OJV transitorio
REM   ("Tabla de tramitacion vacia"), que ocurre por timing
REM   del WAF del PJUD bajo volumen alto.
REM ============================================================

setlocal
cd /d D:\Remates

echo ============================================
echo   FASE 1 / 2: Run normal (merge + F1 + F2 + F3)
echo ============================================
echo.

python filtrador_saldos.py --workers 5
set RC=%errorlevel%

echo.
if %RC% NEQ 0 (
    echo ============================================
    echo   ERROR en Fase 1 ^(exit code %RC%^)
    echo   No se ejecutara el reaudit.
    echo   Revise el log antes de continuar.
    echo ============================================
    echo.
    echo   Presione una tecla para cerrar.
    pause
    exit /b %RC%
)

echo ============================================
echo   FASE 2 / 2: Reaudit ^(recuperacion errores OJV^)
echo ============================================
echo.
echo   Esto reprocesa PENDIENTE_ACTA y PENDIENTE_LIQUIDACION
echo   con sesiones ya calientes y volumen menor.
echo   Tipicamente recupera ^>75%% de los errores transitorios.
echo.

python filtrador_saldos.py --reaudit --workers 5
set RC2=%errorlevel%

echo.
if %RC2% NEQ 0 (
    echo ============================================
    echo   ATENCION: Reaudit termino con error ^(exit code %RC2%^)
    echo   El run normal SI completo. Revisar log del reaudit.
    echo ============================================
    echo.
)

echo.
echo ============================================
echo   Ejecucion completa finalizada.
echo ============================================
echo.
echo ============================================
echo   PASO SIGUIENTE
echo ============================================
echo.
echo   Adjuntar el Excel actualizado a Claude
echo   para generar el Informe de Auditoria
echo   para los abogados.
echo.
echo   Excel: Causas con liq\Causas_posible_saldo.xlsx
echo.
echo ============================================
echo.
echo   Presione una tecla para cerrar.
pause
endlocal
