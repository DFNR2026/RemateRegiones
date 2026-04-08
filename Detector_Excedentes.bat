@echo off
echo ============================================
echo   FILTRADOR DE SALDOS - Ejecucion completa
echo ============================================
echo.
cd /d D:\Remates
python filtrador_saldos.py --workers 5
echo.
echo ============================================
echo   Finalizado.
echo ============================================
echo.
echo ============================================
echo   IMPORTANTE: PASO SIGUIENTE
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
