@echo off
echo ============================================
echo   FILTRADOR DE SALDOS - Ejecucion completa
echo ============================================
echo.
cd /d D:\Remates
python filtrador_saldos.py --workers 5
echo.
echo ============================================
echo   Finalizado. Presione una tecla para cerrar.
echo ============================================
pause
