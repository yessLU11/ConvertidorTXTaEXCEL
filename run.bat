@echo off
echo ========================================
echo Convertidor TXT → Excel
echo ========================================
echo.
echo Configuracion: Max archivos 1000MB
echo.
streamlit run app.py --server.maxUploadSize 1000 --server.enableCORS false --server.enableXsrfProtection false
pause