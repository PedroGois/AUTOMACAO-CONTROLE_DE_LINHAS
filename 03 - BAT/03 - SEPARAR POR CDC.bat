@echo off
setlocal
cd /d "%~dp0\.."
echo ETAPA 3 DE 4 - SEPARAR AS LINHAS VERIFICAR POR CDC
echo.
python "02 - SCRIPTS\separar_verificar_por_centro_custo.py"
set "CODIGO_ERRO=%ERRORLEVEL%"
if not "%CODIGO_ERRO%"=="0" endlocal & exit /b %CODIGO_ERRO%
echo.
echo CONCLUIDO: confira a pasta 04 - SAIDAS\VERIFICAR POR CENTRO DE CUSTO.
pause
endlocal & exit /b 0