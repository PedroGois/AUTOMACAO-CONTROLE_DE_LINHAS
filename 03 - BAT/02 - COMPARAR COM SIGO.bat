@echo off
setlocal
cd /d "%~dp0\.."
echo ETAPA 2 DE 4 - COMPARAR A TELEFONIA COM O SIGO
echo.
if /I "%~1"=="--agendado" goto executar
echo Feche a planilha TELEFONIA antes de continuar.
pause

:executar
python "02 - SCRIPTS\2_comparar_telefonia_sigo.py"
set "CODIGO_ERRO=%ERRORLEVEL%"
if not "%CODIGO_ERRO%"=="0" endlocal & exit /b %CODIGO_ERRO%
echo.
echo CONCLUIDO: CPF, Nome e Status atualizados na aba Planos.
if /I not "%~1"=="--agendado" pause
endlocal & exit /b 0
