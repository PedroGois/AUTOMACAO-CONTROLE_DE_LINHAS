@echo off
setlocal
cd /d "%~dp0\.."
echo ETAPA 2 DE 4 - COMPARAR A TELEFONIA COM O SIGO
echo.
echo Feche a planilha TELEFONIA antes de continuar.
pause
python "02 - SCRIPTS\comparar_telefonia_sigo.py"
set "CODIGO_ERRO=%ERRORLEVEL%"
if not "%CODIGO_ERRO%"=="0" endlocal & exit /b %CODIGO_ERRO%
echo.
echo CONCLUIDO: CPF, Nome e Status atualizados na aba Planos.
pause
endlocal & exit /b 0