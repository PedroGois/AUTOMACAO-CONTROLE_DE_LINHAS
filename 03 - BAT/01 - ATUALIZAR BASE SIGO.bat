@echo off
setlocal
cd /d "%~dp0\.."
echo ETAPA 1 DE 4 - ATUALIZAR A BASE SIGO
echo.
if /I "%~1"=="--agendado" goto executar
echo Digite seu documento e sua senha quando solicitado, caso nao estejam configurados no arquivo .env.
echo.

:executar
python "02 - SCRIPTS\1_baixar_base_sigo.py"
set "CODIGO_ERRO=%ERRORLEVEL%"
if not "%CODIGO_ERRO%"=="0" (
  echo.
  echo ERRO: a base SIGO nao foi atualizada.
  endlocal & exit /b %CODIGO_ERRO%
)
echo.
echo CONCLUIDO: BASE_SIGO.xlsx foi atualizada em 04 - SAIDAS.
if /I not "%~1"=="--agendado" pause
endlocal & exit /b 0
