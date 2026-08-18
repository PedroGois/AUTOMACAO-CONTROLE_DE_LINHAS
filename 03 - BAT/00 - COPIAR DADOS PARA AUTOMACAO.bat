@echo off
setlocal EnableExtensions
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0\00 - COPIAR DADOS PARA AUTOMACAO.ps1"
set "CODIGO_ERRO=%ERRORLEVEL%"

if not "%CODIGO_ERRO%"=="0" (
  echo.
  echo ERRO: consulte "04 - SAIDAS\LOGS\copia_dados.log"
  if /I not "%~1"=="--agendado" pause
  endlocal & exit /b %CODIGO_ERRO%
)

echo Planilhas atualizadas em 01 - DADOS.
if /I not "%~1"=="--agendado" pause
endlocal & exit /b 0
