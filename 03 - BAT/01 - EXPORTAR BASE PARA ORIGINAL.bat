@echo off
setlocal EnableExtensions
cd /d "%~dp0\.."

echo Esta acao substitui a planilha corporativa original pela base atualizada da automacao.
echo Um backup da planilha original sera criado em 04 - SAIDAS\BACKUPS\EXPORTACAO_BASE_ORIGINAL.
choice /c SN /n /m "Deseja continuar? (S/N): "
if errorlevel 2 endlocal ^& exit /b 0

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "03 - BAT\01 - EXPORTAR BASE PARA ORIGINAL.ps1"
set "CODIGO_ERRO=%ERRORLEVEL%"
if not "%CODIGO_ERRO%"=="0" (
  echo.
  echo ERRO: consulte "04 - SAIDAS\LOGS\exportacao_base_original.log"
  pause
  endlocal ^& exit /b %CODIGO_ERRO%
)

echo.
echo CONCLUIDO: a planilha original foi atualizada.
pause
endlocal ^& exit /b 0
