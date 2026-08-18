@echo off
setlocal
cd /d "%~dp0\.."
echo ATUALIZANDO BASE DO DASHBOARD...
python "02 - SCRIPTS\5_gerar_dashboard.py"
set "CODIGO_ERRO=%ERRORLEVEL%"
if not "%CODIGO_ERRO%"=="0" (
  echo.
  echo ERRO: nao foi possivel atualizar a base do dashboard.
  endlocal & exit /b %CODIGO_ERRO%
)
echo.
echo CONCLUIDO: Base atualizada com sucesso! Abra DASH.html para visualizar.
pause
endlocal & exit /b 0