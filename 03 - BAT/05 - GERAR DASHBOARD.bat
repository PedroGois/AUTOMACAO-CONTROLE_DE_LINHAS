@echo off
setlocal
cd /d "%~dp0\.."
echo GERANDO DASHBOARD GERENCIAL...
python "02 - SCRIPTS\gerar_dashboard.py"
set "CODIGO_ERRO=%ERRORLEVEL%"
if not "%CODIGO_ERRO%"=="0" (
  echo.
  echo ERRO: nao foi possivel gerar o dashboard.
  endlocal & exit /b %CODIGO_ERRO%
)
echo.
echo CONCLUIDO: Dashboard atualizado na raiz (DASH.html).
pause
endlocal & exit /b 0