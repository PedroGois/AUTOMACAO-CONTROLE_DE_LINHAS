@echo off
setlocal EnableExtensions
cd /d "%~dp0\.."

echo Esta acao anonimiza nomes, linhas e centros de custo do Dashboard.
echo Use-a somente depois de concluir a conferencia e antes de enviar o Dashboard ao Git.
echo Backups locais serao criados automaticamente antes da alteracao.
choice /c SN /n /m "Deseja continuar? (S/N): "
if errorlevel 2 endlocal & exit /b 0

python "00 - DASHBOARD\dados\anonimizar_base_dashboard.py"
set "CODIGO_ERRO=%ERRORLEVEL%"
if not "%CODIGO_ERRO%"=="0" (
  echo.
  echo ERRO: nao foi possivel anonimizar a base do Dashboard.
  pause
  endlocal & exit /b %CODIGO_ERRO%
)

echo.
echo CONCLUIDO: a base do Dashboard foi anonimizada e esta pronta para revisao antes do envio ao Git.
pause
endlocal & exit /b 0
