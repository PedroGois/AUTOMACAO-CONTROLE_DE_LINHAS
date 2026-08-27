@echo off
setlocal
cd /d "%~dp0\.."
set "OPERADORA=%~1"
set "FILTRO_EMPRESA="

if /I "%OPERADORA%"=="VIVO" set "FILTRO_EMPRESA=--empresa VIVO"
if /I "%OPERADORA%"=="TIM" set "FILTRO_EMPRESA=--empresa TIM"

echo.
set /p PRAZO="Informe o prazo para resposta no corpo do e-mail (DD/MM/AAAA): "
if "%PRAZO%"=="" set "PRAZO=25/08/2026"

echo.
if defined FILTRO_EMPRESA (
  echo Gerando rascunhos EML da %OPERADORA% com prazo ate %PRAZO%...
) else (
  echo Gerando rascunhos EML da VIVO e da TIM com prazo ate %PRAZO%...
)
python "02 - SCRIPTS\4_cobranca_verificar.py" --telefonia "01 - DADOS\TELEFONIA.xlsx" --contatos "01 - DADOS\CONTATO CDC.xlsx" --saida "04 - SAIDAS\COBRANCA E-MAILS" --modo eml --prazo "%PRAZO%" %FILTRO_EMPRESA%
set "CODIGO_ERRO=%ERRORLEVEL%"
if not "%CODIGO_ERRO%"=="0" (
  echo.
  echo ERRO: os rascunhos nao foram gerados.
  endlocal & exit /b %CODIGO_ERRO%
)
echo.
echo CONCLUIDO: confira a pasta 04 - SAIDAS\COBRANCA E-MAILS\rascunhos_eml.
pause
endlocal & exit /b 0
