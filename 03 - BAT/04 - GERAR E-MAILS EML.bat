@echo off
setlocal
cd /d "%~dp0\.."
set "OPERADORA=%~1"

if /I "%OPERADORA%"=="VIVO" goto pedir_prazo
if /I "%OPERADORA%"=="TIM" goto pedir_prazo

cls
echo GERAR E-MAILS PARA CONFERENCIA
echo.
echo  1 - Gerar e-mails somente da VIVO
echo  2 - Gerar e-mails somente da TIM
echo.
choice /c 12 /n /m "Escolha a operadora: "
if errorlevel 2 set "OPERADORA=TIM"
if errorlevel 1 set "OPERADORA=VIVO"

:pedir_prazo
echo.
set /p PRAZO="Informe o prazo para resposta no corpo do e-mail (DD/MM/AAAA): "
if "%PRAZO%"=="" set "PRAZO=25/08/2026"

echo.
echo Gerando rascunhos EML da %OPERADORA% com prazo ate %PRAZO%...
python "02 - SCRIPTS\4_cobranca_verificar.py" --telefonia "01 - DADOS\TELEFONIA.xlsx" --contatos "01 - DADOS\CONTATO CDC-TESTE.xlsx" --saida "04 - SAIDAS\COBRANCA E-MAILS" --modo eml --prazo "%PRAZO%" --empresa "%OPERADORA%"
set "CODIGO_ERRO=%ERRORLEVEL%"
if not "%CODIGO_ERRO%"=="0" (
  echo.
  echo ERRO: os rascunhos da %OPERADORA% nao foram gerados.
  endlocal & exit /b %CODIGO_ERRO%
)
echo.
echo CONCLUIDO: confira a pasta 04 - SAIDAS\COBRANCA E-MAILS\rascunhos_eml.
pause
endlocal & exit /b 0