@echo off
setlocal
cd /d "%~dp0"

:menu
cls
set "TOTAL_VERIFICADO="
if exist "04 - SAIDAS\resultado_comparacao.txt" (
  for /f "usebackq tokens=1,2 delims==" %%A in ("04 - SAIDAS\resultado_comparacao.txt") do (
    if /I "%%A"=="TOTAL_VERIFICADO" set "TOTAL_VERIFICADO=%%B"
  )
)
echo ======================================================
echo          MENU - AUTOMACAO DE TELEFONIA
echo ======================================================
if defined TOTAL_VERIFICADO echo Ultima comparacao: %TOTAL_VERIFICADO% linhas verificadas (inclui FROTA)
echo.
echo  1 - Atualizar dados e Dashboard Gerencial
echo  2 - Exportar base atualizada para a planilha original
echo  3 - Separar linhas VERIFICAR por CDC
echo  4 - Gerar e-mails EML da VIVO e TIM
echo  5 - Executar fluxo completo
echo  6 - Anonimizar base do Dashboard para Git
echo.
echo  0 - Sair
echo.
choice /c 1234560 /n /m "Escolha uma opcao: "

if errorlevel 7 goto sair
if errorlevel 6 goto anonimizar_dashboard
if errorlevel 5 goto fluxo_completo
if errorlevel 4 goto gerar_emails
if errorlevel 3 goto separar_cdc
if errorlevel 2 goto exportar_base_original
if errorlevel 1 goto atualizar_dashboard

:atualizar_dashboard
call "03 - BAT\06 - ATUALIZAR DASHBOARD AUTOMATICO.bat"
goto menu

:exportar_base_original
call "03 - BAT\01 - EXPORTAR BASE PARA ORIGINAL.bat"
goto menu

:separar_cdc
call "03 - BAT\03 - SEPARAR POR CDC.bat"
goto menu

:gerar_emails
call "03 - BAT\04 - GERAR E-MAILS EML.bat"
goto menu

:anonimizar_dashboard
call "03 - BAT\07 - ANONIMIZAR BASE DASHBOARD.bat"
goto menu

:fluxo_completo
cls
echo FLUXO COMPLETO
echo.
call "03 - BAT\06 - ATUALIZAR DASHBOARD AUTOMATICO.bat"
if errorlevel 1 goto falhou
call "03 - BAT\03 - SEPARAR POR CDC.bat"
if errorlevel 1 goto falhou
echo.
call "03 - BAT\04 - GERAR E-MAILS EML.bat"
echo.
echo FLUXO COMPLETO FINALIZADO!
pause
goto menu

:falhou
echo.
echo ======================================================
echo O processo foi interrompido por causa de um erro.
echo ======================================================
pause
goto menu

:sair
endlocal
exit /b 0
