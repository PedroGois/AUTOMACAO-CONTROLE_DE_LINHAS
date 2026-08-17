@echo off
setlocal
cd /d "%~dp0"

:menu
cls
set "TOTAL_VERIFICADO="
if exist "04 - SAIDAS\resultado_comparacao.txt" (
  for /f "tokens=1,2 delims==" %%A in (04 - SAIDAS\resultado_comparacao.txt) do (
    if /I "%%A"=="TOTAL_VERIFICADO" set "TOTAL_VERIFICADO=%%B"
  )
)
echo ======================================================
echo          MENU - AUTOMACAO DE TELEFONIA
echo ======================================================
if defined TOTAL_VERIFICADO echo Ultima comparacao: %TOTAL_VERIFICADO% linhas verificadas (inclui FROTA)
echo.
echo  1 - Atualizar a base SIGO
echo  2 - Comparar Telefonia com SIGO
echo  3 - Separar linhas VERIFICAR por CDC
echo  4 - Gerar e-mails EML da VIVO
echo  5 - Gerar e-mails EML da TIM
echo  6 - Executar fluxo completo
echo  7 - Atualizar Dashboard Gerencial
echo.
echo  0 - Sair
echo.
choice /c 12345670 /n /m "Escolha uma opcao: "

if errorlevel 8 goto sair
if errorlevel 7 goto gerar_dashboard
if errorlevel 6 goto fluxo_completo
if errorlevel 5 goto gerar_eml_tim
if errorlevel 4 goto gerar_eml_vivo
if errorlevel 3 goto separar_cdc
if errorlevel 2 goto comparar_sigo
if errorlevel 1 goto atualizar_sigo

:gerar_dashboard
call "03 - BAT\05 - GERAR DASHBOARD.bat"
goto menu

:atualizar_sigo
call "03 - BAT\01 - ATUALIZAR BASE SIGO.bat"
goto menu

:comparar_sigo
call "03 - BAT\02 - COMPARAR COM SIGO.bat"
goto menu

:separar_cdc
call "03 - BAT\03 - SEPARAR POR CDC.bat"
goto menu

:gerar_eml_vivo
call "03 - BAT\04 - GERAR E-MAILS EML.bat" VIVO
goto menu

:gerar_eml_tim
call "03 - BAT\04 - GERAR E-MAILS EML.bat" TIM
goto menu

:fluxo_completo
cls
echo FLUXO COMPLETO
echo.
call "03 - BAT\01 - ATUALIZAR BASE SIGO.bat"
if errorlevel 1 goto falhou
call "03 - BAT\02 - COMPARAR COM SIGO.bat"
if errorlevel 1 goto falhou
call "03 - BAT\03 - SEPARAR POR CDC.bat"
if errorlevel 1 goto falhou
echo.
echo Qual operadora voce quer cobrar agora? (1 - VIVO / 2 - TIM)
choice /c 12 /n /m "Escolha a operadora: "
if errorlevel 2 call "03 - BAT\04 - GERAR E-MAILS EML.bat" TIM
if errorlevel 1 call "03 - BAT\04 - GERAR E-MAILS EML.bat" VIVO
call "03 - BAT\05 - GERAR DASHBOARD.bat"
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