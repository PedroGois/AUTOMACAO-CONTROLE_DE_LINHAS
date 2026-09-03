@echo off
setlocal EnableExtensions
cd /d "%~dp0"
set "PYTHONIOENCODING=utf-8"
if /I "%~1"=="--agendado" goto atualizar

:menu
cls
echo ======================================================
echo              CONTROLE DE LINHAS
echo ======================================================
echo  1 - Atualizar dados e dashboard
echo  2 - Exportar base atualizada para a planilha original
echo  3 - Separar linhas para verificar por centro de custo
echo  4 - Gerar rascunhos de e-mail
echo  5 - Executar fluxo completo
echo  6 - Anonimizar dados do dashboard antes de publicar
echo  7 - Configurar atualizacao automatica
echo.
echo  0 - Sair
choice /c 12345670 /n /m "Escolha uma opcao: "
if errorlevel 8 goto sair
if errorlevel 7 goto agendar
if errorlevel 6 goto anonimizar
if errorlevel 5 goto completo
if errorlevel 4 goto emails
if errorlevel 3 goto separar
if errorlevel 2 goto exportar
if errorlevel 1 goto atualizar_interativo

:atualizar_interativo
call :atualizar
pause
goto menu

:atualizar
set "LOG_DIR=data\saidas\LOGS"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"
set "LOG_FILE=%LOG_DIR%\atualizacao_dashboard.log"
echo [%date% %time%] Inicio da atualizacao>> "%LOG_FILE%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "scripts\copiar_dados.ps1" >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto falhou
python "src\baixar_base_sigo.py" >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto falhou
python "src\comparar_telefonia_sigo.py" >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto falhou
python "src\gerar_dashboard.py" >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto falhou
echo [%date% %time%] Concluido com sucesso>> "%LOG_FILE%"
echo Atualizacao concluida. Dashboard: "dashboard\index.html"
exit /b 0

:exportar
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "scripts\exportar_base_original.ps1"
pause
goto menu

:separar
python "src\separar_por_centro_custo.py"
pause
goto menu

:emails
set /p PRAZO="Prazo para resposta (DD/MM/AAAA): "
if "%PRAZO%"=="" set "PRAZO=25/08/2026"
python "src\gerar_cobrancas.py" --telefonia "data\entrada\TELEFONIA.xlsx" --contatos "data\entrada\CONTATO CDC.xlsx" --saida "data\saidas\COBRANCA E-MAILS" --modo eml --prazo "%PRAZO%"
pause
goto menu

:anonimizar
python "dashboard\dados\anonimizar_base_dashboard.py"
pause
goto menu

:agendar
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "scripts\agendar_atualizacao.ps1"
pause
goto menu

:completo
call :atualizar
if errorlevel 1 goto erro_menu
python "src\separar_por_centro_custo.py"
if errorlevel 1 goto erro_menu
goto emails

:falhou
echo [%date% %time%] ERRO na etapa anterior>> "%LOG_FILE%"
echo Atualizacao interrompida. Consulte "%LOG_FILE%".
exit /b 1

:erro_menu
echo Operacao interrompida por um erro.
pause
goto menu

:sair
endlocal
exit /b 0
