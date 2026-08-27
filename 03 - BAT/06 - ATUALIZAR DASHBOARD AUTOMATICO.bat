@echo off
setlocal EnableExtensions
cd /d "%~dp0\.."
set "PYTHONIOENCODING=utf-8"

set "LOG_DIR=04 - SAIDAS\LOGS"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"
set "LOG_FILE=%LOG_DIR%\atualizacao_dashboard.log"

echo ======================================================>> "%LOG_FILE%"
echo [%date% %time%] Inicio da atualizacao automatica>> "%LOG_FILE%"

echo Copiando as planilhas de origem...
call "03 - BAT\00 - COPIAR DADOS PARA AUTOMACAO.bat" --agendado >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto falhou

echo Atualizando a base SIGO...
call "03 - BAT\01 - ATUALIZAR BASE SIGO.bat" --agendado >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto falhou

echo Comparando a telefonia com o SIGO...
call "03 - BAT\02 - COMPARAR COM SIGO.bat" --agendado >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto falhou

echo Atualizando os dados do Dashboard...
call "03 - BAT\05 - GERAR DASHBOARD.bat" --agendado >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto falhou

echo [%date% %time%] Concluido com sucesso>> "%LOG_FILE%"
if /I not "%~1"=="--agendado" pause
endlocal & exit /b 0

:falhou
echo [%date% %time%] ERRO na etapa anterior. Consulte o log.>> "%LOG_FILE%"
echo.
echo ERRO: consulte "%LOG_FILE%"
if /I not "%~1"=="--agendado" pause
endlocal & exit /b 1
