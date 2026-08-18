@echo off
setlocal EnableExtensions
cd /d "%~dp0\.."
set "PYTHONIOENCODING=utf-8"

set "LOG_DIR=04 - SAIDAS\LOGS"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"
set "LOG_FILE=%LOG_DIR%\atualizacao_dashboard.log"

echo ======================================================>> "%LOG_FILE%"
echo [%date% %time%] Inicio da atualizacao automatica>> "%LOG_FILE%"

echo Atualizando a base SIGO...
python "02 - SCRIPTS\1_baixar_base_sigo.py" >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto falhou

echo Comparando a telefonia com o SIGO...
python "02 - SCRIPTS\2_comparar_telefonia_sigo.py" >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto falhou

echo Atualizando os dados do Dashboard...
python "02 - SCRIPTS\5_gerar_dashboard.py" >> "%LOG_FILE%" 2>&1
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
