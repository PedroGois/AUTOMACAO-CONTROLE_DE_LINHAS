@echo off
setlocal EnableExtensions

set "TAREFA=Telefonia - Atualizacao Completa"
set "TAREFA_ANTIGA_COPIA=Telefonia - Copiar Dados de Origem"
set "TAREFA_ANTIGA_DASHBOARD=Telefonia - Atualizar Dashboard"
rem Este arquivo fica em 03 - BAT\AGENDAR; volte uma pasta para executar o fluxo completo.
set "EXECUTAR=%~dp0..\06 - ATUALIZAR DASHBOARD AUTOMATICO.bat"
set "AGENDADOR_PS=%TEMP%\agendar_atualizacao_completa_telefonia.ps1"

> "%AGENDADOR_PS%" echo $acao = New-ScheduledTaskAction -Execute $env:ComSpec -Argument ('/c ""{0}"" --agendado' -f '%EXECUTAR%')
>> "%AGENDADOR_PS%" echo $inicio = (Get-Date).AddMinutes(1)
>> "%AGENDADOR_PS%" echo $gatilho = New-ScheduledTaskTrigger -Once -At $inicio -RepetitionInterval (New-TimeSpan -Minutes 30) -RepetitionDuration (New-TimeSpan -Days 3650)
>> "%AGENDADOR_PS%" echo Unregister-ScheduledTask -TaskName '%TAREFA_ANTIGA_COPIA%' -Confirm:$false -ErrorAction SilentlyContinue
>> "%AGENDADOR_PS%" echo Unregister-ScheduledTask -TaskName '%TAREFA_ANTIGA_DASHBOARD%' -Confirm:$false -ErrorAction SilentlyContinue
>> "%AGENDADOR_PS%" echo Register-ScheduledTask -TaskName '%TAREFA%' -Action $acao -Trigger $gatilho -Description 'Copia as planilhas, atualiza SIGO, compara a telefonia e gera o Dashboard a cada 30 minutos.' -Force ^| Out-Null

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%AGENDADOR_PS%"
set "CODIGO_ERRO=%ERRORLEVEL%"
del "%AGENDADOR_PS%" >nul 2>&1

if not "%CODIGO_ERRO%"=="0" (
  echo.
  echo Nao foi possivel criar a tarefa. Execute este arquivo como o usuario que usara a automacao.
  pause
  endlocal & exit /b %CODIGO_ERRO%
)

echo.
echo Tarefa "%TAREFA%" criada ou atualizada.
echo As tarefas antigas de copia e dashboard foram removidas, caso existissem.
echo O fluxo completo sera executado a cada 30 minutos enquanto o computador estiver ligado.
echo.
pause
endlocal & exit /b 0
