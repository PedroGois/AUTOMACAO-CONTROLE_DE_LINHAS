@echo off
setlocal EnableExtensions

set "TAREFA=Telefonia - Atualizar Dashboard"
set "EXECUTAR=%~dp0\06 - ATUALIZAR DASHBOARD AUTOMATICO.bat"
set "AGENDADOR_PS=%TEMP%\agendar_dashboard_telefonia.ps1"

> "%AGENDADOR_PS%" echo $acao = New-ScheduledTaskAction -Execute $env:ComSpec -Argument ('/c ""{0}"" --agendado' -f $env:EXECUTAR)
>> "%AGENDADOR_PS%" echo $inicio = (Get-Date).AddMinutes(1)
>> "%AGENDADOR_PS%" echo $gatilho = New-ScheduledTaskTrigger -Once -At $inicio -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration (New-TimeSpan -Days 3650)
>> "%AGENDADOR_PS%" echo Register-ScheduledTask -TaskName $env:TAREFA -Action $acao -Trigger $gatilho -Description 'Atualiza SIGO, compara a telefonia e gera o Dashboard a cada hora.' -Force ^| Out-Null

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
echo O fluxo sera executado a cada uma hora enquanto o computador estiver ligado.
echo.
pause
endlocal & exit /b 0
