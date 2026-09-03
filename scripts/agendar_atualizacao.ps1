$ErrorActionPreference = 'Stop'

$NomeTarefa = 'Telefonia - Atualizacao Completa'
$TarefasAntigas = @('Telefonia - Copiar Dados de Origem', 'Telefonia - Atualizar Dashboard')

Add-Type @'
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class ShortPath {
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetShortPathName(string longPath, StringBuilder shortPath, int buffer);

    public static string Get(string longPath) {
        var buffer = new StringBuilder(32768);
        var result = GetShortPathName(longPath, buffer, buffer.Capacity);
        if (result == 0) throw new InvalidOperationException("Nao foi possivel obter o caminho curto do arquivo de atualizacao.");
        return buffer.ToString();
    }
}
'@

$Fluxo = Join-Path (Split-Path $PSScriptRoot -Parent) 'executar.bat'
if (-not (Test-Path -LiteralPath $Fluxo)) {
    throw "Arquivo de atualizacao nao encontrado: $Fluxo"
}

$FluxoCurto = [ShortPath]::Get($Fluxo)
$Acao = New-ScheduledTaskAction -Execute $env:ComSpec -Argument ('/d /c ""{0}"" --agendado' -f $FluxoCurto)
$Inicio = (Get-Date).AddMinutes(1)
$GatilhoRecorrente = New-ScheduledTaskTrigger -Once -At $Inicio -RepetitionInterval (New-TimeSpan -Minutes 30) -RepetitionDuration (New-TimeSpan -Days 3650)
$GatilhoLogon = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$Configuracoes = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive

foreach ($Tarefa in $TarefasAntigas) {
    Unregister-ScheduledTask -TaskName $Tarefa -Confirm:$false -ErrorAction SilentlyContinue
}

Register-ScheduledTask -TaskName $NomeTarefa -Action $Acao -Trigger @($GatilhoRecorrente, $GatilhoLogon) -Settings $Configuracoes -Principal $Principal -Description 'Atualiza a telefonia a cada 30 minutos e tambem apos o logon do usuario.' -Force | Out-Null

Write-Host 'Tarefa criada para executar o fluxo periodicamente e apos o logon.'
