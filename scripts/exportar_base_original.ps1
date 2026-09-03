# Exporta a planilha de telefonia atualizada pela automacao para a origem corporativa.

$PastaScript = Split-Path -Parent $MyInvocation.MyCommand.Path
$PastaProjeto = Split-Path $PastaScript -Parent
$PastaTi = Split-Path (Split-Path (Split-Path $PastaProjeto -Parent) -Parent) -Parent

$BaseAtual = Join-Path $PastaProjeto "data\entrada\TELEFONIA.xlsx"
$BaseOriginal = Join-Path $PastaTi "Telefonia e Internet\TELEFONIA\telefonia.xlsx"
$PastaBackup = Join-Path $PastaProjeto "data\saidas\BACKUPS\EXPORTACAO_BASE_ORIGINAL"
$PastaLog = Join-Path $PastaProjeto "data\saidas\LOGS"
$ArquivoLog = Join-Path $PastaLog "exportacao_base_original.log"

try {
    if (-not (Test-Path -LiteralPath $BaseAtual)) {
        throw "Base atualizada nao encontrada: $BaseAtual"
    }
    if (-not (Test-Path -LiteralPath $BaseOriginal)) {
        throw "Planilha original nao encontrada: $BaseOriginal"
    }

    New-Item -ItemType Directory -Path $PastaBackup -Force | Out-Null
    New-Item -ItemType Directory -Path $PastaLog -Force | Out-Null

    $DataHoraArquivo = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $Backup = Join-Path $PastaBackup "telefonia_antes_da_exportacao_$DataHoraArquivo.xlsx"
    Copy-Item -LiteralPath $BaseOriginal -Destination $Backup -Force
    Copy-Item -LiteralPath $BaseAtual -Destination $BaseOriginal -Force

    $DataHora = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    Add-Content -LiteralPath $ArquivoLog -Value "$DataHora - SUCESSO: base exportada para a planilha original. Backup: $Backup"
    Write-Output "Base original atualizada: $BaseOriginal"
    Write-Output "Backup criado: $Backup"
    exit 0
}
catch {
    New-Item -ItemType Directory -Path $PastaLog -Force | Out-Null
    $DataHora = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    Add-Content -LiteralPath $ArquivoLog -Value "$DataHora - ERRO: $($_.Exception.Message)"
    Write-Error $_.Exception.Message
    exit 1
}
