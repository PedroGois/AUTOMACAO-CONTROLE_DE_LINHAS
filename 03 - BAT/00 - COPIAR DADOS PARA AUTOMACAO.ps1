# ============================================================
# COPIAR AS PLANILHAS DE ORIGEM PARA A PASTA DA AUTOMACAO
# ============================================================

$PastaScript = Split-Path -Parent $MyInvocation.MyCommand.Path

# 03 - BAT
#   └── .. = AUTO - CONTROLE DE LINHAS
$PastaProjeto = Split-Path $PastaScript -Parent

# Origem:
# Ti\Telefonia e Internet\TELEFONIA\telefonia.xlsx
$PastaTi = Split-Path (Split-Path (Split-Path $PastaProjeto -Parent) -Parent) -Parent

$PastaOrigem = Join-Path $PastaTi "Telefonia e Internet\TELEFONIA"
$PastaDestino = Join-Path $PastaProjeto "01 - DADOS"
$Arquivos = @(
    @{ Nome = "TELEFONIA.xlsx"; Origem = (Join-Path $PastaOrigem "telefonia.xlsx") },
    @{ Nome = "CONTATO CDC.xlsx"; Origem = (Join-Path $PastaOrigem "CONTATO CDC.xlsx") }
)

# Log:
# AUTO - CONTROLE DE LINHAS\04 - SAIDAS\LOGS
$PastaLog = Join-Path $PastaProjeto "04 - SAIDAS\LOGS"
$ArquivoLog = Join-Path $PastaLog "copia_dados.log"

try {

    # Criar pasta de logs
    if (-not (Test-Path $PastaLog)) {
        New-Item -ItemType Directory -Path $PastaLog -Force | Out-Null
    }

    # Criar pasta de destino
    if (-not (Test-Path $PastaDestino)) {
        New-Item -ItemType Directory -Path $PastaDestino -Force | Out-Null
    }

    foreach ($Arquivo in $Arquivos) {
        if (-not (Test-Path $Arquivo.Origem)) {
            throw "Arquivo de origem não encontrado: $($Arquivo.Origem)"
        }
        $Destino = Join-Path $PastaDestino $Arquivo.Nome
        Copy-Item -LiteralPath $Arquivo.Origem -Destination $Destino -Force
        if (-not (Test-Path $Destino)) {
            throw "O arquivo não foi encontrado no destino após a cópia: $($Arquivo.Nome)"
        }
    }

    $DataHora = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

    Add-Content -Path $ArquivoLog `
        -Value "$DataHora - SUCESSO: TELEFONIA.xlsx e CONTATO CDC.xlsx copiados para 01 - DADOS."

    exit 0
}
catch {

    $DataHora = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

    Add-Content -Path $ArquivoLog `
        -Value "$DataHora - ERRO: $($_.Exception.Message)"

    exit 1
}
