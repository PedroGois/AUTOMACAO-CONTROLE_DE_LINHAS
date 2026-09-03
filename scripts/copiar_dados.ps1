# ============================================================
# COPIAR AS PLANILHAS DE ORIGEM PARA A PASTA DA AUTOMACAO
# ============================================================

$PastaScript = Split-Path -Parent $MyInvocation.MyCommand.Path

# A pasta scripts fica diretamente abaixo da raiz do projeto.
$PastaProjeto = Split-Path $PastaScript -Parent

# Origem:
# Ti\Telefonia e Internet\TELEFONIA\telefonia.xlsx
$PastaTi = Split-Path (Split-Path (Split-Path $PastaProjeto -Parent) -Parent) -Parent

$PastaOrigem = Join-Path $PastaTi "Telefonia e Internet\TELEFONIA"
$PastaDestino = Join-Path $PastaProjeto "data\entrada"
$Arquivos = @(
    @{ Nome = "TELEFONIA.xlsx"; Origem = (Join-Path $PastaOrigem "telefonia.xlsx") },
    @{ Nome = "CONTATO CDC.xlsx"; Origem = (Join-Path $PastaOrigem "CONTATO CDC.xlsx") }
)

# Log:
# Logs locais do fluxo.
$PastaLog = Join-Path $PastaProjeto "data\saidas\LOGS"
$ArquivoLog = Join-Path $PastaLog "copia_dados.log"

function Esperar-ArquivoPronto {
    param(
        [Parameter(Mandatory = $true)][string]$Caminho,
        [Parameter(Mandatory = $true)][string]$Nome
    )

    # Após o logon, o OneDrive pode levar alguns minutos para disponibilizar os arquivos.
    # Só copia quando o tamanho se mantém estável entre duas leituras.
    for ($Tentativa = 1; $Tentativa -le 20; $Tentativa++) {
        if (Test-Path -LiteralPath $Caminho) {
            $PrimeiraLeitura = Get-Item -LiteralPath $Caminho
            Start-Sleep -Seconds 10
            $SegundaLeitura = Get-Item -LiteralPath $Caminho -ErrorAction SilentlyContinue

            if ($SegundaLeitura -and $PrimeiraLeitura.Length -eq $SegundaLeitura.Length -and $SegundaLeitura.Length -gt 0) {
                return
            }
        }

        Start-Sleep -Seconds 20
    }

    throw "Arquivo de origem não ficou disponível para sincronização: $Nome"
}

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
        Esperar-ArquivoPronto -Caminho $Arquivo.Origem -Nome $Arquivo.Nome
        $Destino = Join-Path $PastaDestino $Arquivo.Nome
        Copy-Item -LiteralPath $Arquivo.Origem -Destination $Destino -Force
        if (-not (Test-Path $Destino)) {
            throw "O arquivo não foi encontrado no destino após a cópia: $($Arquivo.Nome)"
        }
    }

    $DataHora = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

    Add-Content -Path $ArquivoLog `
        -Value "$DataHora - SUCESSO: TELEFONIA.xlsx e CONTATO CDC.xlsx copiados para data\\entrada."

    exit 0
}
catch {

    $DataHora = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

    Add-Content -Path $ArquivoLog `
        -Value "$DataHora - ERRO: $($_.Exception.Message)"

    exit 1
}
