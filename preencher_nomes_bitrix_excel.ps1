<#
  Versao SEM modulos externos: usa o Excel instalado no computador (COM).

  Preenche a coluna B (Nome) com o nome do usuario do Bitrix24,
  buscando pelo ID da coluna A (author_id).

  Uso:
    .\preencher_nomes_bitrix_excel.ps1 -Entrada "C:\Automacao\planilha.xlsx"

  Webhook:
    $env:BITRIX_WEBHOOK = "https://villela.bitrix24.com.br/rest/USUARIO/CODIGO/"
  Se a variavel nao existir, o script pede a URL ao executar.

  O original NAO e alterado: o resultado sai em <nome>_preenchido.xlsx
  na mesma pasta. Feche a planilha no Excel antes de rodar.
#>

param(
    [Parameter(Mandatory = $true)][string]$Entrada,
    [string]$Saida,
    [int]$ColunaId = 1,     # coluna A
    [int]$ColunaNome = 2,   # coluna B
    [int]$LinhaInicial = 2  # linha 1 e o cabecalho
)

$ErrorActionPreference = 'Stop'

# --- Webhook ---
$webhook = $env:BITRIX_WEBHOOK
if ([string]::IsNullOrWhiteSpace($webhook)) {
    $webhook = Read-Host "Cole a URL do webhook de entrada do Bitrix24"
}
$webhook = $webhook.Trim()
if (-not $webhook.EndsWith('/')) { $webhook += '/' }

# --- Arquivos (o Excel exige caminho completo) ---
if (-not (Test-Path $Entrada)) { throw "Arquivo nao encontrado: $Entrada" }
$Entrada = (Resolve-Path $Entrada).Path
if ([string]::IsNullOrWhiteSpace($Saida)) {
    $pasta = Split-Path -Parent $Entrada
    $base = [IO.Path]::GetFileNameWithoutExtension($Entrada)
    $Saida = Join-Path $pasta ($base + '_preenchido.xlsx')
}
Copy-Item -Path $Entrada -Destination $Saida -Force
$Saida = (Resolve-Path $Saida).Path

# --- Consulta ao Bitrix24 (com cache para nao repetir chamadas) ---
$cache = @{}
$naoEncontrado = "N$([char]0xE3)o encontrado"
$semUsuario = "Sem usu$([char]0xE1)rio"

function Get-NomeBitrix {
    param([string]$Id)

    if ($cache.ContainsKey($Id)) { return $cache[$Id] }

    $nome = $naoEncontrado
    try {
        # Usuarios ativos
        $r = Invoke-RestMethod -Uri ($webhook + "user.get.json?filter[ID]=$Id") -Method Get
        $u = $r.result | Select-Object -First 1

        # Se nao achou, tenta entre os inativos (desligados)
        if (-not $u) {
            Start-Sleep -Milliseconds 550
            $r = Invoke-RestMethod -Uri ($webhook + "user.get.json?filter[ID]=$Id&filter[ACTIVE]=false") -Method Get
            $u = $r.result | Select-Object -First 1
        }

        if ($u) {
            $nome = ("$($u.NAME) $($u.LAST_NAME)").Trim()
            if ([string]::IsNullOrWhiteSpace($nome)) { $nome = $naoEncontrado }
        }
    }
    catch {
        $nome = "Erro: $($_.Exception.Message)"
    }

    $cache[$Id] = $nome
    Start-Sleep -Milliseconds 550   # respeita o limite de ~2 requisicoes/segundo
    return $nome
}

# --- Preenche a planilha pelo Excel ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null

try {
    $wb = $excel.Workbooks.Open($Saida)
    $ws = $wb.Worksheets.Item(1)

    # Ultima linha preenchida da coluna A (-4162 = xlUp)
    $ultima = $ws.Cells.Item($ws.Rows.Count, $ColunaId).End(-4162).Row
    $total = [Math]::Max(1, $ultima - $LinhaInicial + 1)
    $feitos = 0

    for ($r = $LinhaInicial; $r -le $ultima; $r++) {
        $id = "$($ws.Cells.Item($r, $ColunaId).Text)".Trim()
        $feitos++

        if ([string]::IsNullOrWhiteSpace($id)) { continue }

        if ($id -notmatch '^\d+$' -or $id -eq '0') {
            $ws.Cells.Item($r, $ColunaNome).Value2 = $semUsuario
        }
        else {
            $ws.Cells.Item($r, $ColunaNome).Value2 = Get-NomeBitrix -Id $id
        }

        Write-Progress -Activity "Buscando nomes no Bitrix24" -Status "Linha $r de $ultima" -PercentComplete (($feitos / $total) * 100)
    }

    $wb.Save()
}
finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
}

Write-Host "Pronto. Arquivo gerado em: $Saida"
