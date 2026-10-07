<#
.SYNOPSIS
    Roda os seis relatórios de ponta a ponta.

.DESCRIPTION
    Sobe o servidor, extrai tudo e monta as seis planilhas em
    ..\EXTRACAO. Leva poucos minutos; o mais demorado é o livro caixa
    detalhado (880 mil lançamentos).

.EXAMPLE
    .\EXTRAIR-TUDO.ps1
    .\EXTRAIR-TUDO.ps1 -SemDetalheCaixa   # pula as 880 mil linhas do caixa
    .\EXTRAIR-TUDO.ps1 -Pular 04,06       # pula relatórios específicos
    .\EXTRAIR-TUDO.ps1 -PararAoFim        # derruba o servidor no final
#>
[CmdletBinding()]
param(
    [string[]]$Pular = @(),
    [switch]$SemDetalheCaixa,
    [switch]$PararAoFim
)

. "$PSScriptRoot\_lib\dominio.ps1"

$inicio = Get-Date

Write-Host ""
Write-Host ("#" * 72) -ForegroundColor Magenta
Write-Host "#  EXTRAÇÃO COMPLETA DO DOMÍNIO" -ForegroundColor Magenta
Write-Host "#  período $($DOM.DataIni) a $($DOM.DataFim)" -ForegroundColor Magenta
Write-Host "#  saída   $($DOM.PastaSaida)" -ForegroundColor Magenta
Write-Host ("#" * 72) -ForegroundColor Magenta

# 05 depende de 01, 02 e 04, então a ordem importa.
$relatorios = @(
    @{ id = '01'; script = '01-FISCAL.ps1';          saida = '01-FISCAL.xlsx' },
    @{ id = '02'; script = '02-FOLHA.ps1';           saida = '02-FOLHA-E-ENCARGOS.xlsx' },
    @{ id = '03'; script = '03-CONTABIL.ps1';        saida = '03-CONTABIL-BALANCOS.xlsx' },
    @{ id = '04'; script = '04-LIVRO-CAIXA.ps1';     saida = '04-LIVRO-CAIXA.xlsx' },
    @{ id = '05'; script = '05-PRODUTOR-RURAL.ps1';  saida = '05-PRODUTOR-RURAL.xlsx' },
    @{ id = '06'; script = '06-PRODUTIVIDADE.ps1';   saida = '06-PRODUTIVIDADE-E-TEMPO.xlsx' }
)

# Avisa ANTES de extrair: descobrir no fim que o Excel travou o arquivo
# custa a extração inteira. Só checa o que este comando vai mesmo gravar.
$abertos = $relatorios |
    Where-Object { $Pular -notcontains $_.id } |
    ForEach-Object { Join-Path $DOM.PastaSaida $_.saida } |
    Where-Object { Test-ArquivoEmUso $_ }
if ($abertos) {
    Write-Host ""
    Write-Host "  Estas planilhas estão abertas no Excel e não poderão ser gravadas:" -ForegroundColor Red
    $abertos | ForEach-Object { Write-Host "    $(Split-Path $_ -Leaf)" -ForegroundColor Red }
    throw "Feche as planilhas acima e rode de novo (ou use -Pular)."
}

if (-not (Test-DomServidor)) {
    if (-not (Start-DomServidor)) { throw "Não foi possível subir o servidor do Domínio." }
}

$resultado = @()
foreach ($r in $relatorios) {
    if ($Pular -contains $r.id) {
        Write-Host "`n  -- $($r.script) pulado" -ForegroundColor DarkGray
        $resultado += [pscustomobject]@{ Relatório = $r.script; Situação = 'pulado'; Tempo = '' }
        continue
    }
    $t0 = Get-Date
    try {
        if ($r.id -eq '04' -and $SemDetalheCaixa) {
            & "$PSScriptRoot\$($r.script)" -SemDetalhe
        } else {
            & "$PSScriptRoot\$($r.script)"
        }
        $resultado += [pscustomobject]@{
            Relatório = $r.script; Situação = 'ok'
            Tempo = "{0:N0}s" -f ((Get-Date) - $t0).TotalSeconds
        }
    } catch {
        Write-Host "  FALHOU $($r.script): $($_.Exception.Message)" -ForegroundColor Red
        $resultado += [pscustomobject]@{
            Relatório = $r.script; Situação = 'FALHOU'
            Tempo = "{0:N0}s" -f ((Get-Date) - $t0).TotalSeconds
        }
    }
}

Write-Host ""
Write-Host ("#" * 72) -ForegroundColor Magenta
$resultado | Format-Table -AutoSize
Write-Host ("  tempo total: {0:N1} min" -f ((Get-Date) - $inicio).TotalMinutes) -ForegroundColor Magenta

if (Test-Path $DOM.PastaSaida) {
    Write-Host ""
    Get-ChildItem $DOM.PastaSaida -Filter *.xlsx | Sort-Object Name | ForEach-Object {
        Write-Host ("  {0,-34} {1,8:N1} MB" -f $_.Name, ($_.Length / 1MB)) -ForegroundColor Green
    }
}

if ($PararAoFim) { Write-Host ""; Stop-DomServidor }
Write-Host ""
