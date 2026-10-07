<#
.SYNOPSIS
    Produtores rurais: faturamento, impostos, livro caixa e folha só deles.

.DESCRIPTION
    Gera 05-PRODUTOR-RURAL.xlsx com quatro abas:
      Produtores identificados  cadastro + critério + totais + FUNRURAL
      Faturamento e impostos    mensal, uma coluna por tributo
      Livro caixa               entradas, saídas, saldo acumulado
      Folha e encargos          empregados e encargos

    COMO OS PRODUTORES SÃO IDENTIFICADOS

    O banco não tem marca de produtor rural. Verificado:
      geempre.CNPJ_PRODUTOR_RURAL_EMP  vazio nas 655 empresas
      geempre.ucxa_emp                 0 em todas
      ctcontas.LCDPR_CTA               0 nas 308 mil contas
    Ou seja, o Livro Caixa Digital do Produtor Rural não está configurado.

    A identificação usa dois critérios objetivos:
      1. CNAE 2.0 nas divisões 01, 02 e 03 (agricultura, pecuária,
         produção florestal, pesca) — campo geempre.i_cnae20;
      2. FUNRURAL apurado em efsdoimp.

    A coluna "Critério de identificação" da planilha diz qual regra pegou
    cada linha. CONFIRA essa lista: ela vale o que valer o cadastro de
    CNAE no Domínio.

    Este script não extrai nada novo — ele recorta o que 01, 02 e 04 já
    trouxeram. Rode-os antes, ou rode EXTRAIR-TUDO.ps1.

.EXAMPLE
    .\05-PRODUTOR-RURAL.ps1
#>
[CmdletBinding()]
param()

. "$PSScriptRoot\_lib\dominio.ps1"

Start-DomRelatorio "05 - PRODUTOR RURAL"
& "$PSScriptRoot\00-CADASTROS.ps1"

# Depende das extrações dos outros relatórios. Se faltar alguma, chama.
$dependencias = @{
    '02_faturamento'   = '01-FISCAL.ps1'
    '03_compras'       = '01-FISCAL.ps1'
    '04_impostos'      = '01-FISCAL.ps1'
    '05_folha_encargos'= '02-FOLHA.ps1'
    '05b_guia_inss'    = '02-FOLHA.ps1'
    '30_livro_caixa'   = '04-LIVRO-CAIXA.ps1'
}
$rodar = @{}
foreach ($d in $dependencias.GetEnumerator()) {
    if (-not (Test-Path (Join-Path $DOM.PastaDados "$($d.Key).tsv"))) { $rodar[$d.Value] = $true }
}
foreach ($s in $rodar.Keys) {
    Write-Host "   faltavam dados — rodando $s primeiro" -ForegroundColor Yellow
    if ($s -eq '04-LIVRO-CAIXA.ps1') { & "$PSScriptRoot\$s" -SemDetalhe } else { & "$PSScriptRoot\$s" }
}

Invoke-DomGerador -Script 'rural.js' -Saida '05-PRODUTOR-RURAL.xlsx'
