<#
.SYNOPSIS
    Livro caixa: movimento das contas de caixa, resumido e lançamento a
    lançamento.

.DESCRIPTION
    Gera 04-LIVRO-CAIXA.xlsx com quatro abas:
      Resumo por empresa      entradas, saídas, saldo, período
      Movimento mensal        mês a mês, com saldo acumulado
      Contas de caixa         quais contas cada empresa marcou como caixa
      Lançamentos (detalhe)   data, contrapartida, histórico, valor, operador

    Quais contas são "caixa" vem de ctcontacaixa. Conta de caixa no débito
    é entrada; no crédito, saída.

    A aba de detalhe passa de 880 mil linhas e o arquivo fica com ~60 MB,
    então o Excel demora bastante para abrir. Use -SemDetalhe para gerar só
    os resumos, que abrem na hora.

.EXAMPLE
    .\04-LIVRO-CAIXA.ps1
    .\04-LIVRO-CAIXA.ps1 -SemDetalhe
#>
[CmdletBinding()]
param([switch]$SemDetalhe)

. "$PSScriptRoot\_lib\dominio.ps1"

Start-DomRelatorio "04 - LIVRO CAIXA"
& "$PSScriptRoot\00-CADASTROS.ps1"

# --- movimento agregado por empresa / mês / conta ---------------------
Export-DomTabela -Nome '30_livro_caixa' -Header "codi_emp`tano`tmes`tcodi_cta`tentradas`tsaidas`tlancamentos" -Sql @"
SELECT codi_emp, ano, mes, conta, SUM(entrada), SUM(saida), SUM(qtd) FROM (
  SELECT l.codi_emp, YEAR(l.data_lan) AS ano, MONTH(l.data_lan) AS mes, l.cdeb_lan AS conta,
         SUM(l.vlor_lan) AS entrada, 0 AS saida, COUNT(*) AS qtd
  FROM bethadba.ctlancto l
  JOIN bethadba.ctcontacaixa c ON c.codi_emp = l.codi_emp AND c.codi_cta = l.cdeb_lan
  WHERE l.data_lan >= '$($DOM.DataIni)' AND l.data_lan <= '$($DOM.DataFim)'
  GROUP BY l.codi_emp, YEAR(l.data_lan), MONTH(l.data_lan), l.cdeb_lan
  UNION ALL
  SELECT l.codi_emp, YEAR(l.data_lan), MONTH(l.data_lan), l.ccre_lan,
         0, SUM(l.vlor_lan), COUNT(*)
  FROM bethadba.ctlancto l
  JOIN bethadba.ctcontacaixa c ON c.codi_emp = l.codi_emp AND c.codi_cta = l.ccre_lan
  WHERE l.data_lan >= '$($DOM.DataIni)' AND l.data_lan <= '$($DOM.DataFim)'
  GROUP BY l.codi_emp, YEAR(l.data_lan), MONTH(l.data_lan), l.ccre_lan
) AS t
GROUP BY codi_emp, ano, mes, conta
"@

# --- lançamento a lançamento ------------------------------------------
# O histórico útil é chis_lan (texto livre); cthispad é só o padrão,
# e nesta base codi_his quase sempre vem nulo.
if (-not $SemDetalhe) {
    Export-DomTabela -Nome '31_caixa_detalhe' `
        -Header "codi_emp`tdata`tnumero`tconta_deb`tconta_cre`tvalor`thistorico`tdocumento`toperador" `
        -Sql @"
SELECT l.codi_emp, l.data_lan, l.nume_lan, l.cdeb_lan, l.ccre_lan, l.vlor_lan,
       COALESCE(NULLIF(l.chis_lan,''), h.desc_his, ''), l.ndoc_lan, l.codi_usu
FROM bethadba.ctlancto l
LEFT JOIN bethadba.cthispad h ON h.codi_emp = l.codi_emp AND h.codi_his = l.codi_his
WHERE l.data_lan >= '$($DOM.DataIni)' AND l.data_lan <= '$($DOM.DataFim)'
  AND (EXISTS (SELECT 1 FROM bethadba.ctcontacaixa c
                WHERE c.codi_emp = l.codi_emp AND c.codi_cta = l.cdeb_lan)
    OR EXISTS (SELECT 1 FROM bethadba.ctcontacaixa c
                WHERE c.codi_emp = l.codi_emp AND c.codi_cta = l.ccre_lan))
"@
} else {
    Write-Host "   detalhe pulado (-SemDetalhe)" -ForegroundColor DarkGray
    $det = Join-Path $DOM.PastaDados '31_caixa_detalhe.tsv'
    if (Test-Path $det) { Remove-Item $det -Force }
}

Invoke-DomGerador -Script 'caixa.js' -Saida '04-LIVRO-CAIXA.xlsx'
