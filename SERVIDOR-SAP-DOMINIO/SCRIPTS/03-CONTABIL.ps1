<#
.SYNOPSIS
    Relatório contábil: balanço patrimonial, DRE e balancete analítico.

.DESCRIPTION
    Gera 03-CONTABIL-BALANCOS.xlsx com quatro abas:
      Resumo por empresa e ano  ativo, passivo, receitas, despesas, resultado
      Balanço Patrimonial       saldo em 31/12 por grupo de 2º nível
      DRE - Resultado           receitas e despesas do ano por grupo
      Balancete analítico       conta a conta, com saldo acumulado

    Este relatório IGNORA DataIni e usa DataIniContabil (1995 por padrão).
    O saldo acumulado das contas patrimoniais só fecha se somar o histórico
    inteiro — cortando em 2024 o balanço sairia sem saldo de abertura.

    Natureza das contas, pelo 1º dígito da classificação:
      1 Ativo e 3 Custos/Despesas -> devedora  (saldo = débito - crédito)
      2 Passivo/PL e 4 Receitas   -> credora   (saldo = crédito - débito)
      5 Apuração, 6 Encerramento

    Confira a coluna "Diferença Ativo - Passivo" do resumo. Onde ela não
    for zero, ou falta saldo de abertura no banco, ou o exercício não foi
    encerrado.

.EXAMPLE
    .\03-CONTABIL.ps1
#>
[CmdletBinding()]
param()

. "$PSScriptRoot\_lib\dominio.ps1"

Start-DomRelatorio "03 - CONTÁBIL / BALANÇOS"
Write-Host ("  balancete desde $($DOM.DataIniContabil) (saldo acumulado)") -ForegroundColor DarkCyan
& "$PSScriptRoot\00-CADASTROS.ps1"

# --- balancete: débitos e créditos por conta e ano --------------------
# ctlancto guarda a partida dobrada em duas colunas (cdeb_lan / ccre_lan),
# então cada lançamento entra duas vezes: uma como débito de uma conta,
# outra como crédito de outra. Daí a UNION ALL.
Export-DomTabela -Nome '22_balancete' -Header "codi_emp`tano`tcodi_cta`tdebito`tcredito`tlancamentos" -Sql @"
SELECT codi_emp, ano, conta, SUM(debito), SUM(credito), SUM(qtd) FROM (
  SELECT codi_emp, YEAR(data_lan) AS ano, cdeb_lan AS conta,
         SUM(vlor_lan) AS debito, 0 AS credito, COUNT(*) AS qtd
  FROM bethadba.ctlancto
  WHERE data_lan >= '$($DOM.DataIniContabil)' AND data_lan <= '$($DOM.DataFim)'
    AND cdeb_lan > 0
  GROUP BY codi_emp, YEAR(data_lan), cdeb_lan
  UNION ALL
  SELECT codi_emp, YEAR(data_lan) AS ano, ccre_lan AS conta,
         0 AS debito, SUM(vlor_lan) AS credito, COUNT(*) AS qtd
  FROM bethadba.ctlancto
  WHERE data_lan >= '$($DOM.DataIniContabil)' AND data_lan <= '$($DOM.DataFim)'
    AND ccre_lan > 0
  GROUP BY codi_emp, YEAR(data_lan), ccre_lan
) AS t
GROUP BY codi_emp, ano, conta
"@

Invoke-DomGerador -Script 'contabil.js' -Saida '03-CONTABIL-BALANCOS.xlsx'
