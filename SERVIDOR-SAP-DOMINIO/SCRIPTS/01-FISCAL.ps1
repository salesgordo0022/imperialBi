<#
.SYNOPSIS
    Relatório fiscal: nome, CNPJ, inscrição estadual, faturamento e
    todos os impostos apurados.

.DESCRIPTION
    Gera 01-FISCAL.xlsx com quatro abas:
      Resumo por empresa   uma linha por empresa, uma coluna por tributo
      Mensal por empresa   a mesma coisa por competência
      Impostos detalhado   empresa x competência x tributo
      Totais por tributo   quanto cada tributo somou no escritório

    Cuidado com duas colunas parecidas:
      * faturamento é vcon_sai (valor contábil), NÃO vprod_sai;
      * imposto é sdev_sim (saldo devedor apurado), NÃO vime_sim,
        que é base de ICMS e dá números muito maiores.

.EXAMPLE
    .\01-FISCAL.ps1
#>
[CmdletBinding()]
param()

. "$PSScriptRoot\_lib\dominio.ps1"

Start-DomRelatorio "01 - FISCAL"
& "$PSScriptRoot\00-CADASTROS.ps1"

# --- faturamento ------------------------------------------------------
Export-DomTabela -Nome '02_faturamento' -Header "codi_emp`tano`tmes`tfaturamento`tvalor_produtos`tnotas" -Sql @"
SELECT codi_emp, YEAR(dsai_sai), MONTH(dsai_sai),
       SUM(ISNULL(vcon_sai,0)), SUM(ISNULL(vprod_sai,0)), COUNT(*)
FROM bethadba.efsaidas
WHERE dsai_sai >= '$($DOM.DataIni)' AND dsai_sai <= '$($DOM.DataFim)'
  AND situacao_sai = 0
  AND (cancelada_sai IS NULL OR cancelada_sai <> 'S')
GROUP BY codi_emp, YEAR(dsai_sai), MONTH(dsai_sai)
"@

# --- compras ----------------------------------------------------------
Export-DomTabela -Nome '03_compras' -Header "codi_emp`tano`tmes`tcompras`tnotas" -Sql @"
SELECT codi_emp, YEAR(ddoc_ent), MONTH(ddoc_ent), SUM(ISNULL(vcon_ent,0)), COUNT(*)
FROM bethadba.efentradas
WHERE ddoc_ent >= '$($DOM.DataIni)' AND ddoc_ent <= '$($DOM.DataFim)'
  AND situacao_ent = 0
GROUP BY codi_emp, YEAR(ddoc_ent), MONTH(ddoc_ent)
"@

# --- impostos apurados, quebrados por tributo -------------------------
# geimposto.codi_sis: 5 = Escrita Fiscal, 6 = Lalur, 12 = Folha.
# efsdoimp só guarda os do fiscal; os da folha estão em foguiainss (ver 02-FOLHA).
Export-DomTabela -Nome '04_impostos' `
    -Header "codi_emp`tcompetencia`tsigla`tnome_imposto`tcodi_sis`tsaldo_devedor`tsaldo_credor`tbase_calculo`tvalor_pago`tapuracoes" `
    -Sql @"
SELECT d.codi_emp, d.data_sim, g.sigl_imp, g.nome_imp, g.codi_sis,
       SUM(ISNULL(d.sdev_sim,0)), SUM(ISNULL(d.scre_sim,0)),
       SUM(ISNULL(d.bcal_sim,0)), SUM(ISNULL(d.vlrpago_sim,0)), COUNT(*)
FROM bethadba.efsdoimp d
JOIN bethadba.geimposto g
  ON g.codi_emp = d.codi_emp AND g.codi_imp = d.codi_imp
WHERE d.data_sim >= '$($DOM.DataIni)' AND d.data_sim <= '$($DOM.DataFim)'
GROUP BY d.codi_emp, d.data_sim, g.sigl_imp, g.nome_imp, g.codi_sis
"@

Invoke-DomGerador -Script 'fiscal.js' -Saida '01-FISCAL.xlsx'
