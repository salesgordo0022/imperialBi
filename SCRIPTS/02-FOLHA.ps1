<#
.SYNOPSIS
    Relatório de folha: encargos patronais, funcionários ativos por mês,
    admissões e demissões.

.DESCRIPTION
    Gera 02-FOLHA-E-ENCARGOS.xlsx com cinco abas:
      Resumo por empresa    totais do período
      Encargos mensais      empresa x competência
      Funcionários por mês  ativos, admissões, demissões, saldo
      Admissões             detalhe por empregado
      Demissões             detalhe por empregado

    De onde vem cada número:
      * encargo patronal  -> foguiainss (empresa/acid_trab/terceiros)
      * PIS sobre folha   -> foguiapis
      * proventos         -> fobasesserv com dias_servico > 0
      * admissão          -> foempregados.admissao
      * demissão          -> forescisoes.demissao (a data mais antiga,
                             porque complemento de rescisão gera outra linha)

    Duas coisas que parecem erro e não são:
      * INSS patronal zerado em boa parte das linhas — empresa do Simples
        recolhe a parte patronal dentro do DAS;
      * NÃO há coluna de FGTS. As colunas valor_fgts/base_fgts de
        fobasesserv estão nulas em toda a base e não existe tabela de guia
        de FGTS: o Domínio calcula ao gerar o SEFIP e não grava.

.EXAMPLE
    .\02-FOLHA.ps1
#>
[CmdletBinding()]
param()

. "$PSScriptRoot\_lib\dominio.ps1"

Start-DomRelatorio "02 - FOLHA E ENCARGOS"
& "$PSScriptRoot\00-CADASTROS.ps1"

# --- proventos e descontos -------------------------------------------
# dias_servico > 0 é obrigatório: cada empregado aparece DUAS vezes na
# mesma competência, uma com dias_servico = 0. Sem o filtro a folha dobra.
# As colunas de encargo desta tabela vêm nulas nesta base — por isso o
# encargo real vem da guia, logo abaixo.
Export-DomTabela -Nome '05_folha_encargos' `
    -Header "codi_emp`tcompetencia`tempregados`tproventos`tdescontos`tbase_inss`tsalario_maternidade" `
    -Sql @"
SELECT codi_emp, competencia,
       COUNT(DISTINCT i_empregados),
       SUM(ISNULL(proventos,0)),
       SUM(ISNULL(descontos,0)),
       SUM(ISNULL(base_inss,0)),
       SUM(ISNULL(salario_maternidade,0))
FROM bethadba.fobasesserv
WHERE competencia >= '$($DOM.DataIni)' AND competencia <= '$($DOM.DataFim)'
  AND dias_servico > 0
GROUP BY codi_emp, competencia
"@

# --- encargos patronais: a guia de INSS -------------------------------
#   empresa    = INSS patronal
#   acid_trab  = RAT/SAT
#   terceiros  = contribuição a terceiros (Sistema S, salário-educação...)
#   segurados  = parte descontada do empregado (não é encargo patronal)
Export-DomTabela -Nome '05b_guia_inss' `
    -Header "codi_emp`tcompetencia`tinss_patronal_guia`trat_guia`tterceiros_guia`tsegurados_guia`tvalor_inss`tmulta`ttotal_guia`tguias" `
    -Sql @"
SELECT codi_emp, competencia,
       SUM(ISNULL(empresa,0)), SUM(ISNULL(acid_trab,0)), SUM(ISNULL(terceiros,0)),
       SUM(ISNULL(segurados,0)), SUM(ISNULL(valor_inss,0)), SUM(ISNULL(multa,0)),
       SUM(ISNULL(total_guia,0)), COUNT(*)
FROM bethadba.foguiainss
WHERE competencia >= '$($DOM.DataIni)' AND competencia <= '$($DOM.DataFim)'
GROUP BY codi_emp, competencia
"@

# --- PIS sobre folha --------------------------------------------------
Export-DomTabela -Nome '05c_guia_pis' -Header "codi_emp`tcompetencia`tpis_folha`tguias" -Sql @"
SELECT codi_emp, competencia, SUM(ISNULL(valor,0)), COUNT(*)
FROM bethadba.foguiapis
WHERE competencia >= '$($DOM.DataIni)' AND competencia <= '$($DOM.DataFim)'
GROUP BY codi_emp, competencia
"@

# --- empregados: sem filtro de data, o headcount precisa do histórico --
Export-DomTabela -Nome '06_empregados' `
    -Header "codi_emp`ti_empregados`tnome`tcpf`tadmissao`tcategoria`tvinculo`tsalario`tsai_rais`tdata_transf" `
    -Sql @"
SELECT codi_emp, i_empregados, nome, cpf, admissao, categoria, vinculo,
       salario, sai_rais, data_transf
FROM bethadba.foempregados
"@

Export-DomTabela -Nome '07_rescisoes' `
    -Header "codi_emp`ti_empregados`tdemissao`tmotivo`tcompetencia`ttipo`tproventos`tdescontos" `
    -Sql @"
SELECT codi_emp, i_empregados, demissao, motivo, COMPETENCIA, tipo, proventos, descontos
FROM bethadba.forescisoes
WHERE demissao >= '2015-01-01' AND demissao <= '$($DOM.DataFim)'
"@

Invoke-DomGerador -Script 'folha.js' -Saida '02-FOLHA-E-ENCARGOS.xlsx'
