<#
.SYNOPSIS
    Produtividade: quanto tempo cada usuário passa em cada empresa e
    quantos lançamentos saem desse tempo.

.DESCRIPTION
    Gera 06-PRODUTIVIDADE-E-TEMPO.xlsx com seis abas:
      Resumo por empresa            horas, sessões, dias, horas por módulo
      Resumo por usuário            horas, empresas atendidas, média h/dia
      Tempo por empresa e usuário   quem passou quanto tempo onde
      Mensal tempo x lançamentos    horas ao lado do volume do mesmo mês
      Sessões (detalhe)             uma linha por sessão
      Lanç. contábeis por operador  quem lançou o quê

    A fonte é geloguser, que grava UMA LINHA POR SESSÃO: usuário, empresa,
    módulo, data/hora de entrada (data_log + tini_log) e de saída
    (dfim_log + tfim_log). A duração é a diferença, somando um dia inteiro
    quando a sessão vira a meia-noite.

    O módulo (sist_log) é traduzido por GEMODULOS:
      1 Contabilidade   5 Escrita Fiscal   12 Folha    3 Honorários
      4 Patrimônio      6 Lalur             8 Protocolos  15 Registro

    LEIA A ABA "Mensal tempo x lançamentos" COM CUIDADO. A coluna
    "lançamentos por hora" compara tempo com volume, mas volume ali é
    contagem de documento, não esforço: empresa com importação de XML
    registra milhares de notas em minutos, empresa digitada à mão registra
    dezenas em horas. Serve para achar o que destoa, não para medir gente.

.EXAMPLE
    .\06-PRODUTIVIDADE.ps1
#>
[CmdletBinding()]
param()

. "$PSScriptRoot\_lib\dominio.ps1"

Start-DomRelatorio "06 - PRODUTIVIDADE E TEMPO"
Write-Host ("  log de acesso desde $($DOM.DataIniLog)") -ForegroundColor DarkCyan
& "$PSScriptRoot\00-CADASTROS.ps1"

# --- o log de sessões -------------------------------------------------
Export-DomTabela -Nome '10_loguser' `
    -Header "nume_log`tcodi_emp`tusuario`tmodulo`tdata_ini`thora_ini`tdata_fim`thora_fim`tconnection_id" `
    -Sql @"
SELECT nume_log, codi_emp, usua_log, sist_log, data_log, tini_log, dfim_log, tfim_log, connection_id
FROM bethadba.geloguser
WHERE data_log >= '$($DOM.DataIniLog)' AND data_log <= '$($DOM.DataFim)'
"@

# --- lançamentos contábeis por operador -------------------------------
# codi_usu em ctlancto é o NOME do operador, não um código
# (geusuarios está vazia nesta base).
Export-DomTabela -Nome '11_lanc_contabeis' -Header "codi_emp`tano`tmes`tcodi_usu`torig_lan`tlancamentos`tvalor" -Sql @"
SELECT codi_emp, YEAR(data_lan), MONTH(data_lan), codi_usu, orig_lan,
       COUNT(*), SUM(ISNULL(vlor_lan,0))
FROM bethadba.ctlancto
WHERE data_lan >= '$($DOM.DataIni)' AND data_lan <= '$($DOM.DataFim)'
GROUP BY codi_emp, YEAR(data_lan), MONTH(data_lan), codi_usu, orig_lan
"@

# --- carimbo de quem apurou e quando ----------------------------------
Export-DomTabela -Nome '12_apuracoes' -Header "codi_emp`tusuario`tdia`tapuracoes" -Sql @"
SELECT codi_emp, USUARIO, CAST(DATA_HORA AS DATE), COUNT(*)
FROM bethadba.efsdoimp
WHERE DATA_HORA >= '$($DOM.DataIniLog)' AND DATA_HORA <= '$($DOM.DataFim)'
GROUP BY codi_emp, USUARIO, CAST(DATA_HORA AS DATE)
"@

Export-DomTabela -Nome '13_folha_calculos' -Header "codi_emp`tcompetencia`tusuario`tdia`tcalculos" -Sql @"
SELECT codi_emp, competencia, USUARIO_CALCULO, CAST(DATA_HORA_CALCULO AS DATE), COUNT(*)
FROM bethadba.fobasesserv
WHERE DATA_HORA_CALCULO >= '$($DOM.DataIniLog)' AND DATA_HORA_CALCULO <= '$($DOM.DataFim)'
GROUP BY codi_emp, competencia, USUARIO_CALCULO, CAST(DATA_HORA_CALCULO AS DATE)
"@

# A aba "Mensal tempo x lançamentos" cruza o tempo com o volume fiscal,
# contábil e de folha. Se esses ainda não foram extraídos, o gerador
# apenas deixa as colunas em zero — mas o cruzamento fica pobre.
foreach ($par in @(@('02_faturamento','01-FISCAL.ps1'), @('05_folha_encargos','02-FOLHA.ps1'))) {
    if (-not (Test-Path (Join-Path $DOM.PastaDados "$($par[0]).tsv"))) {
        Write-Host ("   aviso: $($par[0]) não existe — rode $($par[1]) para cruzar tempo x volume") -ForegroundColor Yellow
    }
}

Invoke-DomGerador -Script 'produtividade.js' -Saida '06-PRODUTIVIDADE-E-TEMPO.xlsx'
