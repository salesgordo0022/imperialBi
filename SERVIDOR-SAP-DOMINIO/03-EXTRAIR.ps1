<#
.SYNOPSIS
    Extrai dados do banco do Domínio para CSV (abre direto no Excel).

.DESCRIPTION
    Usa o servidor configurado pelo 02-CONECTAR.ps1. A extração vai por
    UNLOAD (o servidor grava o arquivo — segundos) e cai para leitura pelo
    driver quando o UNLOAD não é permitido.

    O CSV sai com ';' de separador, vírgula decimal e UTF-8 com BOM, que é
    o que o Excel em português abre sem assistente.

.EXAMPLE
    .\03-EXTRAIR.ps1 -Lista                                  # catálogo: todas as tabelas + nº de linhas
    .\03-EXTRAIR.ps1 -Colunas efsaidas                       # colunas de uma tabela
    .\03-EXTRAIR.ps1 -Tabela geempre                         # tabela inteira
    .\03-EXTRAIR.ps1 -Tabela efsaidas -Onde "dsai_sai >= '2026-01-01'"
    .\03-EXTRAIR.ps1 -Sql "SELECT codi_emp, nome_emp FROM bethadba.geempre WHERE stat_emp='A'" -Nome ativas
    .\03-EXTRAIR.ps1 -ArquivoSql .\minha-consulta.sql
    .\03-EXTRAIR.ps1 -Relatorios                             # os 6 relatórios em .xlsx (SCRIPTS\EXTRAIR-TUDO.ps1)
    .\03-EXTRAIR.ps1 -Relatorios -SemDetalheCaixa -Pular 04
#>
[CmdletBinding(DefaultParameterSetName = 'Tabela')]
param(
    [Parameter(ParameterSetName = 'Lista')][switch]$Lista,
    [Parameter(ParameterSetName = 'Colunas')][string]$Colunas,
    [Parameter(ParameterSetName = 'Tabela')][string]$Tabela,
    [Parameter(ParameterSetName = 'Tabela')][string]$Onde,
    [Parameter(ParameterSetName = 'Sql')][string]$Sql,
    [Parameter(ParameterSetName = 'Arquivo')][string]$ArquivoSql,
    [Parameter(ParameterSetName = 'Relatorios')][switch]$Relatorios,
    [Parameter(ParameterSetName = 'Relatorios')][switch]$SemDetalheCaixa,
    [Parameter(ParameterSetName = 'Relatorios')][string[]]$Pular,
    [string]$Nome,
    [string]$Pasta = (Join-Path $PSScriptRoot 'SAIDA')
)

. (Join-Path $PSScriptRoot 'SCRIPTS\_lib\dominio.ps1')

if ($Relatorios) {
    $a = @{}
    if ($SemDetalheCaixa) { $a.SemDetalheCaixa = $true }
    if ($Pular)           { $a.Pular = $Pular }
    & (Join-Path $PSScriptRoot 'SCRIPTS\EXTRAIR-TUDO.ps1') @a
    return
}

if (-not (Test-DomServidor)) {
    if (-not (Start-DomServidor)) { throw "Sem conexão com o Domínio. Rode .\02-CONECTAR.ps1 primeiro." }
}

# ---------------------------------------------------------------------
function Get-NomesColunas([string]$consulta) {
    $cn = New-Object System.Data.Odbc.OdbcConnection((Get-DomConexao)); $cn.Open()
    try {
        $cmd = $cn.CreateCommand(); $cmd.CommandText = $consulta
        $rd = $cmd.ExecuteReader([System.Data.CommandBehavior]::SchemaOnly)
        $n = for ($i = 0; $i -lt $rd.FieldCount; $i++) { $rd.GetName($i) }
        $rd.Close(); return , $n
    } finally { $cn.Close() }
}

function ConvertTo-CsvExcel([string]$tsv, [string[]]$cab, [string]$csv) {
    <#  .tsv do UNLOAD (\t literal, cp1252, sem cabeçalho) -> CSV ';' UTF-8 BOM.  #>
    $sep = [string][char]92 + 't'
    $num = [regex]'^-?\d+\.\d+$'
    $rd = New-Object IO.StreamReader($tsv, [Text.Encoding]::GetEncoding(1252))
    $wr = New-Object IO.StreamWriter($csv, $false, (New-Object Text.UTF8Encoding($true)))
    $q = { param($s) if ($s -match '[;"\r\n]') { '"' + $s.Replace('"', '""') + '"' } else { $s } }
    $n = 0
    try {
        $wr.WriteLine(($cab | ForEach-Object { & $q $_ }) -join ';')
        while ($null -ne ($l = $rd.ReadLine())) {
            $c = $l.Split(@($sep), [StringSplitOptions]::None)
            for ($i = 0; $i -lt $c.Length; $i++) {
                if ($num.IsMatch($c[$i])) { $c[$i] = $c[$i].Replace('.', ',') } else { $c[$i] = & $q $c[$i] }
            }
            $wr.WriteLine($c -join ';'); $n++
        }
    } finally { $rd.Close(); $wr.Close() }
    return $n
}

function Exportar([string]$consulta, [string]$nomeArq) {
    New-Item -ItemType Directory -Force $Pasta | Out-Null
    $csv = Join-Path $Pasta "$nomeArq.csv"
    if (Test-ArquivoEmUso $csv) { throw "$csv está aberto no Excel. Feche e rode de novo." }
    $cab = Get-NomesColunas $consulta
    Export-DomTabela -Nome $nomeArq -Sql $consulta -Header ($cab -join "`t")
    $tsv = Join-Path $DOM.PastaDados "$nomeArq.tsv"
    if (-not (Test-Path $tsv)) { throw "A extração de $nomeArq falhou (veja a mensagem acima)." }
    $n = ConvertTo-CsvExcel $tsv $cab $csv
    Write-Host ("  {0:N0} linhas  ->  {1}" -f $n, $csv) -ForegroundColor Green
    return $csv
}

# ---------------------------------------------------------------------
if ($Lista) {
    # Sem ORDER BY: UNLOAD com ordenação cara trava. Ordena-se no CSV.
    $q = "SELECT t.table_name AS tabela, t.count AS linhas_aprox, " +
         "(SELECT COUNT(*) FROM SYS.SYSTABCOL c WHERE c.table_id = t.table_id) AS colunas, " +
         "CASE LEFT(LOWER(t.table_name),2) WHEN 'ef' THEN 'Escrita Fiscal' WHEN 'ct' THEN 'Contabilidade' " +
         "WHEN 'fo' THEN 'Folha' WHEN 'ge' THEN 'Geral/Cadastros' WHEN 'hr' THEN 'Honorarios' ELSE '' END AS modulo " +
         "FROM SYS.SYSTAB t JOIN SYS.SYSUSER u ON u.user_id = t.creator " +
         "WHERE u.user_name = 'bethadba' AND t.table_type_str = 'BASE'"
    Exportar $q ($(if ($Nome) { $Nome } else { '_catalogo_tabelas' })) | Out-Null
    return
}

if ($Colunas) {
    Invoke-DomSql ("SELECT c.column_name AS coluna, d.domain_name AS tipo, c.width AS tamanho " +
                   "FROM SYS.SYSTABCOL c JOIN SYS.SYSTAB t ON t.table_id = c.table_id " +
                   "JOIN SYS.SYSUSER u ON u.user_id = t.creator JOIN SYS.SYSDOMAIN d ON d.domain_id = c.domain_id " +
                   "WHERE u.user_name = 'bethadba' AND t.table_name = '$Colunas'") | Format-Table -AutoSize
    return
}

if ($ArquivoSql) { $Sql = Get-Content $ArquivoSql -Raw -Encoding UTF8; if (-not $Nome) { $Nome = [IO.Path]::GetFileNameWithoutExtension($ArquivoSql) } }

if ($Tabela) {
    if (-not (Test-DomTabela $Tabela)) { throw "Tabela bethadba.$Tabela não existe. Veja .\03-EXTRAIR.ps1 -Lista" }
    $Sql = "SELECT * FROM bethadba.$Tabela" + $(if ($Onde) { " WHERE $Onde" } else { '' })
    if (-not $Nome) { $Nome = $Tabela }
}

if (-not $Sql) { Get-Help $PSCommandPath -Examples; return }
$Sql = $Sql.Trim().TrimEnd(';')
if (-not $Nome) { $Nome = 'consulta_' + (Get-Date -Format 'yyyyMMdd_HHmmss') }
Exportar $Sql $Nome | Out-Null
