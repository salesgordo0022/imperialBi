<#
.SYNOPSIS
    Liga, desliga e identifica o banco do Domínio.

.EXAMPLE
    .\00-BANCO.ps1                 # status resumido
    .\00-BANCO.ps1 iniciar         # sobe o servidor (com -gl all, para o UNLOAD funcionar)
    .\00-BANCO.ps1 parar           # derruba o servidor e devolve a RAM
    .\00-BANCO.ps1 identificar     # ficha completa: versão, schema, período dos dados
#>
[CmdletBinding()]
param(
    [ValidateSet('status', 'iniciar', 'parar', 'identificar', 'reiniciar')]
    [string]$Acao = 'status'
)

. "$PSScriptRoot\_lib\dominio.ps1"

function Escreve-Titulo($t) {
    Write-Host ""
    Write-Host ("=" * 72) -ForegroundColor Cyan
    Write-Host "  $t" -ForegroundColor Cyan
    Write-Host ("=" * 72) -ForegroundColor Cyan
}
function Escreve-Item($k, $v, $cor = 'Gray') {
    Write-Host ("  {0,-28} " -f "$k :") -NoNewline
    Write-Host $v -ForegroundColor $cor
}

# =====================================================================
function Mostra-Arquivos {
    Escreve-Titulo "ARQUIVOS"
    $db = $DOM.Banco
    if ($DOM.Remoto) {
        Escreve-Item "Banco" "remoto: $($DOM.Servidor) em $($DOM.Host):$($DOM.Porta)" Green
    } elseif (Test-Path $db) {
        $i = Get-Item $db
        Escreve-Item "Banco"            $db
        Escreve-Item "Tamanho"          ("{0:N2} GB" -f ($i.Length / 1GB)) Green
        Escreve-Item "Modificado em"    $i.LastWriteTime
        $lg = [IO.Path]::ChangeExtension($db, '.log')
        if (Test-Path $lg) {
            Escreve-Item "Log de transações" ("{0:N0} MB" -f ((Get-Item $lg).Length / 1MB))
        }
    } else {
        Escreve-Item "Banco" "NÃO ENCONTRADO em $db" Red
    }

    if (Test-Path $DOM.Engine) {
        $v = (Get-Item $DOM.Engine).VersionInfo.ProductVersion
        Escreve-Item "Engine" "$($DOM.Engine)"
        Escreve-Item "Versão da engine" $v Green
    } else {
        Escreve-Item "Engine" "NÃO ENCONTRADA em $($DOM.Engine)" Red
    }
}

function Mostra-Processo {
    Escreve-Titulo "SERVIDOR"
    $p = Get-DomProcesso
    if ($p) {
        foreach ($x in $p) {
            Escreve-Item "Processo" ("{0} (PID {1})" -f $x.ProcessName, $x.Id) Green
            Escreve-Item "  memória em uso" ("{0:N0} MB" -f ($x.WorkingSet64 / 1MB))
            Escreve-Item "  no ar desde"    $x.StartTime
        }
    } else {
        Escreve-Item "Processo" "parado" Yellow
    }

    $conectou = Test-DomServidor
    Escreve-Item "Conexão ODBC" $(if ($conectou) { "OK  ($($DOM.Usuario)@$($DOM.Servidor):$($DOM.Porta))" } else { "sem conexão" }) `
                 $(if ($conectou) { 'Green' } else { 'Yellow' })

    $ram = Get-CimInstance Win32_OperatingSystem
    Escreve-Item "RAM livre na máquina" ("{0:N0} MB de {1:N0} MB" -f ($ram.FreePhysicalMemory / 1KB), ($ram.TotalVisibleMemorySize / 1KB))

    return $conectou
}

function Mostra-Permissoes {
    Escreve-Titulo "PERMISSÕES"
    $arq = Join-Path $DOM.PastaDados '_teste_unload.tsv'
    $ok = $false
    try {
        New-Item -ItemType Directory -Force $DOM.PastaDados | Out-Null
        $cn = New-Object System.Data.Odbc.OdbcConnection((Get-DomConexao))
        $cn.Open()
        $cmd = $cn.CreateCommand(); $cmd.CommandTimeout = 60
        $alvo = if ($DOM.Remoto) { 'INTO CLIENT FILE' } else { 'TO' }
        $cmd.CommandText = "UNLOAD SELECT 1 $alvo '$(($arq -replace '\\','/'))' QUOTES OFF ESCAPES OFF"
        $cmd.ExecuteNonQuery() | Out-Null
        $cn.Close()
        $ok = $true
        Remove-Item $arq -Force -ErrorAction SilentlyContinue
    } catch {
        $msg = $_.Exception.Message
    }
    if ($ok) {
        Escreve-Item "UNLOAD" "liberado — extração rápida disponível" Green
    } else {
        Escreve-Item "UNLOAD" "BLOQUEADO" Red
        if ($DOM.Remoto) {
            Write-Host "    Sem privilégio WRITE CLIENT FILE: a extração cai para leitura pelo driver (funciona, mais lenta)." -ForegroundColor Yellow
        } else {
            Write-Host "    O servidor precisa subir com -gl all. Rode:  .\00-BANCO.ps1 reiniciar" -ForegroundColor Yellow
        }
    }
    return $ok
}

function Mostra-Identificacao {
    Escreve-Titulo "IDENTIFICAÇÃO DO BANCO"
    $p = Invoke-DomSql @"
SELECT DB_PROPERTY('Name')         AS nome,
       DB_PROPERTY('File')         AS arquivo,
       DB_PROPERTY('PageSize')     AS pagina,
       DB_PROPERTY('CharSet')      AS charset,
       DB_PROPERTY('Collation')    AS collation,
       DB_PROPERTY('CreationTime') AS criado,
       PROPERTY('ProductVersion')  AS versao
"@
    if ($p) {
        Escreve-Item "Nome do banco"  $p[0].nome Green
        Escreve-Item "Arquivo"        $p[0].arquivo
        Escreve-Item "Versão SQL Anywhere" $p[0].versao Green
        Escreve-Item "Página"         "$($p[0].pagina) bytes"
        Escreve-Item "Charset"        $p[0].charset Yellow
        Escreve-Item "Collation"      $p[0].collation
        Escreve-Item "Criado em"      $p[0].criado
    }
    Write-Host "    (o charset acima é por que os .tsv saem em cp1252/latin1)" -ForegroundColor DarkGray

    $nt = Get-DomValor "SELECT COUNT(*) FROM SYS.SYSTAB t JOIN SYS.SYSUSER u ON u.user_id=t.creator WHERE u.user_name='bethadba' AND t.table_type_str='BASE'"
    Escreve-Item "Tabelas do schema bethadba" ("{0:N0}" -f $nt) Green
}

function Mostra-Conteudo {
    Escreve-Titulo "CONTEÚDO — EMPRESAS"
    $e = Invoke-DomSql @"
SELECT COUNT(*) AS total,
       SUM(CASE WHEN stat_emp = 'A' THEN 1 ELSE 0 END) AS ativas,
       SUM(CASE WHEN stat_emp = 'I' THEN 1 ELSE 0 END) AS inativas,
       SUM(CASE WHEN LENGTH(TRIM(cgce_emp)) = 11 THEN 1 ELSE 0 END) AS pessoas_fisicas,
       SUM(CASE WHEN i_cnae20 IS NOT NULL AND i_cnae20 <> '' THEN 1 ELSE 0 END) AS com_cnae
FROM bethadba.geempre
"@
    if ($e) {
        Escreve-Item "Empresas cadastradas" ("{0:N0}" -f $e[0].total) Green
        Escreve-Item "  ativas"             ("{0:N0}" -f $e[0].ativas)
        Escreve-Item "  inativas"           ("{0:N0}" -f $e[0].inativas)
        Escreve-Item "  pessoas físicas"    ("{0:N0}" -f $e[0].pessoas_fisicas)
        Escreve-Item "  com CNAE 2.0"       ("{0:N0}" -f $e[0].com_cnae)
    }

    Escreve-Titulo "CONTEÚDO — PERÍODO DISPONÍVEL POR MÓDULO"
    Write-Host ("  {0,-26} {1,12} {2,12} {3,14}" -f 'Fonte', 'De', 'Até', 'Linhas') -ForegroundColor DarkCyan
    Write-Host ("  " + "-" * 68) -ForegroundColor DarkGray

    $fontes = @(
        @{ n = 'Notas de saída';       t = 'efsaidas';      d = 'dsai_sai' },
        @{ n = 'Notas de entrada';     t = 'efentradas';    d = 'ddoc_ent' },
        @{ n = 'Apuração de impostos'; t = 'efsdoimp';      d = 'data_sim' },
        @{ n = 'Lançamentos contábeis';t = 'ctlancto';      d = 'data_lan' },
        @{ n = 'Bases da folha';       t = 'fobasesserv';   d = 'competencia' },
        @{ n = 'Guias de INSS';        t = 'foguiainss';    d = 'competencia' },
        @{ n = 'Rescisões';            t = 'forescisoes';   d = 'demissao' },
        @{ n = 'Log de acesso';        t = 'geloguser';     d = 'data_log' }
    )
    foreach ($f in $fontes) {
        try {
            # limita o intervalo para não deixar lixo de data (1979, 8320) mascarar o range
            $r = Invoke-DomSql @"
SELECT MIN($($f.d)) AS de, MAX($($f.d)) AS ate, COUNT(*) AS n
FROM bethadba.$($f.t)
WHERE $($f.d) >= '2000-01-01' AND $($f.d) <= '2035-12-31'
"@ -Timeout 900
            $de  = if ($r[0].de)  { ([datetime]$r[0].de).ToString('yyyy-MM-dd') }  else { '-' }
            $ate = if ($r[0].ate) { ([datetime]$r[0].ate).ToString('yyyy-MM-dd') } else { '-' }
            Write-Host ("  {0,-26} {1,12} {2,12} {3,14:N0}" -f $f.n, $de, $ate, $r[0].n)
        } catch {
            Write-Host ("  {0,-26} {1}" -f $f.n, "erro ao ler") -ForegroundColor Red
        }
    }

    Escreve-Titulo "CONTEÚDO — MÓDULOS USADOS (do log de acesso)"
    $m = Invoke-DomSql @"
SELECT m.NOME AS modulo, COUNT(*) AS sessoes, COUNT(DISTINCT l.codi_emp) AS empresas,
       COUNT(DISTINCT l.usua_log) AS usuarios
FROM bethadba.geloguser l
LEFT JOIN bethadba.GEMODULOS m ON m.CODIGO = l.sist_log
WHERE l.data_log >= '$($DOM.DataIniLog)'
GROUP BY m.NOME
"@ -Timeout 900
    Write-Host ("  {0,-22} {1,10} {2,10} {3,10}" -f 'Módulo', 'Sessões', 'Empresas', 'Usuários') -ForegroundColor DarkCyan
    Write-Host ("  " + "-" * 56) -ForegroundColor DarkGray
    foreach ($x in ($m | Sort-Object sessoes -Descending)) {
        $nome = if ($x.modulo) { $x.modulo } else { '(não mapeado)' }
        Write-Host ("  {0,-22} {1,10:N0} {2,10:N0} {3,10:N0}" -f $nome, $x.sessoes, $x.empresas, $x.usuarios)
    }
}

function Mostra-Saidas {
    Escreve-Titulo "SAÍDAS JÁ GERADAS"
    if (Test-Path $DOM.PastaSaida) {
        $f = Get-ChildItem $DOM.PastaSaida -Filter *.xlsx -ErrorAction SilentlyContinue
        if ($f) {
            foreach ($x in $f | Sort-Object Name) {
                Write-Host ("  {0,-34} {1,8:N1} MB   {2}" -f $x.Name, ($x.Length / 1MB), $x.LastWriteTime.ToString('dd/MM HH:mm'))
            }
        } else { Write-Host "  (nenhuma planilha ainda — rode EXTRAIR-TUDO.ps1)" -ForegroundColor DarkGray }
    } else { Write-Host "  (pasta ainda não existe)" -ForegroundColor DarkGray }
}

# =====================================================================
switch ($Acao) {

    'iniciar' {
        Mostra-Arquivos
        if (Start-DomServidor) { Mostra-Processo | Out-Null; Mostra-Permissoes | Out-Null }
    }

    'parar' { Stop-DomServidor }

    'reiniciar' {
        Stop-DomServidor
        Start-Sleep -Seconds 2
        if (Start-DomServidor) { Mostra-Permissoes | Out-Null }
    }

    'status' {
        Mostra-Arquivos
        $ok = Mostra-Processo
        if ($ok) { Mostra-Permissoes | Out-Null }
        Mostra-Saidas
        Write-Host ""
        if (-not $ok) {
            Write-Host "  Para subir o banco:   .\00-BANCO.ps1 iniciar" -ForegroundColor Yellow
        } else {
            Write-Host "  Ficha completa:       .\00-BANCO.ps1 identificar" -ForegroundColor DarkGray
        }
        Write-Host ""
    }

    'identificar' {
        Mostra-Arquivos
        $ok = Mostra-Processo
        if (-not $ok) {
            if (-not (Start-DomServidor)) { Write-Host "Sem servidor, não dá para identificar o conteúdo." -ForegroundColor Red; return }
        }
        Mostra-Permissoes | Out-Null
        Mostra-Identificacao
        Mostra-Conteudo
        Mostra-Saidas
        Write-Host ""
    }
}
