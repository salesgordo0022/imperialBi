# =====================================================================
#  Funções comuns de acesso ao Domínio.
#  Uso:  . "$PSScriptRoot\_lib\dominio.ps1"
# =====================================================================

. "$PSScriptRoot\config.ps1"

function Get-DomConexao {
    "Driver={SQL Anywhere 17};Server=$($DOM.Servidor);DBN=$($DOM.BaseNome);UID=$($DOM.Usuario);PWD=$($DOM.Senha);Host=$($DOM.Host):$($DOM.Porta)"
}

# ---------------------------------------------------------------------
#  Servidor
# ---------------------------------------------------------------------

function Test-DomServidor {
    <#  Retorna $true se dá para conectar e consultar.  #>
    try {
        $cn = New-Object System.Data.Odbc.OdbcConnection((Get-DomConexao))
        $cn.Open()
        $cn.Close()
        return $true
    } catch { return $false }
}

function Get-DomProcesso {
    Get-Process dbsrv17, dbeng17 -ErrorAction SilentlyContinue
}

function Start-DomServidor {
    <#
        Sobe o dbsrv17 com -gl all, que é o que libera o UNLOAD para o
        usuário 'externo'. Sem isso a extração cai para leitura linha a
        linha via ODBC, que nesta base leva DIAS em vez de segundos.

        Sobe por um .bat porque Start-Process do PowerShell quebra o
        caminho do banco quando ele tem espaço ("BECKUP DOMINIO").
    #>
    [CmdletBinding()]
    param([int]$EsperaSegundos = 180)

    if (Test-DomServidor) {
        Write-Host "Servidor já está no ar." -ForegroundColor Green
        return $true
    }

    if ($DOM.Remoto) {
        # Servidor de outra máquina: não dá para ligá-lo daqui.
        Write-Host "Sem conexão com $($DOM.Servidor) em $($DOM.Host):$($DOM.Porta)." -ForegroundColor Red
        Write-Host "  Confira se o Domínio está aberto no servidor, a porta $($DOM.Porta) no firewall e o usuário/senha." -ForegroundColor Yellow
        Write-Host "  Diagnóstico:  ..\02-CONECTAR.ps1 -Procurar" -ForegroundColor Yellow
        return $false
    }

    if (-not (Test-Path $DOM.Banco))  { throw "Banco não encontrado: $($DOM.Banco)" }
    if (-not (Test-Path $DOM.Engine)) { throw "Engine não encontrada: $($DOM.Engine)" }

    New-Item -ItemType Directory -Force $DOM.PastaTmp   | Out-Null
    New-Item -ItemType Directory -Force $DOM.PastaDados | Out-Null

    $log = Join-Path $DOM.PastaTmp 'engine.log'
    $bat = Join-Path $DOM.PastaTmp 'iniciar-dominio.bat'

    $linha = '"' + $DOM.Engine + '"' +
             " -n $($DOM.Servidor)" +
             " -x tcpip(port=$($DOM.Porta))" +
             " -gl all" +
             " -c $($DOM.Cache) -ch $($DOM.CacheMax)" +
             " -ti 0" +
             " -o $log" +
             ' "' + $DOM.Banco + '"'

    Set-Content -Path $bat -Value "@echo off`r`n$linha" -Encoding ascii

    Write-Host "Subindo o servidor (o recovery de um banco de 33 GB leva alguns minutos)..." -ForegroundColor Cyan
    Start-Process -FilePath $bat -WindowStyle Minimized

    $t0 = Get-Date
    while (((Get-Date) - $t0).TotalSeconds -lt $EsperaSegundos) {
        Start-Sleep -Seconds 5
        if (Test-DomServidor) {
            Write-Host ("Servidor no ar em {0:N0}s." -f ((Get-Date) - $t0).TotalSeconds) -ForegroundColor Green
            return $true
        }
    }

    Write-Host "Não conectou em $EsperaSegundos s. Fim do log:" -ForegroundColor Red
    if (Test-Path $log) { Get-Content $log -Tail 15 }
    return $false
}

function Stop-DomServidor {
    if ($DOM.Remoto) {
        Write-Host "Servidor remoto ($($DOM.Host)) — não é derrubado daqui, o escritório usa ele." -ForegroundColor Yellow
        return
    }
    $p = Get-DomProcesso
    if (-not $p) { Write-Host "Servidor não estava rodando."; return }
    $p | Stop-Process -Force
    Start-Sleep -Seconds 3
    Write-Host "Servidor parado." -ForegroundColor Yellow
}

# ---------------------------------------------------------------------
#  Consulta pequena (catálogo, conferência, contagem)
#  NÃO use para extrair volume: o fetch linha a linha desta base faz
#  ~60 linhas/min. Para volume use Export-DomTabela.
# ---------------------------------------------------------------------

function Invoke-DomSql {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Sql,
        [int]$Timeout = 300
    )
    $cn = New-Object System.Data.Odbc.OdbcConnection((Get-DomConexao))
    $cn.Open()
    try {
        $cmd = $cn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = $Timeout
        $rd = $cmd.ExecuteReader()
        $cols = @(); for ($i = 0; $i -lt $rd.FieldCount; $i++) { $cols += $rd.GetName($i) }
        $saida = @()
        while ($rd.Read()) {
            $o = [ordered]@{}
            for ($i = 0; $i -lt $rd.FieldCount; $i++) {
                $v = $rd.GetValue($i)
                $o[$cols[$i]] = if ($v -eq [System.DBNull]::Value) { $null } else { $v }
            }
            $saida += [pscustomobject]$o
        }
        $rd.Close()
        return $saida
    } finally { $cn.Close() }
}

function Get-DomValor {
    <#  Primeira coluna da primeira linha. Para COUNT(*), MIN(), MAX().  #>
    param([Parameter(Mandatory)][string]$Sql, [int]$Timeout = 300)
    $r = Invoke-DomSql -Sql $Sql -Timeout $Timeout
    if (-not $r) { return $null }
    return ($r[0].PSObject.Properties | Select-Object -First 1).Value
}

function Get-DomColunas {
    <#  Nomes das colunas de uma tabela.  #>
    param([Parameter(Mandatory)][string]$Tabela)
    (Invoke-DomSql "SELECT cname FROM SYS.SYSCOLUMNS WHERE creator='bethadba' AND tname='$Tabela'").cname
}

function Test-DomTabela {
    param([Parameter(Mandatory)][string]$Tabela)
    [int](Get-DomValor "SELECT COUNT(*) FROM SYS.SYSTAB t JOIN SYS.SYSUSER u ON u.user_id=t.creator WHERE u.user_name='bethadba' AND t.table_name='$Tabela'") -gt 0
}

# ---------------------------------------------------------------------
#  Extração em volume: UNLOAD server-side
# ---------------------------------------------------------------------

function Save-DomLeitura {
    <#
      Plano B do Export-DomTabela para servidor remoto: lê pelo driver e
      grava no MESMO formato do UNLOAD (separador \ + t literal, cp1252,
      sem cabeçalho), para o lib.js não perceber a diferença.
    #>
    param($Conexao, [string]$Sql, [string]$Destino, [int]$Timeout = 7200)
    $cmd = $Conexao.CreateCommand()
    $cmd.CommandTimeout = $Timeout
    $cmd.CommandText = $Sql
    $rd = $cmd.ExecuteReader()
    $sw = New-Object IO.StreamWriter($Destino, $false, [Text.Encoding]::GetEncoding(1252))
    $sep = [string][char]92 + 't'
    $n = 0
    try {
        $vals = New-Object string[] $rd.FieldCount
        while ($rd.Read()) {
            for ($i = 0; $i -lt $rd.FieldCount; $i++) {
                $v = $rd.GetValue($i)
                $vals[$i] = if ($v -is [DBNull]) { '' }
                            elseif ($v -is [datetime] -and $v.TimeOfDay.Ticks -eq 0) { $v.ToString('yyyy-MM-dd') }
                            elseif ($v -is [datetime]) { $v.ToString('yyyy-MM-dd HH:mm:ss.fff') }
                            elseif ($v -is [decimal] -or $v -is [double]) { $v.ToString([Globalization.CultureInfo]::InvariantCulture) }
                            else { ([string]$v) -replace "`r?`n", ' ' }
            }
            $sw.WriteLine([string]::Join($sep, $vals))
            $n++
            if ($n % 50000 -eq 0) { Write-Host ("      ... {0:N0} linhas" -f $n) -ForegroundColor DarkGray }
        }
    } finally { $sw.Close(); $rd.Close() }
}

function Export-DomTabela {
    <#
      Grava o resultado do SELECT num .tsv, via UNLOAD.
      Quem escreve o arquivo é o SERVIDOR, não o driver — por isso é
      centenas de vezes mais rápido que ler linha a linha.

      Peculiaridades do arquivo gerado (o lib.js já trata):
        * separador é a sequência literal \ + t, não um TAB;
        * encoding cp1252 / latin1;
        * sem cabeçalho — ele vai para um .cols ao lado.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Nome,
        [Parameter(Mandatory)][string]$Sql,
        [Parameter(Mandatory)][string]$Header,
        [int]$Timeout = 7200
    )

    New-Item -ItemType Directory -Force $DOM.PastaDados | Out-Null
    $destino = Join-Path $DOM.PastaDados "$Nome.tsv"
    if (Test-Path $destino) { Remove-Item $destino -Force }

    # Barra normal: a barra invertida se perde no caminho até o SQL.
    $caminhoSql = ($destino -replace '\\', '/')
    $TAB = "'" + [char]92 + "t'"

    $cn = New-Object System.Data.Odbc.OdbcConnection((Get-DomConexao))
    $cn.Open()
    $cronometro = [Diagnostics.Stopwatch]::StartNew()
    try {
        $cmd = $cn.CreateCommand()
        $cmd.CommandTimeout = $Timeout
        if (-not $DOM.Remoto) {
            $cmd.CommandText = "UNLOAD $Sql TO '$caminhoSql' DELIMITED BY $TAB QUOTES OFF ESCAPES OFF"
            $cmd.ExecuteNonQuery() | Out-Null
        } else {
            # Servidor remoto: "TO" gravaria no disco DO SERVIDOR. Tenta
            # INTO CLIENT FILE (o servidor manda o arquivo pela conexão);
            # se o usuário não tiver o privilégio WRITE CLIENT FILE, lê
            # pelo driver e grava no mesmo formato (mais lento).
            try {
                $cmd.CommandText = "UNLOAD $Sql INTO CLIENT FILE '$caminhoSql' DELIMITED BY $TAB QUOTES OFF ESCAPES OFF"
                $cmd.ExecuteNonQuery() | Out-Null
            } catch {
                Write-Host "   (sem UNLOAD para arquivo do cliente; lendo pelo driver)" -ForegroundColor DarkYellow
                Save-DomLeitura -Conexao $cn -Sql $Sql -Destino $destino -Timeout $Timeout
            }
        }

        Set-Content -Path (Join-Path $DOM.PastaDados "$Nome.cols") -Value $Header -Encoding UTF8

        $kb = 0
        if (Test-Path $destino) { $kb = [math]::Round((Get-Item $destino).Length / 1KB) }
        Write-Host ("   [ok]    {0,-24} {1,9:N0} KB  {2,6:N1}s" -f $Nome, $kb, $cronometro.Elapsed.TotalSeconds) -ForegroundColor DarkGray
    } catch {
        # Não interrompe o lote: o gerador avisa depois se faltou arquivo.
        Write-Host ("   [FALHA] {0,-24} {1}" -f $Nome, $_.Exception.Message) -ForegroundColor Red
    } finally { $cn.Close() }
}

# ---------------------------------------------------------------------
#  Geração das planilhas (Node + exceljs)
# ---------------------------------------------------------------------

function Test-ArquivoEmUso {
    <#  True se o arquivo existe e está aberto por outro programa (Excel).  #>
    param([Parameter(Mandatory)][string]$Caminho)
    if (-not (Test-Path $Caminho)) { return $false }
    try {
        $fs = [IO.File]::Open($Caminho, 'Open', 'ReadWrite', 'None')
        $fs.Close(); $fs.Dispose()
        return $false
    } catch { return $true }
}

function Invoke-DomGerador {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Script,   # nome do .js em _gerar
        [Parameter(Mandatory)][string]$Saida,    # nome do .xlsx
        [int]$MemoriaMB = 2400
    )
    $raiz = Split-Path $PSScriptRoot -Parent
    $js   = Join-Path $raiz "_gerar\$Script"
    if (-not (Test-Path $js)) { throw "Gerador não encontrado: $js" }
    if (-not (Test-Path (Join-Path $raiz '_gerar\node_modules'))) {
        throw "Faltam as dependências do Node. Rode:  npm install  dentro de $raiz\_gerar"
    }

    New-Item -ItemType Directory -Force $DOM.PastaSaida | Out-Null
    $destino = Join-Path $DOM.PastaSaida $Saida

    # Excel segura o arquivo aberto e a gravação falha no meio.
    if (Test-ArquivoEmUso $destino) {
        throw "$Saida está aberto no Excel. Feche a planilha e rode de novo."
    }

    Write-Host "   montando $Saida ..." -ForegroundColor DarkGray
    & node "--max-old-space-size=$MemoriaMB" $js ($destino -replace '\\', '/') ($DOM.PastaDados -replace '\\', '/')
    if ($LASTEXITCODE -ne 0) { throw "O gerador $Script falhou (código $LASTEXITCODE)." }
    $mb = [math]::Round((Get-Item $destino).Length / 1MB, 1)
    Write-Host ("   [ok] $destino  ($mb MB)") -ForegroundColor Green
}

function Start-DomRelatorio {
    <#  Cabeçalho padrão + garante servidor no ar.  #>
    param([Parameter(Mandatory)][string]$Titulo)
    Write-Host ""
    Write-Host ("=" * 72) -ForegroundColor Cyan
    Write-Host "  $Titulo" -ForegroundColor Cyan
    Write-Host ("  período $($DOM.DataIni) a $($DOM.DataFim)") -ForegroundColor DarkCyan
    Write-Host ("=" * 72) -ForegroundColor Cyan
    if (-not (Test-DomServidor)) {
        if (-not (Start-DomServidor)) { throw "Servidor do Domínio indisponível." }
    }
}
