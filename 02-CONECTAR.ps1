<#
.SYNOPSIS
    Entra no servidor do Domínio (SQL Anywhere 17) e confere se dá para ler o banco.

.DESCRIPTION
    Testa, em ordem: a porta TCP, o login ODBC, a versão do banco, quantas
    empresas existem e se o UNLOAD rápido está liberado. Com -Salvar, grava
    os dados da conexão em SCRIPTS\_lib\conexao.json e todos os scripts de
    extração passam a usar esse servidor.

    Usuário: o Domínio não entrega a senha do 'bethadba'. Use um usuário do
    tipo EXTERNO, criado no Domínio pelo Gerente em
    Controle > Permissões > Usuários > Novo > Tipo de acesso = Externo.

.EXAMPLE
    .\02-CONECTAR.ps1                                   # usa o que está salvo / banco local
    .\02-CONECTAR.ps1 -Ip 192.168.0.10 -Salvar          # servidor do escritório
    .\02-CONECTAR.ps1 -Ip 192.168.0.10 -Servidor srvcontabil -Usuario externo -Senha 123456 -Salvar
    .\02-CONECTAR.ps1 -Procurar                         # procura servidores na rede (porta 2638)
    .\02-CONECTAR.ps1 -Local -Salvar                    # volta para o banco .db desta máquina
#>
[CmdletBinding()]
param(
    [string]$Ip,
    [string]$Servidor,
    [string]$Base,
    [int]$Porta,
    [string]$Usuario,
    [string]$Senha,
    [switch]$Local,
    [switch]$Salvar,
    [switch]$Procurar
)

$lib = Join-Path $PSScriptRoot 'SCRIPTS\_lib'
. (Join-Path $lib 'dominio.ps1')

if ($Local)    { $DOM.Host = 'localhost' }
if ($Ip)       { $DOM.Host = $Ip }
if ($Servidor) { $DOM.Servidor = $Servidor; if (-not $Base) { $DOM.BaseNome = $Servidor } }
if ($Base)     { $DOM.BaseNome = $Base }
if ($Porta)    { $DOM.Porta = $Porta }
if ($Usuario)  { $DOM.Usuario = $Usuario }
if ($Senha)    { $DOM.Senha = $Senha }
$DOM.Remoto = $DOM.Host -notin @('localhost', '127.0.0.1', '.', $env:COMPUTERNAME)

function Item($k, $v, $cor = 'Gray') {
    Write-Host ("  {0,-26} " -f "$k :") -NoNewline; Write-Host $v -ForegroundColor $cor
}
function Test-Porta([string]$h, [int]$p, [int]$ms = 1500) {
    $c = New-Object Net.Sockets.TcpClient
    try { $ok = $c.BeginConnect($h, $p, $null, $null).AsyncWaitHandle.WaitOne($ms); return ($ok -and $c.Connected) }
    catch { return $false } finally { $c.Close() }
}

# ---------------------------------------------------------------------
if ($Procurar) {
    Write-Host ""
    Write-Host "PROCURANDO SERVIDOR SQL ANYWHERE NA REDE (porta $($DOM.Porta))" -ForegroundColor Cyan
    $ips = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
           Where-Object { $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.254.*' -and $_.PrefixLength -ge 16 }
    foreach ($meu in $ips) {
        $prefixo = ($meu.IPAddress -split '\.')[0..2] -join '.'
        Write-Host "  rede $prefixo.0/24 (esta máquina é $($meu.IPAddress)) ..." -ForegroundColor DarkGray
        $tent = 1..254 | ForEach-Object {
            $c = New-Object Net.Sockets.TcpClient
            [pscustomobject]@{ ip = "$prefixo.$_"; c = $c; t = $c.ConnectAsync("$prefixo.$_", $DOM.Porta) }
        }
        Start-Sleep -Milliseconds 1500
        $achou = $tent | Where-Object { $_.t.Status -eq 'RanToCompletion' -and $_.c.Connected }
        foreach ($x in $tent) { $x.c.Close() }
        foreach ($a in $achou) {
            $nome = try { [Net.Dns]::GetHostEntry($a.ip).HostName } catch { '' }
            Item $a.ip "porta $($DOM.Porta) aberta  $nome" Green
        }
        if (-not $achou) { Write-Host "  nenhum servidor respondeu nessa rede." -ForegroundColor Yellow }
    }
    Write-Host ""
    Write-Host "  Achou? Rode:  .\02-CONECTAR.ps1 -Ip <ip> -Salvar" -ForegroundColor Cyan
    Write-Host "  O nome do servidor de banco aparece no Domínio em 'Gerenciar Servidor de Banco de Dados'." -ForegroundColor DarkGray
    Write-Host ""
    return
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "CONEXÃO COM O DOMÍNIO" -ForegroundColor Cyan
Item "Máquina"  $(if ($DOM.Remoto) { "$($DOM.Host) (remoto)" } else { 'esta máquina (banco local)' })
Item "Servidor / base" "$($DOM.Servidor) / $($DOM.BaseNome)"
Item "Porta"    $DOM.Porta
Item "Usuário"  $DOM.Usuario

if (-not (Get-ItemProperty 'HKLM:\SOFTWARE\ODBC\ODBCINST.INI\SQL Anywhere 17' -ErrorAction SilentlyContinue)) {
    Item "Driver ODBC" "SQL Anywhere 17 não registrado — rode .\01-INSTALAR-SAP.ps1 como Administrador" Red
    return
}

# 1. rede
if ($DOM.Remoto) {
    if (Test-Porta $DOM.Host $DOM.Porta) { Item "Porta TCP" "aberta" Green }
    else {
        Item "Porta TCP" "FECHADA ou servidor desligado" Red
        Write-Host "    - o Domínio/servidor de banco está ligado na máquina $($DOM.Host)?" -ForegroundColor Yellow
        Write-Host "    - o firewall dela libera a porta $($DOM.Porta)?" -ForegroundColor Yellow
        Write-Host "    - fora do escritório: precisa de VPN; AnyDesk não abre porta de banco." -ForegroundColor Yellow
        return
    }
} elseif (-not (Test-DomServidor)) {
    Write-Host "  Banco local parado — subindo..." -ForegroundColor DarkGray
    if (-not (Start-DomServidor)) { return }
}

# 2. login
try {
    $v = Invoke-DomSql "SELECT PROPERTY('ServerName') AS srv, DB_NAME() AS base, PROPERTY('ProductVersion') AS versao, CURRENT USER AS usuario"
    Item "Login" "OK" Green
    Item "Servidor de verdade" "$($v[0].srv) / base $($v[0].base)"
    Item "Versão SQL Anywhere" $v[0].versao
} catch {
    $m = $_.Exception.Message
    Item "Login" "FALHOU" Red
    Write-Host "    $m" -ForegroundColor DarkYellow
    if ($m -match '28000|password|senha')       { Write-Host "    -> usuário ou senha errados. Use um usuário EXTERNO do Domínio." -ForegroundColor Yellow }
    elseif ($m -match 'not found|encontrado')    { Write-Host "    -> nome do servidor errado. Tente -Servidor srvcontabil / contabil, ou veja no Domínio." -ForegroundColor Yellow }
    return
}

# 3. conteúdo
try {
    $e = Invoke-DomSql "SELECT COUNT(*) AS total, SUM(CASE WHEN stat_emp='A' THEN 1 ELSE 0 END) AS ativas FROM bethadba.geempre"
    Item "Empresas" ("{0:N0} ({1:N0} ativas)" -f $e[0].total, $e[0].ativas) Green
    $t = Get-DomValor "SELECT COUNT(*) FROM SYS.SYSTAB t JOIN SYS.SYSUSER u ON u.user_id=t.creator WHERE u.user_name='bethadba' AND t.table_type_str='BASE'"
    Item "Tabelas bethadba" ("{0:N0}" -f $t)
} catch { Item "Leitura" "login entrou mas não leu geempre: $($_.Exception.Message)" Red }

# 4. extração rápida
$arq = Join-Path $DOM.PastaDados '_teste_unload.tsv'
New-Item -ItemType Directory -Force $DOM.PastaDados | Out-Null
try {
    $cn = New-Object System.Data.Odbc.OdbcConnection((Get-DomConexao)); $cn.Open()
    $cmd = $cn.CreateCommand()
    $alvo = if ($DOM.Remoto) { 'INTO CLIENT FILE' } else { 'TO' }
    $cmd.CommandText = "UNLOAD SELECT 1 $alvo '$(($arq -replace '\\','/'))' QUOTES OFF ESCAPES OFF"
    $cmd.ExecuteNonQuery() | Out-Null; $cn.Close()
    Remove-Item $arq -Force -ErrorAction SilentlyContinue
    Item "Extração rápida (UNLOAD)" "liberada" Green
} catch {
    Item "Extração rápida (UNLOAD)" "bloqueada — vai pelo driver (mais lento, funciona)" Yellow
}

if ($Salvar) {
    $dados = [ordered]@{ Host = $DOM.Host; Servidor = $DOM.Servidor; BaseNome = $DOM.BaseNome
                         Porta = $DOM.Porta; Usuario = $DOM.Usuario; Senha = $DOM.Senha }
    $dados | ConvertTo-Json | Set-Content (Join-Path $lib 'conexao.json') -Encoding UTF8
    Item "Perfil" "salvo em SCRIPTS\_lib\conexao.json" Green
}
Write-Host ""
Write-Host "  Extrair:  .\03-EXTRAIR.ps1 -Lista   |   -Tabela geempre   |   -Relatorios" -ForegroundColor Cyan
Write-Host ""
