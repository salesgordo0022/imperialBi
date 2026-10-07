<#
.SYNOPSIS
    Instala o SAP SQL Anywhere 17 (servidor + driver ODBC) a partir da cópia
    que está em APP-SAP-SQLAnywhere17\, sem precisar do instalador do Domínio.

.DESCRIPTION
    1. Copia APP-SAP-SQLAnywhere17\ para C:\Program Files\SQL Anywhere 17
       (se já existir, só confere e não sobrescreve sem -Forcar).
    2. Registra o driver ODBC "SQL Anywhere 17" em 64 bits (Bin64) e 32 bits (Bin32).
    3. Coloca Bin64 e Bin32 no PATH da máquina.

    Precisa rodar como Administrador.

.EXAMPLE
    .\01-INSTALAR-SAP.ps1             # instala / confere
    .\01-INSTALAR-SAP.ps1 -Verificar  # só mostra o que está instalado
    .\01-INSTALAR-SAP.ps1 -Forcar     # recopia por cima (ex.: o desinstalador do Domínio apagou arquivos)
#>
[CmdletBinding()]
param(
    [switch]$Verificar,
    [switch]$Forcar,
    [string]$Destino = 'C:\Program Files\SQL Anywhere 17'
)

$origem = Join-Path $PSScriptRoot 'APP-SAP-SQLAnywhere17'
$driver = 'SQL Anywhere 17'

function Item($k, $v, $cor = 'Gray') {
    Write-Host ("  {0,-30} " -f "$k :") -NoNewline; Write-Host $v -ForegroundColor $cor
}

function Mostra-Situacao {
    Write-Host ""
    Write-Host "SQL ANYWHERE 17 NESTA MÁQUINA" -ForegroundColor Cyan
    $srv = Join-Path $Destino 'Bin64\dbsrv17.exe'
    if (Test-Path $srv) {
        Item "Servidor (dbsrv17)" ("{0}  v{1}" -f $srv, (Get-Item $srv).VersionInfo.ProductVersion) Green
    } else { Item "Servidor (dbsrv17)" "não instalado" Yellow }

    foreach ($r in @(
        @{ n = 'Driver ODBC 64 bits'; k = "HKLM:\SOFTWARE\ODBC\ODBCINST.INI\$driver" },
        @{ n = 'Driver ODBC 32 bits'; k = "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\$driver" })) {
        $d = (Get-ItemProperty $r.k -ErrorAction SilentlyContinue).Driver
        if ($d -and (Test-Path $d)) { Item $r.n $d Green }
        elseif ($d) { Item $r.n "registrado, mas o arquivo sumiu: $d" Red }
        else { Item $r.n "não registrado" Yellow }
    }
    $path = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    Item "No PATH" $(if ($path -like "*$Destino\Bin64*") { 'sim' } else { 'não' }) $(if ($path -like "*$Destino\Bin64*") { 'Green' } else { 'Yellow' })
    Write-Host ""
}

if ($Verificar) { Mostra-Situacao; return }

$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin) {
    Write-Host "Rode este script como Administrador (botão direito no PowerShell > Executar como administrador)." -ForegroundColor Red
    return
}
if (-not (Test-Path (Join-Path $origem 'Bin64\dbsrv17.exe'))) {
    throw "Cópia do SQL Anywhere não encontrada em $origem"
}

# ---- 1. arquivos -----------------------------------------------------
if ((Test-Path (Join-Path $Destino 'Bin64\dbsrv17.exe')) -and -not $Forcar) {
    Write-Host "Arquivos já estão em $Destino (use -Forcar para recopiar)." -ForegroundColor Green
} else {
    if (Get-Process dbsrv17, dbeng17 -ErrorAction SilentlyContinue) {
        throw "Há um servidor SQL Anywhere rodando. Pare antes (SCRIPTS\00-BANCO.ps1 parar)."
    }
    Write-Host "Copiando para $Destino ..." -ForegroundColor Cyan
    robocopy $origem $Destino /E /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "robocopy falhou (código $LASTEXITCODE)" }
}

# ---- 2. driver ODBC --------------------------------------------------
foreach ($r in @(
    @{ raiz = 'HKLM:\SOFTWARE\ODBC\ODBCINST.INI';             dll = "$Destino\Bin64\dbodbc17.dll" },
    @{ raiz = 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI'; dll = "$Destino\Bin32\dbodbc17.dll" })) {
    if (-not (Test-Path $r.dll)) { continue }
    $k = Join-Path $r.raiz $driver
    New-Item -Path $k -Force | Out-Null
    Set-ItemProperty $k -Name Driver -Value $r.dll
    Set-ItemProperty $k -Name Setup  -Value $r.dll
    New-Item -Path (Join-Path $r.raiz 'ODBC Drivers') -Force -ErrorAction SilentlyContinue | Out-Null
    Set-ItemProperty (Join-Path $r.raiz 'ODBC Drivers') -Name $driver -Value 'Installed'
}

# ---- 3. PATH ---------------------------------------------------------
$path = [Environment]::GetEnvironmentVariable('Path', 'Machine')
$novos = @("$Destino\Bin64", "$Destino\Bin32") | Where-Object { $path -notlike "*$_*" }
if ($novos) {
    [Environment]::SetEnvironmentVariable('Path', ($path.TrimEnd(';') + ';' + ($novos -join ';')), 'Machine')
    Write-Host "PATH atualizado (abra um PowerShell novo para valer)." -ForegroundColor DarkGray
}

Mostra-Situacao
Write-Host "Próximo passo:  .\02-CONECTAR.ps1 -Ip <ip do servidor>   (ou sem -Ip para o banco local)" -ForegroundColor Cyan
