# =====================================================================
#  Configuração da extração do Domínio
#  Edite aqui e todos os scripts passam a usar os novos valores.
# =====================================================================

$Global:DOM = @{

    # ---- onde está o banco --------------------------------------------
    # 'localhost' = banco .db nesta máquina (os scripts sobem o dbsrv17).
    # Um IP/nome  = servidor do Domínio de outra máquina (ex. '192.168.0.10').
    #               Nada é iniciado aqui; só conecta.
    # Servidor = nome do servidor de banco. No Domínio aparece em
    #            "Gerenciar Servidor de Banco de Dados" (costuma ser 'contabil'
    #            ou 'srvcontabil'). O 02-CONECTAR.ps1 -Procurar ajuda a achar.
    Host     = 'localhost'
    Servidor = 'contabil'
    BaseNome = 'contabil'
    Porta    = 2638

    # ---- banco local (só quando Host = localhost) ---------------------
    Banco    = 'C:\Users\LENOVO\Documents\BECKUP DOMINIO\extraido\contabil.db'
    Engine   = 'C:\Program Files\SQL Anywhere 17\Bin64\dbsrv17.exe'
    Usuario  = 'externo'
    Senha    = '123456'

    # Cache do servidor. Subir só se a máquina tiver RAM sobrando —
    # com pouca RAM o servidor pagina e fica mais lento, não mais rápido.
    Cache    = '384M'
    CacheMax = '768M'

    # ---- pastas ------------------------------------------------------
    # O UNLOAD é executado PELO SERVIDOR, que grava o arquivo direto no
    # disco. Por isso PastaDados precisa ser um caminho local, curto e
    # SEM espaços — caminho com espaço quebra o UNLOAD.
    PastaTmp   = 'C:\dbtmp'
    PastaDados = 'C:\dbtmp\out'

    # Onde as planilhas .xlsx são gravadas.
    PastaSaida = 'C:\Users\LENOVO\Documents\BECKUP DOMINIO\EXTRACAO'

    # ---- período -----------------------------------------------------
    # Vale para fiscal, folha, livro caixa e produtividade.
    DataIni = '2024-01-01'
    DataFim = '2026-12-31'

    # O balancete contábil ignora DataIni e usa esta data, porque o saldo
    # acumulado só fecha se somar o histórico inteiro.
    DataIniContabil = '1995-01-01'

    # O log de acesso costuma ter mais história que o resto.
    DataIniLog = '2023-01-01'
}

# ---- sobrescrita por variável de ambiente ---------------------------
# Útil para rodar uma extração fora do padrão sem editar este arquivo, e
# para os scripts filhos herdarem a escolha do EXTRAIR-TUDO (cada um deles
# recarrega este config, então passar por variável é o que sobrevive).
#
#   $env:DOM_PASTA_SAIDA = 'D:\teste'; .\EXTRAIR-TUDO.ps1
#   $env:DOM_DATA_INI = '2026-01-01';  .\01-FISCAL.ps1
if ($env:DOM_PASTA_SAIDA) { $Global:DOM.PastaSaida = $env:DOM_PASTA_SAIDA }
if ($env:DOM_PASTA_DADOS) { $Global:DOM.PastaDados = $env:DOM_PASTA_DADOS }
if ($env:DOM_DATA_INI)    { $Global:DOM.DataIni    = $env:DOM_DATA_INI }
if ($env:DOM_DATA_FIM)    { $Global:DOM.DataFim    = $env:DOM_DATA_FIM }
if ($env:DOM_BANCO)       { $Global:DOM.Banco      = $env:DOM_BANCO }
if ($env:DOM_DATA_INI_LOG) { $Global:DOM.DataIniLog = $env:DOM_DATA_INI_LOG }
if ($env:DOM_HOST)        { $Global:DOM.Host       = $env:DOM_HOST }
if ($env:DOM_SERVIDOR)    { $Global:DOM.Servidor   = $env:DOM_SERVIDOR }
if ($env:DOM_BASE)        { $Global:DOM.BaseNome   = $env:DOM_BASE }
if ($env:DOM_PORTA)       { $Global:DOM.Porta      = [int]$env:DOM_PORTA }
if ($env:DOM_USUARIO)     { $Global:DOM.Usuario    = $env:DOM_USUARIO }
if ($env:DOM_SENHA)       { $Global:DOM.Senha      = $env:DOM_SENHA }

# Perfil salvo pelo 02-CONECTAR.ps1 -Salvar (fica ao lado deste arquivo).
# Variável de ambiente ainda ganha do perfil.
$perfil = Join-Path $PSScriptRoot 'conexao.json'
if (Test-Path $perfil) {
    $p = Get-Content $perfil -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($c in 'Host', 'Servidor', 'BaseNome', 'Porta', 'Usuario', 'Senha') {
        $ev = @{ Host='DOM_HOST'; Servidor='DOM_SERVIDOR'; BaseNome='DOM_BASE'; Porta='DOM_PORTA'; Usuario='DOM_USUARIO'; Senha='DOM_SENHA' }[$c]
        if ($p.$c -and -not (Get-Item "env:$ev" -ErrorAction SilentlyContinue)) { $Global:DOM[$c] = $p.$c }
    }
}

$Global:DOM.Remoto = $Global:DOM.Host -notin @('localhost', '127.0.0.1', '.', $env:COMPUTERNAME)

# DOM_SO_ATIVAS=1 faz as planilhas trazerem só empresas com stat_emp = 'A'.
# Quem aplica é o _gerar\lib.js (o node herda a variável); o SQL não muda.
$Global:DOM.SoAtivas = ($env:DOM_SO_ATIVAS -eq '1')

# Deriva competências no formato AAAA-MM
$Global:DOM.CompIni = $Global:DOM.DataIni.Substring(0, 7)
$Global:DOM.CompFim = $Global:DOM.DataFim.Substring(0, 7)
