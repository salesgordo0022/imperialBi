# O código explicado do zero

Este documento supõe que **você nunca programou**. Ele pega cada arquivo do
kit, mostra o código de verdade em pedaços e explica cada pedaço em português.

Os outros dois documentos têm outro papel:

- `LEIA-ME.md` — como **usar** (qual comando rodar)
- `DOCUMENTACAO-TECNICA.md` — de onde sai cada número e as fórmulas
- **este aqui** — o que cada linha do código faz

---

# Parte 0 — O mínimo para entender qualquer linha

## 0.1 O que é um script

Um script é um **arquivo de texto com uma lista de ordens**, escritas numa
linguagem que o computador entende. Ele é lido de cima para baixo, uma ordem
por vez, como uma receita de bolo.

A diferença para um programa "de verdade" é só que o script não precisa ser
compilado: você salva o texto e manda rodar.

Nosso kit tem dois tipos de script:

| Extensão | Linguagem | Quem executa |
|:--|:--|:--|
| `.ps1` | PowerShell | o Windows (já vem instalado) |
| `.js` | JavaScript | o Node (foi instalado à parte) |

E dentro dos `.ps1` há trechos de uma terceira linguagem, **SQL**, que é a
língua que o banco de dados fala.

## 0.2 As oito palavras que você precisa saber

**Variável** — uma caixinha com nome, onde se guarda um valor.

```powershell
$porta = 2638          # PowerShell: nome começa com $
```
```js
const porta = 2638;    // JavaScript: declara com const
```

Depois disso, escrever `$porta` é o mesmo que escrever `2638`.

**Comentário** — texto para humano, que o computador ignora. Em PowerShell
começa com `#`; em JavaScript com `//`. Todo o kit é comentado.

**Função** — um pedaço de código com nome, que você manda rodar quantas vezes
quiser. Como um botão.

```powershell
function Somar($a, $b) {
    return $a + $b
}
Somar 2 3        # devolve 5
```

**Parâmetro** — o que você entrega para a função trabalhar. No exemplo acima,
`$a` e `$b`.

**Condição** — "se isso, faça aquilo".

```powershell
if ($idade -ge 18) { "maior" } else { "menor" }
```

**Laço** — repetir para cada item de uma lista.

```powershell
foreach ($nome in $lista) { Write-Host $nome }
```
```js
for (const nome of lista) { console.log(nome); }
```

**Lista** (ou array) — vários valores em sequência, acessados por posição.

```powershell
$cores = @('azul', 'verde')     # $cores[0] é 'azul'
```

**Dicionário** (hashtable, objeto) — vários valores acessados por **nome**, não
por posição. É a estrutura mais usada no kit.

```powershell
$empresa = @{ Nome = 'ACME'; CNPJ = '123' }
$empresa.Nome                   # devolve 'ACME'
```

## 0.3 O que é um banco de dados e o que é SQL

O Domínio guarda tudo num arquivo de 30 GB chamado `contabil.db`. Dentro dele
há **4.800 tabelas** — cada tabela é como uma planilha gigante, com colunas
fixas e milhões de linhas.

Você não abre esse arquivo no Excel. Você **pergunta** para ele, usando SQL:

```sql
SELECT nome_emp, cgce_emp          -- quais colunas eu quero
FROM bethadba.geempre              -- de qual tabela
WHERE stat_emp = 'A'               -- filtrando só as ativas
```

Traduzindo: *"me dê o nome e o CNPJ, da tabela de empresas, só das ativas"*.

Quem responde é um programa separado, o **servidor** (`dbsrv17.exe`). Ele
precisa estar ligado para qualquer pergunta funcionar — é por isso que existe
o `00-BANCO.ps1`.

## 0.4 O caminho que os dados percorrem

```
   contabil.db          →      C:\dbtmp\out\*.tsv      →     EXTRACAO\*.xlsx
   (banco, 30 GB)              (arquivos de texto)           (planilhas)

        ↑                              ↑                            ↑
   os .ps1 perguntam          o servidor grava             os .js leem,
   em SQL                     direto aqui                  calculam e montam
```

Três etapas, três responsáveis:

1. **O `.ps1` pergunta.** Ele monta o comando SQL e manda para o servidor.
   Ele **não calcula nada**.
2. **O servidor responde gravando um arquivo de texto.** Um `.tsv` é um
   arquivo onde cada linha é um registro e as colunas são separadas por um
   caractere combinado.
3. **O `.js` lê esses arquivos, cruza, soma e monta a planilha.** Ele **não
   fala com o banco**.

Essa separação é proposital: se você quiser mudar uma fórmula, roda só o `.js`
e leva segundos, sem religar o banco de 30 GB.

---

# Parte 1 — Os arquivos de apoio

## 1.1 `_lib/config.ps1` — onde ficam as configurações

Este arquivo **não faz nada**. Ele só guarda valores que todos os outros usam.
É o único arquivo que você precisa editar no uso normal.

```powershell
$Global:DOM = @{
    Banco    = 'C:\Users\LENOVO\Documents\BECKUP DOMINIO\extraido\contabil.db'
    Engine   = 'C:\Program Files\SQL Anywhere 17\Bin64\dbsrv17.exe'
    Servidor = 'contabil'
    Porta    = 2638
    Usuario  = 'externo'
    Senha    = '123456'
```

**Linha por linha:**

- `$Global:DOM` — cria uma variável chamada `DOM`. O prefixo `$Global:` diz
  *"essa variável vale em todo lugar"*, inclusive dentro das funções dos
  outros arquivos. Sem isso, cada arquivo teria a sua e elas não se enxergariam.
- `= @{` — abre um **dicionário**. Tudo entre `@{` e `}` são pares
  `nome = valor`.
- `Banco = 'C:\...'` — o caminho do arquivo do banco. Aspas simples em
  PowerShell significam *"use este texto exatamente assim"*.
- `Porta = 2638` — sem aspas porque é número.

Depois, em qualquer script, `$DOM.Porta` devolve `2638`.

```powershell
    Cache    = '384M'
    CacheMax = '768M'
```

Quanta memória o servidor pode usar. **Cuidado:** se você colocar mais do que a
máquina tem livre, o Windows começa a usar o disco como memória e o servidor
fica **mais lento**, não mais rápido.

```powershell
    PastaTmp   = 'C:\dbtmp'
    PastaDados = 'C:\dbtmp\out'
    PastaSaida = 'C:\Users\LENOVO\Documents\BECKUP DOMINIO\EXTRACAO'
```

Três pastas. `PastaDados` é onde os arquivos de texto intermediários nascem;
`PastaSaida` é onde as planilhas vão.

**`PastaDados` tem que ser um caminho curto e sem espaço.** Quem grava ali é o
servidor, e ele não lida bem com espaço no caminho. Por isso `C:\dbtmp` e não
uma pasta dentro de "BECKUP DOMINIO".

```powershell
    DataIni = '2024-01-01'
    DataFim = '2026-12-31'
    DataIniContabil = '1995-01-01'
    DataIniLog = '2023-01-01'
}
```

O período. **Mudar `DataIni` e `DataFim` muda todos os relatórios de uma vez.**

`DataIniContabil` é separado porque o balanço precisa somar o histórico inteiro
para o saldo fechar — cortar em 2024 daria um Ativo que não é o Ativo real.

O `}` sozinho fecha o dicionário.

```powershell
if ($env:DOM_PASTA_SAIDA) { $Global:DOM.PastaSaida = $env:DOM_PASTA_SAIDA }
```

**Traduzindo:** *"se existir uma variável de ambiente chamada `DOM_PASTA_SAIDA`,
use o valor dela no lugar da pasta configurada acima"*.

`$env:` é como o PowerShell acessa variáveis de ambiente — valores que você
define no terminal e que os programas filhos herdam. Serve para uma rodada
avulsa sem editar o arquivo:

```powershell
$env:DOM_PASTA_SAIDA = 'D:\teste'
.\EXTRAIR-TUDO.ps1
```

```powershell
$Global:DOM.CompIni = $Global:DOM.DataIni.Substring(0, 7)
```

`Substring(0, 7)` pega 7 caracteres a partir da posição 0. De `'2024-01-01'`
sobra `'2024-01'` — a competência. É um valor derivado, calculado uma vez aqui
para os outros não precisarem repetir a conta.

---

## 1.2 `_lib/dominio.ps1` — as ferramentas

Aqui moram as **funções** que todos os scripts usam. Sozinho ele não faz nada;
é uma caixa de ferramentas.

```powershell
. "$PSScriptRoot\config.ps1"
```

Aquele **ponto solitário no começo da linha** é um operador, não sujeira. Ele
manda *"execute esse outro arquivo aqui dentro, como se o conteúdo dele
estivesse escrito nesta linha"*. É assim que o `$DOM` do config fica disponível.

`$PSScriptRoot` é uma variável automática: a pasta onde **este** arquivo está.
Usar ela em vez do caminho fixo faz o script funcionar de qualquer lugar.

### A função que monta o endereço de conexão

```powershell
function Get-DomConexao {
    "Driver={SQL Anywhere 17};Server=$($DOM.Servidor);DBN=$($DOM.Servidor);UID=$($DOM.Usuario);PWD=$($DOM.Senha);Host=localhost:$($DOM.Porta)"
}
```

Uma "string de conexão" é o endereço completo do banco: qual driver usar, qual
servidor, qual usuário, qual senha.

O `$( )` dentro das aspas é uma **subexpressão**. Dentro de aspas duplas,
`$variavel` é substituída automaticamente, mas para acessar uma propriedade
(o `.Servidor` do `$DOM`) precisa dos parênteses. Escrever `"$DOM.Servidor"`
imprimiria o dicionário inteiro seguido do texto `.Servidor`.

**A função não tem `return`.** Em PowerShell, o resultado da última expressão é
devolvido automaticamente.

### A função que testa se o banco está no ar

```powershell
function Test-DomServidor {
    try {
        $cn = New-Object System.Data.Odbc.OdbcConnection((Get-DomConexao))
        $cn.Open()
        $cn.Close()
        return $true
    } catch { return $false }
}
```

`try { } catch { }` é o mecanismo de **tentar e não morrer se der errado**: o
código do `try` roda; se ele falhar, em vez de o script parar, o `catch`
assume.

`New-Object System.Data.Odbc.OdbcConnection` cria um objeto de conexão. Esse
`System.Data.Odbc` não é do PowerShell — é do .NET, a plataforma da Microsoft.
PowerShell consegue usar qualquer coisa do .NET diretamente, e é assim que
falamos com o banco sem instalar biblioteca nenhuma.

Lendo inteiro: *"tenta abrir e fechar a conexão; se conseguiu, devolve
verdadeiro; se deu qualquer erro, devolve falso"*.

### A função que liga o servidor

```powershell
    $linha = '"' + $DOM.Engine + '"' +
             " -n $($DOM.Servidor)" +
             " -x tcpip(port=$($DOM.Porta))" +
             " -gl all" +
             " -c $($DOM.Cache) -ch $($DOM.CacheMax)" +
             " -ti 0" +
             " -o $log" +
             ' "' + $DOM.Banco + '"'

    Set-Content -Path $bat -Value "@echo off`r`n$linha" -Encoding ascii
    Start-Process -FilePath $bat -WindowStyle Minimized
```

Aqui montamos, pedaço por pedaço, a linha de comando que liga o servidor. O `+`
junta textos.

O que cada opção faz:

| Opção | Significado |
|:--|:--|
| `-n contabil` | nome do servidor |
| `-x tcpip(port=2638)` | aceitar conexão na porta 2638 |
| **`-gl all`** | **liberar o `UNLOAD` para qualquer usuário** |
| `-c` e `-ch` | memória mínima e máxima |
| `-ti 0` | nunca desconectar por inatividade |
| `-o` | onde gravar o log |

**O `-gl all` é a opção mais importante do kit inteiro.** Sem ela, o comando
que extrai os dados rápido é recusado e a extração cairia de segundos para
dias.

`Set-Content` grava o texto num arquivo `.bat`, e `Start-Process` executa esse
`.bat`.

**Por que o desvio pelo `.bat`?** Porque o caminho do banco tem um espaço
("BECKUP DOMINIO"), e quando o PowerShell passa isso direto para o programa,
ele quebra no espaço — o servidor reclama que `C:\Users\LENOVO\Documents\BECKUP`
não existe. Escrever num `.bat` e mandar o Windows executar resolve.

```powershell
    $t0 = Get-Date
    while (((Get-Date) - $t0).TotalSeconds -lt $EsperaSegundos) {
        Start-Sleep -Seconds 5
        if (Test-DomServidor) {
            Write-Host ("Servidor no ar em {0:N0}s." -f ((Get-Date) - $t0).TotalSeconds)
            return $true
        }
    }
```

`while (condição) { }` repete enquanto a condição for verdadeira.

Traduzindo: *"marque a hora agora (`$t0`); enquanto não passar do limite,
espere 5 segundos e teste; se conectou, avise e saia"*.

`-lt` é "menor que" (**l**ess **t**han). PowerShell usa palavras em vez de
símbolos: `-eq` igual, `-ne` diferente, `-lt` menor, `-gt` maior.

O `-f` é o **operador de formatação**. `"{0:N0}" -f 12.7` coloca o argumento 0
naquele lugar, formatado como número sem decimais → `13`.

### A função que extrai os dados — o coração do kit

```powershell
function Export-DomTabela {
    param(
        [Parameter(Mandatory)][string]$Nome,
        [Parameter(Mandatory)][string]$Sql,
        [Parameter(Mandatory)][string]$Header,
        [int]$Timeout = 7200
    )
```

O bloco `param()` declara o que a função recebe:

- `[string]` — tem que ser texto; `[int]` — tem que ser número inteiro
- `[Parameter(Mandatory)]` — **obrigatório**; se esquecer, o PowerShell pergunta
- `= 7200` — valor padrão (2 horas de limite), usado se você não informar

```powershell
    $destino = Join-Path $DOM.PastaDados "$Nome.tsv"
    if (Test-Path $destino) { Remove-Item $destino -Force }
```

`Join-Path` gruda pasta e arquivo com a barra certa. `Test-Path` verifica se
existe. Traduzindo: *"monta o caminho do arquivo; se já existir, apaga"*.

```powershell
    $caminhoSql = ($destino -replace '\\', '/')
    $TAB = "'" + [char]92 + "t'"
```

Duas gambiarras necessárias, e vale entender por quê.

**Primeira:** `-replace '\\', '/'` troca as barras invertidas por barras
normais, virando `C:/dbtmp/out/arquivo.tsv`. A barra invertida é caractere de
escape em várias linguagens; no caminho que o texto percorre até chegar ao
banco, ela some. Barra normal atravessa inteira.

**Segunda:** `[char]92` é a barra invertida montada pelo **número dela** na
tabela de caracteres. O resultado é o texto `'\t'`. Escrever isso direto no
código não funcionaria, pelo mesmo motivo acima.

```powershell
        $cmd.CommandText = "UNLOAD $Sql TO '$caminhoSql' DELIMITED BY $TAB QUOTES OFF ESCAPES OFF"
        $cmd.ExecuteNonQuery() | Out-Null
```

Aqui a mágica. O comando final fica assim:

```sql
UNLOAD SELECT ... FROM ... TO 'C:/dbtmp/out/01_empresas.tsv'
DELIMITED BY '\t' QUOTES OFF ESCAPES OFF
```

**`UNLOAD` manda o servidor gravar o resultado num arquivo**, em vez de mandar
as linhas de volta. É a diferença entre segundos e dias: quando as linhas vêm
uma a uma pelo driver, esta base faz cerca de 60 linhas por minuto.

`ExecuteNonQuery()` é para comando que não devolve linhas. O `| Out-Null`
descarta o que ele devolve — senão apareceria um número solto no meio do log.

```powershell
        Set-Content -Path (Join-Path $DOM.PastaDados "$Nome.cols") -Value $Header
```

O `UNLOAD` **não grava cabeçalho**. Então os nomes das colunas vão num arquivo
separado, com extensão `.cols`, e o JavaScript junta os dois depois.

```powershell
    } catch {
        Write-Host ("   [FALHA] {0,-24} {1}" -f $Nome, $_.Exception.Message) -ForegroundColor Red
    } finally { $cn.Close() }
```

`$_` dentro do `catch` é o erro que aconteceu. `finally` roda **sempre**, dando
erro ou não — é onde a conexão é fechada, para não ficar conexão pendurada.

**A falha não interrompe o lote.** É decisão de projeto: se uma consulta falhar,
as outras ainda rodam, e você não perde a extração inteira por causa de uma.

### A função que detecta arquivo aberto no Excel

```powershell
function Test-ArquivoEmUso {
    param([string]$Caminho)
    if (-not (Test-Path $Caminho)) { return $false }
    try {
        $fs = [IO.File]::Open($Caminho, 'Open', 'ReadWrite', 'None')
        $fs.Close(); $fs.Dispose()
        return $false
    } catch { return $true }
}
```

Não existe função pronta para "esse arquivo está aberto?". O jeito é **tentar
abrir com exclusividade**: o `'None'` no fim significa *"não deixo ninguém mais
mexer enquanto eu estiver com ele"*. Se o Excel já estiver segurando o arquivo,
a tentativa falha — e é isso que queríamos descobrir.

Serve para avisar **antes** de começar, em vez de você descobrir depois de
cinco minutos de extração.

### A função que chama o JavaScript

```powershell
    & node "--max-old-space-size=$MemoriaMB" $js ($destino -replace '\\','/') ($DOM.PastaDados -replace '\\','/')
    if ($LASTEXITCODE -ne 0) { throw "O gerador $Script falhou (código $LASTEXITCODE)." }
```

O `&` é o **operador de chamada**: executa o que vem depois como programa.

`--max-old-space-size=2400` dá 2,4 GB de memória ao Node — necessário porque o
livro caixa tem 880 mil linhas e o padrão não aguenta.

Os dois caminhos no fim são os argumentos que o JavaScript recebe: para onde
gravar a planilha, e onde estão os arquivos de dados.

`$LASTEXITCODE` é o código de saída do último programa. Zero significa sucesso;
qualquer outro é erro, e o `throw` interrompe.

---

# Parte 2 — Os scripts de extração

## 2.1 A anatomia comum

Todos os seis relatórios têm a mesma forma:

```powershell
<#
   ... bloco de ajuda ...
#>
[CmdletBinding()]
param()

. "$PSScriptRoot\_lib\dominio.ps1"      # 1. carrega as ferramentas

Start-DomRelatorio "01 - FISCAL"        # 2. imprime título e liga o banco
& "$PSScriptRoot\00-CADASTROS.ps1"      # 3. garante os cadastros

Export-DomTabela -Nome '...' ...        # 4. uma chamada por consulta
Export-DomTabela -Nome '...' ...

Invoke-DomGerador -Script 'fiscal.js' -Saida '01-FISCAL.xlsx'   # 5. monta a planilha
```

O bloco `<# ... #>` no topo é **comentário de várias linhas**. As palavras
`.SYNOPSIS`, `.DESCRIPTION` e `.EXAMPLE` fazem o PowerShell reconhecê-lo como
ajuda oficial — dá para ler com:

```powershell
Get-Help .\01-FISCAL.ps1 -Full
```

## 2.2 `00-CADASTROS.ps1` — a base comum

```powershell
$arquivos = @('01_empresas', '01b_cnae', '24_cnae20', '20_plano_contas', '21_contas_caixa', '23_usuarios')

if (-not $Forcar) {
    $faltando = $arquivos | Where-Object { -not (Test-Path (Join-Path $DOM.PastaDados "$_.tsv")) }
    if (-not $faltando) { return }
}
```

`@( )` cria uma **lista** com os seis nomes de arquivo.

A linha do meio usa o **cano** (`|`), que passa cada item da lista para o
comando seguinte. `Where-Object { ... }` mantém só os que passam no teste.
`$_` é "o item da vez".

Traduzindo: *"a menos que tenham pedido `-Forcar`, veja quais dos seis arquivos
não existem; se não faltar nenhum, saia sem fazer nada"*.

É por isso que você quase nunca precisa rodar este script na mão: ele é chamado
pelos outros e sai na hora quando não tem trabalho.

```powershell
Export-DomTabela -Nome '01_empresas' `
    -Header "codi_emp`tnome_emp`trazao_emp`t..." `
    -Sql @"
SELECT codi_emp, nome_emp, razao_emp, cgce_emp, iest_emp, ...
FROM bethadba.geempre
"@
```

Dois detalhes de escrita:

**A crase no fim da linha** (`` ` ``) é o **caractere de continuação**: diz
*"o comando continua na linha de baixo"*. Sem ela o PowerShell acharia que o
comando acabou.

**`` `t `` dentro de aspas duplas é um TAB.** No PowerShell o escape é a crase,
não a barra invertida. Então `"codi_emp`tnome_emp"` é "codi_emp", TAB,
"nome_emp" — os nomes das colunas separados por tabulação.

**`@"` e `"@`** abrem e fecham um texto de várias linhas. É onde o SQL mora.

> **Regra que quebra tudo:** o `"@` de fechamento tem que estar **colado na
> margem esquerda**, sozinho na linha. Se você indentar, dá erro de sintaxe.

As outras cinco chamadas são iguais, mudando só a tabela e as colunas.

## 2.3 `01-FISCAL.ps1` — faturamento e impostos

```powershell
Export-DomTabela -Nome '02_faturamento' -Header "codi_emp`tano`tmes`tfaturamento`tvalor_produtos`tnotas" -Sql @"
SELECT codi_emp, YEAR(dsai_sai), MONTH(dsai_sai),
       SUM(ISNULL(vcon_sai,0)), SUM(ISNULL(vprod_sai,0)), COUNT(*)
FROM bethadba.efsaidas
WHERE dsai_sai >= '$($DOM.DataIni)' AND dsai_sai <= '$($DOM.DataFim)'
  AND situacao_sai = 0
  AND (cancelada_sai IS NULL OR cancelada_sai <> 'S')
GROUP BY codi_emp, YEAR(dsai_sai), MONTH(dsai_sai)
"@
```

Vamos ler o SQL em português, linha por linha:

**`SELECT codi_emp, YEAR(dsai_sai), MONTH(dsai_sai),`**
*"Me dê o código da empresa, o ano da data de saída e o mês da data de saída."*
`YEAR()` e `MONTH()` extraem pedaços de uma data.

**`SUM(ISNULL(vcon_sai,0)), SUM(ISNULL(vprod_sai,0)), COUNT(*)`**
*"...e some o valor contábil, some o valor dos produtos, e conte quantas notas."*

`SUM()` soma. `COUNT(*)` conta linhas.

`ISNULL(x, 0)` significa *"se `x` estiver vazio, use zero"*. **Isso é
obrigatório**: em SQL, qualquer conta com um valor vazio dá vazio. Uma única
linha em branco pode zerar a coluna inteira. Foi exatamente o erro que fez a
primeira versão do relatório de folha sair toda em branco.

**`FROM bethadba.efsaidas`**
A tabela de notas fiscais de saída. `bethadba` é o "sobrenome" da tabela — o
esquema onde o Domínio guarda tudo.

**`WHERE dsai_sai >= '2024-01-01' AND dsai_sai <= '2026-12-31'`**
*"...só das notas com data dentro do período."*

O `'$($DOM.DataIni)'` é substituído pelo PowerShell **antes** de o SQL sair.
O banco recebe a data já escrita.

**`AND situacao_sai = 0`**
Só notas efetivadas.

**`AND (cancelada_sai IS NULL OR cancelada_sai <> 'S')`**
*"...e que não estejam canceladas."*

`IS NULL` é "está vazio" — e repare que em SQL **não se escreve `= NULL`**,
porque vazio não é igual a nada, nem a outro vazio. `<>` é "diferente de".

Os parênteses importam: sem eles o `OR` se misturaria com os `AND` de cima e o
filtro mudaria de sentido.

**`GROUP BY codi_emp, YEAR(dsai_sai), MONTH(dsai_sai)`**
*"Junte as linhas por empresa, ano e mês."*

Esta é a linha que transforma 5 milhões de notas em algumas milhares de linhas
de resumo. A regra: **tudo que está no `SELECT` e não é `SUM` ou `COUNT` tem
que estar no `GROUP BY`**.

### A consulta dos impostos, com `JOIN`

```sql
SELECT d.codi_emp, d.data_sim, g.sigl_imp, g.nome_imp, g.codi_sis,
       SUM(ISNULL(d.sdev_sim,0)), ...
FROM bethadba.efsdoimp d
JOIN bethadba.geimposto g
  ON g.codi_emp = d.codi_emp AND g.codi_imp = d.codi_imp
```

`JOIN` é **juntar duas tabelas**. A tabela de apuração (`efsdoimp`) guarda o
imposto por um código numérico; a de cadastro (`geimposto`) guarda o nome e a
sigla desse código. O `JOIN` cola as duas.

O `d` e o `g` depois dos nomes são **apelidos**, para não repetir o nome inteiro
da tabela toda hora. `d.codi_emp` é "o `codi_emp` da tabela `d`".

**O detalhe que mais causa erro no Domínio:** repare que o `ON` casa **duas**
colunas, e uma delas é `codi_emp`.

> **Toda tabela do Domínio é dividida por empresa.** O imposto de código 3 não
> é "o imposto 3" — é "o imposto 3 **da empresa tal**". Outra empresa tem o seu
> próprio código 3, que pode ser outro tributo.
>
> Se você esquecer o `codi_emp` no `JOIN`, o SQL **não dá erro** — ele cruza o
> imposto de uma empresa com o nome do imposto de outra e devolve números
> errados em silêncio.

---

## 2.4 `02-FOLHA.ps1` — o filtro que evita dobrar a folha

```sql
SELECT codi_emp, competencia,
       COUNT(DISTINCT i_empregados),
       SUM(ISNULL(proventos,0)),
       ...
FROM bethadba.fobasesserv
WHERE competencia >= '...' AND competencia <= '...'
  AND dias_servico > 0
GROUP BY codi_emp, competencia
```

`COUNT(DISTINCT i_empregados)` conta **quantos empregados diferentes**, não
quantas linhas. Sem o `DISTINCT`, um empregado que aparece duas vezes contaria
dois.

**`AND dias_servico > 0` é obrigatório e não é óbvio.** Nesta tabela cada
empregado aparece **duas vezes** na mesma competência: uma linha com
`dias_servico = 0` e outra com os dias trabalhados. Somar sem o filtro dobra a
folha.

Conferência real: a empresa 345 em 10/2025 dá 4.818,15 com o filtro e 9.636,30
sem.

### De onde vem o encargo patronal

```sql
SELECT codi_emp, competencia,
       SUM(ISNULL(empresa,0)), SUM(ISNULL(acid_trab,0)), SUM(ISNULL(terceiros,0)),
       SUM(ISNULL(segurados,0)), ...
FROM bethadba.foguiainss
```

A tabela é a da **guia de INSS**, não a de bases de cálculo. As colunas:

| Coluna | O que é |
|:--|:--|
| `empresa` | INSS patronal — o que a empresa paga |
| `acid_trab` | RAT/SAT — seguro de acidente |
| `terceiros` | Sistema S, salário-educação |
| `segurados` | o que foi **descontado do empregado** (não é encargo patronal) |

**Por que não usar a tabela de bases?** Porque as colunas de encargo dela
(`val_inss_empresa`, `val_acid_trabalho`, etc.) estão **vazias em toda a base**.
A primeira versão deste script somava elas e produziu uma planilha de zeros.

---

## 2.5 `03-CONTABIL.ps1` — a consulta mais complicada do kit

```sql
SELECT codi_emp, ano, conta, SUM(debito), SUM(credito), SUM(qtd) FROM (

  SELECT codi_emp, YEAR(data_lan) AS ano, cdeb_lan AS conta,
         SUM(vlor_lan) AS debito, 0 AS credito, COUNT(*) AS qtd
  FROM bethadba.ctlancto
  WHERE data_lan >= '1995-01-01' AND cdeb_lan > 0
  GROUP BY codi_emp, YEAR(data_lan), cdeb_lan

  UNION ALL

  SELECT codi_emp, YEAR(data_lan), ccre_lan,
         0, SUM(vlor_lan), COUNT(*)
  FROM bethadba.ctlancto
  WHERE data_lan >= '1995-01-01' AND ccre_lan > 0
  GROUP BY codi_emp, YEAR(data_lan), ccre_lan

) AS t
GROUP BY codi_emp, ano, conta
```

**O problema que ela resolve.** Contabilidade funciona por partida dobrada:
todo lançamento debita uma conta e credita outra. Mas o Domínio guarda isso
numa **linha só**, com duas colunas: `cdeb_lan` (conta debitada) e `ccre_lan`
(conta creditada).

Para virar balancete, cada lançamento precisa entrar **duas vezes** — uma como
débito de uma conta, outra como crédito de outra.

**Como o SQL faz isso.**

A primeira metade olha só a coluna de débito e trata `cdeb_lan` como se fosse
"a conta". A segunda metade faz o mesmo com `ccre_lan`.

`UNION ALL` **empilha os dois resultados**, um embaixo do outro.

> O `ALL` importa. Só `UNION` faria o SQL apagar linhas repetidas — e duas
> contas diferentes com o mesmo valor no mesmo ano viram uma só, perdendo
> dinheiro.

Repare nos **zeros de enchimento**: `0 AS credito` na primeira metade e o `0`
solto na segunda. As duas metades precisam ter **o mesmo número de colunas, na
mesma ordem**. A metade do débito não tem crédito, mas precisa da coluna para
alinhar.

`AS ano`, `AS conta`, `AS debito` dão **apelido** à coluna, para poder
referenciá-la depois.

O `( ... ) AS t` transforma tudo isso numa **tabela temporária**, e o
`GROUP BY` de fora reagrupa: uma conta que apareceu como débito e como crédito
no mesmo ano vira uma linha só, com as duas colunas preenchidas.

> **Por que 1995 e não 2024.** O saldo de uma conta é a soma de tudo desde
> sempre. Se cortássemos em 2024, o Ativo mostraria o movimento de dois anos,
> não o patrimônio da empresa.

---

## 2.6 `04-LIVRO-CAIXA.ps1` — o `EXISTS`

```sql
WHERE l.data_lan >= '...' AND l.data_lan <= '...'
  AND (EXISTS (SELECT 1 FROM bethadba.ctcontacaixa c
                WHERE c.codi_emp = l.codi_emp AND c.codi_cta = l.cdeb_lan)
    OR EXISTS (SELECT 1 FROM bethadba.ctcontacaixa c
                WHERE c.codi_emp = l.codi_emp AND c.codi_cta = l.ccre_lan))
```

`EXISTS (...)` pergunta *"existe pelo menos uma linha que satisfaz isso?"*. O
`SELECT 1` é convenção — não interessa o valor, só se achou algo.

Traduzindo: *"traga o lançamento se a conta de débito **ou** a de crédito
estiver na lista de contas de caixa"*.

**Por que `EXISTS` e não `JOIN`?** Porque o teste é em duas colunas. Com `JOIN`,
um lançamento que mexesse em caixa dos dois lados apareceria duplicado.

```sql
COALESCE(NULLIF(l.chis_lan,''), h.desc_his, '')
```

Duas funções encadeadas, lidas de dentro para fora:

- `NULLIF(a, b)` — *"se `a` for igual a `b`, trate como vazio"*. Aqui: se o
  histórico for string vazia, vire vazio de verdade.
- `COALESCE(x, y, z)` — *"devolva o primeiro que não estiver vazio"*.

Junto: *"use o histórico digitado; se estiver em branco, use o histórico
padrão; se não tiver nenhum, deixe em branco"*.

---

## 2.7 `EXTRAIR-TUDO.ps1` — o maestro

```powershell
$relatorios = @(
    @{ id = '01'; script = '01-FISCAL.ps1';  saida = '01-FISCAL.xlsx' },
    @{ id = '02'; script = '02-FOLHA.ps1';   saida = '02-FOLHA-E-ENCARGOS.xlsx' },
    ...
)
```

Uma **lista de dicionários**: cada item tem três campos. É a tabela de
programação do lote.

**A ordem importa:** o relatório 05 (rural) usa os dados que o 01, 02 e 04
extraíram. Por isso ele vem depois.

```powershell
$abertos = $relatorios |
    Where-Object { $Pular -notcontains $_.id } |
    ForEach-Object { Join-Path $DOM.PastaSaida $_.saida } |
    Where-Object { Test-ArquivoEmUso $_ }
```

Um **cano de três estágios**. Cada `|` passa o resultado adiante:

1. `Where-Object { $Pular -notcontains $_.id }` — tira os que você mandou pular
2. `ForEach-Object { Join-Path ... }` — transforma cada um no caminho da planilha
3. `Where-Object { Test-ArquivoEmUso $_ }` — fica só com os que estão travados

O que sobra em `$abertos` são as planilhas abertas no Excel que atrapalhariam.

```powershell
foreach ($r in $relatorios) {
    if ($Pular -contains $r.id) { ...; continue }
    $t0 = Get-Date
    try {
        & "$PSScriptRoot\$($r.script)"
        $resultado += [pscustomobject]@{ Relatório = $r.script; Situação = 'ok'; ... }
    } catch {
        Write-Host "  FALHOU $($r.script): $($_.Exception.Message)"
        $resultado += [pscustomobject]@{ ...; Situação = 'FALHOU'; ... }
    }
}
```

`foreach` percorre a lista. `continue` pula para o próximo item.

Cada relatório roda dentro de `try/catch`, então **um que falhe não derruba os
outros** — no fim sai uma tabela com `ok` ou `FALHOU` e o tempo de cada um.

`[pscustomobject]@{ }` transforma o dicionário num objeto que o
`Format-Table` sabe imprimir bonitinho.

---

# Parte 3 — Os geradores de planilha (JavaScript)

## 3.1 Como o JavaScript difere do PowerShell

| Coisa | PowerShell | JavaScript |
|:--|:--|:--|
| Variável | `$nome = 'x'` | `const nome = 'x'` |
| Comentário | `# texto` | `// texto` |
| Igual | `-eq` | `===` |
| E / Ou / Não | `-and -or -not` | `&& \|\| !` |
| Lista | `@(1,2)` | `[1,2]` |
| Dicionário | `@{a=1}` | `{a: 1}` |
| Fim de linha | nada | `;` |
| Função | `function F($x) { }` | `const F = x => { }` |

`const` significa *"essa caixinha não vai apontar para outra coisa"*. É o padrão
no kit; `let` é usado quando o valor muda.

## 3.2 `_gerar/lib.js` — as ferramentas do lado JavaScript

### Ler o arquivo que o `UNLOAD` gerou

```js
function read(nome) {
  const file = path.join(OUT, nome + '.tsv');
  if (!fs.existsSync(file)) return { cols: [], rows: [] };

  const colsFile = path.join(OUT, nome + '.cols');
  let cols = [];
  if (fs.existsSync(colsFile)) {
    cols = fs.readFileSync(colsFile, 'utf8').replace(/^\uFEFF/, '').trim().split('\t');
  }

  const txt = fs.readFileSync(file).toString('latin1');
  const rows = [];
  for (const line of txt.split(/\r?\n/)) {
    if (!line) continue;
    rows.push(line.split('\\t').map(v => v.trim()));
  }
  return { cols, rows };
}
```

Esta função tem **três armadilhas embutidas**, e cada uma tem uma linha
dedicada:

**1. O separador não é TAB de verdade.**

```js
line.split('\\t')
```

Repare: **duas** barras invertidas. Em JavaScript `'\t'` é o caractere TAB, mas
`'\\t'` é o texto literal barra-invertida seguida de "t". O `UNLOAD` gravou os
dois caracteres, não um TAB. É por isso que este arquivo não abre direito em
nenhum leitor comum de TSV.

Isso é até vantagem: um TAB de verdade dentro do nome de uma empresa não quebra
a linha.

**2. O texto está em cp1252, não UTF-8.**

```js
fs.readFileSync(file).toString('latin1')
```

`readFileSync` sem o segundo argumento devolve bytes crus; `.toString('latin1')`
os interpreta na codificação antiga do Windows. Sem isso, "GRAJAÚ" viraria
"GRAJA?".

**3. Não tem cabeçalho.** Por isso lemos o `.cols` à parte.

`.replace(/^\uFEFF/, '')` remove o **BOM** — três bytes invisíveis que o Windows
põe no começo de arquivos UTF-8 e que, sem essa limpeza, grudariam no nome da
primeira coluna.

`.split(/\r?\n/)` quebra em linhas aceitando os dois formatos de fim de linha
(Windows usa `\r\n`, Linux usa `\n`). Aquilo entre barras é uma **expressão
regular** — um padrão de busca. `?` quer dizer "opcional".

`.map(v => v.trim())` aplica "tire os espaços das pontas" a cada coluna.
`map` transforma cada item de uma lista.

### Transformar linhas em objetos

```js
function readObj(nome) {
  const { cols, rows } = read(nome);
  return rows.map(r => {
    const o = {};
    cols.forEach((c, i) => { o[c] = r[i] === undefined ? '' : r[i]; });
    return o;
  });
}
```

`const { cols, rows } = read(nome)` é **desestruturação**: a função devolve um
objeto com dois campos e essa linha os coloca em duas variáveis de uma vez.

O resto casa cada nome de coluna com o valor da posição correspondente. Em vez
de `linha[3]`, passa a ser `linha.faturamento` — muito mais difícil de errar.

`a ? b : c` é o **operador ternário**: se `a`, use `b`, senão `c`. Aqui: se a
coluna não existe, use texto vazio.

### Converter texto em número

```js
const num = v => {
  if (v === '' || v === null || v === undefined) return 0;
  const n = parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? 0 : n;
};
```

Tudo que vem do arquivo é texto. `'1234.56'` precisa virar o número 1234,56
para poder somar.

**A parte importante é o que ele faz quando dá errado: devolve 0, nunca
`NaN`.** `NaN` ("not a number") é contagioso — qualquer soma com ele vira
`NaN`, e uma célula problemática estragaria a coluna inteira.

`.replace(',', '.')` troca vírgula por ponto, porque JavaScript só entende
ponto decimal.

### A dimensão empresa

```js
function carregaEmpresas() {
  const cnae = {};
  for (const c of readObj('01b_cnae')) {
    cnae[c.codi_emp] = { cod: (c.i_cnae20 || '').trim(), ramo: (c.ramo_emp || '').trim() };
  }
  ...
  const vazio = c => ({ cod: int(c), nome: '(empresa ' + c + ')', cnpj: '', ... });
  return { emp, cnaeDesc, E: c => emp[c] || vazio(c) };
}
```

Monta uma **tabela de consulta**: dado o código da empresa, devolve nome, CNPJ,
inscrição estadual, cidade. Todos os geradores usam.

`(c.i_cnae20 || '')` é um truque comum: `||` devolve o lado esquerdo se ele não
for vazio, senão o direito. Protege contra campo ausente, que faria o `.trim()`
dar erro.

**`E: c => emp[c] || vazio(c)`** é a proteção final: se aparecer um código de
empresa que não está no cadastro, em vez de quebrar, devolve um registro com
nome `(empresa 123)`. A planilha sai com um aviso visível em vez de um erro.

### Os departamentos

```js
const DEPARTAMENTO = {
  '1': 'Contábil',        // Contabilidade
  '4': 'Contábil',        // Patrimônio
  '5': 'Fiscal',          // Escrita Fiscal
  '6': 'Fiscal',          // Lalur
  '12': 'Pessoal',        // Folha
  '15': 'Societário',     // Registro
  ...
};
const dept = sistLog => DEPARTAMENTO[String(sistLog)] || 'Administrativo';
```

O Domínio **não tem cadastro de departamento**. O que ele grava é o módulo em
que a pessoa estava. Este dicionário é a tradução para o vocabulário do
escritório.

**As escolhas discutíveis:** Lalur é apuração de IRPJ/CSLL, então conta como
Fiscal; Patrimônio é imobilizado, conta como Contábil.

**Se a divisão do seu escritório for outra, edite só este bloco.** As quatro
abas de departamento acompanham sozinhas.

O `|| 'Administrativo'` no fim é o padrão para módulo desconhecido.

---

## 3.3 O padrão que se repete em todo gerador

Antes de ver os geradores um a um, vale entender o molde. Praticamente tudo no
kit é uma variação disto:

```js
const mov = {};

const M = (cod, mes) => {
  const k = cod + '|' + mes;
  if (!mov[k]) mov[k] = { cod, mes, fat: 0, notas: 0 };
  return mov[k];
};

for (const r of readObj('02_faturamento')) {
  const o = M(r.codi_emp, mm(r.ano, r.mes));
  o.fat   += num(r.faturamento);
  o.notas += int(r.notas);
}
```

**Como funciona:**

1. `mov` começa vazio. Vai virar o acumulador.
2. `M(cod, mes)` é uma função auxiliar que **pega ou cria**: monta uma chave
   juntando empresa e mês (`'345|2025-10'`), e se ainda não existir esse balde,
   cria zerado.
3. O laço percorre o arquivo e vai somando no balde certo.

**Por que juntar as duas partes num texto?** Porque JavaScript não tem
dicionário de chave dupla. Concatenar com um separador é a saída padrão. Para
separar de volta: `k.split('|')`.

**A vantagem sobre fazer no SQL:** o mesmo acumulador recebe dados de arquivos
diferentes. `M()` é chamado pelo faturamento, depois pelas compras, depois
pelos impostos — e no fim cada linha tem as três coisas juntas.

`+=` quer dizer "some ao que já estava lá".

---

## 3.4 `fiscal.js` — o mais simples

```js
for (const r of readObj('04_impostos')) {
  const sis = (r.codi_sis || '').trim();
  if (sis !== '5' && sis !== '6') continue;
  ...
}
```

`continue` pula para a próxima volta do laço. Traduzindo: *"se não for do
módulo Fiscal nem do Lalur, ignore"* — os impostos de folha saem no outro
relatório.

### Quais tributos viram coluna

```js
const SIGLAS = Object.entries(siglasTotais)
  .filter(([, v]) => v > 0)
  .sort((a, b) => b[1] - a[1])
  .map(([s]) => s);
```

Três operações encadeadas:

- `Object.entries(x)` transforma o dicionário numa lista de pares
  `[chave, valor]` — aqui, `[['ICMS', 4343220], ['PIS', 81464], ...]`
- `.filter(([, v]) => v > 0)` mantém só os que tiveram valor. O `[, v]` é
  desestruturação **pulando a primeira posição** — só interessa o valor.
- `.sort((a, b) => b[1] - a[1])` ordena. O comparador recebe dois itens e
  devolve negativo, zero ou positivo. `b - a` é **decrescente**; `a - b` seria
  crescente.
- `.map(([s]) => s)` fica só com a sigla.

**As colunas não são fixas.** O código descobre quais tributos existem e cria
uma coluna para cada. Nesta base dá 20. Se você mudar o período e aparecer um
tributo novo, a coluna nasce sozinha.

```js
...SIGLAS.map(s => ({ header: s, key: 'i_' + s, width: 14, style: MOEDA })),
```

O `...` é o **espalhamento**: insere os itens de uma lista dentro de outra. É
assim que as 20 colunas variáveis entram no meio da lista fixa de colunas.

> Repare em `s => ({ ... })`: os parênteses em volta das chaves são
> **obrigatórios**. Sem eles o JavaScript leria `{` como início de bloco de
> código, não como objeto.

```js
impTot: SIGLAS.reduce((a, s) => a + (t.imp[s] || 0), 0)
```

`reduce` **dobra uma lista num valor só**. Recebe uma função `(acumulado, item)`
e um valor inicial (o `0` do final). Aqui: percorre as siglas somando o valor de
cada uma. É o total de impostos da linha.

---

## 3.5 `folha.js` — reconstruindo o quadro de funcionários

O banco **não tem** uma tabela dizendo quantos funcionários a empresa tinha em
março. Isso é reconstruído a partir das datas de admissão e demissão.

### Primeiro, achar a data de desligamento certa

```js
for (const r of readObj('07_rescisoes')) {
  const d = r.demissao || '';
  if (!d) continue;
  const k = r.codi_emp + '|' + r.i_empregados;
  if (!demis[k] || d < demis[k].data) demis[k] = { data: d, motivo: ... };
}
```

*"Para cada rescisão: se ainda não tenho data desse empregado, ou se esta é
mais antiga que a que eu tinha, guarde esta."*

**Por que a mais antiga?** Porque complemento de rescisão gera outra linha para
o mesmo empregado, com data posterior. Se pegássemos a última, o funcionário
apareceria ativo depois de já ter saído.

Datas no formato `'2025-03-11'` podem ser comparadas como texto — `<` funciona
porque ano, mês e dia estão em ordem decrescente de importância.

### Depois, montar a régua de meses

```js
const meses = [];
let [y, m] = INI.split('-').map(Number);
const [fy, fm] = FIM.split('-').map(Number);
while (y < fy || (y === fy && m <= fm)) {
  meses.push(y + '-' + String(m).padStart(2, '0'));
  m++;
  if (m > 12) { m = 1; y++; }
}
```

Gera `['2024-01', '2024-02', ...]`.

`INI.split('-').map(Number)` quebra `'2024-01'` em `['2024','01']` e converte
cada pedaço em número.

`String(m).padStart(2, '0')` transforma `3` em `'03'` — completa com zeros à
esquerda até ter 2 caracteres. Sem isso a ordenação ficaria errada: `'2024-10'`
viria antes de `'2024-3'`.

`m++` soma 1. O `if` faz virar o ano em dezembro.

**`INI` e `FIM` não são fixos** — saem do menor e do maior mês encontrado nos
dados. Assim a planilha acompanha o banco sem editar o código.

### Finalmente, contar

```js
for (const e of empregados) {
  const admMes = e.admissao.slice(0, 7);
  const dm = demis[e.cod + '|' + e.id];
  const demMes = dm ? dm.data.slice(0, 7) : null;

  if (admMes >= INI && admMes <= FIM) H(e.cod, admMes).adm++;
  if (demMes && demMes >= INI && demMes <= FIM) H(e.cod, demMes).dem++;

  for (const m of meses) {
    if (m < admMes) continue;        // ainda não tinha entrado
    if (demMes && m > demMes) break; // já tinha saído
    H(e.cod, m).ativos++;
  }
}
```

`.slice(0, 7)` corta os 7 primeiros caracteres: de `'2025-03-11'` sobra
`'2025-03'`.

O laço de dentro é a regra do headcount, e cabe numa frase:

> **Ativo no mês M se foi admitido até o fim de M e não foi demitido antes do
> começo de M.**

`continue` pula o mês (ainda não tinha sido contratado). `break` **abandona o
laço inteiro** — como os meses estão em ordem, depois da demissão não há mais
nada a contar.

Quem entra e sai no mesmo mês conta como ativo naquele mês, e conta 1 admissão
e 1 demissão.

---

## 3.6 `contabil.js` — natureza e saldo acumulado

```js
const GRUPOS = {
  '1': { nome: 'ATIVO',                        nat: 'D', patrimonial: true },
  '2': { nome: 'PASSIVO E PATRIMÔNIO LÍQUIDO', nat: 'C', patrimonial: true },
  '3': { nome: 'CUSTOS E DESPESAS',            nat: 'D', patrimonial: false },
  '4': { nome: 'RECEITAS',                     nat: 'C', patrimonial: false },
  ...
};

const saldoNat = r => {
  const nat = (GRUPOS[r.grupo] || {}).nat || 'D';
  return nat === 'D' ? r.debito - r.credito : r.credito - r.debito;
};
```

O grupo da conta é o **primeiro dígito** da classificação: a conta `1.1.01`
está no grupo 1, Ativo.

Conta de **natureza devedora** (Ativo, Despesa) tem saldo positivo quando o
débito é maior. Conta de **natureza credora** (Passivo, Receita) é o contrário.
Essa função inverte a subtração conforme o caso.

`(GRUPOS[r.grupo] || {}).nat || 'D'` é defesa em dois níveis: se o grupo não
existir no dicionário, use objeto vazio; se mesmo assim não tiver `nat`, use
devedora. Sem isso, uma conta com classificação estranha derrubaria o gerador.

### O saldo acumulado

```js
raw.sort((a, b) => a.cod - b.cod || a.conta - b.conta || a.ano - b.ano);

let chave = '', acumulado = 0;
for (const r of raw) {
  const k = r.cod + '|' + r.conta;
  if (k !== chave) { chave = k; acumulado = 0; }
  r.saldoAno  = saldoNat(r);
  acumulado  += r.saldoAno;
  r.saldoAcum = acumulado;
}
```

**Este é o segundo padrão do kit: acumulação em sequência ordenada.**

`sort` com `||` encadeia critérios: ordena por empresa; se empatar (o
comparador devolve 0, que é falso), desempata por conta; se empatar de novo,
por ano.

Depois o laço vai somando. **A linha crítica é a reinicialização:**

```js
if (k !== chave) { chave = k; acumulado = 0; }
```

*"Se mudei de conta, zere o acumulador."* Sem isso, o saldo de uma conta
vazaria para a seguinte e todos os números depois da primeira estariam errados.

Ordenar antes é obrigatório — o acumulador só funciona se os anos vierem em
ordem.

---

## 3.7 `caixa.js` — por que este é diferente

```js
const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: DEST, useStyles: true });
```

Os outros geradores montam a planilha inteira na memória e gravam no fim. Este
usa o modo **streaming**: grava conforme vai montando.

**Por quê?** São 881.745 linhas de detalhe. Segurar isso tudo na memória
estoura o limite do Node numa máquina com pouca RAM.

```js
const rl = readline.createInterface({
  input: fs.createReadStream(OUT + '/31_caixa_detalhe.tsv', { encoding: 'latin1' }),
  crlfDelay: Infinity
});

for await (const line of rl) {
  const [cod, data, numero, cdeb, ccre, valor, hist, doc, oper] = line.split('\\t');
  const ehDeb = contasCaixa.has(cod + '|' + cdeb);
  ...
  ws.addRow({ ..., entrada: ehDeb ? v : 0, saida: ehDeb ? 0 : v }).commit();
}
```

A leitura também é em streaming: `createReadStream` + `readline` entregam
**uma linha por vez**, sem nunca carregar os 77 MB.

`for await` é a versão assíncrona do `for` — espera cada linha chegar.

`.commit()` no fim de cada linha escreve e **libera a memória dela**. Sem isso o
streaming não adiantaria nada.

`const [a, b, c] = texto.split('\\t')` é desestruturação de lista: quebra a
linha e já distribui as partes em variáveis nomeadas.

`contasCaixa.has(...)` consulta um **Set** — uma coleção de valores únicos, com
busca instantânea. É como decidimos se aquela conta é caixa.

**A regra do livro caixa:** conta de caixa no débito é **entrada** (dinheiro
entrando); no crédito é **saída**.

---

## 3.8 `produtividade.js` — o cálculo mais elaborado

### Passo 1: colocar tudo numa régua de tempo única

```js
const diaAbs = d => {
  const [y, m, dd] = String(d).split('-').map(Number);
  return Date.UTC(y, m - 1, dd) / 60000;
};

const ini = diaAbs(r.data_ini) + hms(r.hora_ini);
let fim = diaAbs(r.data_fim || r.data_ini) + hms(r.hora_fim);
if (fim < ini) fim = ini;
```

`Date.UTC(...)` devolve os milissegundos desde 1970. Dividido por 60.000, vira
**minutos**. Somando a hora do dia, cada momento vira um número só.

**Por que `Date.UTC` e não a data normal?** Porque ele **não depende de fuso
nem de horário de verão**. Como só interessam diferenças, usar a data local
introduziria um erro de 60 minutos nas viradas de horário de verão.

O `m - 1` existe porque em JavaScript os meses são numerados de 0 a 11 — janeiro
é 0. É uma das pegadinhas clássicas da linguagem.

**Benefício de graça:** a sessão que atravessa a meia-noite funciona sem
tratamento especial. O fim simplesmente cai num número maior. São 104 sessões
assim na base.

### Passo 2: o problema das sessões sobrepostas

Caso real, ALINE em 11/03/2025 na empresa 345:

```
módulo 1 (Contabilidade)   05:30:42 → 06:15:20   45 min
módulo 5 (Escrita Fiscal)  06:03:59 → 06:15:25   11 min
```

Ela abriu o Fiscal **sem fechar** o Contábil. Somar dá 56 minutos numa janela de
relógio de 45. Na base inteira: **41.081 h somadas para 37.675 h reais — 8,3% a
mais**, afetando 55 dos 61 usuários.

### Passo 3: a varredura que corrige

```js
function rateiaConcorrencia(sessoes) {
  const porUsu = {};
  sessoes.forEach((s, i) => { (porUsu[s.usuario] = porUsu[s.usuario] || []).push(i); });

  for (const idxs of Object.values(porUsu)) {
    const pontos = [...new Set(idxs.flatMap(i => [sessoes[i].ini, sessoes[i].fim]))]
      .sort((a, b) => a - b);
    const ordem = idxs.slice().sort((a, b) => sessoes[a].ini - sessoes[b].ini);

    let p = 0;
    const ativos = new Set();
    for (let k = 0; k < pontos.length - 1; k++) {
      const ini = pontos[k], fim = pontos[k + 1], dur = fim - ini;
      while (p < ordem.length && sessoes[ordem[p]].ini <= ini) { ativos.add(ordem[p]); p++; }
      for (const i of ativos) if (sessoes[i].fim <= ini) ativos.delete(i);
      if (dur <= 0 || ativos.size === 0) continue;
      const fatia = dur / ativos.size;
      for (const i of ativos) sessoes[i].minAjust += fatia;
    }
  }
}
```

Esta é a parte mais difícil do kit. Vamos devagar.

**A ideia central:** entre dois eventos consecutivos (uma abertura ou um
fechamento), **o conjunto de sessões abertas não muda**. Então dá para tratar
esse pedaço como um bloco e dividir a duração pelo número de sessões abertas.

**Linha 1–2 — agrupar por usuário.**
```js
sessoes.forEach((s, i) => { (porUsu[s.usuario] = porUsu[s.usuario] || []).push(i); });
```
`forEach` com dois parâmetros dá o item e a **posição** dele. Guardamos a
posição, não a sessão, para poder alterar a original depois.

A sobreposição só existe dentro do mesmo usuário — duas pessoas trabalhando ao
mesmo tempo são duas horas de verdade.

**Linha 3 — reunir os instantes em que algo muda.**
```js
const pontos = [...new Set(idxs.flatMap(i => [sessoes[i].ini, sessoes[i].fim]))].sort(...)
```
- `flatMap` — cada sessão vira dois números (início e fim), e o resultado é uma
  lista única, não uma lista de pares
- `new Set(...)` — remove repetidos
- `[...]` — converte o Set de volta em lista, para poder ordenar
- `.sort((a,b) => a-b)` — ordem crescente

**Linha 4 — a lista ordenada por início**, percorrida por um ponteiro que só
anda para a frente. Evita revarrer tudo a cada trecho.

`idxs.slice()` faz uma **cópia** antes de ordenar — `sort` altera a lista
original, e não queremos bagunçar a de fora.

**O laço principal:**

```js
const ini = pontos[k], fim = pontos[k + 1], dur = fim - ini;
```
Pega o trecho entre dois eventos consecutivos.

```js
while (p < ordem.length && sessoes[ordem[p]].ini <= ini) { ativos.add(ordem[p]); p++; }
```
*"Adicione todas as sessões que já começaram."* O ponteiro `p` nunca volta.

```js
for (const i of ativos) if (sessoes[i].fim <= ini) ativos.delete(i);
```
*"Remova as que já terminaram."*

```js
const fatia = dur / ativos.size;
for (const i of ativos) sessoes[i].minAjust += fatia;
```
**A divisão.** Se só uma sessão está aberta, ela leva o trecho inteiro. Se duas,
cada uma leva metade.

**No exemplo da ALINE:** o trecho 05:30→06:03 tem só a Contabilidade, que fica
com os 33,3 minutos. O trecho 06:03→06:15 tem duas abertas, então cada uma leva
5,7 em vez de 11,4. Soma final: 45 minutos — o relógio.

### Passo 4: o teste

O arquivo `teste-rateio.js` confere o **invariante**: para cada usuário, a soma
das fatias tem que ser igual à **união** dos intervalos dele.

A união é calculada lá por outro método — juntando intervalos que se tocam —
**de propósito**: se dois caminhos diferentes chegam ao mesmo número, é porque
está certo. Se usássemos o mesmo raciocínio, repetiríamos o mesmo erro.

```bash
cd "C:\Users\LENOVO\Documents\BECKUP DOMINIO\SCRIPTS\_gerar" && node teste-rateio.js x C:/dbtmp/out
```

Resultado atual: bate nos 61 usuários, maior erro 2×10⁻¹⁰ minutos — que é só
ruído de arredondamento de casas decimais.

**Rode isso sempre que mexer no rateio.**

---

# Parte 4 — Glossário

| Termo | O que é |
|:--|:--|
| **Argumento** | valor que você entrega a uma função ou programa |
| **Array / Lista** | vários valores em sequência, acessados por posição |
| **BOM** | três bytes invisíveis no começo de um arquivo, que dizem a codificação |
| **Cache** | memória que o servidor usa para não reler o disco |
| **cp1252 / latin1** | codificação antiga do Windows para acentos |
| **Desestruturação** | tirar vários campos de um objeto em uma linha só |
| **Dicionário / Hashtable / Objeto** | valores acessados por nome |
| **Driver ODBC** | tradutor que deixa o Windows falar com o banco |
| **Escape** | caractere que muda o sentido do próximo (`` ` `` no PowerShell, `\` no JS) |
| **Expressão regular** | padrão de busca em texto, escrito entre barras |
| **Função** | bloco de código com nome, que se manda rodar |
| **Invariante** | afirmação que tem que ser sempre verdadeira; base de um teste |
| **JOIN** | juntar duas tabelas do banco pelo campo em comum |
| **Laço** | repetição para cada item |
| **NaN** | "não é número"; resultado de conta inválida, contagioso |
| **NULL** | campo vazio no banco; qualquer conta com ele dá vazio |
| **Parâmetro** | o que a função declara que vai receber |
| **Pipeline / Cano (`\|`)** | passar o resultado de um comando para o seguinte |
| **Schema (`bethadba`)** | o "sobrenome" das tabelas do Domínio |
| **Set** | coleção de valores únicos, com busca instantânea |
| **Streaming** | processar aos poucos, sem carregar tudo na memória |
| **String** | texto |
| **Sweep line / Varredura** | percorrer eventos em ordem cronológica |
| **TSV** | arquivo de texto com colunas separadas |
| **UNLOAD** | comando que manda o servidor gravar o resultado num arquivo |
| **Variável de ambiente** | valor definido no terminal, herdado pelos programas |

---

# Parte 5 — Se você for mexer

**Mudar o período** → só `_lib/config.ps1`.

**Mudar os departamentos** → só o dicionário `DEPARTAMENTO` em `_gerar/lib.js`.

**Mudar uma fórmula da planilha** → o `.js` correspondente. Não precisa religar
o banco: os `.tsv` já estão em `C:\dbtmp\out`. Rode direto:

```bash
cd "C:\Users\LENOVO\Documents\BECKUP DOMINIO\SCRIPTS\_gerar" && node fiscal.js "C:/dbtmp/teste.xlsx" C:/dbtmp/out
```

**Mudar uma consulta** → o `.ps1` correspondente. Aí sim o banco precisa estar
no ar.

**Antes de salvar um `.ps1`**, garanta que ficou em UTF-8 **com BOM**. O
PowerShell 5.1 lê arquivo sem BOM como se fosse codificação antiga, os acentos
viram lixo e o script para de funcionar com um erro que aponta para uma linha
que parece perfeita. Para corrigir:

```powershell
$bom = New-Object System.Text.UTF8Encoding($true)
Get-ChildItem . -Recurse -Filter *.ps1 | ForEach-Object {
  $t = [IO.File]::ReadAllText($_.FullName, [Text.UTF8Encoding]::new($false))
  [IO.File]::WriteAllText($_.FullName, $t, $bom)
}
```

**Depois de qualquer mudança**, confira a empresa 345 em 10/2025:

| Valor | Esperado |
|:--|--:|
| Faturamento | 362.025,28 |
| Compras | 699.955,52 |
| Impostos | 55,84 |
| Folha | 4.818,15 |
