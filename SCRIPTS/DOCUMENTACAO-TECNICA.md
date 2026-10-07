# Documentação técnica da extração do Domínio

Como cada script funciona, de onde ele tira cada número e como calcula cada
coluna das planilhas.

Banco: `contabil.db`, SAP SQL Anywhere 17.0.11.7312, schema `bethadba`,
30,84 GB, charset windows-1252, 4.800 tabelas base.

As seções 1 a 18 explicam o **que** cada código faz e como cada número é
calculado. O **Apêndice A**, no fim, explica as três linguagens envolvidas
(SQL, PowerShell e JavaScript) e cada construção de sintaxe usada.

> **Se você não programa**, comece por [CODIGO-EXPLICADO.md](CODIGO-EXPLICADO.md):
> ele percorre o código de verdade em pedaços, explicando cada um em português,
> sem supor conhecimento prévio. Este documento aqui assume que você já sabe ler
> código e quer saber de onde vem cada número.

---

## 1. Arquitetura

O caminho dos dados tem três etapas, sempre na mesma ordem:

```
  contabil.db                    C:\dbtmp\out\*.tsv              EXTRACAO\*.xlsx
       │                                │                              │
       │   UNLOAD (servidor grava)      │   Node + exceljs             │
       └───────────────►────────────────┴──────────►───────────────────┘
            PowerShell manda o SQL           JS lê, cruza, calcula
```

**Por que duas etapas em vez de uma.** O PowerShell não faz cálculo nenhum:
ele só manda o `SELECT` e o servidor cospe um arquivo. Todo cruzamento,
acumulação e formatação acontece no Node. Isso deixa reprocessar a planilha
sem tocar no banco — se você mudar uma fórmula, roda só o `.js` e leva
segundos, sem religar um servidor de 30 GB.

**Divisão de responsabilidade:**

| Camada | Arquivo | Faz | Não faz |
|:--|:--|:--|:--|
| Configuração | `_lib/config.ps1` | caminhos, período, credenciais | nada executável |
| Acesso | `_lib/dominio.ps1` | sobe servidor, roda SQL, exporta TSV | cálculo |
| Extração | `01..06-*.ps1` | o SQL de cada relatório | formatação |
| Cálculo | `_gerar/*.js` | cruzamentos, saldos, durações | acesso ao banco |
| Leitura | `_gerar/lib.js` | parse do TSV, dimensão empresa, estilo | SQL |

---

## 2. Por que `UNLOAD` e não leitura normal

Este é o ponto que define o desempenho de tudo.

A forma óbvia de extrair seria abrir um `OdbcDataReader` e percorrer as linhas.
**Nesta base isso faz cerca de 60 linhas por minuto.** As 655 empresas de
`geempre` levavam 4 minutos; as 226 mil linhas de `geloguser` levariam dias.
O gargalo não é a consulta — é o transporte linha a linha pelo driver, com o
banco paginando em disco.

A saída é o `UNLOAD`:

```sql
UNLOAD SELECT ... TO 'C:/dbtmp/out/arquivo.tsv' DELIMITED BY '\t' QUOTES OFF ESCAPES OFF
```

Quem escreve o arquivo é o **servidor**, direto no disco. Nada trafega pelo
driver. As mesmas 655 empresas saem em 0,5 s, e o `geloguser` inteiro em 1,8 s.

**Duas condições para funcionar:**

1. O servidor tem que subir com **`-gl all`**. O padrão é `dba`, e o usuário
   `externo` leva `Permission denied: you do not have permission to use the
   "UNLOAD" statement`. O `Start-DomServidor` já passa a flag.
2. O caminho do arquivo tem que ser **local, curto e sem espaço**, com **barra
   normal**. `C:/dbtmp/out/x.tsv` funciona; `C:\dbtmp\out\x.tsv` não — a barra
   invertida se perde no caminho até o SQL.

---

## 3. O arquivo que o `UNLOAD` gera

Ele se chama `.tsv` mas não é um TSV comum. Três diferenças que quebram
qualquer leitor ingênuo:

**O separador não é TAB.** `DELIMITED BY '\t'` é interpretado literalmente: o
separador vira a sequência de **dois caracteres** `\` + `t`. Por isso o
`lib.js` faz `line.split('\\t')` e não `split('\t')`. Isso é até vantagem —
um TAB de verdade dentro de um campo de texto não quebra a linha.

**O encoding é cp1252**, não UTF-8. `GRAJAÚ` termina com o byte `0xDA`. O
`lib.js` lê com `.toString('latin1')`.

**Não tem cabeçalho.** O `Export-DomTabela` grava os nomes de coluna num
arquivo `.cols` ao lado, e o `read()` do `lib.js` junta os dois.

`QUOTES OFF ESCAPES OFF` mantém o texto cru. Funciona porque nenhuma consulta
exporta campo com quebra de linha — campos de observação ficam de fora de
propósito.

---

## 4. `_lib/config.ps1`

Um hashtable global `$DOM`. Não executa nada; é lido por todo o resto.

| Campo | Para que serve |
|:--|:--|
| `Banco`, `Engine` | caminho do `.db` e do `dbsrv17.exe` |
| `Servidor`, `Porta`, `Usuario`, `Senha` | montam a string de conexão ODBC |
| `Cache`, `CacheMax` | `-c` e `-ch` do servidor |
| `PastaTmp`, `PastaDados` | onde o servidor grava (curto, sem espaço) |
| `PastaSaida` | onde os `.xlsx` nascem |
| `DataIni`, `DataFim` | período do fiscal, folha, caixa e produtividade |
| `DataIniContabil` | 1995 — só o balancete usa |
| `DataIniLog` | 2023 — o log de acesso tem mais história |

**Por que `DataIniContabil` é separado.** O saldo acumulado das contas
patrimoniais é a soma de tudo desde o começo. Cortar em 2024 daria um balanço
sem saldo de abertura — o Ativo sairia com o movimento de dois anos, não com o
que a empresa tem.

**Sobrescrita por ambiente.** No fim do arquivo:

```powershell
if ($env:DOM_PASTA_SAIDA) { $Global:DOM.PastaSaida = $env:DOM_PASTA_SAIDA }
```

Existe porque **cada script filho recarrega o `config.ps1`**. Se o
`EXTRAIR-TUDO` mudasse `$DOM` na memória, o `01-FISCAL.ps1` sobrescreveria de
volta ao ser chamado. Variável de ambiente é herdada pelo processo filho e
sobrevive ao recarregamento. Valem `DOM_DATA_INI`, `DOM_DATA_FIM`,
`DOM_PASTA_SAIDA`, `DOM_PASTA_DADOS` e `DOM_BANCO`.

---

## 5. `_lib/dominio.ps1`

### `Get-DomConexao`

Monta a string. Note `Server` e `DBN` com o mesmo valor (`contabil`): o nome
do servidor e o nome do banco coincidem. Usa conexão por driver, não DSN —
a DSN `Contabil` registrada na máquina é de 32 bits e não serve.

### `Test-DomServidor`

Abre e fecha a conexão. Retorna booleano. É o teste que todos usam antes de
decidir se precisam subir o servidor.

### `Start-DomServidor`

Três detalhes não óbvios:

**Sobe por um `.bat`.** `Start-Process` com `-ArgumentList` quebra o caminho
`C:\Users\LENOVO\Documents\BECKUP DOMINIO\extraido\contabil.db` no espaço de
"BECKUP DOMINIO", e o servidor reclama que `C:\Users\LENOVO\Documents\BECKUP`
não existe. Escrever a linha num `.bat` e executar o `.bat` resolve.

**Passa `-gl all`** — sem isso todo o resto do kit não funciona.

**Espera em laço** até `Test-DomServidor` responder, com teto de 180 s. O
recovery de um banco de 30 GB costuma levar 10 s, mas pode levar minutos se o
log de transações estiver grande.

### `Invoke-DomSql` / `Get-DomValor` / `Get-DomColunas` / `Test-DomTabela`

Consulta normal via `OdbcDataReader`, devolvendo objetos PowerShell. **São
para consulta pequena**: catálogo, `COUNT(*)`, conferência. Puxar volume por
aqui é o erro que custa dias. O `00-BANCO.ps1` usa; os relatórios não.

### `Export-DomTabela`

O coração. Recebe `-Nome`, `-Sql` e `-Header`, monta o `UNLOAD` e grava.

```powershell
$caminhoSql = ($destino -replace '\\', '/')   # barra normal
$TAB = "'" + [char]92 + "t'"                  # a barra montada por código
$cmd.CommandText = "UNLOAD $Sql TO '$caminhoSql' DELIMITED BY $TAB QUOTES OFF ESCAPES OFF"
```

`$TAB` é montado com `[char]92` em vez de escrito literalmente porque a barra
invertida dentro de string em PowerShell atravessa várias camadas de
interpretação até chegar ao SQL, e some no caminho.

**Erro não interrompe o lote.** O `catch` imprime `[FALHA]` e segue. A escolha
é deliberada: se uma consulta falhar, os outros arquivos ainda são gerados, e o
gerador avisa depois o que faltou. A alternativa — abortar tudo — desperdiçaria
uma extração inteira por causa de um arquivo.

**A função não retorna nada.** Retornar `$true` faria o PowerShell imprimir
`True` solto no meio do log, porque o valor de retorno não capturado vai para o
stdout.

### `Test-ArquivoEmUso`

Tenta abrir o arquivo com `FileShare.None`. Se o Excel estiver com ele aberto,
a abertura falha e a função devolve `$true`. Usada para abortar **antes** de
extrair, em vez de descobrir o problema depois de cinco minutos de trabalho.

### `Invoke-DomGerador`

Chama `node --max-old-space-size=2400 <gerador>.js <destino> <pastaDados>`.
Os dois argumentos chegam no JS como `process.argv[2]` e `argv[3]`.

O limite de memória é explícito porque o gerador do livro caixa processa
881 mil linhas; o padrão do Node estoura.

---

## 6. `_gerar/lib.js`

### `read(nome)` e `readObj(nome)`

```js
const txt = fs.readFileSync(file).toString('latin1');
rows.push(line.split('\\t').map(v => v.trim()));
```

Lê o `.tsv` em latin1, quebra no separador literal `\t`, casa com os nomes do
`.cols` e devolve array de objetos.

### `num(v)` e `int(v)`

`num` troca vírgula por ponto, faz `parseFloat` e devolve **0** para vazio,
nulo ou não numérico. Nunca `NaN`. Isso é o que impede uma célula vazia de
contaminar uma soma inteira.

### `carregaEmpresas()`

Monta a dimensão empresa usada por todos os geradores, cruzando três arquivos:

- `01_empresas` — nome, CNPJ, IE, IM, cidade, UF, situação
- `01b_cnae` — `i_cnae20` e `ramo_emp`
- `24_cnae20` — descrição do CNAE

Derivações:

| Campo | Regra |
|:--|:--|
| `pf` | CNPJ com 11 dígitos depois de tirar pontuação |
| `cnpj` | formatado conforme 11 ou 14 dígitos |
| `status` | `A` → Ativa, `I` → Inativa |
| `atividade` | `ramo_emp` se houver; senão a descrição do CNAE |

`E(cod)` devolve a empresa ou um objeto vazio com nome `(empresa 123)` — assim
um código órfão no movimento não derruba o gerador.

### `fazAba(wb, nome, cols, linhas, cor)`

Cria a aba com cabeçalho branco em fundo colorido, painel congelado na primeira
linha e nas duas primeiras colunas, e autofiltro. Todas as abas têm o mesmo
tratamento.

---

## 7. `00-BANCO.ps1`

Não extrai nada. Cinco ações: `status`, `iniciar`, `parar`, `reiniciar`,
`identificar`.

`identificar` é o diagnóstico completo:

| Bloco | O que consulta |
|:--|:--|
| Arquivos | tamanho do `.db` e do `.log`, versão do `dbsrv17.exe` |
| Servidor | processo, memória, uptime, conexão, RAM livre da máquina |
| Permissões | faz um `UNLOAD SELECT 1` de teste e apaga |
| Identificação | `DB_PROPERTY('CharSet')`, `PageSize`, `Collation`, versão |
| Empresas | total, ativas, inativas, pessoas físicas, com CNAE |
| Período | `MIN`/`MAX`/`COUNT` de 8 tabelas-chave |
| Módulos | `geloguser` cruzado com `GEMODULOS` |

O bloco de período filtra `>= '2000-01-01' AND <= '2035-12-31'`. Sem isso o
`MIN` volta 1979 e o `MAX` volta 8320 — há lixo de digitação nas datas.

O teste de `UNLOAD` existe porque é a falha mais provável: se alguém subir o
Domínio por fora, sem `-gl all`, todo o kit para.

---

## 8. `00-CADASTROS.ps1`

Extrai o que todos os relatórios precisam. Roda em poucos segundos e sai cedo
se os seis arquivos já existirem, a menos que você passe `-Forcar`.

| Arquivo | Origem | Observação |
|:--|:--|:--|
| `01_empresas` | `geempre` | 655 linhas, sem filtro — inativas incluídas |
| `01b_cnae` | `geempre` | só `i_cnae20` e `ramo_emp` |
| `24_cnae20` | `gecnae20` | tabela de descrições |
| `20_plano_contas` | `ctcontas` | 308.853 linhas |
| `21_contas_caixa` | `ctcontacaixa` | 453 linhas, define o livro caixa |
| `23_usuarios` | `GESECUSUARIOS` | `geusuarios` está **vazia** nesta base |

**Por que `i_cnae20` e não `cnae_emp`.** `cnae_emp` é `double` e está vazio em
quase toda a base; quando tem valor, perde o zero à esquerda. `i_cnae20` é
`char(15)`, está preenchido em 603 das 655 empresas e é o CNAE 2.0 de verdade.

---

## 9. `01-FISCAL.ps1` + `_gerar/fiscal.js`

### Extração

**Faturamento** — `02_faturamento`, agregado por empresa, ano e mês:

```sql
SELECT codi_emp, YEAR(dsai_sai), MONTH(dsai_sai),
       SUM(ISNULL(vcon_sai,0)), SUM(ISNULL(vprod_sai,0)), COUNT(*)
FROM bethadba.efsaidas
WHERE dsai_sai BETWEEN :ini AND :fim
  AND situacao_sai = 0
  AND (cancelada_sai IS NULL OR cancelada_sai <> 'S')
GROUP BY codi_emp, YEAR(dsai_sai), MONTH(dsai_sai)
```

**`vcon_sai`, não `vprod_sai`.** `vprod_sai` é o valor dos produtos;
`vcon_sai` é o valor contábil da nota, que é o que o Domínio mostra como
faturamento. Na empresa 345 em 10/2025: `vprod_sai` = 364.042,70 e
`vcon_sai` = 362.025,28. O segundo é o certo. O primeiro vai junto na planilha,
como coluna separada, para conferência.

Os dois filtros são obrigatórios: `situacao_sai = 0` tira nota não efetivada e
`cancelada_sai` tira cancelada. Sem eles o faturamento infla.

**Compras** — `03_compras`, mesma lógica em `efentradas` com `vcon_ent`,
`ddoc_ent` e `situacao_ent = 0`.

**Impostos** — `04_impostos`, por empresa, competência e tributo:

```sql
SELECT d.codi_emp, d.data_sim, g.sigl_imp, g.nome_imp, g.codi_sis,
       SUM(ISNULL(d.sdev_sim,0)), SUM(ISNULL(d.scre_sim,0)),
       SUM(ISNULL(d.bcal_sim,0)), SUM(ISNULL(d.vlrpago_sim,0)), COUNT(*)
FROM bethadba.efsdoimp d
JOIN bethadba.geimposto g ON g.codi_emp = d.codi_emp AND g.codi_imp = d.codi_imp
GROUP BY d.codi_emp, d.data_sim, g.sigl_imp, g.nome_imp, g.codi_sis
```

**`sdev_sim`, não `vime_sim`.** `sdev_sim` é o saldo devedor apurado — o
imposto. `vime_sim` é base de ICMS. Na empresa 345 em 10/2025, `vime_sim` dá
64.878,01 e `sdev_sim` dá 55,84. O segundo é o imposto.

O `JOIN` com `geimposto` é o que traduz `codi_imp` (um número por empresa) na
sigla legível. `codi_sis` diz de que módulo o tributo é: **5** Escrita Fiscal,
**6** Lalur, **12** Folha.

### Cálculo no `fiscal.js`

**Filtro de módulo.** O gerador descarta `codi_sis = 12`:

```js
if (sis !== '5' && sis !== '6') continue;
```

Tributo de folha sai no relatório 02, não aqui. Na prática `efsdoimp` só tem
`codi_sis = 5` nesta base, mas o filtro documenta a intenção.

**Quais tributos viram coluna.** Não são fixos. O gerador soma cada sigla no
período e promove a coluna só as que tiveram valor, da maior para a menor:

```js
const SIGLAS = Object.entries(siglasTotais).filter(([, v]) => v > 0)
  .sort((a, b) => b[1] - a[1]).map(([s]) => s);
```

Dá 20 colunas nesta base. **Não force ICMS/PIS/COFINS**: a maioria das empresas
é Simples Nacional e apura `SIMPLESN`, não os tributos separados.

**Total de impostos** = soma das colunas de sigla da própria linha:

```js
impTot: SIGLAS.reduce((a, s) => a + (t.imp[s] || 0), 0)
```

O resumo por empresa é a soma das competências. Ordenado por faturamento
decrescente.

---

## 10. `02-FOLHA.ps1` + `_gerar/folha.js`

O relatório com mais armadilha do conjunto.

### De onde vem cada valor

| Valor | Tabela | Coluna |
|:--|:--|:--|
| Proventos, descontos | `fobasesserv` | `proventos`, `descontos` |
| INSS patronal | `foguiainss` | `empresa` |
| RAT/SAT | `foguiainss` | `acid_trab` |
| Terceiros | `foguiainss` | `terceiros` |
| INSS segurados | `foguiainss` | `segurados` |
| PIS sobre folha | `foguiapis` | `valor` |
| Admissão | `foempregados` | `admissao` |
| Demissão | `forescisoes` | `demissao` |

**O encargo vem da guia, não das bases.** As colunas `val_inss_empresa`,
`val_acid_trabalho`, `val_inss_terceiro`, `valor_fgts`, `valor_inss` e
`valor_irrf` existem em `fobasesserv` e estão **nulas em toda a base**. A
primeira versão deste script somava elas e produziu uma planilha de zeros.
O valor real está em `foguiainss`.

### `dias_servico > 0`

```sql
WHERE competencia BETWEEN :ini AND :fim AND dias_servico > 0
```

Cada empregado aparece **duas vezes** na mesma competência em `fobasesserv`:
uma linha com `dias_servico = 0` e outra com os dias trabalhados. Somar sem o
filtro dobra a folha. Conferência: empresa 345 em 10/2025 dá 4.818,15 com o
filtro e 9.636,30 sem.

### `SUM(ISNULL(a,0) + ISNULL(b,0))`

Todas as somas usam `ISNULL`. Em SQL, `SUM(a + b)` vira **NULL** se qualquer
parcela de qualquer linha for nula — uma única linha com nulo zera a coluna
inteira. Foi assim que a primeira extração saiu vazia.

### Cálculo do encargo patronal

```js
patTot = patronal + rat + terceiros + pisFolha
```

`segurados` **não entra**: é a parte descontada do empregado, não custo do
empregador. Vai na planilha em coluna separada, para conferir com a guia.

**Patronal zerado não é erro.** Só 909 de 3.590 competências têm valor, porque
empresa do Simples Nacional recolhe a parte patronal dentro do DAS. Nessas
linhas sobra terceiros e a parte do empregado.

### Headcount: como "funcionários ativos por mês" é calculado

O banco não tem uma tabela de headcount. É reconstruído a partir das datas.

**Primeiro, a data de demissão de cada empregado:**

```js
if (!demis[k] || d < demis[k].data) demis[k] = { data: d, motivo: ... };
```

Complemento de rescisão gera outra linha para o mesmo empregado. Vale a **data
mais antiga**, senão o funcionário aparece ativo depois de desligado.

**Depois, a régua de meses.** `INI` e `FIM` não são fixos: saem do menor e do
maior mês presente em `05_folha_encargos` e `05b_guia_inss`. Assim a planilha
acompanha o banco sem precisar editar o script.

**Então, mês a mês:**

```js
for (const m of meses) {
  if (m < admMes) continue;        // ainda não tinha entrado
  if (demMes && m > demMes) break; // já tinha saído
  H(e.cod, m).ativos++;
}
```

A regra é: **ativo no mês M se foi admitido até o fim de M e não foi demitido
antes do começo de M**. Quem é admitido e demitido no mesmo mês conta como
ativo nesse mês, e conta 1 admissão e 1 demissão.

Admissões e demissões são contadas na competência da própria data.
`saldo = admissões − demissões`.

### O que não tem

**FGTS.** Verificado contra o catálogo completo: `VALOR_FGTS` soma zero em
`FOBASES`, `FOBASESSERV`, `FOCONTRBASES` e `FOVBASES`, e não existe tabela de
guia de FGTS. O Domínio calcula na hora de gerar o SEFIP / FGTS Digital e não
grava o valor. Só sai pelos relatórios do próprio sistema.

**Legenda do motivo de demissão.** `forescisoes.motivo` é um código e não há
tabela de descrição no banco — os textos estão no executável. A planilha traz o
código; a legenda está no Domínio, em Folha > Rescisões.

---

## 11. `03-CONTABIL.ps1` + `_gerar/contabil.js`

### A partida dobrada em duas colunas

`ctlancto` guarda cada lançamento numa linha só, com a conta de débito em
`cdeb_lan` e a de crédito em `ccre_lan`. Para virar balancete, cada lançamento
precisa entrar **duas vezes** — uma como débito de uma conta, outra como
crédito de outra. Daí a `UNION ALL`:

```sql
SELECT codi_emp, ano, conta, SUM(debito), SUM(credito), SUM(qtd) FROM (
  SELECT codi_emp, YEAR(data_lan), cdeb_lan, SUM(vlor_lan), 0, COUNT(*)
  FROM bethadba.ctlancto WHERE ... AND cdeb_lan > 0
  GROUP BY codi_emp, YEAR(data_lan), cdeb_lan
  UNION ALL
  SELECT codi_emp, YEAR(data_lan), ccre_lan, 0, SUM(vlor_lan), COUNT(*)
  FROM bethadba.ctlancto WHERE ... AND ccre_lan > 0
  GROUP BY codi_emp, YEAR(data_lan), ccre_lan
) AS t
GROUP BY codi_emp, ano, conta
```

São duas varreduras da tabela de 5 milhões de linhas, mas num comando só, e o
servidor resolve em 4 segundos.

### Grupo e natureza

O grupo é o **primeiro dígito de `clas_cta`**:

| Dígito | Grupo | Natureza | Saldo |
|:--|:--|:--|:--|
| 1 | Ativo | devedora | débito − crédito |
| 2 | Passivo e PL | credora | crédito − débito |
| 3 | Custos e despesas | devedora | débito − crédito |
| 4 | Receitas | credora | crédito − débito |
| 5 | Apuração do resultado | credora | crédito − débito |
| 6 | Encerramento | credora | crédito − débito |

```js
const saldoNat = r => {
  const nat = (GRUPOS[r.grupo] || {}).nat || 'D';
  return nat === 'D' ? r.debito - r.credito : r.credito - r.debito;
};
```

Contas sem classificação caem em devedora, por segurança.

### Saldo acumulado

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

Ordena por empresa, conta e ano, e vai somando. O acumulador zera a cada conta
nova. É por isso que a extração começa em 1995: cortar em 2024 daria um saldo
que não é saldo.

### As quatro abas

| Aba | Usa | Agrupa por |
|:--|:--|:--|
| Balancete analítico | as duas colunas de saldo | nada, é a linha crua |
| Balanço Patrimonial | `saldoAcum`, grupos 1 e 2 | empresa, ano, `clas_cta[0:2]` |
| DRE | `saldoAno`, grupos 3 e 4 | empresa, ano, `clas_cta[0:2]` |
| Resumo | ambos | empresa, ano |

**Patrimonial usa acumulado, resultado usa o ano.** Conta de resultado é zerada
no encerramento do exercício; o que importa nela é o movimento do período.
Conta patrimonial carrega saldo de um ano para o outro.

No resumo:

```
Ativo      = Σ saldoAcum das contas do grupo 1
Passivo+PL = Σ saldoAcum das contas do grupo 2
Receitas   = Σ saldoAno  das contas do grupo 4
Despesas   = Σ saldoAno  das contas do grupo 3
Resultado  = Receitas − Despesas
Diferença  = Ativo − Passivo
```

**A coluna "Diferença Ativo − Passivo" é a sua conferência.** Onde ela não for
zero, ou o banco não tem saldo de abertura daquela empresa, ou o exercício não
foi encerrado. Deixei explícita em vez de esconder.

A descrição de cada classificação sai de `nomeClas`, que guarda o nome da
primeira conta encontrada com aquela classificação — normalmente a sintética.

---

## 12. `04-LIVRO-CAIXA.ps1` + `_gerar/caixa.js`

### O que é "conta de caixa"

Vem de `ctcontacaixa`: 453 contas em 415 empresas. A extração faz `JOIN` de
`ctlancto` com ela, pelos dois lados:

```sql
JOIN bethadba.ctcontacaixa c ON c.codi_emp = l.codi_emp AND c.codi_cta = l.cdeb_lan
-- e, na outra metade da UNION:
JOIN bethadba.ctcontacaixa c ON c.codi_emp = l.codi_emp AND c.codi_cta = l.ccre_lan
```

**Conta de caixa no débito = entrada. No crédito = saída.** Dinheiro entrando
debita o caixa; saindo, credita.

No detalhe o mesmo teste é refeito no JS, para descobrir qual das duas contas
do lançamento é o caixa e qual é a contrapartida:

```js
const ehDeb = contasCaixa.has(cod + '|' + cdeb);
entrada: ehDeb ? v : 0,
saida:   ehDeb ? 0 : v
```

### Histórico

```sql
COALESCE(NULLIF(l.chis_lan,''), h.desc_his, '')
```

`chis_lan` é o texto livre digitado no lançamento e é o que tem conteúdo útil.
`cthispad.desc_his` é o histórico padrão, usado como reserva — nesta base
`codi_his` vem quase sempre nulo, então o `LEFT JOIN` raramente traz algo.

### Saldo acumulado

```js
if (o.cod !== atual) { atual = o.cod; saldo = 0; }
saldo += o.ent - o.sai;
o.saldoAcum = saldo;
```

Acumula por empresa, na ordem dos meses, zerando a cada empresa nova.
**É saldo do período extraído, não o saldo real da conta** — não tem saldo de
abertura anterior a `DataIni`.

### Escrita em streaming

Este gerador é o único que usa `ExcelJS.stream.xlsx.WorkbookWriter` em vez de
`Workbook`. São 881.745 linhas de detalhe; montar tudo em memória estoura o
heap do Node numa máquina com pouca RAM.

O TSV também é lido em streaming, linha a linha, com `readline` — nunca há o
arquivo de 77 MB inteiro na memória:

```js
const rl = readline.createInterface({
  input: fs.createReadStream(OUT + '/31_caixa_detalhe.tsv', { encoding: 'latin1' }),
  crlfDelay: Infinity
});
for await (const line of rl) { ... ws.addRow(...).commit(); }
```

Cada linha é escrita e liberada na hora com `.commit()`.

Há um corte em `LIMITE_EXCEL = 1040000`, abaixo do teto de 1.048.576 linhas da
planilha. Se bater, o log avisa `(TRUNCADO no limite do Excel)`. Hoje não bate.

A opção `-SemDetalhe` apaga o TSV de detalhe e o gerador pula a aba — o
`existe('31_caixa_detalhe')` cuida disso. Gera um arquivo de poucos KB em vez
de 58 MB.

---

## 13. `05-PRODUTOR-RURAL.ps1` + `_gerar/rural.js`

### Não extrai nada

É o único relatório que só recorta o que os outros já trouxeram. Se faltar
algum TSV, ele chama o script responsável:

```powershell
$dependencias = @{
    '02_faturamento' = '01-FISCAL.ps1'
    '05b_guia_inss'  = '02-FOLHA.ps1'
    '30_livro_caixa' = '04-LIVRO-CAIXA.ps1'
    ...
}
```

### Como os produtores são identificados

**O banco não tem marca de produtor rural.** Verificado:

| Campo que deveria servir | Situação |
|:--|:--|
| `geempre.CNPJ_PRODUTOR_RURAL_EMP` | vazio nas 655 empresas |
| `geempre.ucxa_emp` (usa livro caixa) | `0` em todas |
| `ctcontas.LCDPR_CTA` | `0` nas 308.853 contas |
| `fosefip.prod_rural_pf` / `prod_rural_pj` | nenhuma linha |

Ou seja, o Livro Caixa Digital do Produtor Rural não está configurado nesta
base. A identificação é por dedução, com dois critérios:

```js
const AGRO = /^0[123]/;
if (AGRO.test(e.cnae))            c.push('CNAE agropecuário');
if ((funrural[cod] || 0) > 0)     c.push('FUNRURAL apurado');
if (e.pf && AGRO.test(e.cnae))    c.push('Pessoa física');
```

1. **CNAE 2.0 nas divisões 01, 02 e 03** — agricultura e pecuária, produção
   florestal, pesca. 73 empresas.
2. **FUNRURAL apurado** em `efsdoimp`. 6 empresas.

União dos dois: 74 produtores. "Pessoa física" não é critério sozinho — é um
rótulo extra em quem já entrou pelo CNAE.

**A coluna "Critério de identificação" mostra qual regra pegou cada linha.
Confira essa lista** — ela vale o que valer o cadastro de CNAE no Domínio.

Os 74 respondem por R$ 933 mi dos R$ 1,54 bi de faturamento do escritório.

---

## 14. `06-PRODUTIVIDADE.ps1` + `_gerar/produtividade.js`

O relatório mais elaborado do conjunto. É o único que faz um cálculo não
trivial — o rateio por concorrência — e o único que tem teste automatizado.

### A tabela que responde "quanto tempo a pessoa fica na empresa"

`geloguser` grava **uma linha por sessão**:

| Coluna | Conteúdo |
|:--|:--|
| `usua_log` | nome do usuário |
| `codi_emp` | empresa aberta |
| `sist_log` | módulo (vira o departamento) |
| `data_log` + `tini_log` | quando entrou |
| `dfim_log` + `tfim_log` | quando saiu |

284.175 linhas na base, 226.771 no recorte de 2023 em diante, 61 usuários,
632 empresas.

### Passo 1 — duração bruta de cada sessão

As horas viram minutos absolutos numa linha do tempo única:

```js
const diaAbs = d => {
  const [y, m, dd] = String(d).split('-').map(Number);
  return Date.UTC(y, m - 1, dd) / 60000;
};
const ini = diaAbs(r.data_ini) + hms(r.hora_ini);
let fim = diaAbs(r.data_fim || r.data_ini) + hms(r.hora_fim);
if (fim < ini) fim = ini;
```

`hms` converte `hh:mm:ss` em minutos com fração. `Date.UTC` é usado de
propósito em vez de `new Date(texto)`: ele **não depende de fuso nem de
horário de verão**, e aqui só interessam diferenças. Usar a data local
introduziria um erro de 60 minutos nas viradas de horário de verão.

Somar `diaAbs` resolve a sessão que **vira a meia-noite** sem caso especial:
o fim simplesmente cai num dia posterior na mesma régua. São 104 sessões
assim na base.

**Qualidade do dado:** nenhuma duração negativa, nenhuma zerada, e só
**4 sessões acima de 10 horas** em 226.771 — todas atravessando a noite
(ex.: ARIEL, empresa 84, 28/06/2023 17:38 → 29/06 08:09, 14,5 h). É sistema
deixado aberto, não trabalho. Elas ficam na planilha, não são descartadas, e o
log avisa quantas são.

### Passo 2 — rateio por concorrência

**O problema.** Somar as durações conta o mesmo minuto mais de uma vez, porque
o usuário abre mais de uma sessão ao mesmo tempo. Caso real, ALINE em
11/03/2025 na empresa 345:

```
  módulo 1 (Contabilidade)   05:30:42 -> 06:15:20   45 min
  módulo 5 (Escrita Fiscal)  06:03:59 -> 06:15:25   11 min
```

Ela abriu o Fiscal sem fechar o Contábil. A soma dá 56 minutos numa janela de
relógio de 45. **Na base inteira: 41.081 h somadas para 37.675 h reais, 8,3%
a mais.** Afeta 55 dos 61 usuários.

**A solução.** Varredura (*sweep line*) na linha do tempo de cada usuário. Em
cada trecho entre dois eventos, a duração é dividida igualmente entre as
sessões que estavam abertas naquele instante:

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

Como funciona, linha a linha:

1. **Agrupa por usuário.** A concorrência só existe dentro do mesmo usuário —
   duas pessoas trabalhando ao mesmo tempo são duas horas de verdade.
2. **`pontos`** reúne todos os instantes em que algo muda: todo início e todo
   fim. `new Set` tira repetidos, `sort` põe em ordem. Entre dois pontos
   consecutivos o conjunto de sessões abertas **não muda** — é isso que
   permite dividir por um número fixo.
3. **`ordem`** é a lista de sessões ordenada por início, percorrida pelo
   ponteiro `p` que só anda para a frente. Isso evita varrer todas as sessões
   a cada trecho.
4. Em cada trecho: entra quem já começou (`while`), sai quem já terminou
   (`for ... delete`), e o que sobra em `ativos` divide a duração.
5. **`fatia = dur / ativos.size`.** Uma sessão sozinha leva o trecho inteiro;
   duas concorrentes levam metade cada.

No exemplo da ALINE: o trecho 06:03:59→06:15:20 tem duas sessões abertas, então
cada uma leva 5,7 min em vez de 11,4. O trecho 05:30:42→06:03:59 tem só a
Contabilidade, que leva os 33,3 min inteiros. Soma final: 45 min, o relógio.

**Custo.** O(n log n) por usuário. 226 mil sessões processam em menos de um
segundo.

### O teste do rateio

`_gerar/teste-rateio.js` verifica o **invariante** do algoritmo: para cada
usuário, a soma das fatias tem que ser exatamente igual à **união** dos
intervalos dele. A união é calculada ali por outro caminho — merge de
intervalos ordenados — justamente para não repetir o mesmo raciocínio e o
mesmo erro.

```bash
cd "C:\Users\LENOVO\Documents\BECKUP DOMINIO\SCRIPTS\_gerar" && node teste-rateio.js x C:/dbtmp/out
```

Resultado atual: bate nos 61 usuários, maior erro 2×10⁻¹⁰ minutos (ruído de
ponto flutuante), nenhuma fatia negativa ou maior que a própria sessão.

Rode isso depois de mexer no rateio. É barato e pega erro silencioso.

### Como as duas colunas de hora aparecem nas planilhas

| Onde | Coluna | O que é |
|:--|:--|:--|
| Por empresa | Horas atribuídas | soma das fatias — **é a coluna de referência** |
| Por empresa | Horas brutas | soma das durações, para comparar |
| Por usuário | Horas reais | tempo de relógio do usuário |
| Por usuário | Horas brutas, Horas sobrepostas, % sobreposto | o tamanho da distorção |
| Sessões | Minutos (bruto) / Minutos atribuídos | a fatia daquela sessão |
| Sessões | Sessões simultâneas | `bruto ÷ atribuído` — 1,0 é sessão sozinha; 2,0 é ter tido, em média, outra aberta o tempo todo |

O total de "Horas atribuídas" de todas as empresas fecha com o total de
"Horas reais" de todos os usuários. As brutas não fecham com nada — estão lá
só para você ver a diferença.

### Departamento

O Domínio **não tem cadastro de departamento**. O que ele grava é o módulo em
que o usuário estava. O mapa está em `_gerar/lib.js`, num único lugar:

```js
const DEPARTAMENTO = {
  '1': 'Contábil',        // Contabilidade
  '4': 'Contábil',        // Patrimônio
  '5': 'Fiscal',          // Escrita Fiscal
  '6': 'Fiscal',          // Lalur
  '12': 'Pessoal',        // Folha
  '13': 'Pessoal',        // Ponto Eletrônico
  '3': 'Honorários',
  '15': 'Societário',     // Registro
  ...                     // o resto cai em Administrativo
};
```

As escolhas discutíveis: **Lalur** é apuração de IRPJ/CSLL, então conta como
Fiscal; **Patrimônio** é imobilizado, conta como Contábil; **Ponto Eletrônico**
conta como Pessoal. Se a divisão do escritório for outra, **edite esse mapa** e
as quatro abas de departamento acompanham.

Os lançamentos são roteados ao departamento que os produz:

```js
const lancDoDept = (l, d) => {
  if (d === 'Fiscal')   return l.fiscalSai + l.fiscalEnt + l.apuracoes;
  if (d === 'Contábil') return l.contabil;
  if (d === 'Pessoal')  return l.folha;
  return 0;
};
```

Resultado no período 2024-01 a 2026-12:

| Departamento | Horas | % do tempo | Usuários | Empresas | Lançamentos | Lanç./hora |
|:--|--:|--:|--:|--:|--:|--:|
| Fiscal | 20.166,1 | 53,5% | 53 | 581 | 2.737.035 | 135,7 |
| Pessoal | 10.069,9 | 26,7% | 39 | 577 | 24.681 | 2,5 |
| Contábil | 6.956,7 | 18,5% | 36 | 468 | 2.785.959 | 400,5 |
| Societário | 177,5 | 0,5% | 17 | 288 | — | — |
| Administrativo | 243,5 | 0,6% | 42 | 22 | — | — |
| Honorários | 60,9 | 0,2% | 14 | 23 | — | — |

**Não compare "lanç./hora" entre departamentos.** As unidades são diferentes:
no Fiscal é documento, no Contábil é partida, no Pessoal é empregado
processado. Os 2,5 do Pessoal não significam lentidão — significam que a
unidade dali é outra. A comparação que vale é **a mesma coluna entre empresas
do mesmo departamento**.

### As dez abas

| Aba | Responde |
|:--|:--|
| Por departamento | quanto tempo e quanto volume em cada departamento |
| Departamento por mês | como isso varia ao longo do ano |
| Departamento por empresa | qual departamento gasta tempo em qual cliente |
| Departamento por usuário | quem é do Fiscal, quem é do Pessoal, quem divide |
| Resumo por empresa | custo de tempo de cada cliente |
| Resumo por usuário | carga de cada pessoa, e quanto do tempo dela é sobreposto |
| Tempo por empresa e usuário | quem atende quem |
| Mensal tempo x lançamentos | tempo ao lado do volume, mês a mês |
| Sessões (detalhe) | as 226.771 sessões, uma por linha |
| Lanç. contábeis por operador | quem lançou o quê |

### Outras fontes de carimbo de tempo

Além do log de sessão, o script extrai quem processou o quê:

| Arquivo | Origem | Conteúdo |
|:--|:--|:--|
| `11_lanc_contabeis` | `ctlancto.codi_usu` | quem lançou — é o **nome**, não código |
| `12_apuracoes` | `efsdoimp.USUARIO` + `DATA_HORA` | quem apurou imposto e quando |
| `13_folha_calculos` | `fobasesserv.USUARIO_CALCULO` + `DATA_HORA_CALCULO` | quem calculou folha |

## 15. `EXTRAIR-TUDO.ps1`

Roda os seis na ordem. **A ordem importa**: o 05 depende dos TSVs do 01, 02 e
04.

Antes de começar, checa se alguma das planilhas que ele vai gravar está aberta
no Excel — e só as que vai gravar, respeitando o `-Pular`:

```powershell
$abertos = $relatorios |
    Where-Object { $Pular -notcontains $_.id } |
    ForEach-Object { Join-Path $DOM.PastaSaida $_.saida } |
    Where-Object { Test-ArquivoEmUso $_ }
```

Cada relatório roda dentro de `try/catch`: um que falhe não derruba os outros,
e a tabela do fim mostra `ok` ou `FALHOU` com o tempo de cada um.

Parâmetros: `-Pular 04,06`, `-SemDetalheCaixa`, `-PararAoFim`.

Tempo medido na máquina atual: **3,4 minutos** para os seis.

| Relatório | Tempo |
|:--|--:|
| 01-FISCAL | 25 s |
| 02-FOLHA | 7 s |
| 03-CONTABIL | 13 s |
| 04-LIVRO-CAIXA | 103 s |
| 05-PRODUTOR-RURAL | 3 s |
| 06-PRODUTIVIDADE | 52 s |

---

## 16. Dicionário de fórmulas

Referência rápida: coluna da planilha → como é calculada → de onde vem.

### Fiscal

| Coluna | Fórmula | Origem |
|:--|:--|:--|
| Faturamento | `SUM(vcon_sai)`, situação 0, não cancelada | `efsaidas` |
| Notas de saída | `COUNT(*)` com os mesmos filtros | `efsaidas` |
| Compras | `SUM(vcon_ent)`, situação 0 | `efentradas` |
| `<SIGLA>` | `SUM(sdev_sim)` do tributo | `efsdoimp` + `geimposto` |
| Total impostos apurados | soma de todas as colunas de sigla | derivado |
| Base de cálculo | `SUM(bcal_sim)` | `efsdoimp` |
| Valor pago | `SUM(vlrpago_sim)` | `efsdoimp` |

### Folha

| Coluna | Fórmula | Origem |
|:--|:--|:--|
| Proventos | `SUM(proventos)`, `dias_servico > 0` | `fobasesserv` |
| INSS patronal | `SUM(empresa)` | `foguiainss` |
| RAT/SAT | `SUM(acid_trab)` | `foguiainss` |
| Terceiros | `SUM(terceiros)` | `foguiainss` |
| PIS sobre folha | `SUM(valor)` | `foguiapis` |
| **Total encargo patronal** | patronal + RAT + terceiros + PIS | derivado |
| INSS segurados | `SUM(segurados)` | `foguiainss` |
| Funcionários ativos | admissão ≤ fim do mês e sem demissão antes do início | `foempregados` + `forescisoes` |
| Admissões | contagem por mês de `admissao` | `foempregados` |
| Demissões | contagem por mês da **menor** `demissao` | `forescisoes` |
| Saldo do mês | admissões − demissões | derivado |

### Contábil

| Coluna | Fórmula | Origem |
|:--|:--|:--|
| Débito | `SUM(vlor_lan)` onde a conta é `cdeb_lan` | `ctlancto` |
| Crédito | `SUM(vlor_lan)` onde a conta é `ccre_lan` | `ctlancto` |
| Saldo do ano | devedora: D−C; credora: C−D | derivado |
| Saldo acumulado | soma dos saldos do ano, por conta, desde 1995 | derivado |
| Ativo | Σ saldo acumulado do grupo 1 | derivado |
| Passivo + PL | Σ saldo acumulado do grupo 2 | derivado |
| Receitas do ano | Σ saldo do ano do grupo 4 | derivado |
| Despesas do ano | Σ saldo do ano do grupo 3 | derivado |
| Resultado | receitas − despesas | derivado |

### Livro caixa

| Coluna | Fórmula |
|:--|:--|
| Entradas | `SUM(vlor_lan)` com conta de caixa no **débito** |
| Saídas | `SUM(vlor_lan)` com conta de caixa no **crédito** |
| Saldo do mês | entradas − saídas |
| Saldo acumulado | soma corrida por empresa, do primeiro mês extraído |

### Produtividade

| Coluna | Fórmula |
|:--|:--|
| Minutos (bruto) da sessão | `tfim − tini`, + 1440 por dia virado, piso em 0 |
| Minutos atribuídos | soma das fatias: em cada trecho, `duração ÷ sessões abertas` |
| Sessões simultâneas | minutos bruto ÷ minutos atribuídos |
| Horas atribuídas | minutos atribuídos ÷ 60 — **é a coluna de referência** |
| Horas brutas | minutos bruto ÷ 60 |
| Horas sobrepostas | horas brutas − horas atribuídas |
| % sobreposto | horas sobrepostas ÷ horas brutas |
| Dias com acesso | contagem de datas distintas |
| Média h/dia | minutos atribuídos ÷ dias distintos ÷ 60 |
| Média min/sessão | minutos bruto ÷ nº de sessões |
| Total lançamentos | notas saída + notas entrada + contábeis + folha |
| Lanç. por hora | total lançamentos ÷ horas atribuídas |

### Departamento

| Coluna | Fórmula |
|:--|:--|
| Departamento | `sist_log` traduzido pelo mapa `DEPARTAMENTO` em `lib.js` |
| Horas | soma dos minutos atribuídos das sessões do departamento ÷ 60 |
| % do tempo total | horas do departamento ÷ horas de todos |
| Lançamentos (Fiscal) | notas de saída + notas de entrada + apurações |
| Lançamentos (Contábil) | lançamentos contábeis |
| Lançamentos (Pessoal) | empregados processados na folha |
| Lanç. por hora | lançamentos do departamento ÷ horas do departamento |
| % do tempo do usuário | horas do usuário naquele departamento ÷ total dele |

**As unidades de "lançamentos" diferem entre departamentos** — documento no
Fiscal, partida no Contábil, empregado no Pessoal. Compare a mesma coluna entre
empresas do mesmo departamento, nunca entre departamentos.

---

## 17. As armadilhas, reunidas

Para quem for mexer no código.

**No banco:**

1. `vcon_sai` é faturamento; `vprod_sai` é valor dos produtos.
2. `sdev_sim` é imposto; `vime_sim` é base de ICMS.
3. `fobasesserv` duplica cada empregado por competência — filtre `dias_servico > 0`.
4. As colunas de encargo de `fobasesserv` são nulas; use `foguiainss`.
5. `SUM(a + b)` vira NULL com uma parcela nula — sempre `ISNULL`.
6. Datas têm lixo (1979, 8320) — limite o intervalo em qualquer `MIN`/`MAX`.
7. `geusuarios` está vazia; os usuários estão em `GESECUSUARIOS`.
8. `ctlancto.codi_usu` é nome, não código.
9. `cnae_emp` não presta; use `i_cnae20`.
10. `efpagimp.vpag_pim` existe mas a tabela está quase vazia no período.
11. Complemento de rescisão duplica o empregado — pegue a menor data.
12. O catálogo tem 4.800 tabelas base; conte por
    `SYS.SYSTAB ... table_type_str='BASE'`, não confie em listagem parcial.

**Na extração:**

13. `UNLOAD` exige `-gl all` no servidor.
14. Caminho do `UNLOAD`: barra normal, sem espaço.
15. O separador do arquivo é `\`+`t` literal; o encoding é cp1252.
16. Servidor sobe por `.bat` — `Start-Process` quebra caminho com espaço.
17. Cache maior que a RAM livre deixa o servidor **mais lento**, não mais rápido.

**No PowerShell:**

18. `.ps1` precisa de UTF-8 **com BOM**, ou o 5.1 lê como ANSI e o parser quebra.
19. Função que dá `return $true` imprime `True` solto no log.
20. Config de parent não sobrevive ao filho — use variável de ambiente.

**No Excel:**

21. Arquivo aberto trava a gravação — `Test-ArquivoEmUso` checa antes.
22. Acima de ~500 mil linhas, gere em streaming ou o Node estoura.
23. O leitor do exceljs não relê um `.xlsx` de 58 MB (limite de string do V8);
    a gravação funciona, a releitura não. Para validar, inspecione o zip.

---

## 18. Como validar uma extração

Empresa **345** (L B GOMES RAMOS), competência **10/2025**:

| Valor | Esperado | Onde |
|:--|--:|:--|
| Faturamento | 362.025,28 | 01-FISCAL, Mensal por empresa |
| Compras | 699.955,52 | idem |
| Impostos (IRRF-APF) | 55,84 | idem |
| Folha (proventos) | 4.818,15 | 02-FOLHA, Encargos mensais |

Direto dos TSVs, sem abrir planilha:

```bash
cd "C:\Users\LENOVO\Documents\BECKUP DOMINIO\SCRIPTS\_gerar" && node -e "const{readObj,num,comp}=require('./lib');const f=readObj('02_faturamento').find(r=>r.codi_emp==='345'&&r.ano==='2025'&&r.mes==='10');const i=readObj('04_impostos').filter(r=>r.codi_emp==='345'&&comp(r.competencia)==='2025-10');console.log('fat',num(f.faturamento).toFixed(2),'| imp',i.reduce((a,r)=>a+num(r.saldo_devedor),0).toFixed(2))" x C:/dbtmp/out
```

Totais esperados do conjunto, no período 2024-01 a 2026-12:

| Relatório | Números |
|:--|:--|
| 01-FISCAL | 342 empresas, R$ 1.543.494.361,53, R$ 19.888.636,04 de imposto, 20 tributos |
| 02-FOLHA | 230 empresas, 3.021 empregados, R$ 2.311.434,61 de encargo patronal |
| 03-CONTABIL | 343 empresas, 28.778 linhas, 10.023.332 lançamentos |
| 04-LIVRO-CAIXA | 188 empresas, 881.745 lançamentos |
| 05-RURAL | 74 produtores, R$ 933.321.159,26, FUNRURAL R$ 741.545,55 |
| 06-PRODUTIVIDADE | 226.771 sessões, 632 empresas, 61 usuários, 37.674,54 h reais (41.081,37 brutas, 8,3% sobrepostas) |

E o teste do rateio de horas:

```bash
cd "C:\Users\LENOVO\Documents\BECKUP DOMINIO\SCRIPTS\_gerar" && node teste-rateio.js x C:/dbtmp/out
```

Tem que sair `OK` nas três verificações.

---

# Apêndice A — As linguagens e a sintaxe

Esta parte é para quem vai **mexer** no código. Explica as três linguagens
envolvidas, por que cada uma está onde está, e cada construção de sintaxe que
aparece no kit.

## A.1 Quem faz o quê

| Linguagem | Onde | Por que essa e não outra |
|:--|:--|:--|
| **SQL** (dialeto SQL Anywhere) | dentro dos `.ps1` | é a única forma de falar com o banco; e o `UNLOAD` só existe nele |
| **PowerShell 5.1** | `.ps1` | já vem no Windows, fala ODBC nativamente via .NET, e liga/desliga processo |
| **JavaScript** (Node 22) | `.js` | por causa do `exceljs` — é a biblioteca de `.xlsx` que roda aqui sem Office e sem Python |

Não há Python real nesta máquina (só o atalho da Microsoft Store), e o Excel
não está instalado. Node + exceljs foi a única rota para gerar `.xlsx` de
verdade, com formatação e mais de um milhão de linhas.

**A fronteira é o arquivo `.tsv`.** O PowerShell nunca calcula; o JavaScript
nunca acessa o banco. Quem quiser trocar uma das pontas troca só ela.

---

## A.2 SQL — o dialeto do SQL Anywhere

### `UNLOAD`

```sql
UNLOAD SELECT codi_emp, nome_emp FROM bethadba.geempre
TO 'C:/dbtmp/out/01_empresas.tsv'
DELIMITED BY '\t' QUOTES OFF ESCAPES OFF
```

Não é SQL padrão — é extensão do SQL Anywhere. Manda o **servidor** gravar o
resultado num arquivo, em vez de devolver linhas pelo driver.

- `TO` — caminho **no servidor**. Como o servidor é local, é a mesma máquina.
- `DELIMITED BY` — separador. Aqui vira os dois caracteres `\` e `t` literais.
- `QUOTES OFF` — não põe aspas em volta do texto.
- `ESCAPES OFF` — não escapa nada; o texto sai cru.

### `ISNULL(x, 0)`

```sql
SUM(ISNULL(val_inss_empresa, 0) + ISNULL(inss_empresa_fer, 0))
```

Troca nulo por um valor. **É a função mais importante do kit inteiro.**

Em SQL, qualquer operação aritmética com `NULL` resulta `NULL`. Então
`SUM(a + b)` percorre as linhas, encontra uma em que `b` é nulo, calcula
`a + NULL = NULL`, e o `SUM` de um conjunto que contém `NULL`... na verdade
ignora o nulo, mas a **linha inteira** já virou nulo antes de chegar ao `SUM`.
Na prática: uma única linha com nulo pode zerar ou distorcer a coluna.

Foi exatamente o que aconteceu na primeira versão do relatório de folha, que
saiu com todas as colunas de encargo em branco.

`ISNULL` é do SQL Anywhere e do SQL Server. O equivalente padrão é
`COALESCE(x, 0)`, que também funciona aqui.

### `COALESCE` e `NULLIF`

```sql
COALESCE(NULLIF(l.chis_lan, ''), h.desc_his, '')
```

`NULLIF(a, b)` devolve `NULL` se `a = b`, senão devolve `a`. `COALESCE` devolve
o primeiro não nulo da lista.

Lido junto: *"se `chis_lan` for string vazia trate como nulo; então pegue o
primeiro que não for nulo entre `chis_lan`, `desc_his` e string vazia"*. É
como o histórico do lançamento cai para o padrão quando o texto livre está em
branco.

### `GROUP BY` com função na coluna

```sql
SELECT codi_emp, YEAR(dsai_sai), MONTH(dsai_sai), SUM(vcon_sai), COUNT(*)
FROM bethadba.efsaidas
GROUP BY codi_emp, YEAR(dsai_sai), MONTH(dsai_sai)
```

Tudo que está no `SELECT` e **não** é função de agregação precisa estar no
`GROUP BY`. `YEAR()` e `MONTH()` extraem partes da data; agrupar por elas dá a
série mensal sem precisar de tabela de calendário.

`COUNT(*)` conta linhas do grupo; `SUM(x)` soma a coluna dentro do grupo.

### `UNION ALL` com subconsulta

```sql
SELECT codi_emp, ano, conta, SUM(debito), SUM(credito) FROM (
  SELECT codi_emp, YEAR(data_lan) AS ano, cdeb_lan AS conta,
         SUM(vlor_lan) AS debito, 0 AS credito
  FROM bethadba.ctlancto WHERE cdeb_lan > 0
  GROUP BY codi_emp, YEAR(data_lan), cdeb_lan
  UNION ALL
  SELECT codi_emp, YEAR(data_lan), ccre_lan,
         0, SUM(vlor_lan)
  FROM bethadba.ctlancto WHERE ccre_lan > 0
  GROUP BY codi_emp, YEAR(data_lan), ccre_lan
) AS t
GROUP BY codi_emp, ano, conta
```

`UNION ALL` empilha dois resultados. **`ALL` importa**: sem ele o SQL faria
`DISTINCT` e apagaria linhas repetidas legítimas — duas contas com o mesmo
valor no mesmo ano viram uma só.

As duas metades precisam ter o **mesmo número de colunas, na mesma ordem, com
tipos compatíveis**. Daí os `0 AS credito` e `0` como espaçadores: a metade de
débito não tem crédito, mas precisa da coluna para alinhar.

A subconsulta **precisa de um apelido** (`AS t`) no SQL Anywhere.

O `GROUP BY` de fora reagrupa: uma conta que apareceu como débito e como
crédito no mesmo ano vira uma linha só com as duas colunas preenchidas.

### `EXISTS`

```sql
WHERE EXISTS (SELECT 1 FROM bethadba.ctcontacaixa c
              WHERE c.codi_emp = l.codi_emp AND c.codi_cta = l.cdeb_lan)
```

Testa *"existe pelo menos uma linha que satisfaz isso?"*. O `SELECT 1` é
convenção: não interessa o valor, só se achou algo.

Usado em vez de `JOIN` no livro caixa porque o teste é feito em **duas colunas
diferentes** (`cdeb_lan` ou `ccre_lan`). Com `JOIN` um lançamento que tocasse
caixa nos dois lados apareceria duplicado.

### `JOIN` com chave composta

```sql
JOIN bethadba.geimposto g
  ON g.codi_emp = d.codi_emp AND g.codi_imp = d.codi_imp
```

**Toda tabela do Domínio é particionada por empresa.** `codi_imp = 3` não é um
imposto — é o imposto 3 *da empresa tal*. Outra empresa tem o seu próprio 3,
que pode ser outro tributo.

Por isso **todo join carrega `codi_emp`**. Esquecer isso cruza dado de uma
empresa com cadastro de outra, e o erro é silencioso: não dá erro, dá número
errado.

O mesmo vale para `ctcontas` (`codi_emp` + `codi_cta`), `foempregados`
(`codi_emp` + `i_empregados`) e `efprodutos` (`codi_emp` + `codi_pdi`).

### `CAST(x AS DATE)`

```sql
SELECT codi_emp, USUARIO, CAST(DATA_HORA AS DATE), COUNT(*)
```

`DATA_HORA` tem hora. Sem o `CAST`, agrupar por ele daria um grupo por
segundo. O `CAST` corta a hora e deixa o dia.

### Tabelas de catálogo

| Objeto | Para que |
|:--|:--|
| `SYS.SYSCOLUMNS` | colunas: `creator`, `tname`, `cname`, `coltype`, `length` |
| `SYS.SYSTAB` + `SYS.SYSUSER` | tabelas: filtrar `table_type_str='BASE'` |
| `DB_PROPERTY('CharSet')` | propriedades do banco |
| `PROPERTY('ProductVersion')` | versão do servidor |

Cuidado: a view `SYS.SYSCOLUMNS` (compatibilidade) tem `tname`/`cname`;
`SYS.SYSTABCOL` (moderna) tem `table_id`/`column_name`. Misturar as duas dá
`Column not found`.

---

## A.3 PowerShell — a sintaxe usada aqui

### Cabeçalho de script

```powershell
[CmdletBinding()]
param(
    [ValidateSet('status','iniciar','parar','identificar','reiniciar')]
    [string]$Acao = 'status',
    [switch]$Forcar,
    [string[]]$Pular = @()
)
```

- `param(...)` — tem que ser a **primeira instrução executável** do arquivo
  (comentários e o bloco de ajuda `<# ... #>` podem vir antes).
- `[CmdletBinding()]` — liga recursos de cmdlet: `-Verbose`, `-ErrorAction`.
- `[ValidateSet(...)]` — só aceita esses valores; erra na hora se digitar outro.
- `[switch]` — parâmetro booleano de bandeira: `-Forcar` liga, ausente desliga.
- `[string[]]` — array. `-Pular 04,06` chega como dois elementos.
- `= 'status'` — valor padrão.

### Hashtable

```powershell
$Global:DOM = @{ Banco = 'C:\...'; Porta = 2638 }
$DOM.Porta            # lê
$DOM.Porta = 2639     # escreve
```

`@{ }` é hashtable (dicionário); `@( )` é array. `$Global:` põe no escopo
global, visível em qualquer função ou script do mesmo processo.

### Dot-sourcing

```powershell
. "$PSScriptRoot\_lib\dominio.ps1"
```

O ponto isolado no começo é o **operador de dot-source**: roda o arquivo
**no escopo atual**, então as funções e variáveis dele ficam disponíveis aqui.

Sem o ponto (`& arquivo.ps1`), o script roda num escopo filho e tudo que ele
definiu some quando termina.

`$PSScriptRoot` é a pasta do arquivo que está executando — o que faz os
caminhos funcionarem de qualquer diretório de onde você chame.

### Here-string

```powershell
-Sql @"
SELECT codi_emp, nome_emp
FROM bethadba.geempre
WHERE dsai_sai >= '$($DOM.DataIni)'
"@
```

`@"` abre e `"@` fecha um bloco de texto multilinha.

**Duas regras que quebram tudo se violadas:**

1. O `"@` de fechamento tem que estar **na coluna 0**, sozinho na linha. Indentar
   é erro de sintaxe.
2. `@"..."@` interpola variáveis. `@'...'@` (aspas simples) **não** interpola —
   é o que se usa quando o texto tem `$` literal.

### Interpolação e subexpressão

```powershell
"porta $($DOM.Porta) do servidor $servidor"
```

Dentro de aspas duplas, `$variavel` é substituída. Para **expressão** — acesso
a propriedade, índice, chamada — precisa de `$( )`. Escrever `"$DOM.Porta"`
imprimiria o hashtable inteiro seguido do texto `.Porta`.

### Operador de formatação `-f`

```powershell
Write-Host ("   [ok] {0,-24} {1,9:N0} KB {2,6:N1}s" -f $Nome, $kb, $seg)
```

`{índice,largura:formato}`:

- `{0,-24}` — argumento 0, 24 caracteres, alinhado à **esquerda** (sinal de menos)
- `{1,9:N0}` — argumento 1, 9 caracteres à direita, número com separador de
  milhar e 0 decimais
- `{2,6:N1}` — 1 casa decimal

**Não existe `{0,>10}`** — alinhamento à direita é o padrão, basta o número
positivo. Escrever `>` dá erro de formatação.

### .NET direto

```powershell
$cn = New-Object System.Data.Odbc.OdbcConnection($cs)
$cn.Open()
$cmd = $cn.CreateCommand()
$cmd.CommandTimeout = 7200
$cmd.ExecuteNonQuery() | Out-Null
$cn.Close()
```

PowerShell é uma casca em cima do .NET: dá para instanciar qualquer classe.
`System.Data.Odbc` é o provedor ODBC — é assim que o kit fala com o SQL
Anywhere, sem biblioteca externa.

`ExecuteNonQuery()` para comando sem resultado (o `UNLOAD`);
`ExecuteReader()` para `SELECT` que devolve linhas.

### `| Out-Null`

```powershell
New-Item -ItemType Directory -Force $pasta | Out-Null
```

Em PowerShell, **tudo que uma expressão produz e não é consumido vai para a
saída**. `New-Item` devolve o objeto do diretório criado, que apareceria no
log. `| Out-Null` descarta.

É a mesma razão pela qual `Export-DomTabela` não faz `return $true`: o valor
não capturado seria impresso como `True` no meio do relatório.

### `try / catch / finally`

```powershell
try   { $cmd.ExecuteNonQuery() | Out-Null; ... }
catch { Write-Host "FALHA: $($_.Exception.Message)" -ForegroundColor Red }
finally { $cn.Close() }
```

`$_` dentro do `catch` é o erro. `finally` roda sempre, dando erro ou não — é
onde a conexão é fechada, para não vazar conexão quando a consulta falha.

### Operador de chamada `&`

```powershell
& "$PSScriptRoot\00-CADASTROS.ps1"
& node "--max-old-space-size=2400" $js $destino
```

`&` executa o que vem depois como comando. É necessário quando o nome do
programa está numa variável ou tem espaço no caminho — sem ele o PowerShell
trataria a string como texto a ser impresso.

### Pipeline e filtros

```powershell
$abertos = $relatorios |
    Where-Object { $Pular -notcontains $_.id } |
    ForEach-Object { Join-Path $DOM.PastaSaida $_.saida } |
    Where-Object { Test-ArquivoEmUso $_ }
```

O `|` passa **objetos**, não texto. `$_` é o item atual.
`Where-Object` filtra, `ForEach-Object` transforma.

Comparações são palavras, não símbolos: `-eq -ne -lt -gt -le -ge`,
`-contains -notcontains -match -like -not`. Escrever `==` ou `!=` é erro.

### Codificação do arquivo

**Todo `.ps1` do kit tem que estar em UTF-8 com BOM.** O Windows PowerShell 5.1
lê arquivo sem BOM como ANSI (cp1252). Aí `ç` e `ã` viram dois caracteres, uma
aspa pode ficar desemparelhada e o parser quebra com "Token inesperado" numa
linha que parece perfeita.

Sintoma: erro de sintaxe apontando para uma linha correta, e acentos aparecendo
como `Ã£` no terminal. Correção:

```powershell
$bom = New-Object System.Text.UTF8Encoding($true)
Get-ChildItem . -Recurse -Filter *.ps1 | ForEach-Object {
  $t = [IO.File]::ReadAllText($_.FullName, [Text.UTF8Encoding]::new($false))
  [IO.File]::WriteAllText($_.FullName, $t, $bom)
}
```

---

## A.4 JavaScript — a sintaxe usada aqui

### Módulos CommonJS

```js
const ExcelJS = require('exceljs');   // importa
const L = require('./lib');
module.exports = { read, num, fazAba };   // exporta
```

Node com arquivos `.js` usa CommonJS: `require` para importar,
`module.exports` para exportar. (O sistema moderno é `import`/`export` em
`.mjs`; aqui não é usado.)

### Desestruturação

```js
const { readObj, num, int, comp, MOEDA, fazAba } = L;
const [y, m, dd] = String(d).split('-').map(Number);
const [cod, mes] = k.split('|');
```

Tira campos de um objeto, ou posições de um array, direto em variáveis. A
primeira linha é equivalente a `const readObj = L.readObj; const num = L.num; ...`.

### Arrow function

```js
const h = m => Math.round(m / 60 * 100) / 100;
const num = v => { ... return n; };
```

`x => expressão` é função de um parâmetro que retorna a expressão. Com chaves
precisa de `return` explícito.

`Math.round(x * 100) / 100` é o jeito padrão de arredondar para 2 casas.

### O idioma "pega ou cria"

```js
const o = mapa[chave] || (mapa[chave] = { cod, mes, total: 0 });
o.total += valor;
```

Aparece em **todos** os geradores. Lê-se: *"se já existe a entrada, use; senão,
crie, guarde no mapa e use"*. O `||` só avalia o lado direito quando o esquerdo
é falso, e atribuição é uma expressão que devolve o valor atribuído.

É o acumulador de todos os agrupamentos do kit.

### Chave composta

```js
const k = cod + '|' + mes;
lanc[k] = { ... };
```

JavaScript não tem dicionário de chave múltipla, então as partes são
concatenadas com um separador. `'345' + '|' + '2025-10'` vira `'345|2025-10'`.

**O risco:** se alguma parte contiver o separador, a chave colide. Aqui é
seguro porque as partes são códigos numéricos e datas `AAAA-MM`. Para separar
de volta: `k.split('|')`.

### `Set` e `Map`

```js
e.usus = new Set();
e.usus.add(s.usuario);
e.usus.size          // quantos DIFERENTES
```

`Set` guarda valores únicos. É como "usuários distintos" e "dias com acesso"
são contados: adiciona todo mundo e lê o `.size` no fim, sem precisar ordenar
nem comparar.

### `Object.values` / `Object.entries`

```js
for (const o of Object.values(mov)) { ... }
for (const [sigla, total] of Object.entries(siglasTotais)) { ... }
```

`values` dá a lista de valores; `entries` dá pares `[chave, valor]`, que a
desestruturação abre em duas variáveis.

### `map`, `filter`, `reduce`, `sort`

```js
const SIGLAS = Object.entries(siglasTotais)
  .filter(([, v]) => v > 0)          // só quem tem valor
  .sort((a, b) => b[1] - a[1])       // do maior para o menor
  .map(([s]) => s);                  // fica só a sigla
```

- `filter` — mantém quem passa no teste
- `map` — transforma cada item
- `sort(comparador)` — o comparador devolve negativo, zero ou positivo.
  `a - b` é crescente; `b - a` é decrescente. Para texto,
  `a.localeCompare(b)`.
- `reduce((acumulado, item) => ..., inicial)` — dobra a lista num valor só:

```js
impTot: SIGLAS.reduce((a, s) => a + (t.imp[s] || 0), 0)
```

Note `[, v]` no `filter`: desestruturação pulando a primeira posição.

### Ordenação em vários critérios

```js
.sort((a, b) => a.cod - b.cod || a.mes.localeCompare(b.mes))
```

O `||` encadeia: se o primeiro critério empata (devolve 0, que é falso), passa
para o segundo. É como se ordena por empresa e depois por competência.

### Espalhamento e `flatMap`

```js
const cols = [ {...}, ...SIGLAS.map(s => ({ header: s })) ];
const pontos = [...new Set(idxs.flatMap(i => [s[i].ini, s[i].fim]))];
```

`...` dentro de array insere os elementos de outro array — é assim que as
colunas de tributo, que variam em número, entram na definição fixa da planilha.

`flatMap` mapeia e achata um nível: cada sessão vira dois pontos, e o resultado
é uma lista única de pontos, não uma lista de pares.

`[...conjunto]` transforma um `Set` de volta em array, para poder ordenar.

Cuidado com `s => ({ ... })`: os parênteses em volta do objeto são
obrigatórios, senão a chave `{` é lida como início de bloco de código.

### `async` / `await` e leitura em streaming

```js
(async () => {
  const rl = readline.createInterface({
    input: fs.createReadStream(arquivo, { encoding: 'latin1' }),
    crlfDelay: Infinity
  });
  for await (const line of rl) {
    ws.addRow({...}).commit();
  }
  await wb.commit();
})();
```

Só o `caixa.js` usa. `for await` percorre um fluxo assíncrono linha a linha —
o arquivo de 77 MB nunca fica inteiro na memória.

O `(async () => { ... })()` é uma função anônima assíncrona chamada na hora,
necessária porque `await` só existe dentro de `async`.

`.commit()` escreve a linha e libera a memória dela. Sem isso o `exceljs`
acumularia 880 mil linhas antes de gravar.

### Argumentos de linha de comando

```js
const DEST = process.argv[2];   // destino do .xlsx
const OUT  = process.argv[3];   // pasta dos .tsv
```

`process.argv[0]` é o executável do node, `[1]` é o script, e os do usuário
começam em `[2]`. É assim que o `Invoke-DomGerador` passa os caminhos.

### Formatação de número

```js
valor.toLocaleString('pt-BR', { minimumFractionDigits: 2 })   // 1.543.494.361,53
```

Só para o log do terminal. Dentro da planilha o número entra **cru** e quem
formata é o `numFmt` do exceljs (`'#,##0.00'`), para continuar somável no Excel.

---

## A.5 Os três padrões de cálculo do kit

Tirando o SQL, praticamente todo cálculo do kit cai num destes três moldes.
Entendendo os três, entende-se qualquer gerador.

### Padrão 1 — agregação por chave composta

O mais comum. Percorre linhas, monta uma chave, acumula.

```js
const mov = {};
const M = (cod, mes) => {
  const k = cod + '|' + mes;
  if (!mov[k]) mov[k] = { cod, mes, fat: 0, notas: 0 };
  return mov[k];
};
for (const r of readObj('02_faturamento')) {
  const o = M(r.codi_emp, mm(r.ano, r.mes));
  o.fat += num(r.faturamento);
  o.notas += int(r.notas);
}
```

Onde aparece: faturamento, compras, impostos, encargos, headcount, departamento
— quase tudo.

Vantagem sobre fazer no SQL: o mesmo acumulador recebe dados de **arquivos
diferentes**. `M()` é chamado pelo faturamento, pelas compras e pelos impostos,
e no fim cada linha tem as três coisas juntas.

### Padrão 2 — acumulação em sequência ordenada

Para saldo, que depende do que veio antes.

```js
lista.sort((a, b) => a.cod - b.cod || a.ano - b.ano);
let chave = '', acumulado = 0;
for (const r of lista) {
  if (r.cod !== chave) { chave = r.cod; acumulado = 0; }   // zera na troca
  acumulado += r.saldoAno;
  r.saldoAcum = acumulado;
}
```

**Ordenar primeiro é obrigatório** e a **reinicialização na troca de chave** é o
ponto onde se erra: sem ela, o saldo de uma empresa vaza para a seguinte.

Onde aparece: saldo acumulado do balancete (por empresa+conta) e saldo do
livro caixa (por empresa).

### Padrão 3 — varredura na linha do tempo

Para quando os registros se sobrepõem e somar não serve.

```js
const pontos = [... todos os inícios e fins ...].sort((a,b) => a-b);
for (let k = 0; k < pontos.length - 1; k++) {
  // entre pontos[k] e pontos[k+1] o conjunto de ativos não muda
  // → dá para dividir a duração por quantos são
}
```

Usado uma vez só: o rateio por concorrência das horas (seção 14). A ideia geral
é que **entre dois eventos consecutivos nada muda**, então cada trecho pode ser
tratado como um bloco homogêneo.

---

## A.6 Onde cada armadilha de linguagem mordeu

Histórico real deste projeto, para não repetir:

| Erro | Linguagem | Sintoma | Causa |
|:--|:--|:--|:--|
| Coluna de encargo zerada | SQL | planilha em branco | `SUM(a+b)` sem `ISNULL` |
| `Cannot access file ''` | PowerShell | UNLOAD falha | barra invertida sumiu no caminho |
| `Token 'Nome' inesperado` | PowerShell | não roda | `.ps1` salvo sem BOM |
| `True` no meio do log | PowerShell | ruído | função com `return $true` |
| `Column 'table_id' not found` | SQL | consulta falha | misturou `SYSCOLUMNS` com `SYSTABCOL` |
| `Invalid string length` | JavaScript | não relê o xlsx | string do V8 estourou nos 58 MB |
| Horas 8,3% infladas | lógica | número errado, sem erro | somou sessões sobrepostas |
| Arquivo travado | Windows | gravação falha no fim | Excel aberto na planilha |
