# Scripts de extração do Domínio

Cada relatório tem o seu script. Você roda um, ou roda todos.

Há três documentos, cada um para uma pergunta diferente:

| Documento | Responde |
|:--|:--|
| **este aqui** | qual comando eu rodo |
| [DOCUMENTACAO-TECNICA.md](DOCUMENTACAO-TECNICA.md) | de onde sai cada número e qual a fórmula |
| [CODIGO-EXPLICADO.md](CODIGO-EXPLICADO.md) | o que cada linha do código faz — escrito para quem nunca programou |

## Antes de tudo

```powershell
cd "C:\Users\LENOVO\Documents\BECKUP DOMINIO\SCRIPTS"
.\00-BANCO.ps1 identificar
```

Isso sobe o banco e mostra a ficha dele: versão, tamanho, quantas empresas,
até que data cada módulo tem dado, e se o `UNLOAD` está liberado. É o primeiro
comando a rodar sempre que voltar ao projeto.

Se o PowerShell reclamar de política de execução:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

## Os comandos

### Banco

```powershell
.\00-BANCO.ps1              # status resumido (não sobe nada)
.\00-BANCO.ps1 iniciar      # sobe o servidor
.\00-BANCO.ps1 identificar  # ficha completa do banco e do conteúdo
.\00-BANCO.ps1 reiniciar    # derruba e sobe de novo
.\00-BANCO.ps1 parar        # derruba e devolve a RAM
```

### Relatórios

```powershell
.\01-FISCAL.ps1           # faturamento, CNPJ, IE e impostos apurados
.\02-FOLHA.ps1            # encargos patronais, headcount, admissões, demissões
.\03-CONTABIL.ps1         # balanço patrimonial, DRE, balancete
.\04-LIVRO-CAIXA.ps1      # movimento de caixa (use -SemDetalhe para pular as 880 mil linhas)
.\05-PRODUTOR-RURAL.ps1   # recorte dos produtores rurais
.\06-PRODUTIVIDADE.ps1    # tempo por empresa e por usuário
```

### Tudo de uma vez

```powershell
.\EXTRAIR-TUDO.ps1
.\EXTRAIR-TUDO.ps1 -SemDetalheCaixa      # bem mais rápido
.\EXTRAIR-TUDO.ps1 -Pular 04,06
.\EXTRAIR-TUDO.ps1 -PararAoFim           # derruba o servidor no final
```

Cada script sobe o servidor sozinho se ele estiver parado, e chama
`00-CADASTROS.ps1` sozinho se os cadastros ainda não tiverem sido extraídos.
Você não precisa rodar nada na mão antes.

## Mudar o período ou os caminhos

Tudo está em **`_lib\config.ps1`**. Os campos que você mais vai mexer:

```powershell
DataIni = '2024-01-01'     # início do período (fiscal, folha, caixa)
DataFim = '2026-12-31'     # fim
DataIniLog = '2023-01-01'  # o log de acesso costuma ter mais história
```

`DataIniContabil` fica em 1995 de propósito: o saldo acumulado do balanço só
fecha se somar o histórico inteiro.

`PastaDados` precisa ser um caminho local **curto e sem espaço** — quem grava
o arquivo é o servidor, e caminho com espaço quebra o `UNLOAD`.

Para uma rodada avulsa, sem mexer no arquivo, dá para sobrescrever por
variável de ambiente:

```powershell
$env:DOM_DATA_INI = '2026-01-01'
$env:DOM_PASTA_SAIDA = 'D:\so-2026'
.\EXTRAIR-TUDO.ps1
```

Valem `DOM_DATA_INI`, `DOM_DATA_FIM`, `DOM_PASTA_SAIDA`, `DOM_PASTA_DADOS` e
`DOM_BANCO`. Limpe com `Remove-Item Env:\DOM_DATA_INI` quando terminar.

## Se der erro

**"está aberto no Excel"** — o script checa antes de começar, para não
descobrir isso depois de cinco minutos de extração. Feche a planilha, ou use
`-Pular` para saltar aquele relatório.

**"Permission denied ... UNLOAD"** — alguém subiu o Domínio por fora, sem
`-gl all`. Rode `.\00-BANCO.ps1 reiniciar`.

**"Token inesperado" ou acento virando `Ã£`** — o `.ps1` perdeu o BOM. Veja
o comando de reaplicar BOM mais abaixo.

**Servidor não sobe** — confira a RAM livre no `.\00-BANCO.ps1`. Se estiver
abaixo de uns 400 MB, feche o Excel e o navegador antes.

## Como está organizado

```
SCRIPTS/
  00-BANCO.ps1            liga, desliga e identifica o banco
  00-CADASTROS.ps1        empresas, CNAE, plano de contas (base comum)
  01..06-*.ps1            um script por relatório
  EXTRAIR-TUDO.ps1        roda os seis na ordem certa
  _lib/
    config.ps1            TODA a configuração
    dominio.ps1           funções: subir servidor, consultar, exportar
  _gerar/
    lib.js                leitura dos .tsv + dimensão empresa + formatação
    fiscal.js  folha.js  contabil.js  caixa.js  rural.js  produtividade.js
    node_modules/         exceljs (se sumir: npm install aqui dentro)
```

O fluxo de cada relatório é o mesmo: o `.ps1` extrai os `.tsv` para
`C:\dbtmp\out` e chama o `.js` que monta o `.xlsx` em `..\EXTRACAO`.

## Por que é rápido

A extração usa **`UNLOAD SELECT ... TO arquivo`**: quem escreve o arquivo é o
servidor, direto no disco. Ler linha a linha pelo ODBC nesta base faz umas
60 linhas por minuto — as 226 mil linhas do log de acesso levariam dias. Com
`UNLOAD` levam 2 segundos.

Para o `UNLOAD` funcionar o servidor precisa subir com **`-gl all`**. O
`00-BANCO.ps1` já faz isso. Se alguém subir o Domínio por fora e o `UNLOAD`
der `Permission denied`, rode `.\00-BANCO.ps1 reiniciar`.

## Detalhes que economizam tempo depois

**Os `.ps1` têm que ficar em UTF-8 com BOM.** O Windows PowerShell 5.1 lê
arquivo sem BOM como ANSI e os acentos viram lixo, o que quebra o parser. Se
editar em algum lugar que salve sem BOM, rode:

```powershell
$bom = New-Object System.Text.UTF8Encoding($true)
Get-ChildItem . -Recurse -Filter *.ps1 | ForEach-Object {
  $t = [IO.File]::ReadAllText($_.FullName, [Text.UTF8Encoding]::new($false))
  [IO.File]::WriteAllText($_.FullName, $t, $bom)
}
```

**Os `.tsv` do UNLOAD não são TSV de verdade.** O separador é a sequência
literal `\` + `t`, não um TAB, e o encoding é cp1252. O `lib.js` já trata os
dois. Se for abrir um desses arquivos em outra ferramenta, lembre disso.

**`SUM(a + b)` vira NULL** se qualquer parcela for nula. Sempre
`SUM(ISNULL(a,0) + ISNULL(b,0))` — foi assim que a coluna de encargos saiu
zerada na primeira tentativa.

**Caminho com barra normal no UNLOAD.** `C:/dbtmp/out/x.tsv`, não
`C:\dbtmp\out\x.tsv`: a barra invertida se perde no caminho até o SQL.

**Com pouca RAM o servidor fica lento, não rápido.** Se `Cache` em
`config.ps1` for maior que a RAM livre, o servidor pagina. O `00-BANCO.ps1`
mostra a RAM livre no status.

## Conferência

O jeito rápido de saber se a extração saiu certa: empresa 345
(L B GOMES RAMOS), competência 10/2025.

| Valor | Esperado | Onde conferir |
|:--|--:|:--|
| Faturamento | 362.025,28 | 01-FISCAL, aba Mensal por empresa |
| Compras | 699.955,52 | idem |
| Impostos (IRRF-APF) | 55,84 | idem |
| Folha (proventos) | 4.818,15 | 02-FOLHA, aba Encargos mensais |

O que cada planilha contém, de onde vem cada número e o que o banco **não**
tem está em `..\EXTRACAO\LEIA-ME.md`.
