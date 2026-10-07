---
name: dominio-servidor-sap
description: Entrar no servidor do Domínio Sistemas (SAP SQL Anywhere 17, schema bethadba), acessar o banco e extrair dados para CSV/Excel com o kit SERVIDOR-SAP-DOMINIO. Use quando o usuário pedir para conectar no banco do Domínio, acessar o servidor do escritório, extrair tabelas/empresas/notas/lançamentos/folha, instalar o SQL Anywhere ou o driver ODBC, ou montar dashboard com dados do Domínio.
---

# Servidor SAP do Domínio — conectar e extrair

Kit: `C:\Users\LENOVO\Documents\BECKUP DOMINIO\SERVIDOR-SAP-DOMINIO\`
(se a pasta for movida, procurar por `02-CONECTAR.ps1`).

```
SERVIDOR-SAP-DOMINIO\
  APP-SAP-SQLAnywhere17\   cópia do SQL Anywhere 17.0.11 (Bin64, Bin32) — servidor + driver ODBC
  01-INSTALAR-SAP.ps1      instala a cópia, registra o ODBC 32/64, PATH  (Administrador)
  02-CONECTAR.ps1          testa porta, login, versão, empresas, UNLOAD; -Salvar grava o perfil
  03-EXTRAIR.ps1           -Lista | -Colunas t | -Tabela t [-Onde] | -Sql | -ArquivoSql | -Relatorios
  SCRIPTS\                 kit dos 6 relatórios .xlsx (fiscal, folha, contábil, caixa, rural, produtividade)
  SCRIPTS\_lib\config.ps1  TODA a configuração; conexao.json (perfil salvo) sobrepõe
  SAIDA\                   CSVs gerados pelo 03-EXTRAIR
```

## Ordem de trabalho

1. **Driver presente?** `.\01-INSTALAR-SAP.ps1 -Verificar`. Se faltar, o usuário
   roda `.\01-INSTALAR-SAP.ps1` como Administrador (não tente elevar sozinho).
2. **Conectar.** `.\02-CONECTAR.ps1 -Ip <ip> -Servidor <nome> -Usuario <u> -Senha <s> -Salvar`.
   Sem `-Ip` = banco `.db` local (o script sobe o `dbsrv17` com `-gl all`).
   Não sabe o IP: `.\02-CONECTAR.ps1 -Procurar` (varre a /24 na porta 2638).
3. **Extrair.** `.\03-EXTRAIR.ps1 -Lista` primeiro (catálogo com nº de linhas),
   depois `-Tabela` / `-Sql`. Relatórios prontos: `-Relatorios`.

Execute sempre com `powershell -ExecutionPolicy Bypass -File <script>` ou
`Set-ExecutionPolicy -Scope Process Bypass` antes.

## Fatos que já custaram horas

- **Senha do `bethadba` não existe para nós.** Login é por usuário **EXTERNO**
  criado no Domínio pelo Gerente: Controle > Permissões > Usuários > Novo >
  Tipo de acesso = Externo, situação Ativo. No backup antigo: `externo` / `123456`.
  Login integrado do Windows é recusado pela base.
- **Connection string** (driver, não DSN — o DSN "Contabil" antigo é 32 bits):
  `Driver={SQL Anywhere 17};Server=<srv>;DBN=<base>;UID=<u>;PWD=<s>;Host=<ip>:2638`
- **Nome do servidor** aparece no Domínio em "Gerenciar Servidor de Banco de
  Dados". Erro "database server not found" = nome errado; SQLSTATE 28000 = usuário/senha.
- **Fora da rede do escritório não conecta.** Precisa VPN; AnyDesk não expõe a
  porta. Alternativa: rodar o kit DENTRO do servidor via AnyDesk (copiar a pasta).
- **Volume: nunca ler linha a linha pelo ODBC** (~60 linhas/min nesta máquina).
  Use `UNLOAD` (`Export-DomTabela`): local grava com `TO`, remoto com
  `INTO CLIENT FILE`; se o usuário não tiver `WRITE CLIENT FILE`, o kit cai
  sozinho para leitura pelo driver gravando no mesmo formato.
- Arquivo do `UNLOAD`: separador é `\` + `t` **literal**, encoding **cp1252**,
  sem cabeçalho, caminho com **barra normal** e sem espaço. O 03-EXTRAIR já
  converte para CSV `;` UTF-8 BOM com vírgula decimal.
- `UNLOAD` com `ORDER BY` caro trava — ordenar depois.
- `SUM(a+b)` vira NULL se uma parcela for NULL: `SUM(ISNULL(a,0)+ISNULL(b,0))`.
- Datas com lixo (1979, 8320): sempre filtrar faixa `>= '2000-01-01' AND <= '2035-12-31'`.
- `.ps1` precisa ser **UTF-8 com BOM** (PowerShell 5.1). Depois de Write/Edit, regravar com BOM.
- Contar tabelas por `SYS.SYSTAB ... table_type_str='BASE'` (são ~4.800).
- Servidor remoto é do escritório: **nunca** parar/reiniciar (o kit já recusa).

## Onde mora cada dado (schema `bethadba`)

| Assunto | Tabela | Colunas-chave |
|:--|:--|:--|
| Empresas | `geempre` | codi_emp, nome_emp, razao_emp, cgce_emp (CNPJ/CPF), stat_emp ('A' ativa), i_cnae20, cida_emp, esta_emp |
| Notas de saída | `efsaidas` | codi_emp, dsai_sai (data), vcon_sai (valor); impostos em tabelas-filho |
| Itens de saída / entrada | `efmvspro` / `efmvepro` | produtos nota a nota |
| Notas de entrada | `efentradas` | ddoc_ent |
| Apuração de impostos | `efsdoimp` | data_sim |
| Modelo do documento | `efespecies.mode_esp` | 36=NF-e 55, 41=NFC-e 65 |
| Lançamentos contábeis | `ctlancto` | data_lan, vlor_lan, codi_usu (operador); sem hora |
| Plano de contas | `ctcontas` | |
| Folha | `fobasesserv`, `foguiainss`, `forescisoes` | competencia; FGTS não existe no banco |
| Log de uso | `geloguser` + `GEMODULOS` | data_log, usua_log, sist_log |
| Honorários | `HRCONTRATO`, `HRCLIENTE` | HRCLIENTE.I_CLIENTE_FIXO = geempre.codi_emp |
| Produtos | usar `codigo_barras_pdi` (ean_pdi vem vazio) | |

Prefixos: `ef` Escrita Fiscal, `ct` Contabilidade, `fo` Folha, `ge` Geral, `hr` Honorários.
Antes de inventar coluna: `.\03-EXTRAIR.ps1 -Colunas <tabela>`.

## Como responder

Português, direto. O usuário mexe em Domínio de cliente real, às vezes via
AnyDesk. Se a conexão falhar, mostre a linha exata do 02-CONECTAR e o que
fazer — não peça a senha do bethadba.
