# Servidor SAP do Domínio — kit de acesso e extração

Tudo o que precisa para entrar no banco do Domínio (SAP SQL Anywhere 17) e
tirar os dados, numa pasta só. Dá para copiar esta pasta para outro PC (ou
para o próprio servidor, via AnyDesk) e rodar lá.

| Pasta / arquivo | O que é |
|:--|:--|
| `APP-SAP-SQLAnywhere17\` | o programa da SAP: servidor `dbsrv17` + driver ODBC (64 e 32 bits), v17.0.11 |
| `01-INSTALAR-SAP.ps1` | instala o programa acima e registra o driver ODBC |
| `02-CONECTAR.ps1` | entra no servidor e confere se dá para ler |
| `03-EXTRAIR.ps1` | extrai tabelas e consultas para CSV, ou os 6 relatórios .xlsx |
| `SCRIPTS\` | kit dos relatórios (fiscal, folha, contábil, caixa, rural, produtividade) |
| `SKILL\dominio-servidor-sap\` | skill do Claude com o passo a passo e o mapa das tabelas |
| `SAIDA\` | onde os CSVs aparecem |

## Passo 1 — abrir o PowerShell nesta pasta

```powershell
cd "C:\Users\LENOVO\Documents\BECKUP DOMINIO\SERVIDOR-SAP-DOMINIO"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

## Passo 2 — instalar o SAP (só uma vez por máquina)

Abra o PowerShell **como Administrador** e rode:

```powershell
.\01-INSTALAR-SAP.ps1
```

Para só conferir o que já está instalado: `.\01-INSTALAR-SAP.ps1 -Verificar`

## Passo 3 — entrar no servidor

Você precisa de um **usuário Externo** do Domínio (o Gerente cria em
*Controle > Permissões > Usuários > Novo > Tipo de acesso = Externo*).

```powershell
# servidor do escritório
.\02-CONECTAR.ps1 -Ip 192.168.0.10 -Servidor contabil -Usuario externo -Senha 123456 -Salvar

# não sabe o IP? procura na rede
.\02-CONECTAR.ps1 -Procurar

# banco .db nesta máquina (caminho em SCRIPTS\_lib\config.ps1)
.\02-CONECTAR.ps1 -Local -Salvar
```

O `-Salvar` guarda a conexão; daí em diante os outros scripts usam ela sozinhos.
O nome do servidor aparece no Domínio em *Gerenciar Servidor de Banco de Dados*.

Se não conectar de fora do escritório: o banco não fica exposto na internet.
Precisa de VPN, ou copie esta pasta para o servidor pelo AnyDesk e rode lá.

## Passo 4 — extrair

```powershell
.\03-EXTRAIR.ps1 -Lista                         # todas as tabelas e quantas linhas cada uma tem
.\03-EXTRAIR.ps1 -Colunas efsaidas              # colunas de uma tabela
.\03-EXTRAIR.ps1 -Tabela geempre                # cadastro de empresas inteiro
.\03-EXTRAIR.ps1 -Tabela efsaidas -Onde "dsai_sai >= '2026-01-01'"
.\03-EXTRAIR.ps1 -Sql "SELECT codi_emp, nome_emp FROM bethadba.geempre WHERE stat_emp='A'" -Nome ativas
.\03-EXTRAIR.ps1 -Relatorios                    # os 6 relatórios em Excel
```

Os CSVs saem em `SAIDA\`, prontos para o Excel (separador `;`, vírgula decimal).
Os relatórios .xlsx saem na pasta configurada em `SCRIPTS\_lib\config.ps1`
(`PastaSaida`); o período também se ajusta lá. O kit de relatórios precisa do
Node instalado.

## Problemas comuns

| Mensagem | O que fazer |
|:--|:--|
| Porta TCP FECHADA | Domínio desligado no servidor, firewall bloqueando a 2638, ou você está fora da rede |
| Login FALHOU, 28000 | usuário ou senha errados — use um usuário Externo |
| database server not found | nome do servidor errado (`-Servidor`) |
| Extração rápida bloqueada | funciona assim mesmo, só que mais devagar em tabelas grandes |
| Token inesperado / acentos estranhos | o `.ps1` perdeu o BOM; regrave como UTF-8 com BOM |
