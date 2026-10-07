Funciona em quatro etapas, nesta ordem:

**1. Baixar** (em qualquer computador)
O `git clone` traz do GitHub o programa da SAP, os scripts e a skill numa pasta só.

**2. Instalar** (uma vez por computador)
O `01-INSTALAR-SAP.ps1` coloca o SQL Anywhere 17 no Windows e registra o driver ODBC. É esse driver que deixa o PC "falar" com o banco do Domínio.

**3. Conectar** (uma vez por servidor)
O `02-CONECTAR.ps1` vai até o servidor do Domínio pela rede, na porta 2638, e entra com o usuário Externo. Ele confere em sequência:
- se a porta está aberta;
- se o usuário e a senha entram;
- se consegue ler as empresas;
- se a extração rápida está liberada.

Com `-Salvar`, ele guarda esses dados. Daí em diante os outros scripts já sabem onde e como entrar.

**4. Extrair** (sempre que precisar)
O `03-EXTRAIR.ps1` pede os dados ao servidor:
- **Caminho rápido:** o próprio servidor monta o arquivo e manda pronto. Leva segundos, mesmo com milhões de linhas.
- **Caminho de reserva:** se o usuário não tiver permissão para o caminho rápido, os dados vêm linha por linha. O resultado é o mesmo, só demora mais.

No fim, o script converte tudo para CSV no formato do Excel em português e grava na pasta `SAIDA\`. Com `-Relatorios`, ele monta as 6 planilhas prontas: fiscal, folha, contábil, caixa, rural e produtividade.

```
GitHub ──clone──▶ PC ──instala──▶ driver SAP
                                    │
                         02-CONECTAR (IP, usuário Externo)
                                    │
                     Servidor do Domínio (porta 2638)
                                    │
                         03-EXTRAIR (UNLOAD rápido / reserva)
                                    │
                     SAIDA\*.csv  e  relatórios .xlsx ──▶ Excel / BI
```

Duas condições valem sempre:
- **Precisa estar na rede do escritório.** De fora, só com VPN, ou rodando o kit dentro do próprio servidor pelo AnyDesk.
- **O kit só lê.** O servidor remoto nunca é desligado nem alterado pelos scripts, porque o escritório está usando o Domínio ao mesmo tempo.

A skill `dominio-servidor-sap` guarda esse passo a passo e o mapa das tabelas. Assim, quando você pedir "extrai as notas da empresa X", eu já sei qual comando e qual tabela usar.
