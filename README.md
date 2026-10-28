# Fluxo de aprovação de APs

Script do Google Apps Script (`aprovacaoAPS.gs`) que acompanha os e-mails de aprovação de AP no Gmail e, quando uma AP é aprovada:

1. Migra os dados do `.xlsm` anexado para a planilha de controle.
2. Gera um PDF consolidado com os anexos `.pdf` (nota fiscal, boleto etc.) e a mensagem de aprovação, e salva no Google Drive.
3. Salva uma cópia original do e-mail de aprovação (`.eml`) como comprovante.
4. Avisa no Google Chat, com o link do PDF.

## Como funciona

O acionador chama `processarEmailsAprovados` periodicamente. A cada execução, o script:

1. Busca no Gmail conversas com **"aprovação"** e **"ap"** no assunto que ainda não tenham as labels `1_Aprovações` ou `1_Aprovações_Erro`.
2. Para cada conversa, olha a **última mensagem**:
   - O remetente precisa estar em `APROVADORES`. Se a lista estiver vazia, qualquer remetente vale.
   - O texto novo (sem o histórico citado) precisa conter **"aprovado/aprovada"**, **"de acordo"** ou **"ok"**, sem negação como "não aprovado".
   - Se não for uma aprovação, a conversa é ignorada e verificada de novo na próxima execução.
3. Procura os anexos `.xlsm` em toda a conversa (normalmente estão na mensagem original) e grava uma linha por arquivo na aba `APs`.
4. Gera o PDF consolidado e o `.eml` na pasta do Drive.
5. Marca a conversa como lida, aplica a label `1_Aprovações`, arquiva e envia a notificação no Chat.

### Dados migrados do `.xlsm`

O script lê a coluna C da primeira aba do `.xlsm` (ou da aba em `NOME_ABA_ANEXO`) e grava nas colunas A a J da próxima linha livre:

| Coluna | Origem |
|---|---|
| A | C77 |
| B | C17 |
| C | C8 |
| D | Iniciativa: C25 sem o prefixo até o `\|`; se C25 for "INICIATIVA NÃO CADASTRADA", usa C27 |
| E | C74 |
| F | (vazio) |
| G | Data do processamento |
| H | (vazio) |
| I | C26, sempre negativo |
| J | C71 |

Se a conversão do `.xlsm` gerar erros de fórmula (`#REF!`, `#N/A` etc.), a linha não é gravada e a conversa entra no tratamento de erro.

### PDF consolidado

Arquivos gerados na pasta de destino para cada AP:

- `merged_docs - <assunto> - <data>.pdf`: anexos `.pdf` da conversa e, **por último**, a mensagem de aprovação.
- `merged_docs - <assunto> - <data>.eml`: cópia exata do e-mail de aprovação, com cabeçalhos e assinatura digital (DKIM) do servidor de quem enviou. Serve como comprovante de autenticidade.

**Mensagem de aprovação:** entra só a última mensagem da conversa, com cabeçalho (De, Para, Cc, Data, Assunto), texto e assinatura. As imagens da assinatura, embutidas no e-mail ou hospedadas na internet, são incorporadas ao PDF.

**Ordem dos anexos:** é definida pelo **conteúdo** de cada PDF, não pelo nome do arquivo. Cada anexo é convertido temporariamente em Google Docs, com reconhecimento de texto (funciona com PDF escaneado), e o script procura as marcas típicas de cada tipo listado em `ORDEM_DOCUMENTOS`:

| Ordem | Tipo | Exemplos de marcas |
|---|---|---|
| 1 | Nota fiscal | DANFSe, DANFE, NFS-e, "Documento auxiliar", "Chave de acesso", Prestador, Tomador |
| 2 | Boleto | "Linha digitável", "Ficha de compensação", "Local de pagamento", Beneficiário, Pagador, "Nosso número" |

- O tipo com mais marcas encontradas vence, com mínimo de `MIN_MARCAS` (2). Assim, um boleto que cita "Ref. NFS-e 283" continua sendo boleto.
- PDFs não identificados vêm depois, na ordem em que foram anexados.
- Para mudar a ordem, troque os blocos de posição em `ORDEM_DOCUMENTOS`. Para incluir outro tipo (contrato, orçamento...), acrescente um bloco com `tipo` e `marcas`.

## Configuração inicial

Faça uma vez, no editor do Apps Script:

1. **Serviços (+)**: adicione **Drive API** (v3) e **Gmail API** (v1). A Gmail API é usada para recuperar as imagens embutidas nas assinaturas.
2. **Configurações do projeto > Propriedades do script > Adicionar propriedade do script**:
   - Propriedade: `WEBHOOK_CHAT`
   - Valor: URL do webhook do espaço do Google Chat (no Chat: nome do espaço > **Apps e integrações** > **Webhooks** > **Adicionar webhook**).

   A URL fica fora do código de propósito: quem tem a URL consegue postar no espaço. Sem essa propriedade, o script funciona, mas não envia notificações.
3. Ajuste o bloco `CONFIG` (veja abaixo), principalmente `PASTA_DESTINO_ID`.
4. Rode `processarEmailsAprovados` uma vez manualmente para autorizar as permissões (Gmail, Drive, Planilhas, Documentos e acesso externo).
5. **Acionadores** (ícone de relógio) > **Adicionar acionador**:
   - Função: `processarEmailsAprovados`
   - Fonte do evento: **Baseado no tempo**, contador de minutos, a cada 10 minutos (por exemplo)
   - Notificação de falha: diária

   O acionador lê o Gmail da conta de quem o criou. Crie só um.

### Como saber se o acionador está ativo

- **Acionadores**: se ele aparece na lista, está ativo. Não há opção de pausar; para desativar, exclua.
- **Execuções**: cada execução automática aparece com o tipo "Acionador" e o status (Concluída / Falha).

## Parâmetros (`CONFIG`)

| Parâmetro | Descrição |
|---|---|
| `SPREADSHEET_ID` | ID da planilha de destino |
| `NOME_DA_ABA` | Aba de destino (`APs`) |
| `NOME_ABA_ANEXO` | Aba do `.xlsm` de onde ler os dados. Vazio = primeira aba |
| `LABEL_PROCESSADO` / `LABEL_ERRO` | Labels do Gmail para conversas processadas e com erro |
| `APROVADORES` | E-mails autorizados a aprovar, ex.: `["gestor@empresa.com"]`. Vazio = qualquer remetente |
| `MAX_THREADS_POR_EXECUCAO` | Conversas analisadas por execução |
| `MAX_TENTATIVAS` | Falhas seguidas antes de aplicar a label de erro |
| `TEMPO_LIMITE_MS` | Para antes do limite de 6 min do Apps Script; o restante fica para a próxima execução |
| `PASTA_DESTINO_ID` | ID da pasta do Drive para os PDFs (o trecho depois de `/folders/` no link da pasta). Vazio = usa ou cria `NOME_PASTA_PADRAO` na raiz do Drive. A conta do acionador precisa ter permissão de edição na pasta |
| `NOME_PASTA_PADRAO` | Nome da pasta usada quando `PASTA_DESTINO_ID` está vazio |
| `PREFIXO_ARQUIVO` | Início do nome dos arquivos gerados (`merged_docs`) |
| `PDF_LIB_URL` | Endereço da biblioteca pdf-lib, usada para juntar os PDFs |
| `MAX_IMAGENS_EXTERNAS` / `MAX_BYTES_IMAGEM` | Limites para baixar imagens da internet presentes no e-mail |
| `ORDEM_DOCUMENTOS` / `MIN_MARCAS` | Ordem e reconhecimento dos tipos de anexo (veja "PDF consolidado") |

## Ter os arquivos no computador

O script roda nos servidores do Google e não grava direto no computador. Instale o **Google Drive para computador** e deixe a pasta de destino disponível off-line: cada PDF gerado aparece automaticamente numa pasta local.

## Erros e reprocessamento

- Se algo falhar (migração de um `.xlsm` ou geração do PDF), a conversa é tentada de novo nas próximas execuções. O que já foi feito é lembrado: `.xlsm` já migrados não são gravados de novo e o PDF não é gerado em duplicidade.
- Depois de `MAX_TENTATIVAS` falhas, a conversa recebe a label `1_Aprovações_Erro` e um aviso vai para o Chat. Para tentar de novo, remova essa label do e-mail.
- **Para refazer uma AP já processada:** remova a label `1_Aprovações` da conversa, apague a linha correspondente na planilha e apague o PDF e o `.eml` gerados.

### Mensagens no registro de execuções

| Mensagem | Significado |
|---|---|
| `"..." ignorada: última mensagem não é uma aprovação` | A conversa ainda aguarda a resposta de aprovação |
| `aprovada, mas sem anexo .xlsm` | Aprovada, mas sem planilha anexada; será verificada de novo |
| `"arquivo.pdf": Boleto (7 marca(s))` | Como cada anexo foi classificado para a ordenação |
| `Imagens no e-mail ...` | Imagens encontradas na mensagem de aprovação |
| `[AVISO] ... imagem ...` | Uma imagem não pôde ser incluída no PDF (o PDF é gerado mesmo assim) |
| `Não foi possível ler o texto de "..."` | O anexo não pôde ser classificado e ficou na ordem original |
| `Propriedade WEBHOOK_CHAT não configurada` | Notificações no Chat desativadas |

## Limitações

- PDFs protegidos por senha podem falhar ao juntar.
- Assinaturas digitais de PDFs anexados deixam de ser válidas no arquivo final (o conteúdo continua visível).
- Um mesmo PDF com nota e boleto juntos é classificado pelo tipo com mais marcas.
- O PDF da mensagem de aprovação fica muito próximo do original, mas layouts muito elaborados podem ter pequenas diferenças de espaçamento ou fonte.
- Cada anexo PDF leva alguns segundos a mais por causa da leitura do conteúdo.
