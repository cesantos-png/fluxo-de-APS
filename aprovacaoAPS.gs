/**
 * Processa e-mails de aprovação de AP: migra os dados do .xlsm anexado para a
 * planilha de destino e notifica no Google Chat.
 *
 * CONFIGURAÇÃO INICIAL (uma vez só):
 * 1. Editor > Serviços (+) > adicionar "Drive API" (v3).
 * 2. Configurações do projeto > Propriedades do script > Adicionar propriedade:
 *      Nome:  WEBHOOK_CHAT
 *      Valor: URL do NOVO webhook do Google Chat (exclua o antigo, que ficou exposto)
 * 3. Preencha APROVADORES abaixo (ou deixe vazio para aceitar qualquer remetente).
 *
 * FUNCIONAMENTO EM CASO DE ERRO:
 * - Se algum anexo falhar, a thread é tentada de novo nas próximas execuções.
 *   Anexos que já foram migrados são lembrados e NÃO são gravados de novo.
 * - Depois de MAX_TENTATIVAS falhas, a thread recebe a label LABEL_ERRO e um
 *   aviso é enviado ao Chat. Para tentar de novo, basta remover essa label
 *   do e-mail (os anexos já migrados continuam sendo pulados).
 */

const CONFIG = {
  SPREADSHEET_ID: "16KRDu8Y2htmGd_mihTaIU3CB3bJzhMQg-Pvn9xMIoVE",
  NOME_DA_ABA: "APs",              // Aba de destino
  NOME_ABA_ANEXO: "",              // Aba dentro do .xlsm de onde ler os dados. Vazio = primeira aba
  LABEL_PROCESSADO: "1_Aprovações",
  LABEL_ERRO: "1_Aprovações_Erro",
  APROVADORES: [],                 // ex.: ["gestor@empresa.com", "diretor@empresa.com"]. Vazio = qualquer remetente
  MAX_THREADS_POR_EXECUCAO: 20,
  MAX_TENTATIVAS: 3,
  TEMPO_LIMITE_MS: 5 * 60 * 1000   // Para antes do limite de 6 min do Apps Script
};

/**
 * Função principal (é esta que o gatilho deve chamar).
 * O lock impede que duas execuções processem as mesmas threads ao mesmo tempo.
 */
function processarEmailsAprovados() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log("Outra execução já está em andamento. Encerrando.");
    return;
  }
  try {
    executarProcessamento();
  } finally {
    lock.releaseLock();
  }
}

function executarProcessamento() {
  const inicio = Date.now();

  const webhook = PropertiesService.getScriptProperties().getProperty("WEBHOOK_CHAT");
  if (!webhook) {
    Logger.log("[AVISO] Propriedade WEBHOOK_CHAT não configurada; notificações no Chat desativadas.");
  }

  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.NOME_DA_ABA);
  if (!sheet) throw new Error(`Aba "${CONFIG.NOME_DA_ABA}" não encontrada na planilha de destino.`);

  const labelProcessado = obterOuCriarLabel(CONFIG.LABEL_PROCESSADO);
  const labelErro = obterOuCriarLabel(CONFIG.LABEL_ERRO);

  // Não depende de "não lido": abrir/responder o e-mail já marca como lido antes do script rodar.
  const query = `subject:aprovação subject:ap -label:${CONFIG.LABEL_PROCESSADO} -label:${CONFIG.LABEL_ERRO}`;
  const threads = GmailApp.search(query, 0, CONFIG.MAX_THREADS_POR_EXECUCAO);
  Logger.log(`[DEBUG] Query: ${query} | Threads encontradas: ${threads.length}`);

  for (const thread of threads) {
    if (Date.now() - inicio > CONFIG.TEMPO_LIMITE_MS) {
      Logger.log("[AVISO] Tempo limite atingido; as threads restantes ficam para a próxima execução.");
      break;
    }
    processarThread(thread, sheet, labelProcessado, labelErro, webhook);
  }
}

/**
 * Processa uma thread: valida a aprovação, migra cada .xlsm (sem repetir os
 * já migrados) e marca a thread como processada ou com erro.
 */
function processarThread(thread, sheet, labelProcessado, labelErro, webhook) {
  const messages = thread.getMessages();
  const ultimaMensagem = messages[messages.length - 1];
  const assunto = ultimaMensagem.getSubject();
  const remetente = ultimaMensagem.getFrom();

  if (!remetenteAutorizado(remetente)) {
    Logger.log(`[DEBUG] "${assunto}" ignorada: remetente não autorizado (${remetente}).`);
    return;
  }

  if (!verificarAprovacao(ultimaMensagem.getPlainBody())) {
    Logger.log(`[DEBUG] "${assunto}" ignorada: última mensagem não é uma aprovação.`);
    return;
  }

  // O .xlsm normalmente vem na mensagem original, não na resposta - procura na thread toda.
  const anexos = obterAnexosXlsmUnicos(messages);
  if (anexos.length === 0) {
    Logger.log(`[DEBUG] "${assunto}" aprovada, mas sem anexo .xlsm. Será verificada de novo na próxima execução.`);
    return;
  }

  // Estado da thread: quais anexos já foram migrados e quantas tentativas falharam.
  const props = PropertiesService.getScriptProperties();
  const chaveEstado = "estado_" + thread.getId();
  const estado = JSON.parse(props.getProperty(chaveEstado) || '{"migrados":[],"tentativas":0}');

  let houveErro = false;

  for (const anexo of anexos) {
    const chave = chaveAnexo(anexo);
    if (estado.migrados.includes(chave)) {
      Logger.log(`[DEBUG] "${anexo.getName()}" já migrado anteriormente; pulando.`);
      continue;
    }
    try {
      migrarAnexo(anexo, sheet);
      estado.migrados.push(chave);
      // Salva imediatamente, para não duplicar caso a execução seja interrompida depois daqui.
      props.setProperty(chaveEstado, JSON.stringify(estado));
      Logger.log(`Dados migrados com sucesso: "${anexo.getName()}" (e-mail "${assunto}")`);
    } catch (erro) {
      houveErro = true;
      Logger.log(`Erro ao processar anexo "${anexo.getName()}": ${erro}`);
    }
  }

  if (!houveErro) {
    props.deleteProperty(chaveEstado);
    // Fica acessível clicando na label, como se fosse uma pasta.
    thread.markRead();
    thread.addLabel(labelProcessado);
    thread.moveToArchive();

    const aprovadoPor = extrairNomeRemetente(remetente);
    enviarNotificacaoGoogleChat(
      webhook,
      `A AP do e-mail "${assunto}" foi aprovada por ${aprovadoPor}\nPróxima ação: entregar ao financeiro`
    );
    return;
  }

  estado.tentativas++;
  if (estado.tentativas >= CONFIG.MAX_TENTATIVAS) {
    // Mantém a lista de migrados, mas zera as tentativas: se alguém remover a
    // label de erro, a thread é tentada de novo sem duplicar o que já entrou.
    estado.tentativas = 0;
    props.setProperty(chaveEstado, JSON.stringify(estado));
    thread.addLabel(labelErro);
    enviarNotificacaoGoogleChat(
      webhook,
      `⚠️ Falha ao migrar a AP do e-mail "${assunto}" após ${CONFIG.MAX_TENTATIVAS} tentativas. ` +
      `${estado.migrados.length} de ${anexos.length} arquivo(s) migrado(s). ` +
      `Verifique o log de execuções e remova a label "${CONFIG.LABEL_ERRO}" para tentar de novo.`
    );
  } else {
    props.setProperty(chaveEstado, JSON.stringify(estado));
    Logger.log(`[DEBUG] "${assunto}" com erro (tentativa ${estado.tentativas}/${CONFIG.MAX_TENTATIVAS}).`);
  }
}

/**
 * Converte o .xlsm, lê as células e grava uma linha na planilha de destino.
 * O arquivo temporário é sempre excluído, mesmo se der erro no meio.
 */
function migrarAnexo(anexo, sheet) {
  let tempFileId = null;
  try {
    tempFileId = converterXlsmParaGoogleSheets(anexo);
    const tempSpreadsheet = SpreadsheetApp.openById(tempFileId);
    const tempSheet = CONFIG.NOME_ABA_ANEXO
      ? tempSpreadsheet.getSheetByName(CONFIG.NOME_ABA_ANEXO)
      : tempSpreadsheet.getSheets()[0];
    if (!tempSheet) throw new Error(`Aba "${CONFIG.NOME_ABA_ANEXO}" não encontrada no anexo.`);

    // Lê a coluna C inteira de uma vez (muito mais rápido que várias getValue()).
    const colunaC = tempSheet.getRange("C1:C77").getValues().map(linha => linha[0]);
    const C = numeroLinha => colunaC[numeroLinha - 1];

    // C25 normalmente é a iniciativa; se não estiver cadastrada, usa C27.
    // Quando não for esse caso, C25 vem como "PREFIXO.| Valor real" - remove até o "|".
    const iniciativaNaoCadastrada = String(C(25)).trim().toUpperCase() === "INICIATIVA NÃO CADASTRADA";
    const valorIniciativa = iniciativaNaoCadastrada ? C(27) : removerPrefixoAntesDoPipe(C(25));

    // C26 sempre entra como negativo, independente do sinal original.
    const valorC26Negativo = converterParaNegativo(C(26));

    // Data de processamento (apenas a data, sem horário).
    const agora = new Date();
    const dataProcessamento = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate());

    // Ordem final: C77, C17, C8, iniciativa (C25 ou C27), C74, [vazio], data, [vazio], C26 (negativo), C71
    const linhaDados = [
      C(77),
      C(17),
      C(8),
      valorIniciativa,
      C(74),
      "",
      dataProcessamento,
      "",
      valorC26Negativo,
      C(71)
    ];

    // Se a conversão quebrou alguma fórmula (macro, vínculo externo), não grava lixo.
    const celulaComErro = linhaDados.find(v => typeof v === "string" && /^#(REF|N\/A|VALUE|DIV\/0|NAME|NUM|ERROR|NULL)/i.test(v));
    if (celulaComErro) {
      throw new Error(`Valor inválido após a conversão do anexo: ${celulaComErro}`);
    }

    // Próxima linha livre considerando as colunas A a J (as que o script grava).
    // Olhar só a coluna A fazia linhas com C77 vazio serem sobrescritas pela gravação seguinte.
    const proximaLinha = obterProximaLinhaLivre(sheet, 1, linhaDados.length);
    sheet.getRange(proximaLinha, 1, 1, linhaDados.length).setValues([linhaDados]);
    SpreadsheetApp.flush();
  } finally {
    if (tempFileId) {
      try {
        DriveApp.getFileById(tempFileId).setTrashed(true);
      } catch (erro) {
        Logger.log(`[AVISO] Não foi possível excluir o arquivo temporário ${tempFileId}: ${erro}`);
      }
    }
  }
}

/**
 * Verifica se o corpo da mensagem é uma aprovação.
 * Ignora o histórico citado, exige palavras inteiras ("ok" não casa com "Outlook")
 * e rejeita negações ("não aprovado", "não está de acordo").
 */
function verificarAprovacao(corpo) {
  const texto = removerTextoCitado(corpo).toLowerCase();
  if (/\bn[aã]o\b[^.\n]{0,30}?(\baprovad[oa]\b|\bde acordo\b|\bok\b)/.test(texto)) return false;
  return /(\baprovad[oa]\b|\bde acordo\b|\bok\b)/.test(texto);
}

/**
 * Remove o histórico citado de uma resposta de e-mail (Gmail e Outlook,
 * em português e inglês), mantendo só o texto novo.
 */
function removerTextoCitado(corpo) {
  let texto = "\n" + String(corpo || "");

  const marcadores = [
    /\n\s*Em\s[^\n]*(?:\n[^\n]*){0,2}?escreveu:/i,                    // Gmail PT
    /\n\s*On\s[^\n]*(?:\n[^\n]*){0,2}?wrote:/i,                         // Gmail EN
    /\n\s*-{2,}\s*(Mensagem original|Original Message)\s*-{2,}/i,       // Outlook antigo
    /\n\s*(De|From):\s[^\n]*\n\s*(Enviad[oa]|Sent|Data|Date)( em)?:/i   // Outlook
  ];

  let corte = texto.length;
  for (const marcador of marcadores) {
    const match = texto.match(marcador);
    if (match && match.index < corte) corte = match.index;
  }
  texto = texto.substring(0, corte);

  // Remove também linhas citadas com ">".
  return texto.split("\n").filter(linha => !linha.trim().startsWith(">")).join("\n");
}

/**
 * Se APROVADORES estiver vazio, aceita qualquer remetente.
 */
function remetenteAutorizado(remetente) {
  if (CONFIG.APROVADORES.length === 0) return true;
  const email = extrairEmail(remetente);
  return CONFIG.APROVADORES.some(aprovador => aprovador.trim().toLowerCase() === email);
}

/**
 * Lista os anexos .xlsm de todas as mensagens da thread, sem repetir
 * o mesmo arquivo (mesmo nome e tamanho) reenviado em respostas/encaminhamentos.
 */
function obterAnexosXlsmUnicos(messages) {
  const vistos = new Set();
  return messages
    .flatMap(msg => msg.getAttachments())
    .filter(anexo => anexo.getName().toLowerCase().endsWith(".xlsm"))
    .filter(anexo => {
      const chave = chaveAnexo(anexo);
      if (vistos.has(chave)) return false;
      vistos.add(chave);
      return true;
    });
}

function chaveAnexo(anexo) {
  return anexo.getName() + "|" + anexo.getSize();
}

/**
 * Converte o valor para negativo. Aceita número ou texto no formato
 * brasileiro ("R$ 1.500,00"). Se não der para converter, grava como veio e loga.
 */
function converterParaNegativo(valor) {
  if (typeof valor === "number") return -Math.abs(valor);
  if (typeof valor === "string" && valor.trim() !== "") {
    const limpo = valor.replace(/R\$|\s/g, "").replace(/\./g, "").replace(",", ".");
    const numero = Number(limpo);
    if (!isNaN(numero)) return -Math.abs(numero);
    Logger.log(`[AVISO] C26 não é numérico ("${valor}"); gravado sem conversão para negativo.`);
  }
  return valor;
}

function obterOuCriarLabel(nome) {
  return GmailApp.getUserLabelByName(nome) || GmailApp.createLabel(nome);
}

/**
 * Extrai só o nome de exibição de "Nome <email@dominio.com>".
 * Se não houver esse formato, retorna o valor original.
 */
function extrairNomeRemetente(remetente) {
  if (typeof remetente !== "string") return remetente;
  const match = remetente.match(/^(.*?)\s*<.*>$/);
  const nome = match ? match[1].replace(/^"|"$/g, "").trim() : remetente.trim();
  return nome || remetente.trim();
}

/**
 * Extrai só o endereço de e-mail de "Nome <email@dominio.com>", em minúsculas.
 */
function extrairEmail(remetente) {
  const texto = String(remetente || "");
  const match = texto.match(/<([^>]+)>/);
  return (match ? match[1] : texto).trim().toLowerCase();
}

/**
 * Envia uma mensagem de texto pro Google Chat via webhook de entrada.
 * Não interrompe o processamento se o envio falhar (só loga o erro).
 */
function enviarNotificacaoGoogleChat(webhookUrl, mensagem) {
  if (!webhookUrl) return;
  try {
    UrlFetchApp.fetch(webhookUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ text: mensagem })
    });
  } catch (erro) {
    Logger.log(`Erro ao enviar notificação pro Google Chat: ${erro}`);
  }
}

/**
 * Remove tudo antes (e incluindo) o "|" de um texto no formato
 * "PREFIXO.| Valor real". Se não houver "|", retorna o texto aparado.
 */
function removerPrefixoAntesDoPipe(texto) {
  if (typeof texto !== "string") return texto;
  const indice = texto.indexOf("|");
  if (indice === -1) return texto.trim();
  return texto.substring(indice + 1).trim();
}

/**
 * Retorna o número da próxima linha vazia considerando o bloco de colunas
 * informado (a partir de colunaIndex, numColunas colunas). Uma linha conta
 * como ocupada se QUALQUER coluna do bloco tiver valor. Conteúdo solto em
 * colunas fora do bloco não afeta o resultado.
 */
function obterProximaLinhaLivre(sheet, colunaIndex, numColunas) {
  const ultimaLinha = Math.max(sheet.getLastRow(), 1);
  const valores = sheet.getRange(1, colunaIndex, ultimaLinha, numColunas || 1).getValues();
  for (let i = valores.length - 1; i >= 0; i--) {
    if (valores[i].some(v => v !== "" && v !== null)) {
      return i + 2;
    }
  }
  return 1;
}

/**
 * Converte o anexo Excel (.xlsm) para Google Sheets nativo.
 * Requer o serviço avançado Drive API (v3) ativado no projeto.
 */
function converterXlsmParaGoogleSheets(blobAnexo) {
  const fileMetadata = {
    name: "TEMP_" + blobAnexo.getName(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  const file = Drive.Files.create(fileMetadata, blobAnexo.copyBlob());
  return file.id;
}
