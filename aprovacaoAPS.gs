/**
 * Fluxo de aprovação de APs. Configuração e funcionamento: ver README.md.
 */

const CONFIG = {
  SPREADSHEET_ID: "16KRDu8Y2htmGd_mihTaIU3CB3bJzhMQg-Pvn9xMIoVE",
  NOME_DA_ABA: "APs",
  NOME_ABA_ANEXO: "",              // Vazio = primeira aba do .xlsm
  LABEL_PROCESSADO: "1_Aprovações",
  LABEL_ERRO: "1_Aprovações_Erro",
  APROVADORES: [],                 // Vazio = qualquer remetente
  MAX_THREADS_POR_EXECUCAO: 20,
  MAX_TENTATIVAS: 3,
  TEMPO_LIMITE_MS: 5 * 60 * 1000,  // Para antes do limite de 6 min do Apps Script
  PASTA_DESTINO_ID: "",            // Vazio = NOME_PASTA_PADRAO na raiz do Drive
  NOME_PASTA_PADRAO: "APs Aprovadas",
  PREFIXO_ARQUIVO: "merged_docs",
  PDF_LIB_URL: "https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js",
  MAX_IMAGENS_EXTERNAS: 20,
  MAX_BYTES_IMAGEM: 2 * 1024 * 1024,
  // Ordem dos anexos no PDF final; ganha o tipo com mais marcas no texto (empate: o primeiro).
  ORDEM_DOCUMENTOS: [
    {
      tipo: "Nota fiscal",
      marcas: [/\bDANFS-?e\b/i, /\bDANFE\b/i, /\bNFS-?e\b/i, /\bNF-?e\b/i, /documento auxiliar/i,
               /chave de acesso/i, /\bprestador\b/i, /\btomador\b/i, /\bemitente\b/i,
               /\bdestinat[aá]rio\b/i, /inscri[cç][aã]o municipal/i]
    },
    {
      tipo: "Boleto",
      marcas: [/linha digit[aá]vel/i, /ficha de compensa[cç][aã]o/i, /local de pagamento/i,
               /\bbenefici[aá]rio\b/i, /\bcedente\b/i, /\bsacado\b/i, /\bpagador\b/i, /nosso n[uú]mero/i,
               /autentica[cç][aã]o mec[aâ]nica/i, /recibo do pagador/i,
               /\d{5}\.?\d{5}\s+\d{5}\.?\d{6}\s+\d{5}\.?\d{6}\s+\d\s+\d{14}/]
    }
  ],
  MIN_MARCAS: 2
};

// Função chamada pelo acionador. Async porque a pdf-lib trabalha com Promises.
async function processarEmailsAprovados() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log("Outra execução já está em andamento. Encerrando.");
    return;
  }
  try {
    await executarProcessamento();
  } finally {
    lock.releaseLock();
  }
}

async function executarProcessamento() {
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
    await processarThread(thread, sheet, labelProcessado, labelErro, webhook);
  }
}

async function processarThread(thread, sheet, labelProcessado, labelErro, webhook) {
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

  // Estado entre execuções: anexos já migrados, tentativas e link do PDF gerado.
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

  // Só gera o PDF depois que todos os .xlsm entraram, e uma vez só por thread.
  if (!houveErro && !estado.documentoUrl) {
    try {
      estado.documentoUrl = await gerarDocumentoConsolidado(thread, messages, assunto);
      props.setProperty(chaveEstado, JSON.stringify(estado));
      Logger.log(`PDF consolidado gerado: ${estado.documentoUrl} (e-mail "${assunto}")`);
    } catch (erro) {
      houveErro = true;
      Logger.log(`Erro ao gerar o PDF consolidado do e-mail "${assunto}": ${erro}`);
    }
  }

  if (!houveErro) {
    props.deleteProperty(chaveEstado);
    thread.markRead();
    thread.addLabel(labelProcessado);
    thread.moveToArchive();

    const aprovadoPor = extrairNomeRemetente(remetente);
    enviarNotificacaoGoogleChat(
      webhook,
      `A AP do e-mail "${assunto}" foi aprovada por ${aprovadoPor}\n` +
      `Documentos: ${estado.documentoUrl}\n` +
      `Próxima ação: entregar ao financeiro`
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
      `${estado.migrados.length} de ${anexos.length} arquivo(s) migrado(s)` +
      `${estado.migrados.length === anexos.length ? " (a falha foi na geração do PDF consolidado)" : ""}. ` +
      `Verifique o log de execuções e remova a label "${CONFIG.LABEL_ERRO}" para tentar de novo.`
    );
  } else {
    props.setProperty(chaveEstado, JSON.stringify(estado));
    Logger.log(`[DEBUG] "${assunto}" com erro (tentativa ${estado.tentativas}/${CONFIG.MAX_TENTATIVAS}).`);
  }
}

function migrarAnexo(anexo, sheet) {
  let tempFileId = null;
  try {
    tempFileId = converterXlsmParaGoogleSheets(anexo);
    const tempSpreadsheet = SpreadsheetApp.openById(tempFileId);
    const tempSheet = CONFIG.NOME_ABA_ANEXO
      ? tempSpreadsheet.getSheetByName(CONFIG.NOME_ABA_ANEXO)
      : tempSpreadsheet.getSheets()[0];
    if (!tempSheet) throw new Error(`Aba "${CONFIG.NOME_ABA_ANEXO}" não encontrada no anexo.`);

    const colunaC = tempSheet.getRange("C1:C77").getValues().map(linha => linha[0]);
    const C = numeroLinha => colunaC[numeroLinha - 1];

    // C25 vem como "PREFIXO.| Iniciativa"; se for "INICIATIVA NÃO CADASTRADA", usa C27.
    const iniciativaNaoCadastrada = String(C(25)).trim().toUpperCase() === "INICIATIVA NÃO CADASTRADA";
    const valorIniciativa = iniciativaNaoCadastrada ? C(27) : removerPrefixoAntesDoPipe(C(25));

    const valorC26Negativo = converterParaNegativo(C(26));

    const agora = new Date();
    const dataProcessamento = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate());

    // Colunas A a J da aba de destino.
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

    // Considera A a J, não só A: linhas com C77 vazio seriam sobrescritas.
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

// Palavras inteiras ("ok" não casa com "Outlook") e rejeita negações ("não aprovado").
function verificarAprovacao(corpo) {
  const texto = removerTextoCitado(corpo).toLowerCase();
  if (/\bn[aã]o\b[^.\n]{0,30}?(\baprovad[oa]\b|\bde acordo\b|\bok\b)/.test(texto)) return false;
  return /(\baprovad[oa]\b|\bde acordo\b|\bok\b)/.test(texto);
}

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

  return texto.split("\n").filter(linha => !linha.trim().startsWith(">")).join("\n");
}

function remetenteAutorizado(remetente) {
  if (CONFIG.APROVADORES.length === 0) return true;
  const email = extrairEmail(remetente);
  return CONFIG.APROVADORES.some(aprovador => aprovador.trim().toLowerCase() === email);
}

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

// Aceita número ou texto no formato "R$ 1.500,00"; se não converter, grava como veio.
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

function extrairNomeRemetente(remetente) {
  if (typeof remetente !== "string") return remetente;
  const match = remetente.match(/^(.*?)\s*<.*>$/);
  const nome = match ? match[1].replace(/^"|"$/g, "").trim() : remetente.trim();
  return nome || remetente.trim();
}

function extrairEmail(remetente) {
  const texto = String(remetente || "");
  const match = texto.match(/<([^>]+)>/);
  return (match ? match[1] : texto).trim().toLowerCase();
}

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

function removerPrefixoAntesDoPipe(texto) {
  if (typeof texto !== "string") return texto;
  const indice = texto.indexOf("|");
  if (indice === -1) return texto.trim();
  return texto.substring(indice + 1).trim();
}

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

function converterXlsmParaGoogleSheets(blobAnexo) {
  const fileMetadata = {
    name: "TEMP_" + blobAnexo.getName(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  const file = Drive.Files.create(fileMetadata, blobAnexo.copyBlob());
  return file.id;
}

async function gerarDocumentoConsolidado(thread, messages, assunto) {
  const ultimaMensagem = messages[messages.length - 1];
  const nomeBase = montarNomeArquivo(assunto);

  const documentos = ordenarDocumentos(obterAnexosPdfUnicos(messages));
  documentos.push(gerarPdfDoEmail(ultimaMensagem));

  const pdfFinal = await juntarPdfs(documentos);
  const pasta = obterPastaDestino();
  const arquivoPdf = pasta.createFile(pdfFinal.setName(nomeBase + ".pdf"));

  // Cópia exata do e-mail de aprovação, com cabeçalhos e assinatura digital (DKIM) do remetente.
  pasta.createFile(Utilities.newBlob(ultimaMensagem.getRawContent(), "message/rfc822", nomeBase + ".eml"));

  return arquivoPdf.getUrl();
}

// A ordenação é estável: não identificados ficam no fim, na ordem original.
function ordenarDocumentos(blobs) {
  if (blobs.length < 2) return blobs;
  const posicoes = new Map(blobs.map(blob => [blob, posicaoPorConteudo(blob)]));
  return blobs.slice().sort((a, b) => posicoes.get(a) - posicoes.get(b));
}

function posicaoPorConteudo(blob) {
  const naoIdentificado = CONFIG.ORDEM_DOCUMENTOS.length;
  let texto;
  try {
    texto = extrairTextoPdf(blob);
  } catch (erro) {
    Logger.log(`[AVISO] Não foi possível ler o texto de "${blob.getName()}"; mantido na ordem original. ${erro}`);
    return naoIdentificado;
  }

  let melhor = { indice: naoIdentificado, marcas: 0 };
  CONFIG.ORDEM_DOCUMENTOS.forEach((tipo, indice) => {
    const marcas = tipo.marcas.filter(marca => marca.test(texto)).length;
    if (marcas >= CONFIG.MIN_MARCAS && marcas > melhor.marcas) melhor = { indice, marcas };
  });

  const nomeTipo = melhor.indice < naoIdentificado ? CONFIG.ORDEM_DOCUMENTOS[melhor.indice].tipo : "não identificado";
  Logger.log(`[DEBUG] "${blob.getName()}": ${nomeTipo} (${melhor.marcas} marca(s)).`);
  return melhor.indice;
}

// A conversão para Google Docs faz OCR, então funciona também com PDF escaneado.
function extrairTextoPdf(blob) {
  let tempFileId = null;
  try {
    tempFileId = Drive.Files.create(
      { name: "TEMP_" + blob.getName(), mimeType: MimeType.GOOGLE_DOCS },
      blob.copyBlob(),
      { ocrLanguage: "pt" }
    ).id;
    return DocumentApp.openById(tempFileId).getBody().getText();
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

function obterAnexosPdfUnicos(messages) {
  const vistos = new Set();
  return messages
    .flatMap(msg => msg.getAttachments())
    .filter(anexo => anexo.getName().toLowerCase().endsWith(".pdf") || anexo.getContentType() === MimeType.PDF)
    .filter(anexo => {
      const chave = chaveAnexo(anexo);
      if (vistos.has(chave)) return false;
      vistos.add(chave);
      return true;
    });
}

// Imagens são incorporadas em base64 porque o conversor não lê cid: nem baixa URLs.
function gerarPdfDoEmail(msg) {
  const fontes = (msg.getBody().match(/<img\b[^>]*?\bsrc\s*=\s*["'][^"']{0,80}/gi) || [])
    .map(tag => tag.replace(/^.*src\s*=\s*["']/i, ""));
  if (fontes.length) Logger.log(`[DEBUG] Imagens no e-mail ${msg.getId()}: ${fontes.join(" | ")}`);
  const corpo = incorporarImagensExternas(incorporarImagensEmbutidas(msg.getBody(), msg.getId()), { externas: 0 });

  const html = `<html><head><meta charset="UTF-8"></head><body>
    <div style="font-family:Arial,sans-serif;font-size:12px;border-bottom:1px solid #999;padding-bottom:8px;margin-bottom:12px">
      <b>De:</b> ${escaparHtml(msg.getFrom())}<br>
      <b>Para:</b> ${escaparHtml(msg.getTo())}<br>
      ${msg.getCc() ? `<b>Cc:</b> ${escaparHtml(msg.getCc())}<br>` : ""}
      <b>Data:</b> ${Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")}<br>
      <b>Assunto:</b> ${escaparHtml(msg.getSubject())}
    </div>
    <div>${corpo}</div>
  </body></html>`;
  return Utilities.newBlob(html, MimeType.HTML, "email.html").getAs(MimeType.PDF);
}

// Usa a Gmail API porque o GmailApp não informa o Content-ID das imagens.
function incorporarImagensEmbutidas(html, messageId) {
  if (!/cid:/i.test(html)) return html;
  try {
    const imagens = {};
    const percorrer = parte => {
      const cabecalho = (parte.headers || []).find(h => /^(content-id|x-attachment-id)$/i.test(h.name));
      if (cabecalho && /^image\//i.test(parte.mimeType)) {
        const dados = parte.body.data ||
          Gmail.Users.Messages.Attachments.get("me", messageId, parte.body.attachmentId).data;
        // O serviço avançado devolve os bytes já decodificados (array); a API REST devolve base64url.
        const base64 = typeof dados === "string"
          ? dados.replace(/-/g, "+").replace(/_/g, "/")
          : Utilities.base64Encode(dados);
        imagens[cabecalho.value.replace(/[<>]/g, "").trim()] = `data:${parte.mimeType};base64,${base64}`;
      }
      (parte.parts || []).forEach(percorrer);
    };
    percorrer(Gmail.Users.Messages.get("me", messageId, { format: "full" }).payload);

    return html.replace(/(["'])cid:([^"']+)\1/gi, (original, aspas, cid) => {
      const dataUri = imagens[decodeURIComponent(cid).trim()];
      if (!dataUri) {
        Logger.log(`[AVISO] Imagem "cid:${cid}" não encontrada no e-mail ${messageId}. ` +
          `Content-IDs disponíveis: ${Object.keys(imagens).join(", ") || "nenhum"}`);
        return original;
      }
      return aspas + dataUri + aspas;
    });
  } catch (erro) {
    Logger.log(`[AVISO] Não foi possível incorporar as imagens embutidas do e-mail ${messageId}: ${erro}`);
    return html;
  }
}

function incorporarImagensExternas(html, contador) {
  return html.replace(/(<img\b[^>]*?\bsrc\s*=\s*)(["'])(https?:\/\/[^"']+)\2/gi, (original, prefixo, aspas, url) => {
    if (contador.externas >= CONFIG.MAX_IMAGENS_EXTERNAS) return original;
    contador.externas++;
    try {
      const resposta = UrlFetchApp.fetch(url.replace(/&amp;/g, "&"), { muteHttpExceptions: true, followRedirects: true });
      const blob = resposta.getBlob();
      const tipo = String(blob.getContentType() || "");
      if (resposta.getResponseCode() !== 200 || !/^image\//i.test(tipo) || blob.getBytes().length > CONFIG.MAX_BYTES_IMAGEM) {
        Logger.log(`[AVISO] Imagem ${url} não incorporada: HTTP ${resposta.getResponseCode()}, ` +
          `tipo "${tipo}", ${blob.getBytes().length} bytes.`);
        return original;
      }
      return `${prefixo}${aspas}data:${tipo};base64,${Utilities.base64Encode(blob.getBytes())}${aspas}`;
    } catch (erro) {
      Logger.log(`[AVISO] Não foi possível baixar a imagem ${url}: ${erro}`);
      return original;
    }
  });
}

async function juntarPdfs(blobs) {
  const PDFLib = carregarPdfLib();
  const pdfFinal = await PDFLib.PDFDocument.create();
  for (const blob of blobs) {
    let doc;
    try {
      doc = await PDFLib.PDFDocument.load(new Uint8Array(blob.getBytes()), { ignoreEncryption: true });
    } catch (erro) {
      throw new Error(`Não foi possível ler o PDF "${blob.getName()}": ${erro}`);
    }
    const paginas = await pdfFinal.copyPages(doc, doc.getPageIndices());
    paginas.forEach(pagina => pdfFinal.addPage(pagina));
  }
  const bytes = await pdfFinal.save();
  return Utilities.newBlob(Array.from(new Int8Array(bytes.buffer, bytes.byteOffset, bytes.length)), MimeType.PDF);
}

// O Apps Script não tem setTimeout, que a pdf-lib usa entre etapas; executar na hora basta.
let pdfLibCarregada = null;
function carregarPdfLib() {
  if (pdfLibCarregada) return pdfLibCarregada;
  if (typeof globalThis.setTimeout !== "function") {
    globalThis.setTimeout = (funcao, ms) => {
      if (ms) Utilities.sleep(ms);
      funcao();
      return 0;
    };
  }
  const codigo = UrlFetchApp.fetch(CONFIG.PDF_LIB_URL).getContentText();
  (0, eval)(codigo); // eval indireto: roda no escopo global e define globalThis.PDFLib
  if (!globalThis.PDFLib) throw new Error("Falha ao carregar a pdf-lib.");
  pdfLibCarregada = globalThis.PDFLib;
  return pdfLibCarregada;
}

function obterPastaDestino() {
  if (CONFIG.PASTA_DESTINO_ID) return DriveApp.getFolderById(CONFIG.PASTA_DESTINO_ID);
  const pastas = DriveApp.getRootFolder().getFoldersByName(CONFIG.NOME_PASTA_PADRAO);
  return pastas.hasNext() ? pastas.next() : DriveApp.getRootFolder().createFolder(CONFIG.NOME_PASTA_PADRAO);
}

function montarNomeArquivo(assunto) {
  const data = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const assuntoLimpo = String(assunto || "sem assunto")
    .replace(/^((re|res|fw|fwd|enc|tr)\s*:\s*)+/i, "")
    .replace(/[\\\/:*?"<>|\r\n]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .substring(0, 120);
  return `${CONFIG.PREFIXO_ARQUIVO} - ${assuntoLimpo} - ${data}`;
}

function escaparHtml(texto) {
  return String(texto || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
