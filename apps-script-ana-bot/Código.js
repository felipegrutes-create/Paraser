// ================================================================
// ANA Bot — Atendente IA WhatsApp do Instituto Paraser
// Webhook Z-API -> memória (Sheets) -> Claude API -> resposta WhatsApp
// Fase 1 / Semana 1-2: fundação (sem pré-agendamento Feegow ainda).
//
// Tokens NUNCA no código — tudo via PropertiesService.
// Configurar uma vez em Project Settings > Script Properties:
//   ANTHROPIC_KEY, ZAPI_INSTANCE_ID, ZAPI_TOKEN, ZAPI_CLIENT_TOKEN, SLACK_TOKEN
// (SPREADSHEET_ID é criado automático no primeiro uso)
// ================================================================

var _P = PropertiesService.getScriptProperties();

const ANTHROPIC_KEY     = _P.getProperty('ANTHROPIC_KEY');
const ANTHROPIC_MODEL   = 'claude-sonnet-4-6';
const ZAPI_INSTANCE_ID  = _P.getProperty('ZAPI_INSTANCE_ID');
const ZAPI_TOKEN        = _P.getProperty('ZAPI_TOKEN');
const ZAPI_CLIENT_TOKEN = _P.getProperty('ZAPI_CLIENT_TOKEN');
const SLACK_TOKEN       = _P.getProperty('SLACK_TOKEN');
const SLACK_CHANNEL     = 'atendimento';

const MEMORIA_MAX = 10; // últimas N mensagens guardadas por telefone

// ----------------------------------------------------------------
// Persona / cérebro da ANA (system prompt — da spec aprovada)
// ----------------------------------------------------------------
const SYSTEM_PROMPT =
  'Você é a ANA, atendente do Instituto Paraser (clínica de fertilidade em Botafogo, RJ).\n\n' +
  'PERSONA:\n' +
  '- Acolhedora, calorosa, mas profissional.\n' +
  '- Trata por "você" (nunca senhora).\n' +
  '- Usa 💜 com parcimônia (1 por conversa, no máximo).\n' +
  '- Frases curtas. Sem listas longas no WhatsApp.\n\n' +
  'VOCÊ PODE:\n' +
  '- Responder dúvidas sobre tratamentos, preços, médicos, processos\n' +
  '- Qualificar perfil (tempo tentando, idade, etc)\n' +
  '- Contar histórias / educar\n\n' +
  'VOCÊ NÃO PODE:\n' +
  '- Dar diagnóstico médico\n' +
  '- Confirmar agendamento sozinha\n' +
  '- Negociar preço (sempre passa pra humano)\n' +
  '- Falar sobre paciente existente\n' +
  '- Acessar prontuário/laudo\n\n' +
  'QUANDO PERGUNTAREM SE VOCÊ É HUMANA OU IA: admita — "aqui é a Ana, assistente ' +
  'da Paraser, e respondo com IA. Mas qualquer momento a equipe humana entra."\n\n' +
  'ENDEREÇO: Rua Prof. Álvaro Rodrigues, 352, 10º andar, Botafogo (perto do metrô Botafogo, saída E).';

// Resposta de reserva enquanto a chave Anthropic não está configurada.
// Permite testar o ciclo WhatsApp ponta-a-ponta antes de plugar o cérebro.
const RESPOSTA_RESERVA =
  'Oi! 💜 Aqui é a Ana, do Instituto Paraser. Recebi sua mensagem e em ' +
  'instantes a equipe te responde. Obrigada pelo contato!';

// Palavras-gatilho de escalação imediata pra humano (red flags).
// Lista mínima por segurança — análise por IA vem na Fase 2.
const RED_FLAGS = [
  'não aguento mais', 'nao aguento mais', 'não vejo saída', 'nao vejo saida',
  'me matar', 'suicíd', 'suicid', 'sangramento intenso', 'sangrando muito',
  'dor forte', 'febre alta', 'ohss', 'vou denunciar', 'estou indignada'
];

// ================================================================
// WEBHOOK — Z-API chama esta URL a cada mensagem recebida
// ================================================================
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    // Ignora mensagens enviadas por nós mesmos e eventos sem texto
    if (body.fromMe === true) return _ok();
    var phone = body.phone;
    var texto = body.text && body.text.message ? String(body.text.message).trim() : '';
    if (!phone || !texto) return _ok();

    var nome = body.senderName || body.chatName || '';

    processarMensagem(phone, nome, texto);
    return _ok();
  } catch (err) {
    log('doPost', 'erro: ' + err.message);
    return _ok(); // sempre 200 pro Z-API não reenviar em loop
  }
}

function doGet() {
  // Healthcheck simples
  return ContentService.createTextOutput('ANA bot online').setMimeType(ContentService.MimeType.TEXT);
}

// ================================================================
// Núcleo: recebe -> memória -> red flag -> Claude -> responde
// ================================================================
function processarMensagem(phone, nome, texto) {
  salvarMensagem(phone, 'user', texto);

  // 1) Red flag? Para de responder, marca prioridade, chama humano.
  if (ehRedFlag(texto)) {
    sendWhatsApp(phone,
      'Entendi. Vou pedir pra uma pessoa da nossa equipe falar com você agora, tá? 💜');
    notificarSlack('🚨 *RED FLAG — atenção humana AGORA*\n' +
      'Telefone: ' + phone + (nome ? ' (' + nome + ')' : '') + '\n' +
      'Mensagem: "' + texto + '"');
    log('redflag', phone + ' :: ' + texto);
    return;
  }

  // 2) Resposta da ANA (Claude se tiver chave, senão reserva)
  var resposta;
  if (ANTHROPIC_KEY) {
    try {
      resposta = chamarClaude(phone);
    } catch (err) {
      log('claude', 'erro: ' + err.message);
      resposta = RESPOSTA_RESERVA;
      notificarSlack('⚠️ ANA caiu pra resposta de reserva (erro Claude): ' + err.message);
    }
  } else {
    resposta = RESPOSTA_RESERVA;
  }

  // 3) Envia e guarda
  sendWhatsApp(phone, resposta);
  salvarMensagem(phone, 'assistant', resposta);
}

function ehRedFlag(texto) {
  var t = texto.toLowerCase();
  for (var i = 0; i < RED_FLAGS.length; i++) {
    if (t.indexOf(RED_FLAGS[i]) !== -1) return true;
  }
  return false;
}

// ================================================================
// Claude API — chamada HTTP direta (Apps Script não tem SDK)
// ================================================================
function chamarClaude(phone) {
  var historico = carregarHistorico(phone);   // [{role, content}, ...]
  var system    = SYSTEM_PROMPT + '\n\n' + montarFAQ();

  var payload = {
    model:      ANTHROPIC_MODEL,
    max_tokens: 1024,
    system:     system,
    messages:   historico
  };

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:      'post',
    contentType: 'application/json',
    headers: {
      'x-api-key':         ANTHROPIC_KEY,
      'anthropic-version': '2023-06-01'
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Anthropic HTTP ' + code + ': ' + resp.getContentText().substring(0, 300));
  }

  var data = JSON.parse(resp.getContentText());
  // Concatena os blocos de texto da resposta
  var texto = '';
  if (data.content) {
    for (var i = 0; i < data.content.length; i++) {
      if (data.content[i].type === 'text') texto += data.content[i].text;
    }
  }
  return texto.trim() || RESPOSTA_RESERVA;
}

// Monta o bloco de FAQ pra injetar no system prompt
function montarFAQ() {
  try {
    var sh = planilha().getSheetByName('ANA_FAQ');
    if (!sh || sh.getLastRow() < 2) return '';
    var dados = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    var linhas = [];
    for (var i = 0; i < dados.length; i++) {
      if (dados[i][0]) linhas.push('P: ' + dados[i][0] + '\nR: ' + dados[i][1]);
    }
    if (!linhas.length) return '';
    return 'BASE DE CONHECIMENTO (FAQ):\n' + linhas.join('\n\n');
  } catch (err) {
    log('montarFAQ', 'erro: ' + err.message);
    return '';
  }
}

// ================================================================
// Memória por telefone — aba ANA_Conversas
// ================================================================
function salvarMensagem(phone, role, content) {
  try {
    var sh = planilha().getSheetByName('ANA_Conversas');
    sh.appendRow([new Date(), phone, role, content]);
  } catch (err) {
    log('salvarMensagem', 'erro: ' + err.message);
  }
}

// Retorna as últimas MEMORIA_MAX mensagens deste telefone como [{role, content}]
function carregarHistorico(phone) {
  var sh = planilha().getSheetByName('ANA_Conversas');
  if (!sh || sh.getLastRow() < 2) return [];
  var dados = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  var msgs = [];
  for (var i = 0; i < dados.length; i++) {
    if (String(dados[i][1]) === String(phone)) {
      msgs.push({ role: dados[i][2], content: String(dados[i][3]) });
    }
  }
  // só as últimas N
  if (msgs.length > MEMORIA_MAX) msgs = msgs.slice(msgs.length - MEMORIA_MAX);
  // a API exige começar por 'user' — corta prefixo de assistant se houver
  while (msgs.length && msgs[0].role !== 'user') msgs.shift();
  return msgs;
}

// ================================================================
// Z-API — envio de mensagem (mesmo padrão do script Confirmações)
// ================================================================
function sendWhatsApp(phone, message) {
  var url = 'https://api.z-api.io/instances/' + ZAPI_INSTANCE_ID + '/token/' + ZAPI_TOKEN + '/send-text';
  var headers = {};
  if (ZAPI_CLIENT_TOKEN) headers['Client-Token'] = ZAPI_CLIENT_TOKEN;
  var resp = UrlFetchApp.fetch(url, {
    method:             'post',
    contentType:        'application/json',
    headers:            headers,
    payload:            JSON.stringify({ phone: phone, message: message }),
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    log('sendWhatsApp', 'Z-API HTTP ' + code + ': ' + resp.getContentText().substring(0, 200));
  }
}

// ================================================================
// Slack — notifica a equipe (mesmo canal das confirmações)
// ================================================================
function notificarSlack(texto) {
  try {
    if (!SLACK_TOKEN) return;
    var channelId = slackGetChannelId(SLACK_CHANNEL);
    if (!channelId) return;
    UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method:      'post',
      contentType: 'application/json; charset=utf-8',
      headers:     { Authorization: 'Bearer ' + SLACK_TOKEN },
      payload:     JSON.stringify({ channel: channelId, text: texto }),
      muteHttpExceptions: true
    });
  } catch (err) {
    log('notificarSlack', 'erro: ' + err.message);
  }
}

function slackGetChannelId(channelName) {
  try {
    var url = 'https://slack.com/api/conversations.list?limit=200&exclude_archived=true&types=public_channel,private_channel';
    var data = JSON.parse(UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + SLACK_TOKEN }, muteHttpExceptions: true
    }).getContentText());
    if (!data.ok || !data.channels) return null;
    for (var i = 0; i < data.channels.length; i++) {
      if (data.channels[i].name === channelName) return data.channels[i].id;
    }
    return null;
  } catch (err) {
    return null;
  }
}

// ================================================================
// Planilha ANA — criada/encontrada automaticamente (lazy setup)
// ================================================================
function planilha() {
  var id = _P.getProperty('SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (err) { /* recria abaixo */ }
  }
  return criarPlanilhaAna_();
}

function criarPlanilhaAna_() {
  var ss = SpreadsheetApp.create('ANA Paraser — Dados');
  _P.setProperty('SPREADSHEET_ID', ss.getId());

  // Abas da arquitetura (spec §4). Cabeçalhos nas que o código usa já.
  _aba(ss, 'ANA_Config',           ['chave', 'valor']);
  _aba(ss, 'ANA_FAQ',              ['pergunta', 'resposta']);
  _aba(ss, 'ANA_Conversas',        ['timestamp', 'telefone', 'role', 'mensagem']);
  _aba(ss, 'ANA_Quiz_Submissoes',  ['timestamp', 'telefone', 'nome', 'perfil']);
  _aba(ss, 'ANA_Pre_Agendamentos', ['timestamp', 'telefone', 'nome', 'proposta', 'status']);
  _aba(ss, 'ANA_Logs',             ['timestamp', 'origem', 'detalhe']);

  // remove a "Sheet1"/"Página1" default
  var def = ss.getSheets()[0];
  if (def && def.getName().match(/^(Sheet1|Página1|Sheet|Página)/)) ss.deleteSheet(def);

  // Seed de FAQ (a recepção cura depois pelo Dashboard)
  var faq = ss.getSheetByName('ANA_FAQ');
  faq.appendRow(['Onde fica a clínica?', 'No Rio, em Botafogo: Rua Prof. Álvaro Rodrigues, 352, 10º andar, pertinho do metrô Botafogo (saída E).']);
  faq.appendRow(['Vocês atendem por plano de saúde?', 'A primeira consulta é particular. Alguns exames e procedimentos podem ter cobertura dependendo do seu plano — a equipe te explica direitinho.']);
  faq.appendRow(['Quais tratamentos vocês fazem?', 'Cuidamos de toda a jornada da fertilidade: investigação, indução, inseminação, FIV e preservação. Na consulta o médico monta o melhor caminho pra você.']);

  return ss;
}

function _aba(ss, nome, cabecalho) {
  var sh = ss.insertSheet(nome);
  sh.appendRow(cabecalho);
  sh.getRange(1, 1, 1, cabecalho.length).setFontWeight('bold');
  return sh;
}

// ================================================================
// Log interno (aba ANA_Logs) — best-effort
// ================================================================
function log(origem, detalhe) {
  try {
    planilha().getSheetByName('ANA_Logs').appendRow([new Date(), origem, detalhe]);
  } catch (err) {
    Logger.log(origem + ': ' + detalhe);
  }
}

function _ok() {
  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}
