// ================================================================
// Paraser — Meta Conversions API (CAPI) — Eventos do Feegow → Meta
// ================================================================
// O QUE FAZ:
//   Lê agendamentos novos no Feegow e dispara eventos pro Meta
//   (Schedule = consulta marcada, CompleteRegistration = consulta realizada).
//   Isso ensina o Meta a otimizar pelos leads que viram PACIENTE real,
//   não só os que preenchem formulário.
//
// FLUXO:
//   1. Cron diário roda dispatchMetaCapi()
//   2. Busca agendamentos das últimas 48h via /appoints/search
//   3. Pra cada novo: pega email/celular do paciente via /patient/search
//   4. Hasheia (SHA-256) e POST pra Graph API
//   5. Loga tudo em "MetaCapi_Log" pra evitar duplicar
//
// SETUP (Felipe — só uma vez):
//   1. Business Manager → Conjuntos de dados → "Leads Paraser" → Configurações
//      → Tokens de acesso → Gerar Token (com permissões: ads_management,
//      business_management). Copia o token.
//   2. Execute setupMetaCapi() no editor, cole o token quando pedir.
//   3. Execute enviarEventoTeste() pra mandar 1 evento de teste.
//   4. Confere no Meta → Eventos de Teste se chegou.
//   5. Execute criarTriggerDiarioMetaCapi() pra ativar o cron 2x ao dia.
// ================================================================

// ---- CONFIG via Script Properties ----
var _MCP = PropertiesService.getScriptProperties();

const MCP_API_VERSION    = 'v22.0';
const MCP_DATASET_ID     = '920108941023871';  // "Leads Paraser" no BM Instituto ParaSer
const MCP_FEEGOW_BASE    = 'https://api.feegow.com/v1/api';

// ↓ Lidos do Properties (rodar setupMetaCapi pra preencher)
const MCP_ACCESS_TOKEN   = _MCP.getProperty('META_CAPI_TOKEN');
const MCP_FEEGOW_TOKEN   = _MCP.getProperty('FEEGOW_TOKEN');
const MCP_SPREADSHEET_ID = _MCP.getProperty('SPREADSHEET_ID');
const MCP_TEST_CODE      = _MCP.getProperty('META_CAPI_TEST_CODE') || '';  // ex: TEST12345 — só pra modo teste

const MCP_LOG_SHEET      = 'MetaCapi_Log';

// status_id Feegow:
//   1 = Marcado não confirmado
//   3 = Atendido (consulta realizada)
//   6 = Não compareceu
//   7 = Marcado confirmado
//   11 = Desmarcado pelo paciente
//   15 = Remarcado
//   22 = Cancelado pelo profissional
const MCP_STATUS_SCHEDULED = [1, 7];   // dispara Schedule
const MCP_STATUS_ATTENDED  = [3];      // dispara CompleteRegistration


// ================================================================
// HASHING (SHA-256 hex lowercase — exigido pelo Meta)
// ================================================================
function mcpSha256(str) {
  if (!str) return null;
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str);
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var v = bytes[i];
    if (v < 0) v += 256;
    var h = v.toString(16);
    hex += (h.length === 1 ? '0' + h : h);
  }
  return hex;
}

function mcpStripAccents(s) {
  return (s || '').normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function mcpNormEmail(email) {
  if (!email) return null;
  var e = String(email).trim().toLowerCase();
  if (!e.match(/^[^@\s]+@[^@\s]+\.[^@\s]+$/)) return null;
  return mcpSha256(e);
}

function mcpNormPhone(phone) {
  if (!phone) return null;
  // Só dígitos
  var p = String(phone).replace(/\D/g, '');
  if (!p) return null;
  // Adiciona código BR (55) se vier só DDD+número (10-11 dígitos)
  if (p.length === 10 || p.length === 11) p = '55' + p;
  // Tira "0" extra que às vezes vem antes do DDD (ex: 021987654321 → 55021987... corrige)
  // já tratado acima na lógica de length
  if (p.length < 12 || p.length > 13) return null;  // formato BR inválido
  return mcpSha256(p);
}

function mcpNormName(name) {
  if (!name) return null;
  var n = mcpStripAccents(String(name)).trim().toLowerCase();
  n = n.replace(/[^a-z\s]/g, '').replace(/\s+/g, ' ');
  if (!n) return null;
  return mcpSha256(n);
}

function mcpSplitName(fullName) {
  var parts = String(fullName || '').trim().split(/\s+/);
  if (!parts.length) return { fn: null, ln: null };
  var first = parts[0];
  var last  = parts.length > 1 ? parts[parts.length - 1] : '';
  return {
    fn: mcpNormName(first),
    ln: last ? mcpNormName(last) : null
  };
}


// ================================================================
// FEEGOW — Buscar agendamentos das últimas N horas
// ================================================================
function mcpFmtData(d) {
  return Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd');
}

function mcpGetAgendamentos(daysBack) {
  daysBack = daysBack || 2;
  var hoje = new Date();
  var inicio = new Date(); inicio.setDate(inicio.getDate() - daysBack);
  var ds = mcpFmtData(inicio);
  var de = mcpFmtData(hoje);

  var url = MCP_FEEGOW_BASE + '/appoints/search?data_start=' + ds + '&data_end=' + de;
  var resp = UrlFetchApp.fetch(url, {
    headers: { 'x-access-token': MCP_FEEGOW_TOKEN },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() >= 400) {
    Logger.log('Feegow appoints/search HTTP ' + resp.getResponseCode() + ': ' + resp.getContentText().slice(0, 400));
    return [];
  }
  var json = JSON.parse(resp.getContentText());
  return Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);
}

function mcpGetPatientData(pacienteId) {
  if (!pacienteId) return null;
  var url = MCP_FEEGOW_BASE + '/patient/search?paciente_id=' + pacienteId;
  var resp = UrlFetchApp.fetch(url, {
    headers: { 'x-access-token': MCP_FEEGOW_TOKEN },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() >= 400) return null;
  var json = JSON.parse(resp.getContentText());
  var p = json.content || json;
  if (!p) return null;
  var cel = p.celular || (p.celulares && p.celulares[0]) || p.telefone || p.telefone_celular || '';
  var em  = p.email || (p.emails && p.emails[0]) || '';
  return {
    nome:    p.nome || p.paciente_nome || p.name || '',
    email:   em,
    celular: cel,
    cidade:  p.cidade || '',
    estado:  p.estado || p.uf || ''
  };
}


// ================================================================
// CAPI — Monta e envia evento
// ================================================================
function mcpBuildUserData(patient) {
  var ud = {};
  var em = mcpNormEmail(patient.email);
  var ph = mcpNormPhone(patient.celular);
  var nm = mcpSplitName(patient.nome);
  if (em) ud.em = [em];
  if (ph) ud.ph = [ph];
  if (nm.fn) ud.fn = [nm.fn];
  if (nm.ln) ud.ln = [nm.ln];
  if (patient.cidade) ud.ct = [mcpSha256(mcpStripAccents(patient.cidade).toLowerCase().replace(/\s/g, ''))];
  if (patient.estado) ud.st = [mcpSha256(String(patient.estado).toLowerCase().trim())];
  ud.country = [mcpSha256('br')];
  return ud;
}

function mcpSendEvent(eventName, ag, patient, customData) {
  if (!MCP_ACCESS_TOKEN) throw new Error('META_CAPI_TOKEN não configurado. Rode setupMetaCapi().');

  var userData = mcpBuildUserData(patient);
  // Evento só vai se tiver pelo menos email OU phone (Meta exige)
  if (!userData.em && !userData.ph) {
    return { skipped: true, reason: 'sem_email_e_telefone' };
  }

  var eventTime = Math.floor(Date.now() / 1000);
  var eventId   = 'feegow_appt_' + ag.agendamento_id + '_' + eventName.toLowerCase();

  var event = {
    event_name:    eventName,
    event_time:    eventTime,
    event_id:      eventId,
    action_source: 'system_generated',
    user_data:     userData
  };
  if (customData) event.custom_data = customData;

  var payload = { data: [event] };
  if (MCP_TEST_CODE) payload.test_event_code = MCP_TEST_CODE;

  var url = 'https://graph.facebook.com/' + MCP_API_VERSION + '/'
          + MCP_DATASET_ID + '/events?access_token=' + encodeURIComponent(MCP_ACCESS_TOKEN);

  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  var body = resp.getContentText();
  return {
    ok: code >= 200 && code < 300,
    code: code,
    body: body,
    eventId: eventId
  };
}


// ================================================================
// LOG — Sheet "MetaCapi_Log" (anti-duplicação)
// ================================================================
function mcpGetLogSheet() {
  if (!MCP_SPREADSHEET_ID) throw new Error('SPREADSHEET_ID não configurado.');
  var ss = SpreadsheetApp.openById(MCP_SPREADSHEET_ID);
  var sh = ss.getSheetByName(MCP_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MCP_LOG_SHEET);
    sh.appendRow([
      'Timestamp', 'EventName', 'EventID', 'AgendamentoID',
      'PacienteID', 'PacienteNome', 'Status', 'HttpCode', 'Response'
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function mcpAlreadySent(eventId) {
  var sh = mcpGetLogSheet();
  var last = sh.getLastRow();
  if (last <= 1) return false;
  // Coluna C = EventID, lê só ela
  var ids = sh.getRange(2, 3, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === eventId) return true;
  }
  return false;
}

function mcpLogEvent(eventName, eventId, ag, patient, status, httpCode, responseText) {
  var sh = mcpGetLogSheet();
  sh.appendRow([
    new Date(),
    eventName,
    eventId,
    ag.agendamento_id || '',
    ag.paciente_id    || '',
    (patient && patient.nome) || '',
    status,
    httpCode || '',
    (responseText || '').slice(0, 500)
  ]);
}


// ================================================================
// MAIN — dispatch (rode pelo trigger 2x/dia)
// ================================================================
function dispatchMetaCapi() {
  if (!MCP_ACCESS_TOKEN) {
    Logger.log('META_CAPI_TOKEN não configurado. Pulando.');
    return;
  }
  var ags = mcpGetAgendamentos(2);
  Logger.log('Agendamentos últimos 2 dias: ' + ags.length);

  var enviados = 0, pulados = 0, erros = 0;

  ags.forEach(function(ag) {
    try {
      var sid = Number(ag.status_id);
      var eventName = null;
      if (MCP_STATUS_SCHEDULED.indexOf(sid) >= 0)      eventName = 'Schedule';
      else if (MCP_STATUS_ATTENDED.indexOf(sid) >= 0)  eventName = 'CompleteRegistration';
      else return;  // status não relevante (cancelado, remarcado, etc)

      var eventId = 'feegow_appt_' + ag.agendamento_id + '_' + eventName.toLowerCase();
      if (mcpAlreadySent(eventId)) { pulados++; return; }

      var patient = mcpGetPatientData(ag.paciente_id);
      if (!patient) {
        mcpLogEvent(eventName, eventId, ag, null, 'PACIENTE_NAO_ENCONTRADO', '', '');
        erros++; return;
      }

      // Pausa pequena pra não estourar rate limit
      Utilities.sleep(250);

      var result = mcpSendEvent(eventName, ag, patient, null);
      if (result.skipped) {
        mcpLogEvent(eventName, eventId, ag, patient, 'SKIPPED_' + result.reason, '', '');
        pulados++;
      } else if (result.ok) {
        mcpLogEvent(eventName, eventId, ag, patient, 'ENVIADO', result.code, result.body);
        enviados++;
      } else {
        mcpLogEvent(eventName, eventId, ag, patient, 'ERRO_HTTP', result.code, result.body);
        erros++;
      }
    } catch (e) {
      Logger.log('Erro ag ' + ag.agendamento_id + ': ' + e.message);
      mcpLogEvent('ERRO', '-', ag, null, 'EXCECAO', '', e.message);
      erros++;
    }
  });

  Logger.log('META CAPI dispatch — enviados:' + enviados + ' pulados:' + pulados + ' erros:' + erros);
}


// ================================================================
// SETUP — rodar 1× pra configurar
// ================================================================
function setupMetaCapi() {
  // Standalone script — usa Browser.inputBox (não SpreadsheetApp.getUi)
  var token = Browser.inputBox(
    'Cole o Access Token gerado em Business Manager → Conjuntos de Dados → Leads Paraser → Configurações → Tokens de acesso',
    Browser.Buttons.OK_CANCEL
  );
  if (!token || token === 'cancel') { Logger.log('Setup cancelado'); return; }
  _MCP.setProperty('META_CAPI_TOKEN', token.trim());
  Logger.log('META_CAPI_TOKEN salvo (len=' + token.length + ').');

  if (!_MCP.getProperty('FEEGOW_TOKEN')) {
    Logger.log('⚠️ FEEGOW_TOKEN ainda não está nas Properties — copie do script "Confirmações Agenda".');
  }
  if (!_MCP.getProperty('SPREADSHEET_ID')) {
    Logger.log('⚠️ SPREADSHEET_ID ainda não está nas Properties — copie do script "Confirmações Agenda".');
  }
  Logger.log('✅ Setup concluído. Próximo passo: enviarEventoTeste()');
}

function setarFeegowToken() {
  var t = Browser.inputBox('FEEGOW_TOKEN (copia do script Confirmações Agenda)', Browser.Buttons.OK_CANCEL);
  if (!t || t === 'cancel') return;
  _MCP.setProperty('FEEGOW_TOKEN', t.trim());
  Logger.log('FEEGOW_TOKEN salvo.');
}

function setarSpreadsheetId() {
  var t = Browser.inputBox('SPREADSHEET_ID (copia do script Confirmações Agenda)', Browser.Buttons.OK_CANCEL);
  if (!t || t === 'cancel') return;
  _MCP.setProperty('SPREADSHEET_ID', t.trim());
  Logger.log('SPREADSHEET_ID salvo.');
}

function setarTestEventCode() {
  var t = Browser.inputBox(
    'Test Event Code (Meta → Conjuntos de Dados → Eventos de Teste → Test ID).\n' +
    'Deixe em branco pra desligar modo teste.',
    Browser.Buttons.OK_CANCEL
  );
  if (t === 'cancel') return;
  _MCP.setProperty('META_CAPI_TEST_CODE', (t || '').trim());
  Logger.log('META_CAPI_TEST_CODE = "' + (t || '') + '"');
}


// ================================================================
// TESTE — envia 1 evento fake pra Meta (com test_event_code se setado)
// ================================================================
function enviarEventoTeste() {
  if (!MCP_ACCESS_TOKEN) { Logger.log('❌ Sem token. Rode setupMetaCapi() primeiro.'); return; }
  var fakeAg = { agendamento_id: 'TEST_' + Date.now(), paciente_id: 0 };
  var fakePatient = {
    nome: 'Maria da Silva Teste',
    email: 'maria.teste+capi@paraser.com.br',
    celular: '21998765432',
    cidade: 'Rio de Janeiro',
    estado: 'RJ'
  };
  var result = mcpSendEvent('Schedule', fakeAg, fakePatient, null);
  Logger.log('Teste — code:' + result.code + ' body:' + result.body);
  if (result.ok) {
    Logger.log('✅ Evento de teste enviado. Vai aparecer em Meta → Conjuntos de Dados → Eventos de Teste em ~30s.');
    if (!MCP_TEST_CODE) {
      Logger.log('⚠️ Sem TEST_CODE setado — evento foi pra produção. Pra modo teste, rode setarTestEventCode() com o ID do Test ID.');
    }
  } else {
    Logger.log('❌ Falhou. Veja o body acima — provavelmente token inválido ou sem permissão.');
  }
}


// ================================================================
// DEBUG — listar agendamentos das últimas 48h sem enviar nada
// ================================================================
function debugListarAgendamentos() {
  if (!MCP_FEEGOW_TOKEN) { Logger.log('❌ Sem FEEGOW_TOKEN. Rode setarFeegowToken().'); return; }
  var ags = mcpGetAgendamentos(2);
  Logger.log('Total nas últimas 48h: ' + ags.length);
  ags.slice(0, 10).forEach(function(ag) {
    Logger.log(JSON.stringify({
      id: ag.agendamento_id,
      pac: ag.paciente_id,
      prof: ag.profissional_id,
      proc: ag.procedimento_id,
      data: ag.data,
      hora: ag.horario,
      status: ag.status_id
    }));
  });
}


// ================================================================
// TRIGGERS
// ================================================================
function criarTriggerDiarioMetaCapi() {
  // limpa triggers antigas
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dispatchMetaCapi') ScriptApp.deleteTrigger(t);
  });
  // Roda 2× ao dia: 12h e 22h (depois das principais ondas de marcação)
  ScriptApp.newTrigger('dispatchMetaCapi').timeBased().atHour(12).everyDays(1).create();
  ScriptApp.newTrigger('dispatchMetaCapi').timeBased().atHour(22).everyDays(1).create();
  Logger.log('✅ Trigger criado: dispatchMetaCapi roda às 12h e 22h diariamente.');
}

function removerTriggersMetaCapi() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dispatchMetaCapi') { ScriptApp.deleteTrigger(t); n++; }
  });
  Logger.log('Removidos ' + n + ' trigger(s).');
}
