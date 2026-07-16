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
//      → Tokens de acesso → Gerar Token. Copia o token (EAA...).
//   2. No Apps Script: ícone ⚙ Configurações do projeto (menu lateral esquerdo)
//      → Propriedades do script → Adicionar propriedade. Cria estas 4:
//         META_CAPI_TOKEN       = (cole o token do passo 1)
//         FEEGOW_TOKEN          = (copia do script Confirmações Agenda)
//         SPREADSHEET_ID        = 1uthRnuWMk2A26dZ8GaMXinuvPyxY50EH8RX85NemXwg
//         META_CAPI_TEST_CODE   = (opcional: TEST_xxx do BM Eventos de Teste)
//      Salvar.
//   3. Execute verificarConfig() pra confirmar que tudo está setado.
//   4. Execute enviarEventoTeste() pra mandar 1 evento.
//   5. Confere no Meta → Eventos de Teste se chegou.
//   6. Execute criarTriggerDiarioMetaCapi() pra ativar o cron 2x ao dia.
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

// Lista de Clientes (Custom Audience) que se mantém atualizada sozinha.
// Default = "Controle Agendamentos - Cópia de Dados Diários.csv" (a que tem o lookalike de 1,2M pendurado).
// Pode trocar via Script Property META_AUDIENCE_ID.
const MCP_AUDIENCE_ID    = _MCP.getProperty('META_AUDIENCE_ID') || '120242388951550375';

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
// CUSTOM AUDIENCE — mantém a "Lista de Clientes" sempre atualizada
// ================================================================
// Por que: a lista subida por CSV é uma FOTO CONGELADA. Aqui o robô
// adiciona automaticamente cada paciente novo na mesma lista (casado
// por email/telefone), então o lookalike pendurado nela fica sempre fresco.
// Usa o token de Ads (System User, não expira) — precisa scope ads_management.
// ================================================================
function mcpAudienceRow(patient) {
  var em = mcpNormEmail(patient.email);
  var ph = mcpNormPhone(patient.celular);
  if (!em && !ph) return null;  // Meta exige pelo menos 1 chave
  var nm = mcpSplitName(patient.nome);
  var ct = patient.cidade ? mcpSha256(mcpStripAccents(patient.cidade).toLowerCase().replace(/\s/g, '')) : '';
  var st = patient.estado ? mcpSha256(String(patient.estado).toLowerCase().trim()) : '';
  // ordem = schema EMAIL,PHONE,FN,LN,CT,ST,COUNTRY
  return [ em || '', ph || '', nm.fn || '', nm.ln || '', ct, st, mcpSha256('br') ];
}

function mcpAddUsersToAudience(rows) {
  if (!rows || !rows.length) return { ok: true, skipped: true, reason: 'sem_linhas' };
  var token = _mktToken_();  // prefere META_ADS_TOKEN (System User)
  if (!token) return { ok: false, reason: 'sem_token' };
  var payload = {
    schema: ['EMAIL', 'PHONE', 'FN', 'LN', 'CT', 'ST', 'COUNTRY'],
    data: rows
  };
  var url = 'https://graph.facebook.com/' + MCP_API_VERSION + '/' + MCP_AUDIENCE_ID + '/users';
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    muteHttpExceptions: true,
    payload: {
      access_token: token,
      payload: JSON.stringify(payload)
    }
  });
  var code = resp.getResponseCode();
  var body = resp.getContentText();
  return { ok: code >= 200 && code < 300, code: code, body: body, count: rows.length };
}

// Busca agendamentos numa JANELA [fromDaysAgo, toDaysAgo] (ex: de 60 a 30 dias atrás).
function mcpGetAgendamentosWindow(fromDaysAgo, toDaysAgo) {
  var ini = new Date(); ini.setDate(ini.getDate() - fromDaysAgo);
  var fim = new Date(); fim.setDate(fim.getDate() - (toDaysAgo || 0));
  var url = MCP_FEEGOW_BASE + '/appoints/search?data_start=' + mcpFmtData(ini) + '&data_end=' + mcpFmtData(fim);
  var resp = UrlFetchApp.fetch(url, { headers: { 'x-access-token': MCP_FEEGOW_TOKEN }, muteHttpExceptions: true });
  if (resp.getResponseCode() >= 400) return [];
  var json = JSON.parse(resp.getContentText());
  return Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);
}

// Backfill eficiente: dedup por paciente_id ANTES de consultar o Feegow (corta muita chamada).
// Janela [from, to] em dias atrás. Use janelas de ~30 dias pra não estourar o limite de 6 min.
function backfillAudienceWindow(fromDaysAgo, toDaysAgo) {
  fromDaysAgo = fromDaysAgo || 30;
  toDaysAgo   = toDaysAgo   || 0;
  var ags = mcpGetAgendamentosWindow(fromDaysAgo, toDaysAgo);

  // 1) coleta paciente_ids únicos dos status relevantes (sem chamar Feegow ainda)
  var pacIds = {}, totalRelev = 0;
  ags.forEach(function(ag) {
    var sid = Number(ag.status_id);
    if (MCP_STATUS_SCHEDULED.indexOf(sid) < 0 && MCP_STATUS_ATTENDED.indexOf(sid) < 0) return;
    if (!ag.paciente_id) return;
    totalRelev++;
    pacIds[ag.paciente_id] = true;
  });
  var ids = Object.keys(pacIds);

  // 2) consulta cada paciente UMA vez e monta as linhas
  var rows = [], seen = {}, fetched = 0;
  for (var k = 0; k < ids.length; k++) {
    var patient = mcpGetPatientData(ids[k]);
    fetched++;
    if (!patient) continue;
    Utilities.sleep(80);
    var arow = mcpAudienceRow(patient);
    if (!arow) continue;
    var key = arow[0] + '|' + arow[1];
    if (seen[key]) continue;
    seen[key] = true;
    rows.push(arow);
  }

  // 3) envia em lotes
  var totalOk = 0, lastResp = '', lastCode = 0;
  for (var i = 0; i < rows.length; i += 1000) {
    var ar = mcpAddUsersToAudience(rows.slice(i, i + 1000));
    lastResp = ar.body; lastCode = ar.code;
    if (ar.ok) totalOk += Math.min(1000, rows.length - i);
    Utilities.sleep(300);
  }
  var out = { window: fromDaysAgo + '→' + toDaysAgo + ' dias atrás', relevantes: totalRelev, pacientesUnicos: ids.length, fetched: fetched, linhas: rows.length, addedOk: totalOk, code: lastCode, lastResp: (lastResp || '').slice(0, 300) };
  Logger.log('Backfill window: ' + JSON.stringify(out));
  return out;
}

// Atalho: últimos N dias (janela [N, 0]).
function backfillAudience(daysBack) {
  return backfillAudienceWindow(daysBack || 14, 0);
}

// ---- Backfill RETOMÁVEL (resiste ao limite de 6 min) ----
// Salva os paciente_ids já processados na aba "Audience_Synced".
// Rode quantas vezes precisar com a mesma janela; cada chamada avança e persiste.
function _audienceDoneSheet_() {
  var ss = SpreadsheetApp.openById(MCP_SPREADSHEET_ID);
  var sh = ss.getSheetByName('Audience_Synced');
  if (!sh) { sh = ss.insertSheet('Audience_Synced'); sh.appendRow(['paciente_id', 'ts']); sh.setFrozenRows(1); }
  return sh;
}
function _audienceDoneSet_(sh) {
  var last = sh.getLastRow();
  var set = {};
  if (last > 1) {
    var vals = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) set[String(vals[i][0])] = true;
  }
  return set;
}
function backfillResumable(fromDaysAgo, toDaysAgo, maxSeconds) {
  fromDaysAgo = fromDaysAgo || 90;
  toDaysAgo   = toDaysAgo   || 0;
  maxSeconds  = maxSeconds  || 280;
  var start = Date.now();
  var ags = mcpGetAgendamentosWindow(fromDaysAgo, toDaysAgo);
  var doneSheet = _audienceDoneSheet_();
  var done = _audienceDoneSet_(doneSheet);

  // paciente_ids únicos relevantes que ainda NÃO foram processados
  var pend = {}, relev = 0;
  ags.forEach(function(ag) {
    var sid = Number(ag.status_id);
    if (MCP_STATUS_SCHEDULED.indexOf(sid) < 0 && MCP_STATUS_ATTENDED.indexOf(sid) < 0) return;
    if (!ag.paciente_id) return;
    relev++;
    if (!done[String(ag.paciente_id)]) pend[ag.paciente_id] = true;
  });
  var ids = Object.keys(pend);

  var rows = [], doneRows = [], processed = 0, addedOk = 0, timedOut = false;
  function flush() {
    if (rows.length) {
      var ar = mcpAddUsersToAudience(rows);
      if (ar.ok) addedOk += rows.length;
      rows = [];
    }
    if (doneRows.length) {
      doneSheet.getRange(doneSheet.getLastRow() + 1, 1, doneRows.length, 2).setValues(doneRows);
      doneRows = [];
    }
  }
  for (var k = 0; k < ids.length; k++) {
    if ((Date.now() - start) / 1000 > maxSeconds) { timedOut = true; break; }
    var pid = ids[k];
    var patient = mcpGetPatientData(pid);
    processed++;
    var arow = patient ? mcpAudienceRow(patient) : null;
    if (arow) rows.push(arow);
    doneRows.push([pid, new Date()]);  // marca processado mesmo sem email/tel (não reprocessa)
    if (doneRows.length >= 200) flush();
    Utilities.sleep(40);
  }
  flush();
  var out = {
    window: fromDaysAgo + '→' + toDaysAgo + ' dias atrás',
    relevantes: relev, pendentesNoInicio: ids.length,
    processadosAgora: processed, adicionados: addedOk,
    faltam: ids.length - processed, timedOut: timedOut
  };
  Logger.log('backfillResumable: ' + JSON.stringify(out));
  return out;
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
  var audienceRows = [], audienceSeen = {};  // pacientes novas pra Lista de Clientes

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

      // Junta a paciente pra Lista de Clientes (dedup por email|telefone)
      var arow = mcpAudienceRow(patient);
      if (arow) {
        var akey = arow[0] + '|' + arow[1];
        if (!audienceSeen[akey]) { audienceSeen[akey] = true; audienceRows.push(arow); }
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

  // Flush: adiciona as pacientes novas na Lista de Clientes (mantém o lookalike fresco)
  if (audienceRows.length) {
    var ar = mcpAddUsersToAudience(audienceRows);
    Logger.log('Audience sync — ' + (ar.ok ? 'OK' : 'ERRO') + ' (' + audienceRows.length + ' pacientes) code:' + ar.code + ' body:' + (ar.body || '').slice(0, 300));
    mcpLogEvent('AUDIENCE_SYNC', 'audience_' + MCP_AUDIENCE_ID + '_' + Date.now(),
      { agendamento_id: '', paciente_id: '' },
      { nome: audienceRows.length + ' pacientes' },
      ar.ok ? 'ENVIADO' : 'ERRO_HTTP', ar.code, ar.body);
  }

  Logger.log('META CAPI dispatch — enviados:' + enviados + ' pulados:' + pulados + ' erros:' + erros);
}


// ================================================================
// SETUP — verificar se as Script Properties estão configuradas
// (As propriedades devem ser adicionadas via UI:
//  ⚙ Configurações do projeto → Propriedades do script → Adicionar propriedade)
// ================================================================
function verificarConfig() {
  var props = {
    META_CAPI_TOKEN:     _MCP.getProperty('META_CAPI_TOKEN'),
    FEEGOW_TOKEN:        _MCP.getProperty('FEEGOW_TOKEN'),
    SPREADSHEET_ID:      _MCP.getProperty('SPREADSHEET_ID'),
    META_CAPI_TEST_CODE: _MCP.getProperty('META_CAPI_TEST_CODE')
  };
  var ok = true;
  Logger.log('=== Script Properties ===');
  if (props.META_CAPI_TOKEN) {
    Logger.log('✅ META_CAPI_TOKEN: setado (len=' + props.META_CAPI_TOKEN.length + ', começa com "' + props.META_CAPI_TOKEN.slice(0, 6) + '...")');
  } else {
    Logger.log('❌ META_CAPI_TOKEN: FALTA — adicione em ⚙ Configurações do projeto → Propriedades do script');
    ok = false;
  }
  if (props.FEEGOW_TOKEN) {
    Logger.log('✅ FEEGOW_TOKEN: setado (len=' + props.FEEGOW_TOKEN.length + ')');
  } else {
    Logger.log('❌ FEEGOW_TOKEN: FALTA — copie do script "Confirmações Agenda"');
    ok = false;
  }
  if (props.SPREADSHEET_ID) {
    Logger.log('✅ SPREADSHEET_ID: ' + props.SPREADSHEET_ID);
  } else {
    Logger.log('❌ SPREADSHEET_ID: FALTA — use 1uthRnuWMk2A26dZ8GaMXinuvPyxY50EH8RX85NemXwg');
    ok = false;
  }
  if (props.META_CAPI_TEST_CODE) {
    Logger.log('🧪 META_CAPI_TEST_CODE: "' + props.META_CAPI_TEST_CODE + '" (modo teste ATIVO — eventos vão pra "Eventos de Teste", não pra produção)');
  } else {
    Logger.log('ℹ️ META_CAPI_TEST_CODE: vazio (modo PRODUÇÃO — eventos contam pra otimização)');
  }
  Logger.log(ok ? '\n✅ Tudo configurado. Próximo passo: enviarEventoTeste()' : '\n❌ Faltam propriedades. Configure e rode esta função novamente.');
  return ok;
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


// ================================================================
// MARKETING DASHBOARD — endpoints + snapshot funil + triggers
// ================================================================
// O QUE FAZ:
//   Endpoint web app pra alimentar a aba "📊 Marketing" no dashboard.
//     ?action=marketing-now    → live: gasto today/7d/this_month + campanhas ativas
//     ?action=marketing-funnel → snapshot diário: impressões → cliques → schedule → procedimento
//   gravarFunilSnapshot() roda 2× ao dia (8h + 18h), grava aba MARKETING_FUNNEL.
//
// SETUP:
//   Property META_ADS_TOKEN (se omitir, reusa META_CAPI_TOKEN — precisa scope ads_read)
//   Property META_AD_ACCOUNT_ID (default act_849546694121439 = "CA - Anunciante 3")
//   Roda setupMarketingDashboard() pra ativar trigger.
// ================================================================

const MKT_AUTH_KEY        = 'paraser2026';
const MKT_DEFAULT_ACCOUNT = 'act_849546694121439';
const MKT_FUNNEL_SHEET    = 'MARKETING_FUNNEL';

function _mktToken_() {
  return _MCP.getProperty('META_ADS_TOKEN') || _MCP.getProperty('META_CAPI_TOKEN');
}
function _mktAccount_() {
  var id = _MCP.getProperty('META_AD_ACCOUNT_ID') || MKT_DEFAULT_ACCOUNT;
  if (id.indexOf('act_') !== 0) id = 'act_' + id;
  return id;
}
function _mktApiBase_() { return 'https://graph.facebook.com/' + MCP_API_VERSION; }

function _mktFetch_(path, params) {
  var token = _mktToken_();
  if (!token) throw new Error('META_ADS_TOKEN/META_CAPI_TOKEN ausente');
  var q = ['access_token=' + encodeURIComponent(token)];
  Object.keys(params || {}).forEach(function(k) {
    q.push(encodeURIComponent(k) + '=' + encodeURIComponent(params[k]));
  });
  var url = _mktApiBase_() + path + '?' + q.join('&');
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var code = resp.getResponseCode();
  var body = resp.getContentText();
  if (code >= 400) throw new Error('Meta API ' + code + ' em ' + path + ': ' + body.substring(0, 300));
  return JSON.parse(body);
}

function _mktActionVal_(actions, type) {
  if (!actions || !actions.length) return 0;
  for (var i = 0; i < actions.length; i++) {
    if (actions[i].action_type === type) return Number(actions[i].value || 0);
  }
  return 0;
}

// "Chamaram no WhatsApp": campanhas de lead/form contam isso como conversa de
// mensagem (messaging_conversation_started), NÃO como o evento Contact do Pixel.
// Tenta os nomes de mensagem primeiro, cai pro Contact do Pixel como reserva.
function _mktContactVal_(actions) {
  return _mktActionVal_(actions, 'onsite_conversion.messaging_conversation_started_7d')
      || _mktActionVal_(actions, 'onsite_conversion.total_messaging_connection')
      || _mktActionVal_(actions, 'contact_total')
      || _mktActionVal_(actions, 'contact');
}

function _mktAccountInsights_(datePreset) {
  var r = _mktFetch_('/' + _mktAccount_() + '/insights', {
    level: 'account',
    date_preset: datePreset,
    fields: 'spend,impressions,clicks,ctr,cpc,actions'
  });
  var row = (r.data && r.data[0]) || {};
  var actions = row.actions || [];
  return {
    spend:       Number(row.spend || 0),
    impressions: Number(row.impressions || 0),
    clicks:      Number(row.clicks || 0),
    ctr:         Number(row.ctr || 0),
    cpc:         Number(row.cpc || 0),
    lead:        _mktActionVal_(actions, 'lead'),
    contact:     _mktContactVal_(actions)
  };
}

function _mktCampaigns_() {
  var r = _mktFetch_('/' + _mktAccount_() + '/campaigns', {
    fields: 'id,name,status,effective_status,daily_budget,insights.date_preset(last_7d){spend,impressions,clicks,ctr,actions}',
    effective_status: '["ACTIVE","PAUSED","CAMPAIGN_PAUSED","ADSET_PAUSED","WITH_ISSUES"]',
    limit: 30
  });
  var out = [];
  (r.data || []).forEach(function(c) {
    var ins = (c.insights && c.insights.data && c.insights.data[0]) || null;
    var actions = (ins && ins.actions) || [];
    var leads = _mktActionVal_(actions, 'lead');
    var contacts = _mktContactVal_(actions);
    var spend = ins ? Number(ins.spend || 0) : 0;
    out.push({
      id: c.id,
      name: c.name,
      status: c.status,
      effectiveStatus: c.effective_status,
      dailyBudget: c.daily_budget ? Number(c.daily_budget) / 100 : null,
      spend: spend,
      impressions: ins ? Number(ins.impressions || 0) : 0,
      clicks: ins ? Number(ins.clicks || 0) : 0,
      ctr: ins ? Number(ins.ctr || 0) : 0,
      leads: leads,
      contacts: contacts,
      cpl: leads ? spend / leads : null
    });
  });
  // mostra ativas + qualquer pausada que tenha gasto nos últimos 7d
  out = out.filter(function(c) { return c.effectiveStatus === 'ACTIVE' || c.spend > 0; });
  // ordena: ativas primeiro, depois por gasto desc
  out.sort(function(a, b) {
    var aActive = a.effectiveStatus === 'ACTIVE' ? 1 : 0;
    var bActive = b.effectiveStatus === 'ACTIVE' ? 1 : 0;
    if (aActive !== bActive) return bActive - aActive;
    return b.spend - a.spend;
  });
  return out;
}

function _marketingNow_() {
  var today = _mktAccountInsights_('today');
  var d7    = _mktAccountInsights_('last_7d');
  var month = _mktAccountInsights_('this_month');
  var campaigns = _mktCampaigns_();
  return {
    ok: true,
    fetchedAt: new Date().toISOString(),
    accountId: _mktAccount_(),
    kpis: {
      spendToday: today.spend,
      spend7d: d7.spend,
      spendMonth: month.spend,
      leads7d: d7.lead,
      contacts7d: d7.contact,
      cpl7d: d7.lead ? d7.spend / d7.lead : null,
      cpc7d: d7.cpc,
      ctr7d: d7.ctr
    },
    campaigns: campaigns
  };
}

// ---- Funil agregado: lê Meta Ads + MetaCapi_Log + Feegow status_id=3 ----
function _mktSheet_() {
  var ss = SpreadsheetApp.openById(MCP_SPREADSHEET_ID);
  var sh = ss.getSheetByName(MKT_FUNNEL_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MKT_FUNNEL_SHEET);
    sh.appendRow(['snapshotAt','windowDays','spend','impressions','clicks','ctr','contacts_pixel','schedule_capi','complete_reg_capi','procedimentos_feegow','cpl_schedule','cpa_procedimento','consultas_1avez','procedimentos_fiv']);
    sh.setFrozenRows(1);
    return sh;
  }
  // garante as 2 colunas novas no fim (planilha antiga)
  var lastCol = sh.getLastColumn();
  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('consultas_1avez') < 0)   { sh.getRange(1, lastCol + 1).setValue('consultas_1avez'); lastCol++; }
  if (header.indexOf('procedimentos_fiv') < 0) { sh.getRange(1, lastCol + 1).setValue('procedimentos_fiv'); }
  return sh;
}

function _mktCountCapiEvents_(eventName, days) {
  var ss = SpreadsheetApp.openById(MCP_SPREADSHEET_ID);
  var sh = ss.getSheetByName(MCP_LOG_SHEET);
  if (!sh) return 0;
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return 0;
  var head = data[0];
  // suporta os 2 schemas possíveis (PascalCase atual + snake_case legado)
  var idxEvent  = head.indexOf('EventName');  if (idxEvent  < 0) idxEvent  = head.indexOf('event_name');
  var idxTs     = head.indexOf('Timestamp');  if (idxTs     < 0) idxTs     = head.indexOf('timestamp');
  var idxStatus = head.indexOf('Status');     if (idxStatus < 0) idxStatus = head.indexOf('status');
  if (idxEvent < 0 || idxTs < 0) return 0;
  var since = new Date();
  since.setDate(since.getDate() - days);
  var n = 0;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[idxEvent] !== eventName) continue;
    var ts = new Date(row[idxTs]);
    if (isNaN(ts.getTime()) || ts < since) continue;
    // aceita ENVIADO (atual) e OK (legado)
    if (idxStatus >= 0 && row[idxStatus] !== 'ENVIADO' && row[idxStatus] !== 'OK') continue;
    n++;
  }
  return n;
}

// Conta os atendidos (status_id=3) nos últimos N dias, classificando por tipo:
//   total           = todos os atendimentos realizados
//   consultas1aVez  = consultas de primeira vez (paciente novo)
//   fiv             = procedimentos FIV de verdade (punção, TEC) — exclui USG/preparo
// Nomes de procedimentos (cache 6h)
function _feegowProcNames_() {
  var names = {};
  var pc = CacheService.getScriptCache();
  var hit = pc.get('feegow_procnames');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  try {
    var r = UrlFetchApp.fetch(MCP_FEEGOW_BASE + '/procedures/list', { headers: { 'x-access-token': MCP_FEEGOW_TOKEN }, muteHttpExceptions: true });
    ((JSON.parse(r.getContentText()).content) || []).forEach(function(p) { names[p.id || p.procedimento_id] = (p.nome || p.procedimento || p.name || ''); });
    pc.put('feegow_procnames', JSON.stringify(names), 21600);
  } catch (e) {}
  return names;
}

// Nomes de profissionais (cache 6h)
function _feegowProfNames_() {
  var names = {};
  var pc = CacheService.getScriptCache();
  var hit = pc.get('feegow_profnames');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  try {
    var r = UrlFetchApp.fetch(MCP_FEEGOW_BASE + '/professional/list', { headers: { 'x-access-token': MCP_FEEGOW_TOKEN }, muteHttpExceptions: true });
    ((JSON.parse(r.getContentText()).content) || []).forEach(function(p) { names[p.id || p.profissional_id] = (p.nome || p.name || ''); });
    pc.put('feegow_profnames', JSON.stringify(names), 21600);
  } catch (e) {}
  return names;
}

// "R$ 2 000,00" / "R$ 1.234,56" -> número
function _parseBRL_(s) {
  if (!s) return 0;
  var t = String(s).replace(/R\$/i, '').replace(/\s/g, '').replace(/\./g, '').replace(',', '.');
  var n = parseFloat(t);
  return isNaN(n) ? 0 : n;
}

function _mktCountByType_(sinceDate, untilDate) {
  var out = { total: 0, consultas1aVez: 0, fiv: 0, atendidosBreakdown: [], fivPorMedico: [], outrosDetalhe: [] };
  if (!MCP_FEEGOW_TOKEN) return out;
  var until = untilDate || new Date();
  var fmt = function(d) { return Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd'); };
  try {
    var resp = UrlFetchApp.fetch(MCP_FEEGOW_BASE + '/appoints/search?data_start=' + fmt(sinceDate) + '&data_end=' + fmt(until) + '&status_id=3&per_page=1000',
      { headers: { 'x-access-token': MCP_FEEGOW_TOKEN }, muteHttpExceptions: true });
    if (resp.getResponseCode() >= 400) return out;
    var content = (JSON.parse(resp.getContentText()).content) || [];
    out.total = content.length;
    var names = _feegowProcNames_();
    var profs = _feegowProfNames_();
    var cats = {}, fivMed = {}, outros = {}, catsVal = {}, fivMedVal = {}, outrosVal = {};
    content.forEach(function(a) {
      var v = _parseBRL_(a.valor);
      var nm = (names[a.procedimento_id] || '').normalize('NFKD').replace(/[̀-ͯ]/g, '').toUpperCase();
      var isUSG = nm.indexOf('USG') >= 0;
      // Injúria: por nome OU pelo id legado 69 (que saiu da lista de procedimentos)
      var isInjuria = !isUSG && (nm.indexOf('INJURIA') >= 0 || Number(a.procedimento_id) === 69);
      var isCons = nm.indexOf('CONSULTA') >= 0;
      var is1a = isCons && !isUSG && (nm.indexOf('1A VEZ') >= 0 || nm.indexOf('PRIMEIRA') >= 0);
      // Retorno: por nome OU id legado 34 (Consulta Presencial de retorno não cobrada, R$0)
      var isRet = (isCons && !isUSG && nm.indexOf('RETORNO') >= 0) || Number(a.procedimento_id) === 34;
      // Triagem: conversa OU avaliação de doadora/receptora
      var isTriagem = nm.indexOf('CONVERSA') >= 0 || (nm.indexOf('AVALIACAO') >= 0 && (nm.indexOf('DOADORA') >= 0 || nm.indexOf('RECEPTORA') >= 0));
      var isPuncao = nm.indexOf('PUNCAO') >= 0;
      var isTEC = nm.indexOf('TEC') >= 0 && !isUSG && nm.indexOf('PREPARO') < 0;
      var isFIV = isPuncao || isTEC;
      var cat;
      if (isFIV)            cat = 'Procedimento FIV';
      else if (isUSG)       cat = 'Ultrassom';
      else if (isInjuria)   cat = 'Injúria endometrial';
      else if (is1a)        cat = 'Consulta 1ª vez';
      else if (isRet)       cat = 'Consulta retorno';
      else if (isTriagem)   cat = 'Triagem (doadora/receptora)';
      else if (isCons)      cat = 'Consulta (outras)';
      else                  cat = 'Outros';
      cats[cat] = (cats[cat] || 0) + 1;
      catsVal[cat] = (catsVal[cat] || 0) + v;
      if (cat === 'Outros') {
        var rn = names[a.procedimento_id] || ('Sem nome (id ' + a.procedimento_id + ')');
        outros[rn] = (outros[rn] || 0) + 1;
        outrosVal[rn] = (outrosVal[rn] || 0) + v;
      }
      if (is1a) out.consultas1aVez++;
      if (isFIV) {
        out.fiv++;
        var med = profs[a.profissional_id] || ('Profissional ' + (a.profissional_id || '?'));
        fivMed[med] = (fivMed[med] || 0) + 1;
        fivMedVal[med] = (fivMedVal[med] || 0) + v;
      }
    });
    out.atendidosBreakdown = Object.keys(cats).map(function(k) { return { label: k, qtd: cats[k], valor: catsVal[k] || 0 }; }).sort(function(a, b) {
      if (a.label === 'Outros') return 1;
      if (b.label === 'Outros') return -1;
      return b.qtd - a.qtd;
    });
    out.fivPorMedico = Object.keys(fivMed).map(function(k) { return { medico: k, qtd: fivMed[k], valor: fivMedVal[k] || 0 }; }).sort(function(a, b) { return b.qtd - a.qtd; });
    out.outrosDetalhe = Object.keys(outros).map(function(k) { return { nome: k, qtd: outros[k], valor: outrosVal[k] || 0 }; }).sort(function(a, b) { return b.qtd - a.qtd; });
  } catch (e) {}
  return out;
}

// Início do mês corrente (fuso São Paulo, à meia-noite local)
function _mktMonthStart_() {
  var s = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM') + '-01T00:00:00';
  return new Date(s);
}

// Conta eventos do CAPI log a partir de uma data (em vez de "últimos N dias")
function _mktCountCapiSince_(eventName, sinceDate) {
  var ss = SpreadsheetApp.openById(MCP_SPREADSHEET_ID);
  var sh = ss.getSheetByName(MCP_LOG_SHEET);
  if (!sh) return 0;
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return 0;
  var head = data[0];
  var idxEvent  = head.indexOf('EventName');  if (idxEvent  < 0) idxEvent  = head.indexOf('event_name');
  var idxTs     = head.indexOf('Timestamp');  if (idxTs     < 0) idxTs     = head.indexOf('timestamp');
  var idxStatus = head.indexOf('Status');     if (idxStatus < 0) idxStatus = head.indexOf('status');
  if (idxEvent < 0 || idxTs < 0) return 0;
  var n = 0;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[idxEvent] !== eventName) continue;
    var ts = new Date(row[idxTs]);
    if (isNaN(ts.getTime()) || ts < sinceDate) continue;
    if (idxStatus >= 0 && row[idxStatus] !== 'ENVIADO' && row[idxStatus] !== 'OK') continue;
    n++;
  }
  return n;
}

function gravarFunilSnapshot() {
  var since = _mktMonthStart_();
  var periodo = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'MM/yyyy');
  var ins = _mktAccountInsights_('this_month');
  var schedule    = _mktCountCapiSince_('Schedule', since);
  var completeReg = _mktCountCapiSince_('CompleteRegistration', since);
  var byType = _mktCountByType_(since);
  var procedimentos = byType.total;
  var row = [
    new Date().toISOString(),
    periodo,
    ins.spend,
    ins.impressions,
    ins.clicks,
    ins.ctr,
    ins.contact,
    schedule,
    completeReg,
    procedimentos,
    schedule ? (ins.spend / schedule) : null,
    procedimentos ? (ins.spend / procedimentos) : null,
    byType.consultas1aVez,
    byType.fiv
  ];
  _mktSheet_().appendRow(row);
  Logger.log('Funil snapshot gravado: ' + JSON.stringify(row));
  return row;
}

function _marketingFunnel_() {
  var sh = _mktSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) {
    return { ok: true, hasData: false, msg: 'Sem snapshot ainda. Rode gravarFunilSnapshot() ou aguarde a trigger.' };
  }
  var head = data[0];
  var last = data[data.length - 1];
  var rec = {};
  head.forEach(function(h, i) { rec[h] = last[i]; });

  // últimas 14 linhas pra mini-histórico
  var hist = data.slice(Math.max(1, data.length - 14)).map(function(r) {
    var o = {}; head.forEach(function(h, i) { o[h] = r[i]; }); return o;
  });

  return {
    ok: true,
    hasData: true,
    snapshot: rec,
    history: hist,
    funnelSteps: [
      { label: 'Impressões Meta',      value: Number(rec.impressions || 0),         conv: null },
      { label: 'Cliques',              value: Number(rec.clicks || 0),              conv: rec.impressions ? (rec.clicks / rec.impressions * 100) : null },
      { label: 'Conversas WhatsApp',   value: Number(rec.contacts_pixel || 0),      conv: rec.clicks ? (rec.contacts_pixel / rec.clicks * 100) : null },
      { label: 'Schedule (CAPI)',      value: Number(rec.schedule_capi || 0),       conv: rec.contacts_pixel ? (rec.schedule_capi / rec.contacts_pixel * 100) : null },
      { label: 'Cadastro completo',    value: Number(rec.complete_reg_capi || 0),   conv: rec.schedule_capi ? (rec.complete_reg_capi / rec.schedule_capi * 100) : null },
      { label: 'Consultas 1ª vez',     value: Number(rec.consultas_1avez || 0),     conv: rec.complete_reg_capi ? (rec.consultas_1avez / rec.complete_reg_capi * 100) : null },
      { label: 'Procedimentos FIV',    value: Number(rec.procedimentos_fiv || 0),   conv: rec.consultas_1avez ? (rec.procedimentos_fiv / rec.consultas_1avez * 100) : null }
    ]
  };
}

// ---- ESTRUTURA: dados da "máquina" (públicos ao vivo + eventos) ----
function _datasetEventCount_(eventName, days) {
  try {
    var end = Math.floor(Date.now() / 1000);
    var start = end - (days || 7) * 86400;
    var r = _mktFetch_('/' + MCP_DATASET_ID + '/stats', { aggregation: 'event', start_time: start, end_time: end });
    var total = 0;
    ((r && r.data) || []).forEach(function(bucket) {
      ((bucket && bucket.data) || []).forEach(function(d) {
        if (d.value === eventName) total += Number(d.count || 0);
      });
    });
    return total;
  } catch (e) { return null; }
}

function _datasetEventCountSince_(eventName, sinceDate, untilDate) {
  try {
    var end = Math.floor((untilDate ? untilDate.getTime() : Date.now()) / 1000);
    var start = Math.floor(sinceDate.getTime() / 1000);
    var r = _mktFetch_('/' + MCP_DATASET_ID + '/stats', { aggregation: 'event', start_time: start, end_time: end });
    var total = 0;
    ((r && r.data) || []).forEach(function(bucket) {
      ((bucket && bucket.data) || []).forEach(function(d) { if (d.value === eventName) total += Number(d.count || 0); });
    });
    return total;
  } catch (e) { return null; }
}

function _audInfo_(id) {
  try {
    var r = _mktFetch_('/' + id, { fields: 'name,approximate_count_lower_bound,approximate_count_upper_bound,time_updated' });
    return {
      id: id, name: r.name || '',
      count: Number(r.approximate_count_lower_bound || 0),
      countUpper: Number(r.approximate_count_upper_bound || 0),
      updatedAt: r.time_updated || null
    };
  } catch (e) { return { id: id, error: String(e.message || e) }; }
}

function _marketingSetup_() {
  return {
    ok: true,
    fetchedAt: new Date().toISOString(),
    config: { engajado: 'CompleteRegistration (cadastro)', cliente: 'Controle Agendamentos (Schedule)' },
    eventsMes: {
      pageview:             _datasetEventCountSince_('PageView', _mktMonthStart_()),
      completeRegistration: _mktCountCapiSince_('CompleteRegistration', _mktMonthStart_()),
      schedule:             _mktCountCapiSince_('Schedule', _mktMonthStart_())
    },
    audiences: {
      clientes:  _audInfo_('120242388951550375'),  // Controle Agendamentos (auto-sync)
      lookalike: _audInfo_('120242388972990375'),  // Semelhante 1% Controle Agendamentos
      engajados: _audInfo_('120247022590510375')   // Engajados Paraser (CompleteRegistration)
    }
  };
}

// ---- ESTRUTURA por mês (navegável) ----
var _MESES_PT = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

function _ymBounds_(ym) {
  var now = new Date();
  var nowYm = Utilities.formatDate(now, 'America/Sao_Paulo', 'yyyy-MM');
  if (!ym || !/^\d{4}-\d{2}$/.test(ym)) ym = nowYm;
  var y = Number(ym.slice(0, 4)), m = Number(ym.slice(5, 7));  // m = 1..12
  var since = new Date(ym + '-01T00:00:00');
  var isCurrent = (ym === nowYm);
  var until = isCurrent ? now : new Date(y, m, 0, 23, 59, 59);  // new Date(y, m, 0) = último dia do mês m (1-based)
  return { ym: ym, since: since, until: until, isCurrent: isCurrent, label: _MESES_PT[m - 1] + ' de ' + y };
}

function _mktAccountInsightsRange_(since, until) {
  var fmt = function(d) { return Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd'); };
  var r = _mktFetch_('/' + _mktAccount_() + '/insights', {
    level: 'account',
    time_range: JSON.stringify({ since: fmt(since), until: fmt(until) }),
    fields: 'spend,impressions,clicks,ctr,cpc,actions'
  });
  var row = (r.data && r.data[0]) || {};
  var actions = row.actions || [];
  return {
    spend: Number(row.spend || 0), impressions: Number(row.impressions || 0),
    clicks: Number(row.clicks || 0), ctr: Number(row.ctr || 0),
    lead: _mktActionVal_(actions, 'lead'), contact: _mktContactVal_(actions)
  };
}

function _mktCountCapiRange_(eventName, since, until) {
  var ss = SpreadsheetApp.openById(MCP_SPREADSHEET_ID);
  var sh = ss.getSheetByName(MCP_LOG_SHEET);
  if (!sh) return 0;
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return 0;
  var head = data[0];
  var idxEvent  = head.indexOf('EventName');  if (idxEvent  < 0) idxEvent  = head.indexOf('event_name');
  var idxTs     = head.indexOf('Timestamp');  if (idxTs     < 0) idxTs     = head.indexOf('timestamp');
  var idxStatus = head.indexOf('Status');     if (idxStatus < 0) idxStatus = head.indexOf('status');
  if (idxEvent < 0 || idxTs < 0) return 0;
  var n = 0;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[idxEvent] !== eventName) continue;
    var ts = new Date(row[idxTs]);
    if (isNaN(ts.getTime()) || ts < since || ts > until) continue;
    if (idxStatus >= 0 && row[idxStatus] !== 'ENVIADO' && row[idxStatus] !== 'OK') continue;
    n++;
  }
  return n;
}

function _estruturaMonth_(ym) {
  var b = _ymBounds_(ym);
  var cache = CacheService.getScriptCache();
  var key = 'estrut8_' + b.ym;
  var hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var ins      = _mktAccountInsightsRange_(b.since, b.until);
  var schedule = _mktCountCapiRange_('Schedule', b.since, b.until);
  var cadastro = _mktCountCapiRange_('CompleteRegistration', b.since, b.until);
  var byType   = _mktCountByType_(b.since, b.until);
  var pageview = _datasetEventCountSince_('PageView', b.since, b.until);
  var pct = function(a, base) { return base ? (a / base * 100) : null; };
  var result = {
    ok: true, ym: b.ym, label: b.label, isCurrent: b.isCurrent, fetchedAt: new Date().toISOString(),
    config: { engajado: 'CompleteRegistration (cadastro)', cliente: 'Controle Agendamentos (Schedule)' },
    eventsMes: { pageview: pageview, completeRegistration: cadastro, schedule: schedule },
    audiences: {
      clientes:  _audInfo_('120242388951550375'),
      lookalike: _audInfo_('120242388972990375'),
      engajados: _audInfo_('120247022590510375')
    },
    // Funil de CAPTAÇÃO do mês (mesma turma, defasagem de dias/semanas) — termina na 1ª consulta
    funnelSteps: [
      { label: 'Impressões Meta',    value: ins.impressions, conv: null },
      { label: 'Cliques',            value: ins.clicks,      conv: pct(ins.clicks, ins.impressions) },
      { label: 'Conversas WhatsApp', value: ins.contact,     conv: pct(ins.contact, ins.clicks) },
      { label: 'Schedule (CAPI)',    value: schedule,        conv: pct(schedule, ins.contact) },
      { label: 'Consultas 1ª vez',   value: byType.consultas1aVez, conv: pct(byType.consultas1aVez, schedule) }
    ],
    // PRODUÇÃO da clínica no mês (volume; pacientes entraram em meses variados — NÃO é conversão deste mês)
    producao: { atendidos: byType.total, fiv: byType.fiv, atendidosBreakdown: byType.atendidosBreakdown, fivPorMedico: byType.fivPorMedico, outrosDetalhe: byType.outrosDetalhe }
  };
  try { cache.put(key, JSON.stringify(result), b.isCurrent ? 120 : 21600); } catch (e) {}
  return result;
}

// ================================================================
// CLARITY — coleta diária + agente que lê os números e propõe plano de ação
// ================================================================
// O QUE FAZ:
//   1× ao dia (trigger 7h) puxa as métricas de comportamento do Microsoft
//   Clarity (rage/dead click, scroll, engajamento, páginas, dispositivo),
//   grava uma linha na aba CLARITY_HISTORICO e chama o Claude pra escrever
//   um plano de ação em cima dos números + da variação vs. a semana passada.
//
//     ?action=clarity        → últimos números + variação + plano (lê da aba)
//     ?action=clarity-now    → força uma coleta agora (gasta 1 das 10 chamadas)
//     coletarClarity()       → o que a trigger das 7h roda
//
// ⚠️ POR QUE O DASHBOARD NÃO CHAMA A API DO CLARITY DIRETO:
//   A API do Clarity permite só 10 chamadas/dia por projeto (limite da
//   Microsoft). Se cada abertura da aba Marketing gastasse uma, a cota
//   morria antes do meio-dia e o painel quebrava. Por isso só a trigger
//   chama (1/dia) e o dashboard lê a aba · instantâneo e sem cota.
//   Mesma razão pro plano do agente: gerado 1×/dia (~22s, ~US$ 0,02) e
//   gravado. Se fosse por abertura, seria 22s de espera e custo por clique.
//
// ⚠️ A API só devolve os ÚLTIMOS 1-3 DIAS. Não existe histórico do lado da
//   Microsoft · é por isso que gravamos snapshot: sem ele o agente nunca
//   consegue dizer "piorou", só "está assim".
//
// SETUP (uma vez): ?action=clarity-setup&key=...  (cria a trigger das 7h)
//   Properties: CLARITY_TOKEN (Settings > Data Export no clarity.microsoft.com)
//               ANTHROPIC_KEY (console.anthropic.com)
// ================================================================

const CLR_SHEET = 'CLARITY_HISTORICO';
const CLR_API   = 'https://www.clarity.ms/export-data/api/v1/project-live-insights';
const CLR_MODEL = 'claude-opus-4-8';
const CLR_HEAD  = ['ts', 'data', 'sessoes', 'bots', 'paginas_por_sessao', 'scroll_medio_pct',
                   'tempo_ativo_s', 'dead_pct', 'rage_pct', 'quickback_pct', 'scripterror_pct',
                   'mobile_pct', 'top_pagina', 'top_referrer', 'plano_md'];

function _clrSheet_() {
  var ss = SpreadsheetApp.openById(MCP_SPREADSHEET_ID);
  var sh = ss.getSheetByName(CLR_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CLR_SHEET);
    sh.appendRow(CLR_HEAD);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ---- 1. Puxa a API do Clarity e achata o retorno num objeto legível ----
function _clarityFetch_(numOfDays) {
  var token = _MCP.getProperty('CLARITY_TOKEN');
  if (!token) throw new Error('CLARITY_TOKEN ausente nas Properties do script');
  var url = CLR_API + '?numOfDays=' + (numOfDays || 3);
  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  var body = resp.getContentText();
  if (code >= 400) throw new Error('Clarity API ' + code + ': ' + body.substring(0, 300));
  return JSON.parse(body);
}

function _clarityResumo_(raw) {
  var o = { paginas: [], dispositivos: {}, navegadores: {} };
  (raw || []).forEach(function(m) {
    var n = m.metricName, info = m.information || [];
    if (!info.length) return;
    var first = info[0];
    if (n === 'Traffic') {
      o.sessoes = Number(first.totalSessionCount || 0);
      o.bots = Number(first.totalBotSessionCount || 0);
      o.paginasPorSessao = Math.round((first.pagesPerSessionPercentage || 0) * 100) / 100;
    } else if (n === 'ScrollDepth') {
      o.scrollMedio = Math.round((first.averageScrollDepth || 0) * 10) / 10;
    } else if (n === 'EngagementTime') {
      o.tempoAtivo = Number(first.activeTime || 0);
      o.tempoTotal = Number(first.totalTime || 0);
    } else if (n === 'DeadClickCount')   { o.dead = _clrPct_(first); }
    else if (n === 'RageClickCount')     { o.rage = _clrPct_(first); }
    else if (n === 'QuickbackClick')     { o.quickback = _clrPct_(first); }
    else if (n === 'ScriptErrorCount')   { o.scriptError = _clrPct_(first); }
    else if (n === 'ErrorClickCount')    { o.errorClick = _clrPct_(first); }
    else if (n === 'PopularPages') {
      o.paginas = info.slice(0, 5).map(function(r) { return { url: r.url, visitas: Number(r.visitsCount || 0) }; });
    } else if (n === 'Device') {
      info.forEach(function(r) { o.dispositivos[r.name] = Number(r.sessionsCount || 0); });
    } else if (n === 'Browser') {
      info.slice(0, 5).forEach(function(r) { o.navegadores[r.name] = Number(r.sessionsCount || 0); });
    } else if (n === 'ReferrerUrl') {
      o.topReferrer = { url: first.name, sessoes: Number(first.sessionsCount || 0) };
    }
  });
  var mob = o.dispositivos.Mobile || 0;
  o.mobilePct = o.sessoes ? Math.round(mob * 1000 / o.sessoes) / 10 : 0;
  return o;
}

function _clrPct_(info) {
  return { pct: info.sessionsWithMetricPercentage || 0, total: Number(info.subTotal || 0) };
}

// ---- 2. O agente: manda os números pro Claude e recebe o plano ----
function _clarityAgente_(resumo, historico) {
  var key = _MCP.getProperty('ANTHROPIC_KEY');
  if (!key) return '_(ANTHROPIC_KEY ausente · plano não gerado)_';

  var sistema = 'Você é analista de CRO da Paraser, clínica de fertilidade no Rio de Janeiro. ' +
    'Recebe métricas do Microsoft Clarity do site paraser.com.br e propõe planos de ação.\n' +
    'REGRAS DE ESCRITA: português do Brasil natural, como gente fala. NUNCA use travessões (— ou –); ' +
    'use vírgula, ponto, parênteses ou "·".\n' +
    'Seja específico e acionável: cada achado vira uma ação concreta que alguém executa amanhã. ' +
    'Nada de conselho genérico de marketing.\n' +
    'Se um número não sustenta uma conclusão, diga que não dá pra concluir em vez de inventar. ' +
    'Se a variação vs. a semana passada for o mais relevante, lidere com ela.\n' +
    'CONTEXTO: ~90% do tráfego é mobile. /contato/ é a página de conversão (formulário Fale Conosco). ' +
    'O Google traz a maior parte das visitas. A clínica quer mais consultas agendadas.';

  var user = 'Métricas do Clarity (últimos 3 dias):\n' + JSON.stringify(resumo, null, 1) + '\n\n' +
    (historico && historico.length
      ? 'Histórico dos últimos dias (pra você ver tendência):\n' + JSON.stringify(historico, null, 1) + '\n\n'
      : 'Sem histórico ainda (primeira coleta). Não invente tendência.\n\n') +
    'Escreva o plano de ação: no máximo 3 achados, do mais urgente ao menos. ' +
    'Para cada um: o que o número mostra, por que importa pra clínica, e a ação concreta. ' +
    'Máximo 200 palavras no total. Markdown simples.';

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    // effort low + max_tokens curto: o UrlFetchApp corta em ~60s e a chamada
    // medida ficou em ~22s. Não subir sem medir de novo.
    payload: JSON.stringify({
      model: CLR_MODEL,
      max_tokens: 2000,
      thinking: { type: 'adaptive' },
      output_config: { effort: 'low' },
      system: sistema,
      messages: [{ role: 'user', content: user }]
    }),
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  if (code >= 400) {
    Logger.log('Claude API ' + code + ': ' + resp.getContentText().substring(0, 300));
    return '_(erro ao gerar plano: HTTP ' + code + ')_';
  }
  var data = JSON.parse(resp.getContentText());
  if (data.stop_reason === 'refusal') return '_(o modelo recusou a análise)_';
  var txt = '';
  (data.content || []).forEach(function(b) { if (b.type === 'text') txt += b.text; });
  return txt || '_(plano vazio)_';
}

// ---- 3. O que a trigger das 7h roda ----
function coletarClarity() {
  var resumo = _clarityResumo_(_clarityFetch_(3));
  var hist = _clarityLerHistorico_(7);
  var plano = _clarityAgente_(resumo, hist);
  var topPag = resumo.paginas.length ? resumo.paginas[0].url + ' (' + resumo.paginas[0].visitas + ')' : '';
  var row = [
    new Date().toISOString(),
    Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy'),
    resumo.sessoes || 0,
    resumo.bots || 0,
    resumo.paginasPorSessao || 0,
    resumo.scrollMedio || 0,
    resumo.tempoAtivo || 0,
    resumo.dead ? resumo.dead.pct : 0,
    resumo.rage ? resumo.rage.pct : 0,
    resumo.quickback ? resumo.quickback.pct : 0,
    resumo.scriptError ? resumo.scriptError.pct : 0,
    resumo.mobilePct || 0,
    topPag,
    resumo.topReferrer ? resumo.topReferrer.url : '',
    plano
  ];
  _clrSheet_().appendRow(row);
  Logger.log('Clarity coletado: ' + resumo.sessoes + ' sessões, plano com ' + plano.length + ' chars');
  return row;
}

// ---- 4. Leitura (o dashboard só usa isto · zero cota, instantâneo) ----
function _clarityLerHistorico_(dias) {
  var sh = _clrSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  var head = data[0];
  var out = [];
  for (var i = Math.max(1, data.length - dias); i < data.length; i++) {
    var rec = {};
    head.forEach(function(h, j) { if (h !== 'plano_md' && h !== 'ts') rec[h] = data[i][j]; });
    out.push(rec);
  }
  return out;
}

function _clarity_() {
  var sh = _clrSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) {
    return { ok: true, hasData: false, msg: 'Sem coleta ainda. Roda ?action=clarity-now ou espera a trigger das 7h.' };
  }
  var head = data[0];
  var last = data[data.length - 1];
  var rec = {};
  head.forEach(function(h, j) { rec[h] = last[j]; });

  // variação vs ~7 dias atrás (só quando já existe histórico suficiente)
  var ref = data.length >= 8 ? data[data.length - 8] : null;
  var delta = null;
  if (ref) {
    delta = {};
    ['sessoes', 'scroll_medio_pct', 'tempo_ativo_s', 'scripterror_pct', 'dead_pct'].forEach(function(c) {
      var j = head.indexOf(c);
      if (j >= 0 && typeof last[j] === 'number' && typeof ref[j] === 'number') {
        delta[c] = Math.round((last[j] - ref[j]) * 10) / 10;
      }
    });
  }
  return { ok: true, hasData: true, atual: rec, delta: delta, dias: data.length - 1 };
}

function _claritySetup_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'coletarClarity') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('coletarClarity').timeBased().atHour(7).everyDays(1).create();
  return {
    ok: true,
    trigger: 'coletarClarity 7h/dia',
    clarityToken: _MCP.getProperty('CLARITY_TOKEN') ? 'ok' : 'AUSENTE',
    anthropicKey: _MCP.getProperty('ANTHROPIC_KEY') ? 'ok' : 'AUSENTE'
  };
}


// ---- Web app ----
// POST: grava Properties (segredos vão no BODY, nunca na URL · a URL fica em log)
function doPost(e) {
  var json = function(o) {
    return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  };
  var body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json({ ok: false, error: 'body inválido' }); }
  if (body.key !== MKT_AUTH_KEY) return json({ ok: false, error: 'unauthorized' });
  if (body.action !== 'set-props') return json({ ok: false, error: 'ação desconhecida' });
  var gravadas = [];
  Object.keys(body.props || {}).forEach(function(k) {
    _MCP.setProperty(k, String(body.props[k]));
    gravadas.push(k);
  });
  return json({ ok: true, gravadas: gravadas });  // nunca devolve o valor
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var json = function(o) {
    return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  };
  if (p.key !== MKT_AUTH_KEY) return json({ ok: false, error: 'unauthorized' });
  try {
    if (p.action === 'prof-names')    return json({ ok: true, profs: _feegowProfNames_() });
                  return json({ ok: true, procurando: alvo3, endpoints: achados3 });
    }
                if (p.action === 'clarity')       return json(_clarity_());
    if (p.action === 'clarity-now')   return json({ ok: true, row: coletarClarity() });
    if (p.action === 'clarity-setup') return json(_claritySetup_());
    if (p.action === 'marketing-now')     return json(_marketingNow_());
    if (p.action === 'marketing-funnel')  return json(_marketingFunnel_());
    if (p.action === 'marketing-setup')   return json(_marketingSetup_());
    if (p.action === 'estrutura-month')   return json(_estruturaMonth_(p.ym));
    if (p.action === 'marketing-snapshot-now') return json({ ok: true, row: gravarFunilSnapshot() });
    if (p.action === 'audience-info') return json({ ok: true, audienceId: MCP_AUDIENCE_ID, tokenSource: _MCP.getProperty('META_ADS_TOKEN') ? 'META_ADS_TOKEN' : 'META_CAPI_TOKEN' });
    if (p.action === 'audience-test') {
      var fake = { nome: 'Maria Teste Sync', email: 'maria.sync+test@paraser.com.br', celular: '21998765432', cidade: 'Rio de Janeiro', estado: 'RJ' };
      var ar = mcpAddUsersToAudience([mcpAudienceRow(fake)]);
      return json({ ok: ar.ok, code: ar.code, body: ar.body });
    }
    if (p.action === 'audience-sync-now') return json(backfillAudience(Number(p.days || 14)));
    if (p.action === 'audience-sync-window') return json(backfillAudienceWindow(Number(p.from || 30), Number(p.to || 0)));
    if (p.action === 'audience-backfill') return json(backfillResumable(Number(p.from || 90), Number(p.to || 0), Number(p.max || 280)));
    if (p.action === 'token-check') {
      var out = {};
      ['META_ADS_TOKEN', 'META_CAPI_TOKEN'].forEach(function(propName) {
        var tk = _MCP.getProperty(propName);
        if (!tk) { out[propName] = 'ausente'; return; }
        var meR  = UrlFetchApp.fetch('https://graph.facebook.com/' + MCP_API_VERSION + '/me?fields=id,name&access_token=' + encodeURIComponent(tk), { muteHttpExceptions: true });
        var audR = UrlFetchApp.fetch('https://graph.facebook.com/' + MCP_API_VERSION + '/' + MCP_AUDIENCE_ID + '?fields=id,name,approximate_count_lower_bound&access_token=' + encodeURIComponent(tk), { muteHttpExceptions: true });
        out[propName] = { me: meR.getContentText().slice(0, 200), audRead_code: audR.getResponseCode(), audRead: audR.getContentText().slice(0, 300) };
      });
      return json(out);
    }
    if (p.action === 'setup-trigger') { setupMarketingDashboard(); return json({ ok: true, msg: 'triggers 8h+18h ativos' }); }
    if (p.action === 'debug-capi-log') {
      var ss = SpreadsheetApp.openById(MCP_SPREADSHEET_ID);
      var sh = ss.getSheetByName(MCP_LOG_SHEET);
      if (!sh) return json({ ok: false, error: 'sheet not found' });
      var data = sh.getDataRange().getValues();
      return json({ ok: true, header: data[0], totalRows: data.length - 1, lastRow: data[data.length - 1] });
    }
    return json({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  }
}

// ---- Trigger 2x/dia ----
function setupMarketingDashboard() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'gravarFunilSnapshot') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('gravarFunilSnapshot').timeBased().atHour(8).everyDays(1).create();
  ScriptApp.newTrigger('gravarFunilSnapshot').timeBased().atHour(18).everyDays(1).create();
  Logger.log('✅ Triggers gravarFunilSnapshot: 8h e 18h diárias');
}

function removerMarketingDashboardTriggers() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'gravarFunilSnapshot') { ScriptApp.deleteTrigger(t); n++; }
  });
  Logger.log('Removidos ' + n + ' trigger(s) gravarFunilSnapshot');
}
