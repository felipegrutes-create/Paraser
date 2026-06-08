var LIST_A_ID    = 4;    // Reativacao-2026-MetaLeads
var LIST_B_ID    = 5;    // Reativacao-2026-Active
var MAX_POR_REQ  = 500;  // limite GET da API Brevo
var CHUNK_ADD    = 150;  // limite POST /contacts/add da API Brevo
var RUNS_POR_VEZ = 3;    // 3 × 500 = 1.500 por semana
var OFFSET_KEY   = 'BREVO_LOTE_OFFSET';

function enviarLoteSemanal() {
  var props  = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('BREVO_API_KEY');
  var offset = parseInt(props.getProperty(OFFSET_KEY) || '0');
  var totalEnviados = 0;

  for (var i = 0; i < RUNS_POR_VEZ; i++) {
    // 1. Busca até 500 da Lista A
    var resp = UrlFetchApp.fetch(
      'https://api.brevo.com/v3/contacts/lists/' + LIST_A_ID +
      '/contacts?limit=' + MAX_POR_REQ + '&offset=' + offset + '&sort=asc',
      { headers: { 'api-key': apiKey, 'content-type': 'application/json' } }
    );

    var contacts = JSON.parse(resp.getContentText()).contacts || [];

    if (contacts.length === 0) {
      Logger.log('Todos os lotes processados — desativando trigger.');
      _desativarTrigger();
      _notificarSlack('✅ *Brevo Reativação* — todos os 18.593 leads foram processados. Trigger desativado.');
      return;
    }

    // 2. Adiciona à Lista B em chunks de 150 (limite do endpoint)
    var emails = contacts.map(function(c){ return c.email; });
    for (var j = 0; j < emails.length; j += CHUNK_ADD) {
      var chunk = emails.slice(j, j + CHUNK_ADD);
      UrlFetchApp.fetch(
        'https://api.brevo.com/v3/contacts/lists/' + LIST_B_ID + '/contacts/add',
        {
          method: 'POST',
          headers: { 'api-key': apiKey, 'content-type': 'application/json' },
          payload: JSON.stringify({ emails: chunk })
        }
      );
    }

    offset += contacts.length;
    totalEnviados += contacts.length;
    props.setProperty(OFFSET_KEY, String(offset));
  }

  var semana = Math.ceil(offset / (MAX_POR_REQ * RUNS_POR_VEZ));
  Logger.log('Semana ' + semana + ': ' + totalEnviados + ' leads enviados. Total: ' + offset + '/18593');
  _notificarSlack(
    '📧 *Brevo Reativação* — Semana ' + semana + ': *' + totalEnviados +
    ' leads* adicionados à sequência.\nTotal processado: ' + offset + ' / 18.593'
  );
}

// ── Setup ────────────────────────────────────────────────────

function configurarTriggerSemanal() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'enviarLoteSemanal') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('enviarLoteSemanal')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  Logger.log('Trigger criado: toda segunda-feira às 8h.');
}

function testarLote() {
  enviarLoteSemanal();
}

function resetarOffset() {
  PropertiesService.getScriptProperties().setProperty(OFFSET_KEY, '0');
  Logger.log('Offset zerado.');
}

function verStatus() {
  var props  = PropertiesService.getScriptProperties();
  var offset = parseInt(props.getProperty(OFFSET_KEY) || '0');
  Logger.log('Leads processados: ' + offset + ' / 18593');
  Logger.log('Próxima semana começa no offset: ' + offset);
}

// ── Internos ─────────────────────────────────────────────────

function _desativarTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'enviarLoteSemanal') ScriptApp.deleteTrigger(t);
  });
}

function _notificarSlack(msg) {
  var webhook = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!webhook) return;
  UrlFetchApp.fetch(webhook, {
    method: 'POST',
    payload: JSON.stringify({ text: msg })
  });
}
