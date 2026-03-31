// =========================================================
// Feegow CRM — Google Apps Script Web App
// Cole este código no Apps Script da planilha e faça deploy
// como Web App: Execute as = você, Who has access = Anyone
// =========================================================

// =========================================================
// Feegow API Proxy — para contornar CORS no navegador
// =========================================================
const FEEGOW_API_BASE  = 'https://api.feegow.com/v1/api';
const FEEGOW_API_TOKEN = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJmZWVnb3ciLCJhdWQiOiJwdWJsaWNhcGkiLCJpYXQiOjE3NDM0NzEyNDIsImxpY2Vuc2VJRCI6MTQ0MzR9.oh2VSWT5UPEfYRrPCv34IM1NuP8Aq_ehFYWhE8f5MuU';

// Fichas Paciente
const FICHA_SHEET_NAME = 'Fichas_Pacientes';
const FICHA_HEADERS = [
  'paciente_id','paciente_nome','data_consulta','medico','nascimento','idade',
  'conjuge','nasc_conjuge','idade_conjuge','data_max_retorno','tratamento',
  'obs_caso','urologista','avisos1',
  'reg_doadora','status_tto','amh','preparo_esp',
  'semen','tto_urologista','frag_dna','data_frag_dna','obs_semen',
  'data_inducao','medicacao','foliculos','ovulos_aspirados','ovulos_maduros',
  'ovulos_fertilizados','blastocistos','class_palhetas',
  'pgta','result_pgta','hemato','aviso_hemato',
  'data_te','tipo_te','blastos_transf','sobrou_blastos','blastos_sobrando','laboratorio',
  'injuria','result_injuria','filgastrim','tractocile','tipo_preparo','med_preparo','clexane','obs_te',
  'data_beta','dosagens_beta','sg','bce','obs_1_usg','data_2_usg','obs_2_usg',
  'pre_natal','tipo_gravidez','tipo_aborto','pre_natal_paraser',
  'data_parto','tipo_parto','idade_gest','obs_parto',
  'nasc_vivos','genero_bebe1','estatura_bebe1','peso_bebe1','nome_bebe1',
  'genero_bebe2','estatura_bebe2','peso_bebe2','nome_bebe2','obs_bebe',
  'depoimento','info_depoimento',
  'indicacao_pacote','orcamento_feito','data_orcamento','orcamento','obs_orcamento',
  'status_retorno','data_retorno','sem_retorno_motivo','nao_vai_tratar',
  'aguardando_motivo','avisos2','pagamento','aviso_pagamento',
  'contato_psi','resultado','consulta1','consulta2','info_psi','updated_at'
];

const CRM_SHEET_NAME = 'CRM_Pendentes';
const CRM_LOG_SHEET  = 'CRM_Log';
const CRM_HEADERS = [
  'paciente_key','contato1_data','contato2_data',
  'vendedora','valor','retorno_marcado','proposta_feita',
  'parc_pago','quitado','observacoes','classificacao','ultima_atualizacao'
];
const LOG_HEADERS = [
  'timestamp','acao','paciente_key','usuario',
  'dados_anteriores','coluna_origem','coluna_destino'
];
const DONOR_SHEET = 'Doadoras_Perfis';
const DONOR_HEADERS = ['id','codigo_perfil','dados','ultima_atualizacao'];

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || '';

    // Proxy Feegow API
    if (action === 'feegow_proxy') {
      return handleFeegowProxy(e.parameter);
    }

    // Teste de diagnóstico
    if (action === 'test_feegow') {
      return handleFeegowTest(e.parameter);
    }

    // Retornar fichas de pacientes
    if (action === 'get_fichas') {
      const fSheet = getOrCreateFichaSheet();
      const fData = fSheet.getDataRange().getValues();
      if (fData.length <= 1) return jsonOk({ fichas: [] });
      const fHeaders = fData[0];
      const fichas = fData.slice(1).map(row => {
        const obj = {};
        fHeaders.forEach((h, i) => { obj[h] = row[i]; });
        return obj;
      });
      return jsonOk({ fichas });
    }

    // Retornar log de exclusões
    if (action === 'get_log') {
      const logSheet = getOrCreateLogSheet();
      const data = logSheet.getDataRange().getValues();
      if (data.length <= 1) return jsonOk({ logs: [] });
      const headers = data[0];
      const logs = data.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i]; });
        return obj;
      });
      return jsonOk({ logs });
    }

    // Retornar perfis de doadoras
    if (action === 'get_donors') {
      const dSheet = getOrCreateDonorSheet();
      const dData = dSheet.getDataRange().getValues();
      if (dData.length <= 1) return jsonOk({ profiles: [] });
      const dHeaders = dData[0];
      const profiles = dData.slice(1).map(row => {
        const obj = {};
        dHeaders.forEach((h, i) => { obj[h] = row[i]; });
        return obj;
      });
      return jsonOk({ profiles });
    }

    // Retornar registros CRM (padrão)
    const sheet = getOrCreateSheet();
    const data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return jsonOk({ records: [] });

    const headers = data[0];
    const records = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
    return jsonOk({ records });
  } catch(err) {
    return jsonErr(err.message);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || 'save';

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      if (action === 'save_donor') {
        return handleSaveDonor(body);
      }

      if (action === 'save_ficha') {
        return handleSaveFicha(body);
      }

      if (action === 'upload_photo') {
        return handleUploadPhoto(body);
      }

      const key = (body.paciente_key || '').trim();
      if (!key) return jsonErr('paciente_key obrigatório');

      if (action === 'delete') {
        return handleDelete(body, key);
      } else if (action === 'undo_delete') {
        return handleUndoDelete(body, key);
      } else {
        return handleSave(body, key);
      }
    } finally {
      lock.releaseLock();
    }
  } catch(err) {
    return jsonErr(err.message);
  }
}

function handleSave(body, key) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length === 0) return jsonErr('Planilha sem cabeçalho');
  const headers = data[0];
  const keyCol  = headers.indexOf('paciente_key');
  if (keyCol === -1) return jsonErr('Header paciente_key não encontrado');

  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyCol]).trim() === key) {
      targetRow = i + 1;
      break;
    }
  }

  body.ultima_atualizacao = new Date().toISOString();
  const rowData = headers.map(h => body[h] !== undefined ? body[h] : '');

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return jsonOk({ success: true });
}

function handleDelete(body, key) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length === 0) return jsonErr('Planilha sem cabeçalho');
  const headers = data[0];
  const keyCol  = headers.indexOf('paciente_key');
  if (keyCol === -1) return jsonErr('Header paciente_key não encontrado');

  // Buscar dados anteriores
  let targetRow = -1;
  let oldData = {};
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyCol]).trim() === key) {
      targetRow = i + 1;
      headers.forEach((h, j) => { oldData[h] = data[i][j]; });
      break;
    }
  }

  // Gravar log antes de apagar
  const logSheet = getOrCreateLogSheet();
  logSheet.appendRow([
    new Date().toISOString(),
    'delete',
    key,
    body.usuario || 'desconhecido',
    JSON.stringify(oldData),
    body.coluna_origem || '',
    body.coluna_destino || ''
  ]);

  // Limpar registro CRM (manter apenas paciente_key)
  if (targetRow > 0) {
    const emptyRow = headers.map(h => h === 'paciente_key' ? key : (h === 'ultima_atualizacao' ? new Date().toISOString() : ''));
    sheet.getRange(targetRow, 1, 1, emptyRow.length).setValues([emptyRow]);
  }

  return jsonOk({ success: true, logged: true });
}

function handleUndoDelete(body, key) {
  // Restaurar dados de um log de exclusão
  const previousData = body.dados_anteriores;
  if (!previousData) return jsonErr('dados_anteriores obrigatório para undo');

  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol  = headers.indexOf('paciente_key');

  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyCol]).trim() === key) {
      targetRow = i + 1;
      break;
    }
  }

  previousData.ultima_atualizacao = new Date().toISOString();
  const rowData = headers.map(h => previousData[h] !== undefined ? previousData[h] : '');

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  // Gravar log do undo
  const logSheet = getOrCreateLogSheet();
  logSheet.appendRow([
    new Date().toISOString(),
    'undo_delete',
    key,
    body.usuario || 'desconhecido',
    JSON.stringify(previousData),
    '',
    ''
  ]);

  return jsonOk({ success: true });
}

function handleSaveDonor(body) {
  const id = (body.id || '').trim();
  if (!id) return jsonErr('id obrigatório para perfil');

  const sheet = getOrCreateDonorSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === id) {
      targetRow = i + 1;
      break;
    }
  }

  body.ultima_atualizacao = new Date().toISOString();
  const rowData = headers.map(h => body[h] !== undefined ? body[h] : '');

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return jsonOk({ success: true });
}

function getOrCreateDonorSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DONOR_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DONOR_SHEET);
    sheet.getRange(1, 1, 1, DONOR_HEADERS.length).setValues([DONOR_HEADERS]);
    sheet.getRange(1, 1, 1, DONOR_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CRM_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CRM_SHEET_NAME);
    sheet.getRange(1, 1, 1, CRM_HEADERS.length).setValues([CRM_HEADERS]);
    sheet.getRange(1, 1, 1, CRM_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CRM_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CRM_LOG_SHEET);
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// =========================================================
// Upload de Foto para Google Drive
// =========================================================
const PHOTO_FOLDER_NAME = 'Paraser_Fotos_Doadoras';

function getOrCreatePhotoFolder() {
  const folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function handleUploadPhoto(body) {
  const base64Data = body.foto || '';
  const fileName = body.file_name || ('foto_' + Date.now() + '.jpg');

  if (!base64Data) return jsonErr('foto obrigatória');

  // Remove data URI prefix: "data:image/jpeg;base64,..."
  const parts = base64Data.split(',');
  const raw = parts.length > 1 ? parts[1] : parts[0];
  const mimeMatch = base64Data.match(/data:([^;]+);/);
  const mime = mimeMatch ? mimeMatch[1] : 'image/jpeg';

  const decoded = Utilities.base64Decode(raw);
  const blob = Utilities.newBlob(decoded, mime, fileName);

  const folder = getOrCreatePhotoFolder();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const viewUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;

  return jsonOk({ success: true, url: viewUrl, file_id: fileId });
}

// =========================================================
// Feegow API Proxy
// =========================================================
function handleFeegowProxy(params) {
  const endpoint = params.endpoint || '';
  if (!endpoint) return jsonErr('endpoint obrigatório');

  // Build URL with remaining params (exclude action and endpoint)
  const url = FEEGOW_API_BASE + endpoint;
  const queryParts = [];
  Object.keys(params).forEach(k => {
    if (k !== 'action' && k !== 'endpoint') {
      queryParts.push(encodeURIComponent(k) + '=' + encodeURIComponent(params[k]));
    }
  });
  const fullUrl = queryParts.length > 0 ? url + '?' + queryParts.join('&') : url;

  try {
    const resp = UrlFetchApp.fetch(fullUrl, {
      method: 'get',
      headers: { 'x-access-token': FEEGOW_API_TOKEN },
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    const body = resp.getContentText();

    if (code >= 200 && code < 300) {
      return ContentService
        .createTextOutput(body)
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      return jsonErr('Feegow API retornou ' + code + ': ' + body);
    }
  } catch(err) {
    return jsonErr('Erro ao chamar Feegow: ' + err.message);
  }
}

// =========================================================
// Fichas Pacientes
// =========================================================
function handleSaveFicha(body) {
  const pid = (body.paciente_id || '').toString().trim();
  if (!pid) return jsonErr('paciente_id obrigatório');

  const sheet = getOrCreateFichaSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('paciente_id');

  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === pid) {
      targetRow = i + 1;
      break;
    }
  }

  body.updated_at = new Date().toISOString();
  const rowData = headers.map(h => body[h] !== undefined ? body[h] : '');

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return jsonOk({ success: true });
}

function getOrCreateFichaSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(FICHA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FICHA_SHEET_NAME);
    sheet.getRange(1, 1, 1, FICHA_HEADERS.length).setValues([FICHA_HEADERS]);
    sheet.getRange(1, 1, 1, FICHA_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Teste de diagnóstico - acesse via ?action=test_feegow&paciente_id=1001962
function handleFeegowTest(params) {
  const pid = params.paciente_id || '';
  if (!pid) return jsonErr('paciente_id obrigatório');

  const results = {};

  // Test 1: Patient search
  try {
    const r1 = UrlFetchApp.fetch(FEEGOW_API_BASE + '/patient/search?paciente_id=' + pid, {
      headers: { 'x-access-token': FEEGOW_API_TOKEN }, muteHttpExceptions: true
    });
    results.patient = JSON.parse(r1.getContentText());
  } catch(e) { results.patient_error = e.message; }

  // Test 2: Appointments (max 6 months interval)
  try {
    var now = new Date();
    var fiveAgo = new Date(now);
    fiveAgo.setMonth(fiveAgo.getMonth() - 5);
    var pad = function(n) { return n < 10 ? '0'+n : ''+n; };
    var ds = pad(fiveAgo.getDate()) + '-' + pad(fiveAgo.getMonth()+1) + '-' + fiveAgo.getFullYear();
    var de = pad(now.getDate()) + '-' + pad(now.getMonth()+1) + '-' + now.getFullYear();
    const r2 = UrlFetchApp.fetch(FEEGOW_API_BASE + '/appoints/search?paciente_id=' + pid + '&data_start=' + ds + '&data_end=' + de, {
      headers: { 'x-access-token': FEEGOW_API_TOKEN }, muteHttpExceptions: true
    });
    results.appoints = JSON.parse(r2.getContentText());
  } catch(e) { results.appoints_error = e.message; }

  // Test 3: Proposals
  try {
    const r3 = UrlFetchApp.fetch(FEEGOW_API_BASE + '/proposal/list?paciente_id=' + pid, {
      headers: { 'x-access-token': FEEGOW_API_TOKEN }, muteHttpExceptions: true
    });
    results.proposals = JSON.parse(r3.getContentText());
  } catch(e) { results.proposals_error = e.message; }

  // Test 4: Laudos (needs start_date and end_date in YYYY-MM-DD)
  try {
    var twoYearsAgo = new Date();
    twoYearsAgo.setFullYear(twoYearsAgo.getFullYear() - 2);
    var ldStart = twoYearsAgo.toISOString().split('T')[0];
    var ldEnd = new Date().toISOString().split('T')[0];
    const r4 = UrlFetchApp.fetch(FEEGOW_API_BASE + '/medical-reports/get-laudos-list?patient_id=' + pid + '&start_date=' + ldStart + '&end_date=' + ldEnd, {
      headers: { 'x-access-token': FEEGOW_API_TOKEN }, muteHttpExceptions: true
    });
    results.laudos = JSON.parse(r4.getContentText());
  } catch(e) { results.laudos_error = e.message; }

  // Test 5: Exam requests
  try {
    const r5 = UrlFetchApp.fetch(FEEGOW_API_BASE + '/patient/exam-requests?paciente_id=' + pid, {
      headers: { 'x-access-token': FEEGOW_API_TOKEN }, muteHttpExceptions: true
    });
    results.exams = JSON.parse(r5.getContentText());
  } catch(e) { results.exams_error = e.message; }

  return ContentService
    .createTextOutput(JSON.stringify(results, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonErr(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
