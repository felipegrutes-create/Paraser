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
  'parc_pago','quitado','projeto_ana','observacoes','classificacao','ultima_atualizacao'
];
const LOG_HEADERS = [
  'timestamp','acao','paciente_key','usuario',
  'dados_anteriores','coluna_origem','coluna_destino'
];
const DONOR_SHEET = 'Doadoras_Perfis';
const CAIXA_SHEET_NAME = 'Caixa_Recepcao';
const CAIXA_HEADERS = [
  'timestamp','data','paciente_nome','profissional','procedimento',
  'valor_total','valor_antecipado','valor_pago','forma_pagamento',
  'observacoes','registrado_por','nf_emitida','nf_numero'
];

// Mapa estático de procedimentos — preencha caso o list-sales não traga nomes
// Exemplo: var PROCEDURE_NAMES = { '4': 'FIV com Óvulos Próprios', '7': 'Criopreservação' };
var PROCEDURE_NAMES = {};
const BACKUP_SHEET_NAME = 'Backup_Agendamentos';
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
    if (action === 'test_financial') {
      return handleTestFinancial(e.parameter);
    }
    if (action === 'get_financial') {
      return handleGetFinancial(e.parameter);
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

    // Buscar respostas do questionário de anamnese por nome
    if (action === 'get_form_responses') {
      return handleGetFormResponses(e.parameter);
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

    // Retornar pagamentos do Caixa Recepção
    if (action === 'get_caixa') {
      const cxSheet = getOrCreateCaixaSheet();
      const cxData = cxSheet.getDataRange().getValues();
      if (cxData.length <= 1) return jsonOk({ pagamentos: [] });
      const cxHeaders = cxData[0];
      const filterDate = (e && e.parameter && e.parameter.data) || '';
      const pagamentos = cxData.slice(1).filter(row => {
        if (filterDate) {
          const rowDate = String(row[cxHeaders.indexOf('data')] || '');
          return rowDate === filterDate;
        }
        return true;
      }).map(row => {
        const obj = {};
        cxHeaders.forEach((h, i) => { obj[h] = row[i]; });
        return obj;
      });
      return jsonOk({ pagamentos });
    }

    // Retornar backup de agendamentos
    if (action === 'get_backup') {
      const bSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BACKUP_SHEET_NAME);
      if (!bSheet) return jsonOk({ backup: null });
      const bData = bSheet.getDataRange().getValues();
      const meta = bSheet.getRange('A1').getNote(); // timestamp no note da célula A1
      return jsonOk({ backup: bData, timestamp: meta || '' });
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
      if (action === 'save_caixa') {
        return handleSaveCaixa(body);
      }

      if (action === 'update_caixa_nf') {
        return handleUpdateCaixaNF(body);
      }

      if (action === 'save_backup') {
        return handleSaveBackup(body);
      }

      if (action === 'save_donor') {
        return handleSaveDonor(body);
      }

      if (action === 'delete_donor') {
        return handleDeleteDonor(body);
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
  } else {
    // Add missing columns (e.g. projeto_ana added later)
    const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    CRM_HEADERS.forEach((h, i) => {
      if (!existing.includes(h)) {
        const newCol = existing.length + 1;
        sheet.getRange(1, newCol).setValue(h).setFontWeight('bold');
        existing.push(h);
      }
    });
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
  const pacienteNome = (body.paciente_nome || '').trim();

  if (!base64Data) return jsonErr('foto obrigatória');

  // Remove data URI prefix: "data:image/jpeg;base64,..."
  const parts = base64Data.split(',');
  const raw = parts.length > 1 ? parts[1] : parts[0];
  const mimeMatch = base64Data.match(/data:([^;]+);/);
  const mime = mimeMatch ? mimeMatch[1] : 'image/jpeg';

  const decoded = Utilities.base64Decode(raw);
  const blob = Utilities.newBlob(decoded, mime, fileName);

  // Pasta raiz → subpasta com nome da paciente
  const rootFolder = getOrCreatePhotoFolder();
  let targetFolder = rootFolder;

  if (pacienteNome) {
    // Buscar ou criar subpasta com nome da paciente
    const subFolders = rootFolder.getFoldersByName(pacienteNome);
    if (subFolders.hasNext()) {
      targetFolder = subFolders.next();
    } else {
      targetFolder = rootFolder.createFolder(pacienteNome);
    }
  }

  const file = targetFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const viewUrl = 'https://lh3.googleusercontent.com/d/' + fileId;

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

// Teste endpoints financeiros - acesse via ?action=test_financial&data_start=01/01/2026&data_end=31/01/2026
function handleTestFinancial(params) {
  const ds = params.data_start || '01/01/2026';
  const de = params.data_end   || '31/01/2026';
  const headers = { 'x-access-token': FEEGOW_API_TOKEN };
  const results = {};

  const ds2 = ds.split('/').reverse().join('-'); // 01/01/2026 → 2026-01-01
  const de2 = de.split('/').reverse().join('-');

  const base = 'https://api.feegow.com/v1/api';

  // Endpoint documentado: /financial/list-sales
  const eps = [
    '/financial/list-sales?date_start=' + ds2 + '&date_end=' + de2 + '&unidade_id=0',
  ];

  eps.forEach(ep => {
    try {
      const r = UrlFetchApp.fetch(base + ep, { headers, muteHttpExceptions: true });
      results[ep] = { status: r.getResponseCode(), body: r.getContentText().slice(0, 500) };
    } catch(e) {
      results[ep] = { error: e.message };
    }
  });

  return ContentService
    .createTextOutput(JSON.stringify(results, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

// =========================================================
// Contas a Receber — lista faturas com nome do paciente
// Acesse via ?action=get_financial&data_start=01/01/2026&data_end=31/01/2026
// =========================================================
function handleGetFinancial(params) {
  var ds = params.data_start || '';
  var de = params.data_end   || '';
  if (!ds || !de) return jsonErr('data_start e data_end obrigatórios (formato DD/MM/YYYY)');

  // Converte DD/MM/YYYY → YYYY-MM-DD
  var ds2 = ds.split('/').reverse().join('-');
  var de2 = de.split('/').reverse().join('-');

  var reqHeaders = { 'x-access-token': FEEGOW_API_TOKEN };
  var base = FEEGOW_API_BASE;

  // Passo 1: buscar faturas do período
  var invoicesUrl = base + '/financial/list-invoice?data_start=' + ds2 + '&data_end=' + de2 + '&tipo_transacao=C&unidade_id=0';
  var r1;
  try {
    r1 = UrlFetchApp.fetch(invoicesUrl, { headers: reqHeaders, muteHttpExceptions: true });
  } catch(e) {
    return jsonErr('Erro ao chamar list-invoice: ' + e.message);
  }
  if (r1.getResponseCode() !== 200) {
    return jsonErr('list-invoice retornou ' + r1.getResponseCode() + ': ' + r1.getContentText().slice(0, 400));
  }

  var invoicesData;
  try {
    invoicesData = JSON.parse(r1.getContentText());
  } catch(e) {
    return jsonErr('Erro ao parsear list-invoice: ' + e.message);
  }

  var invoices = (invoicesData && invoicesData.content) ? invoicesData.content : [];
  if (!invoices.length) return jsonOk({ success: true, total: 0, data: [] });

  // Passo 2: carregar cache de nomes (pacientes, procedimentos, produtos)
  var CACHE_SHEET = 'Cache_Nomes';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cacheSheet = ss.getSheetByName(CACHE_SHEET);
  var nameCache = {}; // tipo_id → nome
  if (cacheSheet) {
    var cacheData = cacheSheet.getDataRange().getValues();
    for (var ci = 1; ci < cacheData.length; ci++) {
      var ck = String(cacheData[ci][0] || '');
      var cv = String(cacheData[ci][1] || '');
      if (ck && cv) nameCache[ck] = cv;
    }
  } else {
    cacheSheet = ss.insertSheet(CACHE_SHEET);
    cacheSheet.getRange(1, 1, 1, 2).setValues([['chave', 'nome']]);
    cacheSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    cacheSheet.setFrozenRows(1);
  }

  // Coletar IDs únicos que precisam de lookup
  var procMap = {}, prodMap = {};
  Object.keys(PROCEDURE_NAMES).forEach(function(k) { procMap[k] = PROCEDURE_NAMES[k]; });
  var needFetchProc = [], needFetchProd = [], needFetchPat = [];

  invoices.forEach(function(invoice) {
    (invoice.itens || []).forEach(function(item) {
      var key = item.tipo + '_' + item.procedimento_id;
      if (nameCache[key]) {
        if (item.tipo === 'S') procMap[item.procedimento_id] = nameCache[key];
        if (item.tipo === 'M') prodMap[item.procedimento_id] = nameCache[key];
      } else {
        if (item.tipo === 'S' && item.procedimento_id && !procMap[item.procedimento_id]) needFetchProc.push(item.procedimento_id);
        if (item.tipo === 'M' && item.procedimento_id && !prodMap[item.procedimento_id]) needFetchProd.push(item.procedimento_id);
      }
    });
    var det = invoice.detalhes && invoice.detalhes[0];
    if (det && det.tipo_conta === 3 && det.conta_id) {
      var patKey = 'PAT_' + det.conta_id;
      if (!nameCache[patKey]) needFetchPat.push(det.conta_id);
    }
  });

  // Deduplica
  needFetchProc = [...new Set(needFetchProc)];
  needFetchProd = [...new Set(needFetchProd)];
  needFetchPat = [...new Set(needFetchPat)];

  var newCacheRows = [];

  // Buscar procedimentos não cacheados (em lotes de 30)
  for (var pi = 0; pi < needFetchProc.length; pi += 30) {
    var batch = needFetchProc.slice(pi, pi + 30);
    var reqs = batch.map(function(id) {
      return { url: base + '/procedures/list?procedimento_id=' + id, headers: reqHeaders, muteHttpExceptions: true };
    });
    try {
      var resps = UrlFetchApp.fetchAll(reqs);
      resps.forEach(function(resp, i) {
        if (resp.getResponseCode() === 200) {
          try {
            var pd = JSON.parse(resp.getContentText());
            var nome = pd && pd.content && pd.content[0] && pd.content[0].nome;
            if (nome) { procMap[batch[i]] = nome; newCacheRows.push(['S_' + batch[i], nome]); }
          } catch(e) {}
        }
      });
    } catch(e) { /* rate limit - skip */ }
  }

  // Buscar produtos não cacheados (em lotes de 30)
  for (var mi = 0; mi < needFetchProd.length; mi += 30) {
    var batchM = needFetchProd.slice(mi, mi + 30);
    var reqsM = batchM.map(function(id) {
      return { url: base + '/core/financial/base/product/list', method: 'post', contentType: 'application/json',
               payload: JSON.stringify({ id: parseInt(id, 10), page: 1, perPage: 1 }), headers: reqHeaders, muteHttpExceptions: true };
    });
    try {
      var respsM = UrlFetchApp.fetchAll(reqsM);
      respsM.forEach(function(resp, i) {
        if (resp.getResponseCode() === 200) {
          try {
            var pd = JSON.parse(resp.getContentText());
            var nome = pd && pd.data && pd.data[0] && pd.data[0].NomeProduto;
            if (nome) { prodMap[batchM[i]] = nome; newCacheRows.push(['M_' + batchM[i], nome]); }
          } catch(e) {}
        }
      });
    } catch(e) { /* rate limit - skip */ }
  }

  // Buscar pacientes não cacheados (em lotes de 30)
  var patNameMap = {};
  for (var pti = 0; pti < needFetchPat.length; pti += 30) {
    var batchP = needFetchPat.slice(pti, pti + 30);
    var reqsP = batchP.map(function(id) {
      return { url: base + '/patient/search?paciente_id=' + id, headers: reqHeaders, muteHttpExceptions: true };
    });
    try {
      var respsP = UrlFetchApp.fetchAll(reqsP);
      respsP.forEach(function(resp, i) {
        try {
          if (resp.getResponseCode() === 200) {
            var pd = JSON.parse(resp.getContentText());
            var nome = (pd && pd.content && pd.content.nome) || null;
            if (nome) { patNameMap[batchP[i]] = nome; newCacheRows.push(['PAT_' + batchP[i], nome]); }
          }
        } catch(e) {}
      });
    } catch(e) { /* rate limit - skip */ }
  }

  // Salvar novos nomes no cache
  if (newCacheRows.length > 0) {
    var lastRow = cacheSheet.getLastRow();
    cacheSheet.getRange(lastRow + 1, 1, newCacheRows.length, 2).setValues(newCacheRows);
  }

  // Montar nameMap (índice da fatura → nome paciente)
  var nameMap = {};
  invoices.forEach(function(invoice, idx) {
    var det = invoice.detalhes && invoice.detalhes[0];
    if (det && det.tipo_conta === 3 && det.conta_id) {
      var patKey = 'PAT_' + det.conta_id;
      nameMap[idx] = patNameMap[det.conta_id] || nameCache[patKey] || null;
    }
  });

  // Montar resultado final
  var result = invoices.map(function(invoice, idx) {
    var out = {};
    for (var k in invoice) {
      if (invoice.hasOwnProperty(k)) out[k] = invoice[k];
    }
    out.nomePaciente = nameMap.hasOwnProperty(idx) ? nameMap[idx] : null;
    // Enriquecer itens com categoria e nome resolvido
    out.itens = (invoice.itens || []).map(function(item) {
      var enriched = {};
      for (var ki in item) { if (item.hasOwnProperty(ki)) enriched[ki] = item[ki]; }
      if (item.tipo === 'S') {
        enriched.categoria = 'Procedimento';
        enriched.nomeResolvido = procMap[item.procedimento_id] || null;
      } else if (item.tipo === 'M') {
        enriched.categoria = 'Medicamento';
        enriched.nomeResolvido = prodMap[item.procedimento_id] || null;
      } else if (item.tipo === 'K') {
        enriched.categoria = 'Kit';
        enriched.nomeResolvido = item.descricao || null;
      } else {
        enriched.categoria = item.tipo || 'Outros';
        enriched.nomeResolvido = item.descricao || null;
      }
      return enriched;
    });
    return out;
  });

  return jsonOk({ success: true, total: result.length, data: result });
}

function handleDeleteDonor(body) {
  const id = (body.id || '').trim();
  if (!id) return jsonErr('id obrigatório para exclusão');

  const sheet = getOrCreateDonorSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  let targetRow = -1;
  let oldData = {};
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === id) {
      targetRow = i + 1;
      headers.forEach((h, j) => { oldData[h] = data[i][j]; });
      break;
    }
  }

  if (targetRow < 0) return jsonErr('Perfil não encontrado');

  // Gravar log antes de apagar
  const logSheet = getOrCreateLogSheet();
  logSheet.appendRow([
    new Date().toISOString(),
    'delete_donor',
    id,
    body.usuario || 'desconhecido',
    body.dados_excluidos || JSON.stringify(oldData),
    '',
    ''
  ]);

  // Remover a linha da planilha
  sheet.deleteRow(targetRow);

  return jsonOk({ success: true, logged: true });
}

function getOrCreateCaixaSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CAIXA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CAIXA_SHEET_NAME);
    sheet.getRange(1, 1, 1, CAIXA_HEADERS.length).setValues([CAIXA_HEADERS]);
    sheet.getRange(1, 1, 1, CAIXA_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function handleSaveCaixa(body) {
  const sheet = getOrCreateCaixaSheet();
  body.timestamp = new Date().toISOString();
  if (!body.data) {
    const now = new Date();
    body.data = String(now.getDate()).padStart(2,'0') + '/' + String(now.getMonth()+1).padStart(2,'0') + '/' + now.getFullYear();
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = headers.map(h => body[h] !== undefined ? body[h] : '');
  sheet.appendRow(rowData);
  return jsonOk({ success: true });
}

function handleUpdateCaixaNF(body) {
  const sheet = getOrCreateCaixaSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const tsCol = headers.indexOf('timestamp');
  const nfEmCol = headers.indexOf('nf_emitida');
  const nfNumCol = headers.indexOf('nf_numero');

  const ts = (body.timestamp || '').trim();
  if (!ts) return jsonErr('timestamp obrigatório');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][tsCol]).trim() === ts) {
      if (nfEmCol >= 0) sheet.getRange(i+1, nfEmCol+1).setValue(body.nf_emitida || 'Sim');
      if (nfNumCol >= 0) sheet.getRange(i+1, nfNumCol+1).setValue(body.nf_numero || '');
      return jsonOk({ success: true });
    }
  }
  return jsonErr('Registro não encontrado');
}

function handleSaveBackup(body) {
  const rows = body.rows;
  if (!rows || !rows.length) return jsonErr('rows obrigatório');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BACKUP_SHEET_NAME);

  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(BACKUP_SHEET_NAME);
  }

  // Escrever dados em batch (muito mais rápido que appendRow)
  if (rows.length > 0) {
    const maxCols = Math.max(...rows.map(r => (r || []).length));
    // Normalizar todas as linhas para ter o mesmo número de colunas
    const normalized = rows.map(r => {
      const row = r || [];
      while (row.length < maxCols) row.push('');
      return row;
    });
    sheet.getRange(1, 1, normalized.length, maxCols).setValues(normalized);
    sheet.setFrozenRows(1);
  }

  // Salvar timestamp como note na célula A1
  sheet.getRange('A1').setNote('Backup: ' + new Date().toISOString() + ' | ' + rows.length + ' linhas');

  return jsonOk({ success: true, rows_saved: rows.length });
}

// =========================================================
// Questionário de Anamnese — busca por nome na planilha do Forms
// =========================================================
function handleGetFormResponses(params) {
  var rawName = (params.nome || '').toString().trim();
  if (!rawName || rawName.length < 2) return jsonErr('nome deve ter pelo menos 2 caracteres');
  var name = rawName.toLowerCase();

  try {
    var FORM_SS_ID = '1h9Jhjpv-SRmOGX6rXypHgTy_9YaAhr2WvqA8G8EpwrY';
    var ss = SpreadsheetApp.openById(FORM_SS_ID);
    var allSheets = ss.getSheets();

    // Lista todas as abas para diagnóstico
    var sheetList = allSheets.map(function(s) {
      return s.getName() + '(id:' + s.getSheetId() + ')';
    });

    var FORM_SHEET_GID = 1452374010;
    var sheet = null;
    for (var si = 0; si < allSheets.length; si++) {
      if (allSheets[si].getSheetId() === FORM_SHEET_GID) { sheet = allSheets[si]; break; }
    }
    if (!sheet) sheet = ss.getSheetByName('Form_Responses1');
    if (!sheet) sheet = ss.getSheetByName('Respostas ao formulário 1');
    if (!sheet) sheet = allSheets[0];

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return jsonOk({ responses: [], total: 0, debug: 'sheet vazia', debug_all_sheets: sheetList });

    var headers = data[0].map(function(h) { return h ? h.toString().trim() : ''; });

    // Descoberta dinâmica da coluna de nome — não depende de índice fixo
    var NAME_COL = -1;
    var normalizeStr = function(s) {
      return s.toLowerCase()
               .replace(/[àáâãä]/g,'a').replace(/[èéêë]/g,'e')
               .replace(/[ìíîï]/g,'i').replace(/[òóôõö]/g,'o').replace(/[ùúûü]/g,'u')
               .replace(/[ç]/g,'c')
               .replace(/[^a-z0-9 ]/g,' ')
               .replace(/\s+/g,' ')
               .trim();
    };
    // Prioridade 1: "nome completo" ou "nome da doadora/paciente"
    for (var hi = 0; hi < headers.length; hi++) {
      var hn = normalizeStr(headers[hi]);
      if (hn.indexOf('nome completo') >= 0 || hn.indexOf('nome da doadora') >= 0 || hn.indexOf('nome da paciente') >= 0) {
        NAME_COL = hi; break;
      }
    }
    // Prioridade 2: qualquer coluna com "nome" exceto parceiro/companheiro/conjuge
    if (NAME_COL < 0) {
      for (var hi2 = 0; hi2 < headers.length; hi2++) {
        var hn2 = normalizeStr(headers[hi2]);
        if (hn2.indexOf('nome') >= 0 && !/parceiro|companheiro|conjuge|esposo|esposa|marido/.test(hn2)) {
          NAME_COL = hi2; break;
        }
      }
    }
    // Fallback: col 2 (padrão Google Forms com coleta de e-mail ativa)
    if (NAME_COL < 0) NAME_COL = 2;

    var nameParts = name.split(/\s+/).filter(function(p) { return p.length > 2; });

    // Amostra das 3 primeiras linhas de dados (cols 0..5) para diagnóstico
    var sampleRows = [];
    for (var si2 = 1; si2 <= Math.min(3, data.length - 1); si2++) {
      var sr = [];
      for (var sc = 0; sc < Math.min(6, data[si2].length); sc++) {
        var sv = data[si2][sc];
        sr.push(sv instanceof Date ? Utilities.formatDate(sv,'America/Sao_Paulo','dd/MM/yyyy') : (sv||'').toString());
      }
      sampleRows.push(sr);
    }

    var matches = [];
    for (var i = 1; i < data.length; i++) {
      var cell = data[i][NAME_COL];
      if (!cell) continue;
      var rowName = normalizeStr(cell.toString());
      if (!rowName) continue;
      var normSearch = normalizeStr(name);
      var hit = rowName.indexOf(normSearch) >= 0;
      if (!hit && nameParts.length >= 2) {
        var normParts = nameParts.map(normalizeStr);
        hit = normParts.every(function(p) { return rowName.indexOf(p) >= 0; });
      }
      if (!hit) continue;
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        var val = data[i][j];
        if (val instanceof Date) {
          obj[j] = Utilities.formatDate(val, 'America/Sao_Paulo', 'dd/MM/yyyy');
        } else {
          obj[j] = (val !== undefined && val !== null) ? val.toString() : '';
        }
      }
      matches.push(obj);
    }

    // Últimas 10 entradas da coluna de nome para diagnóstico
    var lastNames = [];
    for (var li = Math.max(1, data.length - 10); li < data.length; li++) {
      var lv = data[li][NAME_COL];
      lastNames.push((lv || '').toString().trim());
    }

    return jsonOk({ success: true, responses: matches, total: matches.length, headers: headers,
                    debug_rows: data.length, debug_sheet: sheet.getName(),
                    debug_all_sheets: sheetList,
                    debug_name_col: NAME_COL,
                    debug_name_col_header: headers[NAME_COL] || '?',
                    debug_sample: sampleRows,
                    debug_last_names: lastNames });
  } catch(err) {
    return jsonErr('Erro ao ler formulário: ' + err.message);
  }
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
