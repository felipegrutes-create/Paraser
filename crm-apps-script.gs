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
  'parc_pago','quitado','projeto_ana','observacoes','classificacao',
  'engravidou','desistiu','ultima_atualizacao'
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

    // Grade de Horários — salas por médico+dia (compartilhado entre todos os aparelhos)
    if (action === 'grade_salas_get') {
      var raw = PropertiesService.getScriptProperties().getProperty('GRADE_SALAS') || '{}';
      return jsonOk({ salas: JSON.parse(raw) });
    }
    if (action === 'grade_salas_set') {
      var med = String(e.parameter.med || '');
      var dia = String(e.parameter.dia || '');
      var sala = String(e.parameter.sala || '');
      if (!med || !dia) return jsonOk({ ok: false, error: 'faltou med/dia' });
      var lock = LockService.getScriptLock();
      try { lock.waitLock(8000); } catch (le) { return jsonOk({ ok: false, error: 'lock' }); }
      try {
        var props = PropertiesService.getScriptProperties();
        var map = JSON.parse(props.getProperty('GRADE_SALAS') || '{}');
        var key = med + '|' + dia;
        if (sala) map[key] = sala; else delete map[key];
        props.setProperty('GRADE_SALAS', JSON.stringify(map));
        return jsonOk({ ok: true, salas: map });
      } finally { lock.releaseLock(); }
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

    if (action === 'get_estoque') return handleGetEstoque();
    if (action === 'parse_nfs')   return handleParseNFs();

    // Meta comercial — barra da meta lida pelo dashboard
    if (action === 'get_meta') {
      const _mes = (e.parameter.mes || Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM'));
      const props = PropertiesService.getScriptProperties();
      let c = null; try { c = JSON.parse(props.getProperty('META_CACHE') || 'null'); } catch (ce) { c = null; }
      const agora = new Date().getTime();
      if (!c || c.mes !== _mes || (agora - (c.ts || 0)) > 1200000 || e.parameter.fresh) {
        const m = computarMetaMes_(_mes);
        c = { mes: m.mes, meta: m.meta, total: m.total, comercial: m.comercial, outros: m.outros,
              cartao: m.cartao, pixLinkado: m.pixLinkado, aConferirValor: m.aConferirValor,
              aConferirQtd: m.aConferirQtd, porVendedora: m.porVendedora, ts: agora };
        props.setProperty('META_CACHE', JSON.stringify(c));
      }
      return jsonOk({ ok: true, mes: c.mes, meta: c.meta, total: c.total, comercial: c.comercial,
        outros: c.outros, cartao: c.cartao, pixLinkado: c.pixLinkado,
        aConferirValor: c.aConferirValor, aConferirQtd: c.aConferirQtd,
        porVendedora: c.porVendedora, confirmado: c.comercial });
    }

    // Lista os PIX pendentes de conferência (não linkaram por CPF).
    if (action === 'get_aconferir') {
      const sh = getOrCreateSheetGen_(PIX_CONFERIR_SHEET, PIXC_HEADERS);
      const dd = sh.getDataRange().getValues();
      const Hc = {}; PIXC_HEADERS.forEach(function(h, i){ Hc[h] = i; });
      const itens = [];
      for (let i = 1; i < dd.length; i++) {
        if (String(dd[i][Hc.status]) === 'PENDENTE') {
          itens.push({ fitid: String(dd[i][Hc.fitid]), data: String(dd[i][Hc.data]), valor: Number(dd[i][Hc.valor]),
            pagador: String(dd[i][Hc.pagador]), cpf: String(dd[i][Hc.cpf]), sugestao: String(dd[i][Hc.sugestao]) });
        }
      }
      return jsonOk({ ok: true, itens: itens });
    }

    // Entradas manuais de card (paciente sem agenda, pedido pela recepção).
    if (action === 'get_entradas_manuais') return handleGetEntradasManuais();

    if (action === 'wpp_admin') return handleWppAdmin(e.parameter);

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

    // Webhook do Z-API (WhatsApp comercial): roteado por query param porque o
    // payload deles não tem 'action'. Fora do lock — não disputa com o dashboard.
    if (e && e.parameter && e.parameter.action === 'zapi_webhook') {
      return handleZapiWebhook(body, e.parameter);
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      if (action === 'ajuste_estoque')       return handleAjusteEstoque(body);
      if (action === 'enviar_pipeline_slack') return handleEnviarPipelineSlack(body);
      if (action === 'inventario_inicial')   return handleInventarioInicial(body);

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

      if (action === 'marcar_venda_fechada') return handleMarcarVendaFechada(body);
      if (action === 'set_meta')             return handleSetMeta(body);
      if (action === 'conferir_pix')         return handleConferirPix(body);
      if (action === 'add_entrada_manual')    return handleAddEntradaManual(body);
      if (action === 'remove_entrada_manual') return handleRemoveEntradaManual(body);

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

  // Carregar lista de profissionais p/ resolver item.executante_id → nome
  var profNameMap = {};
  try {
    var rProf = UrlFetchApp.fetch(base + '/professional/list', { headers: reqHeaders, muteHttpExceptions: true });
    if (rProf.getResponseCode() === 200) {
      var pj = JSON.parse(rProf.getContentText());
      (pj.content || []).forEach(function(p) {
        var pid = p.profissional_id || p.id;
        if (pid) profNameMap[pid] = p.nome || p.name || '';
      });
    }
  } catch(e) { /* sem prof → cai no fallback de nome */ }

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
    // Enriquecer itens com categoria, nome resolvido e nome do executante
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
      // executante (profissional que executou o procedimento) — usado p/ atribuir receita ao médico
      enriched.executanteNome = (item.executante_id && profNameMap[item.executante_id]) || null;
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

// =========================================================
// ESTOQUE DE MEDICAMENTOS
// =========================================================
const PASTA_NFS_XML_ID      = '1hxONrvuoUVCmwlwcSswWB90bcaXf0EmH';
const ESTOQUE_NF_SHEET      = 'Entradas_NF';
const ESTOQUE_AJUSTES_SHEET = 'Ajustes_Estoque';
const ESTOQUE_NF_HEADERS    = ['nf_numero','data_nf','fornecedor','cnpj_fornecedor','produto','quantidade','unidade','valor_unit','valor_total','arquivo_xml','importado_em'];
const ESTOQUE_AJUSTES_HEADERS = ['data','produto','quantidade','tipo','motivo','usuario','observacoes'];

function getOrCreateEstoqueNFSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ESTOQUE_NF_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ESTOQUE_NF_SHEET);
    sh.getRange(1, 1, 1, ESTOQUE_NF_HEADERS.length).setValues([ESTOQUE_NF_HEADERS]);
    sh.getRange(1, 1, 1, ESTOQUE_NF_HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getOrCreateAjustesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ESTOQUE_AJUSTES_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ESTOQUE_AJUSTES_SHEET);
    sh.getRange(1, 1, 1, ESTOQUE_AJUSTES_HEADERS.length).setValues([ESTOQUE_AJUSTES_HEADERS]);
    sh.getRange(1, 1, 1, ESTOQUE_AJUSTES_HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function parseNFeXML(content, filename) {
  try {
    const doc  = XmlService.parse(content);
    const root = doc.getRootElement();

    // NFS-e (serviços municipais) — ignora, não é NF-e de produto
    const rootName = root.getName();
    if (/CompNfse|ConsultarNfseResposta|GerarNfseResposta|nfse/i.test(rootName)) return { isNfse: true };

    const ns   = XmlService.getNamespace('http://www.portalfiscal.inf.br/nfe');

    // Suporta procNFe (wrapped) e NFe (direto)
    let infNFe = root.getChild('infNFe', ns);
    if (!infNFe) {
      const nfe = root.getChild('NFe', ns);
      if (nfe) infNFe = nfe.getChild('infNFe', ns);
    }
    if (!infNFe) return null;

    const ide  = infNFe.getChild('ide', ns);
    const emit = infNFe.getChild('emit', ns);
    if (!ide || !emit) return null;

    const nNF        = ide.getChildText('nNF', ns) || '';
    const dhEmi      = ide.getChildText('dhEmi', ns) || ide.getChildText('dEmi', ns) || '';
    const data_nf    = dhEmi ? dhEmi.substring(0, 10) : '';
    const fornecedor = emit.getChildText('xNome', ns) || '';
    const cnpj       = emit.getChildText('CNPJ', ns) || emit.getChildText('CPF', ns) || '';

    const dets = infNFe.getChildren('det', ns);
    const itens = [];
    dets.forEach(function(det) {
      const prod = det.getChild('prod', ns);
      if (!prod) return;
      const xProd  = prod.getChildText('xProd', ns) || '';
      const qCom   = parseFloat(prod.getChildText('qCom', ns)   || '0');
      const uCom   = prod.getChildText('uCom', ns) || '';
      const vUnCom = parseFloat(prod.getChildText('vUnCom', ns) || '0');
      const vProd  = parseFloat(prod.getChildText('vProd', ns)  || '0');
      if (xProd && qCom > 0) {
        itens.push({ produto: xProd, quantidade: qCom, unidade: uCom, valor_unit: vUnCom, valor_total: vProd });
      }
    });

    return { nNF, data_nf, fornecedor, cnpj, itens };
  } catch(e) {
    Logger.log('Erro ao parsear ' + filename + ': ' + e.message);
    return null;
  }
}

function handleParseNFs() {
  try {
    const pasta   = DriveApp.getFolderById(PASTA_NFS_XML_ID);
    const nfSheet = getOrCreateEstoqueNFSheet();

    // Coleta chaves já importadas (nf_numero + arquivo_xml) para evitar duplicatas
    const existData = nfSheet.getDataRange().getValues();
    const existKeys = new Set();
    for (var i = 1; i < existData.length; i++) {
      existKeys.add(String(existData[i][0]) + '|' + String(existData[i][9]));
    }

    const arquivos = pasta.getFiles(); // getFiles() pega todos os MIMEs (xml pode ser application/xml ou text/xml)
    const novas = [];
    var processados = 0, erros = 0, ignorados = 0;

    while (arquivos.hasNext()) {
      const arq  = arquivos.next();
      const nome = arq.getName();
      if (!nome.toLowerCase().endsWith('.xml')) continue;
      // Ignora CC-e e DF-e pelo nome; NFS-e é detectada pelo parser (sem infNFe)
      if (/-CCe|^DFE/i.test(nome)) { ignorados++; continue; }

      const content = arq.getBlob().getDataAsString('UTF-8');
      const parsed  = parseNFeXML(content, nome);
      if (!parsed) { erros++; continue; }
      if (parsed.isNfse) { ignorados++; continue; }
      if (!parsed.itens.length) { erros++; continue; }

      const agora = new Date().toISOString();
      parsed.itens.forEach(function(it) {
        const key = parsed.nNF + '|' + nome;
        if (existKeys.has(key)) return;
        novas.push([parsed.nNF, parsed.data_nf, parsed.fornecedor, parsed.cnpj,
                    it.produto, it.quantidade, it.unidade, it.valor_unit, it.valor_total,
                    nome, agora]);
        existKeys.add(key);
      });
      processados++;
    }

    if (novas.length > 0) {
      nfSheet.getRange(nfSheet.getLastRow() + 1, 1, novas.length, ESTOQUE_NF_HEADERS.length).setValues(novas);
    }

    return jsonOk({ success: true, processados, novas_linhas: novas.length, ignorados, erros });
  } catch(e) {
    return jsonErr('Erro ao processar NFs: ' + e.message);
  }
}

// Extrai as 2 primeiras palavras do nome como chave de matching (ex: "ORGALUTRAN 0,25MG/0,5ML")
function _estoqueKey(nome) {
  var partes = String(nome || '').trim().toUpperCase().split(/[\s–—\-]+/);
  return partes.slice(0, 2).filter(Boolean).join(' ');
}

function handleGetEstoque() {
  try {
    const nfData = getOrCreateEstoqueNFSheet().getDataRange().getValues();
    const ajData = getOrCreateAjustesSheet().getDataRange().getValues();

    const estoque = {};
    const keyToNF = {}; // chave curta → nome completo da NF (para matching de ajustes)

    // Soma entradas das NFs (agrupado pelo nome exato da NF)
    for (var i = 1; i < nfData.length; i++) {
      const row     = nfData[i];
      const produto = String(row[4] || '').trim();
      if (!produto) continue;
      if (!estoque[produto]) {
        estoque[produto] = { produto, unidade: String(row[6] || ''), qtd_nf: 0, qtd_ajuste: 0, valor_unit: 0, ultima_entrada: '', fornecedor: '' };
        const k = _estoqueKey(produto);
        if (!keyToNF[k]) keyToNF[k] = produto; // registra chave curta → nome NF
      }
      estoque[produto].qtd_nf    += parseFloat(row[5] || 0);
      estoque[produto].valor_unit = parseFloat(row[7] || 0);
      const d = String(row[1] || '');
      if (d > estoque[produto].ultima_entrada) {
        estoque[produto].ultima_entrada = d;
        estoque[produto].fornecedor     = String(row[2] || '');
      }
    }

    // Aplica ajustes — tenta match exato, depois match por chave curta
    for (var j = 1; j < ajData.length; j++) {
      const row     = ajData[j];
      const produto = String(row[1] || '').trim();
      if (!produto) continue;
      const qtd = parseFloat(row[2] || 0);

      if (estoque[produto]) {
        // Match exato
        estoque[produto].qtd_ajuste += qtd;
      } else {
        // Fallback: match por 2 primeiras palavras
        const nfMatch = keyToNF[_estoqueKey(produto)];
        if (nfMatch) {
          estoque[nfMatch].qtd_ajuste += qtd;
        } else {
          // Sem correspondência — cria entrada standalone
          estoque[produto] = { produto, unidade: '', qtd_nf: 0, qtd_ajuste: qtd, valor_unit: 0, ultima_entrada: '', fornecedor: '' };
        }
      }
    }

    const lista = Object.values(estoque).map(function(e) {
      return Object.assign({}, e, { saldo: e.qtd_nf + e.qtd_ajuste });
    }).sort(function(a, b) { return a.produto.localeCompare(b.produto); });

    const resumo = {
      total_itens:   lista.length,
      estoque_baixo: lista.filter(function(e) { return e.saldo > 0 && e.saldo <= 2; }).length,
      sem_estoque:   lista.filter(function(e) { return e.saldo <= 0; }).length,
      valor_total:   lista.reduce(function(s, e) { return s + (e.saldo > 0 ? e.saldo * e.valor_unit : 0); }, 0)
    };

    return jsonOk({ success: true, estoque: lista, resumo });
  } catch(e) {
    return jsonErr('Erro ao calcular estoque: ' + e.message);
  }
}

function handleAjusteEstoque(body) {
  try {
    const produto = (body.produto || '').trim();
    const qtd     = parseFloat(body.quantidade || 0);
    if (!produto)  return jsonErr('produto obrigatório');
    if (qtd === 0) return jsonErr('quantidade não pode ser zero');

    getOrCreateAjustesSheet().appendRow([
      body.data || new Date().toISOString().substring(0, 10),
      produto, qtd,
      body.tipo || 'ajuste_manual',
      body.motivo   || '',
      body.usuario  || '',
      body.observacoes || ''
    ]);

    return jsonOk({ success: true, message: 'Ajuste registrado: ' + produto + ' (' + (qtd > 0 ? '+' : '') + qtd + ')' });
  } catch(e) {
    return jsonErr('Erro ao registrar ajuste: ' + e.message);
  }
}

function handleInventarioInicial(body) {
  try {
    var dataInv = body.data || new Date().toISOString().substring(0, 10);
    var itens   = body.itens || [];

    var nfSheet = getOrCreateEstoqueNFSheet();
    var ajSheet = getOrCreateAjustesSheet();
    var nfData  = nfSheet.getDataRange().getValues();
    var ajData  = ajSheet.getDataRange().getValues();

    var novasLinhas = [];
    var resultados  = [];

    itens.forEach(function(item) {
      var busca   = (item.busca || '').toUpperCase().trim();
      var qtdReal = Number(item.quantidade);

      // Soma entradas de NF que contenham o termo de busca
      var saldoNF = 0;
      for (var i = 1; i < nfData.length; i++) {
        if (String(nfData[i][4] || '').toUpperCase().indexOf(busca) >= 0) {
          saldoNF += Number(nfData[i][5]) || 0;
        }
      }

      // Soma ajustes existentes
      var saldoAj = 0;
      for (var j = 1; j < ajData.length; j++) {
        if (String(ajData[j][1] || '').toUpperCase().indexOf(busca) >= 0) {
          saldoAj += Number(ajData[j][2]) || 0;
        }
      }

      var saldoAtual = saldoNF + saldoAj;
      var diff       = qtdReal - saldoAtual;

      novasLinhas.push([dataInv, item.nome, diff, 'inventario_inicial',
                        'Contagem física ' + dataInv, 'dashboard', '']);
      resultados.push({ nome: item.nome, busca: busca, qtdReal: qtdReal,
                        saldoNF: saldoNF, saldoAj: saldoAj, saldoAtual: saldoAtual, diff: diff });
    });

    if (novasLinhas.length > 0) {
      ajSheet.getRange(ajSheet.getLastRow() + 1, 1, novasLinhas.length,
                       ESTOQUE_AJUSTES_HEADERS.length).setValues(novasLinhas);
    }

    return jsonOk({ success: true, ajustes: novasLinhas.length, resultados: resultados });
  } catch(e) {
    return jsonErr('Erro no inventário: ' + e.message);
  }
}

function handleEnviarPipelineSlack(body) {
  try {
    const webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_COMERCIAL_WEBHOOK');
    if (!webhookUrl) return jsonErr('SLACK_COMERCIAL_WEBHOOK não configurado. Adicione nas propriedades do script.');

    const periodo    = body.periodo || '';
    const itens      = body.itens   || [];
    const total      = body.total   || 0;

    const fmtR = function(cents) {
      return 'R$ ' + (cents / 100).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    };

    var linhasTodas = itens.map(function(it) {
      var data = (it.data || '').replace(/-/g, '/');
      var proc = (it.procedimento || '—').substring(0, 40);
      var val  = it.valor > 0 ? fmtR(it.valor) : '_sem valor_';
      return '🔸 *' + it.paciente + '* — ' + proc + ' — ' + val + ' — ' + data;
    });

    // Envia cabeçalho + itens em blocos de 30 (limite seguro do Slack)
    var cabecalho = '📊 *Pipeline Comercial — ' + periodo + '*\n*' + itens.length + ' orçamentos · ' + fmtR(total) + '*';
    var blocos = [];
    for (var i = 0; i < linhasTodas.length; i += 30) {
      blocos.push(linhasTodas.slice(i, i + 30).join('\n'));
    }

    // Primeira mensagem: cabeçalho + bloco 1
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: cabecalho + '\n\n' + (blocos[0] || '') })
    });

    // Demais blocos como mensagens separadas
    for (var j = 1; j < blocos.length; j++) {
      Utilities.sleep(500);
      UrlFetchApp.fetch(webhookUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ text: blocos[j] })
      });
    }

    return jsonOk({ success: true, enviados: itens.length });
  } catch(e) {
    return jsonErr('Erro ao enviar para Slack: ' + e.message);
  }
}

// =========================================================
// META COMERCIAL — leitor de extrato OFX do Itaú (PIX recebido)
// A analista larga o arquivo .ofx numa pasta do Drive; o sistema lê
// e extrai os PIX recebidos (valor + data + pagador + CPF) pra casar
// com as vendas marcadas como "Venda fechada". Cartão NÃO vem por aqui
// (vem da API da REDE): o OFX só traz a liquidação em lote, atrasada.
// =========================================================
const OFX_FOLDER_PROP = 'OFX_FOLDER_ID';
const OFX_FOLDER_NAME = 'Extratos Itau OFX';

function getOrCreateOfxFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(OFX_FOLDER_PROP);
  if (id) {
    try { return DriveApp.getFolderById(id); } catch(e) { /* sumiu, recria abaixo */ }
  }
  const folder = DriveApp.createFolder(OFX_FOLDER_NAME);
  props.setProperty(OFX_FOLDER_PROP, folder.getId());
  return folder;
}

// Cria a pasta (1x) e mostra o link. Rodar manual no editor.
function setupPastaOfx() {
  const f = getOrCreateOfxFolder_();
  Logger.log('Pasta OFX pronta: ' + f.getName());
  Logger.log('Link (largue os .ofx aqui): ' + f.getUrl());
  Logger.log('ID: ' + f.getId());
  return f.getUrl();
}

// Lê todos os .ofx da pasta e devolve os PIX recebidos.
function lerOfxRecebimentos_(folder) {
  folder = folder || getOrCreateOfxFolder_();
  const files = folder.getFiles();
  const recebidos = [];
  while (files.hasNext()) {
    const file = files.next();
    if (!/\.ofx$/i.test(file.getName())) continue;
    const texto = file.getBlob().getDataAsString('ISO-8859-1');
    parseOfxPix_(texto, file.getName()).forEach(function(r) { recebidos.push(r); });
  }
  return recebidos;
}

// Extrai PIX RECEBIDO / PIX QR CODE RECEBIDO de um conteúdo OFX.
// Tudo está no MEMO: "PIX RECEBIDO <nomecurto><dd/mm> <NOME COMPLETO> <CPF>".
function parseOfxPix_(texto, arquivo) {
  const out = [];
  const blocos = texto.split('<STMTTRN>').slice(1);
  blocos.forEach(function(b) {
    const tag = function(t) {
      const m = b.match(new RegExp('<' + t + '>([^\\r\\n<]*)'));
      return m ? m[1].trim() : '';
    };
    const valor = parseFloat(tag('TRNAMT'));
    const memo  = tag('MEMO');
    const up    = memo.toUpperCase();
    const ehPix = up.indexOf('PIX RECEBIDO') === 0 || up.indexOf('PIX QR CODE RECEBIDO') === 0;
    if (!ehPix || !(valor > 0)) return;
    const dt  = tag('DTPOSTED').slice(0, 8); // yyyymmdd
    const cpf = (memo.match(/(\d{3}\.\d{3}\.\d{3}-\d{2})/) || [])[1] || '';
    // nome do pagador: tira o prefixo, o "dd/mm" grudado e o CPF
    let nome = memo.replace(/^PIX( QR CODE)? RECEBIDO\s*/i, '');
    if (cpf) nome = nome.split(cpf)[0];
    const md = nome.match(/\d{2}\/\d{2}/);
    if (md) nome = nome.slice(nome.indexOf(md[0]) + md[0].length);
    nome = nome.replace(/\s+/g, ' ').trim();
    out.push({
      data:    dt ? (dt.slice(6,8) + '/' + dt.slice(4,6) + '/' + dt.slice(0,4)) : '',
      valor:   valor,
      pagador: nome,
      cpf:     cpf,
      fitid:   tag('FITID'),
      arquivo: arquivo
    });
  });
  return out;
}

// Teste: lê a pasta e loga o resumo. Rodar manual no editor depois de largar o OFX.
function testarOfx() {
  const recebidos = lerOfxRecebimentos_();
  let soma = 0;
  recebidos.forEach(function(r) { soma += r.valor; });
  Logger.log('PIX recebidos encontrados: ' + recebidos.length + ' = R$' + soma.toFixed(2));
  recebidos.slice(0, 12).forEach(function(r) {
    Logger.log('  ' + r.data + '  R$' + r.valor.toFixed(2) + '  ' + r.pagador + '  [' + r.cpf + ']');
  });
  return recebidos.length;
}

// =========================================================
// PIX via API do Itaú (lido do BigQuery abastecido pelo coletor Cloud Run).
// O coletor puxa o extrato das contas de hora em hora e grava em
// paraser-extrato-itau.extrato.lancamentos. Aqui lemos SÓ os PIX recebidos
// (crédito via PIX_RECEPCAO) no MESMO formato do parser OFX, então o motor de
// conciliação não muda. Toggle Script Property PIX_FONTE: 'bigquery' | 'ofx'.
// =========================================================
const BQ_PROJECT = 'paraser-extrato-itau';
const BQ_PIX_TABLE = '`paraser-extrato-itau.extrato.lancamentos`';

// Fonte do PIX: 'bigquery' = API Itaú via BigQuery; qualquer outro = OFX (pasta Drive).
function lerPixRecebimentos_(startIso, endIso) {
  const fonte = PropertiesService.getScriptProperties().getProperty('PIX_FONTE') || 'ofx';
  if (fonte === 'bigquery') return lerPixBigQuery_(startIso, endIso);
  return lerOfxRecebimentos_(); // OFX lê a pasta inteira (ignora as datas), comportamento antigo
}

// Consulta o BigQuery e devolve os PIX recebidos no shape do OFX:
// { data:'dd/mm/yyyy', valor:Number, pagador, cpf, fitid }.
function lerPixBigQuery_(startIso, endIso) {
  const sql =
    "SELECT FORMAT_DATE('%d/%m/%Y', data_contabil) AS data, valor, " +
    "contraparte_nome AS pagador, " +
    "IF(REGEXP_CONTAINS(contraparte_documento, r'^[0-9]{3}\\.[0-9]{3}\\.[0-9]{3}-[0-9]{2}$'), contraparte_documento, '') AS cpf, " +
    "id AS fitid FROM " + BQ_PIX_TABLE + " " +
    "WHERE operation = 'C' AND origem_operacao = 'PIX_RECEPCAO' " +
    "AND data_contabil BETWEEN @start AND @end";
  const req = {
    query: sql,
    useLegacySql: false,
    parameterMode: 'NAMED',
    timeoutMs: 30000,
    queryParameters: [
      { name: 'start', parameterType: { type: 'DATE' }, parameterValue: { value: startIso } },
      { name: 'end',   parameterType: { type: 'DATE' }, parameterValue: { value: endIso } }
    ]
  };
  let res = BigQuery.Jobs.query(req, BQ_PROJECT);
  const jobId = res.jobReference.jobId;
  const loc = res.jobReference.location;
  // A query pode não terminar dentro do timeout: nesse caso jobComplete=false e
  // rows vem vazio. Esperar terminar (senão a meta viria subestimada e silenciosa).
  let tentativas = 0;
  while (!res.jobComplete) {
    if (++tentativas > 30) throw new Error('BigQuery: consulta do PIX não completou a tempo');
    Utilities.sleep(1000);
    res = BigQuery.Jobs.getQueryResults(BQ_PROJECT, jobId, { location: loc, timeoutMs: 30000 });
  }
  // Concatena todas as páginas (pageToken) — não perder lançamento em silêncio.
  let rows = res.rows || [];
  let pageToken = res.pageToken;
  while (pageToken) {
    res = BigQuery.Jobs.getQueryResults(BQ_PROJECT, jobId, { location: loc, pageToken: pageToken });
    rows = rows.concat(res.rows || []);
    pageToken = res.pageToken;
  }
  const out = [];
  rows.forEach(function(row) {
    const f = row.f;
    out.push({
      data:    f[0].v,
      valor:   Number(f[1].v),
      pagador: f[2].v || '',
      cpf:     f[3].v || '',
      fitid:   f[4].v
    });
  });
  return out;
}

// Janela de datas pra buscar o PIX das vendas PIX ainda AGUARDANDO (menor data -3d .. hoje).
function janelaPixPendentes_(pendentes, H) {
  const hojeIso = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
  const hoje = diaNum_(hojeIso);
  const dias = pendentes
    .filter(function(p) { return String(p.v[H.forma_pgto]) === 'PIX'; })
    .map(function(p) { return diaNum_(normData_(p.v[H.data_venda])); })
    .filter(function(n) { return !isNaN(n); });
  if (!dias.length) return { start: diaToIso_(hoje - 2), end: hojeIso };
  return { start: diaToIso_(Math.min.apply(null, dias) - 3), end: hojeIso };
}

// Primeiro e último dia de um mês 'yyyy-MM'.
function inicioDoMesIso_(mesRe) { return String(mesRe).slice(0, 7) + '-01'; }
function fimDoMesIso_(mesRe) {
  const p = String(mesRe).slice(0, 7).split('-');
  const d = new Date(Number(p[0]), Number(p[1]), 0); // dia 0 do mês seguinte = último dia deste
  return Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd');
}

// Liga/desliga a fonte BigQuery. Rodar no editor 1x (também dispara a autorização do BigQuery).
function ativarPixBigQuery() {
  PropertiesService.getScriptProperties().setProperty('PIX_FONTE', 'bigquery');
  return 'PIX_FONTE = bigquery (API Itaú)';
}
function desativarPixBigQuery() {
  PropertiesService.getScriptProperties().setProperty('PIX_FONTE', 'ofx');
  return 'PIX_FONTE = ofx (fallback pasta Drive)';
}
// Teste manual no editor: lê o PIX do mês atual do BigQuery e loga o resumo.
function testarPixBigQuery() {
  const mes = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM');
  const r = lerPixBigQuery_(inicioDoMesIso_(mes), fimDoMesIso_(mes));
  let soma = 0; r.forEach(function(x) { soma += x.valor; });
  Logger.log('PIX (BigQuery) em ' + mes + ': ' + r.length + ' = R$' + soma.toFixed(2));
  r.slice(0, 12).forEach(function(x) {
    Logger.log('  ' + x.data + '  R$' + Number(x.valor).toFixed(2) + '  ' + x.pagador + '  [' + x.cpf + ']');
  });
  return r.length;
}

// =========================================================
// META COMERCIAL — conector REDE (Gestão de Vendas / cartão)
// Cartão conta "quando a transação passa": usamos amount + saleDate
// das vendas APPROVED. Token OAuth client_credentials (Bearer, ~24min).
// Credenciais em Script Properties (REDE_CLIENT_ID/SECRET/PV), nunca no código.
// =========================================================
const REDE_BASE = 'https://rl7-sandbox-api.useredecloud.com.br'; // SANDBOX — trocar p/ produção depois

// =========================================================
// META COMERCIAL — leitor do CSV de vendas da Rede (cartão, PONTE)
// A analista exporta o relatório de vendas no portal da Rede e larga o
// CSV numa pasta do Drive. Lemos as vendas APROVADAS (valor original +
// data) pra casar com as vendas marcadas, igual o OFX faz com o PIX.
// Ponte enquanto a API de produção não libera (REDE_FONTE=arquivo).
// CSV: separador ';', UTF-8, valores pt-BR (9.750,00).
// =========================================================
const REDE_FOLDER_PROP = 'REDE_FOLDER_ID';
const REDE_FOLDER_NAME = 'Vendas Rede';

function getOrCreateRedeFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(REDE_FOLDER_PROP);
  if (id) { try { return DriveApp.getFolderById(id); } catch(e) { /* recria */ } }
  const folder = DriveApp.createFolder(REDE_FOLDER_NAME);
  props.setProperty(REDE_FOLDER_PROP, folder.getId());
  return folder;
}
function setupPastaRede() {
  const f = getOrCreateRedeFolder_();
  Logger.log('Pasta Vendas Rede: ' + f.getUrl());
  return f.getUrl();
}

function parseValorBR_(s) {
  s = String(s || '').replace(/[^\d.,-]/g, '');
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

// Acha colunas pelo nome (sem acento), filtra status "aprovada" e não cancelada.
function parseRedeCsv_(texto) {
  const linhas = texto.split(/\r?\n/).filter(function(l){ return l.trim(); });
  if (linhas.length < 2) return [];
  const norm = function(s){ return String(s || '').trim().toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu, ""); };
  const head = linhas[0].split(';').map(norm);
  const idx = function(nome){ return head.indexOf(norm(nome)); };
  const cData = idx('data da venda'), cStatus = idx('status da venda'),
        cValor = idx('valor da venda original'), cParc = idx('numero de parcelas'),
        cMod = idx('modalidade'), cNsu = idx('nsu/cv'),
        cPv = idx('numero do estabelecimento'), cCanc = idx('cancelada pelo estabelecimento');
  const out = [];
  for (var i = 1; i < linhas.length; i++) {
    const c = linhas[i].split(';');
    if (norm(c[cStatus]) !== 'aprovada') continue;
    if (cCanc >= 0 && norm(c[cCanc]) === 'sim') continue;
    const valor = parseValorBR_(c[cValor]);
    if (!(valor > 0)) continue;
    out.push({
      valor: valor, data: (c[cData] || '').trim(),
      parcelas: cParc >= 0 ? (c[cParc] || '').trim() : '',
      modalidade: cMod >= 0 ? (c[cMod] || '').trim() : '',
      nsu: cNsu >= 0 ? (c[cNsu] || '').trim() : '',
      pv: cPv >= 0 ? (c[cPv] || '').trim() : ''
    });
  }
  return out;
}

// Lê todos os .csv da pasta e devolve as vendas aprovadas no formato do casamento.
function lerRedeRecebimentos_(folder) {
  folder = folder || getOrCreateRedeFolder_();
  const files = folder.getFiles();
  const out = [];
  while (files.hasNext()) {
    const file = files.next();
    if (!/\.csv$/i.test(file.getName())) continue;
    parseRedeCsv_(file.getBlob().getDataAsString('UTF-8')).forEach(function(r){
      out.push({
        valor: r.valor, dia: diaNum_(normData_(r.data)),
        key: 'rede:' + r.nsu + ':' + r.pv,
        quem: r.modalidade + (Number(r.parcelas) > 1 ? ' ' + r.parcelas + 'x' : '')
      });
    });
  }
  return out;
}

function redeToken_() {
  const p = PropertiesService.getScriptProperties();
  const cid = p.getProperty('REDE_CLIENT_ID'), sec = p.getProperty('REDE_CLIENT_SECRET');
  if (!cid || !sec) throw new Error('REDE_CLIENT_ID/REDE_CLIENT_SECRET não configurados');
  const resp = UrlFetchApp.fetch(REDE_BASE + '/oauth2/token', {
    method: 'post',
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(cid + ':' + sec) },
    payload: { grant_type: 'client_credentials' },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('REDE token HTTP ' + resp.getResponseCode() + ': ' + resp.getContentText().slice(0, 200));
  return JSON.parse(resp.getContentText()).access_token;
}

// Vendas APROVADAS de um dia (Date). Uma página (size 200) basta p/ volume diário de clínica.
function redeVendasDia_(token, dia) {
  const p = PropertiesService.getScriptProperties();
  const pvRaw = p.getProperty('REDE_PV');
  if (!pvRaw) throw new Error('REDE_PV não configurado');
  const ds = Utilities.formatDate(dia, 'America/Sao_Paulo', 'yyyy-MM-dd');
  // A Paraser tem mais de um PV (PARASER SERVICOS + INSTITUTO) — REDE_PV pode ser
  // vários separados por vírgula; consulta cada um e junta.
  const pvs = pvRaw.split(',').map(function(s){ return s.trim(); }).filter(Boolean);
  const out = [];
  pvs.forEach(function(pv) {
    const url = REDE_BASE + '/merchant-statement/v1/sales?parentCompanyNumber=' + pv +
                '&subsidiaries=' + pv + '&startDate=' + ds + '&endDate=' + ds + '&status=APPROVED&size=200';
    const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) throw new Error('REDE vendas HTTP ' + resp.getResponseCode() + ' (PV ' + pv + '): ' + resp.getContentText().slice(0, 200));
    const j = JSON.parse(resp.getContentText());
    const txns = (j.content && j.content.transactions) || [];
    if (j.cursor && j.cursor.hasNextKey) Logger.log('REDE ' + ds + ' PV ' + pv + ': >200 vendas (paginação não tratada)');
    txns.forEach(function(t) {
      out.push({ valor: t.amount, data: t.saleDate, modalidade: t.modality && t.modality.type, parcelas: t.installmentQuantity, nsu: t.nsu, pv: pv });
    });
  });
  return out;
}

// =========================================================
// META COMERCIAL — vendas fechadas + meta mensal + conciliação
// =========================================================
const VENDAS_FECHADAS_SHEET = 'Vendas_Fechadas';
const VF_HEADERS = ['id','timestamp','data_venda','vendedora','paciente','valor','forma_pgto','status','confirmado_em','match_info','mes'];
const METAS_SHEET = 'Metas';
const METAS_HEADERS = ['mes','valor'];

function getOrCreateSheetGen_(nome, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// data_venda em qualquer forma (YYYY-MM-DD ou DD/MM/YYYY) -> 'YYYY-MM-DD'
function normData_(s) {
  // Sheets converte "2026-06-01" em objeto Date ao gravar — tratar isso.
  if (Object.prototype.toString.call(s) === '[object Date]') {
    return Utilities.formatDate(s, 'America/Sao_Paulo', 'yyyy-MM-dd');
  }
  s = String(s || '').trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return m[3] + '-' + m[2] + '-' + m[1];
  return s;
}
function diaNum_(iso) { // 'YYYY-MM-DD' -> número de dias (p/ janela)
  const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return NaN;
  return Math.floor(Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000);
}
// mês 'YYYY-MM' robusto a coerção do Sheets (que vira Date)
function normMes_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, 'America/Sao_Paulo', 'yyyy-MM');
  return String(v || '').trim().slice(0, 7);
}

// Vendedora marca a venda fechada (a transação passou). Fica AGUARDANDO confirmação.
function handleMarcarVendaFechada(body) {
  const sh = getOrCreateSheetGen_(VENDAS_FECHADAS_SHEET, VF_HEADERS);
  const dataVenda = normData_(body.data_venda || new Date());
  const valor = Number(body.valor) || 0;
  if (!valor) return jsonErr('valor obrigatório');
  const forma = (String(body.forma_pgto || '').toUpperCase().indexOf('PIX') >= 0) ? 'PIX' : 'CARTAO';
  const row = {
    id: Utilities.getUuid(), timestamp: new Date().toISOString(), data_venda: dataVenda,
    vendedora: body.vendedora || '', paciente: body.paciente || '', valor: valor,
    forma_pgto: forma, status: 'AGUARDANDO', confirmado_em: '', match_info: '',
    mes: dataVenda.slice(0, 7)
  };
  sh.appendRow(VF_HEADERS.map(function(h){ return row[h]; }));
  return jsonOk({ ok: true, id: row.id });
}

function getMetaMes_(mes) {
  const sh = getOrCreateSheetGen_(METAS_SHEET, METAS_HEADERS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) if (normMes_(data[i][0]) === mes) return Number(data[i][1]) || 0;
  return 0;
}
function handleSetMeta(body) {
  const mes = String(body.mes || '').trim(); // 'YYYY-MM'
  const valor = Number(body.valor) || 0;
  if (!/^\d{4}-\d{2}$/.test(mes)) return jsonErr('mes deve ser YYYY-MM');
  const sh = getOrCreateSheetGen_(METAS_SHEET, METAS_HEADERS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normMes_(data[i][0]) === mes) { sh.getRange(i + 1, 2).setValue(valor); return jsonOk({ ok: true, mes: mes, valor: valor }); }
  }
  sh.appendRow(["'" + mes, valor]);
  return jsonOk({ ok: true, mes: mes, valor: valor });
}

// Normaliza nome pra comparação (minúsculo, sem acento, só letras).
function normNome_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/\p{Diacritic}/gu, '')
    .replace(/[^a-z ]/g, ' ').replace(/\s+/g, ' ').trim();
}
// Quantos nomes (tokens >=3 letras) do paciente aparecem no nome do pagador.
function scoreNome_(paciente, pagador) {
  if (!paciente || !pagador) return 0;
  const tp = paciente.split(' ').filter(function(t){ return t.length >= 3; });
  const tg = ' ' + pagador + ' ';
  let n = 0;
  tp.forEach(function(t){ if (tg.indexOf(' ' + t + ' ') >= 0) n++; });
  return n;
}

// Casa as vendas AGUARDANDO com PIX (OFX) e cartão (REDE). Marca CONFIRMADA.
// postarSlack=false por padrão (modo teste). Janela de ±2 dias, valor exato.
// Desempate: se >1 candidato bate valor+data, o nome do pagador (só PIX) desempata.
function conciliarVendasFechadas_(postarSlack) {
  const sh = getOrCreateSheetGen_(VENDAS_FECHADAS_SHEET, VF_HEADERS);
  const data = sh.getDataRange().getValues();
  const H = {}; VF_HEADERS.forEach(function(h, i){ H[h] = i; });
  const pendentes = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][H.status]) === 'AGUARDANDO') pendentes.push({ row: i + 1, v: data[i] });
  }
  if (!pendentes.length) return { confirmadas: 0, pendentes: 0, detalhes: [] };

  const janelaPix = janelaPixPendentes_(pendentes, H);
  const ofx = lerPixRecebimentos_(janelaPix.start, janelaPix.end).map(function(r){ return { valor: r.valor, dia: diaNum_(normData_(r.data)), key: 'pix:' + r.fitid, quem: r.pagador }; });
  const temCartao = pendentes.some(function(p){ return String(p.v[H.forma_pgto]) === 'CARTAO'; });
  // Fonte do cartão: 'arquivo' = CSV da Rede no Drive (ponte enquanto a API de produção não libera);
  // 'api' = Rede ao vivo (trocar quando vierem as credenciais de produção). Default: arquivo.
  const fonteRede = PropertiesService.getScriptProperties().getProperty('REDE_FONTE') || 'arquivo';
  const redeArquivo = (temCartao && fonteRede === 'arquivo') ? lerRedeRecebimentos_() : null;
  const redeCache = {}; let token = null;
  function redeDoDia(diaIso) {
    if (redeCache[diaIso]) return redeCache[diaIso];
    if (!token) token = redeToken_();
    const d = new Date(diaIso + 'T12:00:00-03:00');
    const lst = redeVendasDia_(token, d).map(function(t){ return { valor: t.valor, dia: diaNum_(t.data), key: 'rede:' + t.nsu, quem: t.modalidade + (t.parcelas > 1 ? ' ' + t.parcelas + 'x' : '') }; });
    redeCache[diaIso] = lst; return lst;
  }

  const usados = {};
  const novas = [];
  pendentes.forEach(function(p) {
    const valor = Number(p.v[H.valor]);
    const baseIso = normData_(p.v[H.data_venda]);
    const diaV = diaNum_(baseIso);
    const forma = String(p.v[H.forma_pgto]);
    let cand = [];
    if (forma === 'PIX') {
      cand = ofx;
    } else if (temCartao) {
      if (fonteRede === 'arquivo') {
        cand = redeArquivo || [];
      } else {
        // api: junta o dia da venda e ±1 (fuso/virada)
        [-1, 0, 1].forEach(function(off) {
          const isoOff = isoComOffset_(baseIso, off);
          if (isoOff) cand = cand.concat(redeDoDia(isoOff));
        });
      }
    }
    const elegiveis = cand.filter(function(c){ return !usados[c.key] && Math.abs(c.valor - valor) < 0.01 && Math.abs(c.dia - diaV) <= 2; });
    let hit = elegiveis[0];
    // desempate por nome do pagador (só o PIX traz nome; cartão não)
    if (elegiveis.length > 1 && forma === 'PIX') {
      const alvo = normNome_(p.v[H.paciente]);
      let melhor = -1;
      elegiveis.forEach(function(c){ const s = scoreNome_(alvo, normNome_(c.quem)); if (s > melhor) { melhor = s; hit = c; } });
    }
    if (hit) {
      usados[hit.key] = true;
      sh.getRange(p.row, H.status + 1).setValue('CONFIRMADA');
      sh.getRange(p.row, H.confirmado_em + 1).setValue(new Date().toISOString());
      sh.getRange(p.row, H.match_info + 1).setValue(hit.key + ' (' + (hit.quem || '') + ')');
      novas.push({ paciente: p.v[H.paciente], vendedora: p.v[H.vendedora], valor: valor, forma: forma, match: hit.key });
    }
  });

  if (postarSlack && novas.length) notificarMetaSlack_(novas);
  return { confirmadas: novas.length, pendentes: pendentes.length - novas.length, detalhes: novas };
}

function isoComOffset_(iso, offDias) {
  const n = diaNum_(normData_(iso));
  if (isNaN(n)) return null;
  const d = new Date(n * 86400000 + offDias * 86400000);
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

// Avisa no #comercial as vendas confirmadas + quanto falta pra meta do mês.
// 'yyyy-MM' -> 'Julho/2026'
function mesPorExtenso_(mes) {
  const M = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const p = String(mes).split('-');
  return (M[parseInt(p[1], 10) - 1] || p[1]) + '/' + p[0];
}
// Barra de progresso em 20 blocos; cor muda conforme % (laranja→amarelo→verde→azul batida).
function barraMeta_(pct) {
  const n = 20, cheios = Math.max(0, Math.min(n, Math.round(pct / 100 * n)));
  const cor = pct >= 100 ? '🟦' : pct >= 75 ? '🟩' : pct >= 40 ? '🟨' : '🟧';
  return cor.repeat(cheios) + '⬜'.repeat(n - cheios);
}

// Card visual da meta no #comercial (Block Kit). Hoje mostra o confirmado do comercial;
// quando o "outros" (cartão limpo + PIX linkado) entrar, o total passa a incluí-lo.
function notificarMetaSlack_(novas, fechamento) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_COMERCIAL_WEBHOOK');
  if (!webhookUrl) return;
  const mes = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM');
  const m = computarMetaMes_(mes);
  const pct = m.meta > 0 ? Math.round(m.total / m.meta * 100) : 0;
  const falta = Math.max(0, m.meta - m.total);
  const fmt = function(v){ return 'R$ ' + Number(v).toLocaleString('pt-BR', { maximumFractionDigits: 0 }); };
  const hora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM HH:mm');

  const blocks = [
    { type: 'header', text: { type: 'plain_text', text: (fechamento ? '🌙 Fechamento do dia · ' : '📊 Meta · ') + mesPorExtenso_(mes), emoji: true } }
  ];
  if (novas && novas.length) {
    let l = '';
    novas.forEach(function(n){ l += '✅ *' + (n.vendedora || 'Venda') + '* fechou ' + fmt(n.valor) + (n.paciente ? ' (' + n.paciente + ')' : '') + '\n'; });
    blocks.push({ type: 'section', text: { type: 'mrkdwn', text: l.trim() } });
  }
  blocks.push({ type: 'section', text: { type: 'mrkdwn', text: '*' + fmt(m.total) + '* de *' + fmt(m.meta) + '*   ·   *' + pct + '%*\n' + barraMeta_(pct) } });
  blocks.push({ type: 'section', text: { type: 'mrkdwn', text: falta > 0 ? '🔴 Faltam *' + fmt(falta) + '* pra bater a meta' : '🎉 *Meta batida!*' } });
  blocks.push({ type: 'divider' });
  let vend = '';
  (m.porVendedora || []).forEach(function(v){ vend += '\n• ' + v.vendedora + '  ' + fmt(v.valor); });
  blocks.push({ type: 'section', fields: [
    { type: 'mrkdwn', text: '🏷️ *Comercial*\n' + fmt(m.comercial) + vend },
    { type: 'mrkdwn', text: '➕ *Outros*\n' + fmt(m.outros) }
  ] });
  if (m.aConferirValor > 0) {
    blocks.push({ type: 'section', text: { type: 'mrkdwn', text: '⏳ *A conferir:* ' + fmt(m.aConferirValor) + ' · ' + m.aConferirQtd + ' PIX sem link no Feegow (dar ok no painel)' } });
  }
  blocks.push({ type: 'context', elements: [{ type: 'mrkdwn', text: '💳 cartão ' + fmt(m.cartao) + ' · 📥 PIX ' + fmt(m.pixLinkado) + ' · atualizado ' + hora } ] });

  UrlFetchApp.fetch(webhookUrl, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ blocks: blocks, text: '📊 Meta ' + mesPorExtenso_(mes) + ': ' + fmt(m.total) + ' de ' + fmt(m.meta) })
  });
}

function somaConfirmadoMes_(mes) {
  const sh = getOrCreateSheetGen_(VENDAS_FECHADAS_SHEET, VF_HEADERS);
  const data = sh.getDataRange().getValues();
  const H = {}; VF_HEADERS.forEach(function(h, i){ H[h] = i; });
  let soma = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][H.status]) === 'CONFIRMADA' && normData_(data[i][H.data_venda]).slice(0, 7) === mes) soma += Number(data[i][H.valor]) || 0;
  }
  return soma;
}

// === META: total que caiu no banco (cartão limpo + PIX linkado por CPF) ===
const PIX_CONFERIR_SHEET = 'PIX_A_Conferir';
const PIXC_HEADERS = ['fitid', 'data', 'valor', 'pagador', 'cpf', 'sugestao', 'status', 'conferido_em'];

// Pergunta ao Feegow se um CPF é de paciente (base inteira). Cache por execução.
function feegowPacientePorCpf_(cpf, cache) {
  cpf = String(cpf || '').replace(/\D/g, '');
  if (cpf.length !== 11) return false;
  if (cache && cpf in cache) return cache[cpf];
  let ok = false;
  try {
    const res = UrlFetchApp.fetch(FEEGOW_API_BASE + '/patient/list?cpf=' + cpf, {
      headers: { 'x-access-token': FEEGOW_API_TOKEN }, muteHttpExceptions: true
    });
    if (res.getResponseCode() === 200) {
      const j = JSON.parse(res.getContentText());
      ok = !!(j && j.success && j.content && j.content.length > 0);
    }
  } catch (e) { ok = false; }
  if (cache) cache[cpf] = ok;
  return ok;
}

// número de dia -> 'yyyy-mm-dd'
function diaToIso_(dia) { return isNaN(dia) ? '' : new Date(dia * 86400000).toISOString().slice(0, 10); }

// Comercial do mês agrupado por vendedora (só vendas CONFIRMADAS), maior primeiro.
function comercialPorVendedora_(mes) {
  const sh = getOrCreateSheetGen_(VENDAS_FECHADAS_SHEET, VF_HEADERS);
  const data = sh.getDataRange().getValues();
  const H = {}; VF_HEADERS.forEach(function(h, i){ H[h] = i; });
  const map = {};
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][H.status]) === 'CONFIRMADA' && normData_(data[i][H.data_venda]).slice(0, 7) === mes) {
      const v = String(data[i][H.vendedora] || '').trim() || 'Sem nome';
      map[v] = (map[v] || 0) + (Number(data[i][H.valor]) || 0);
    }
  }
  return Object.keys(map).map(function(k){ return { vendedora: k, valor: map[k] }; })
    .sort(function(a, b){ return b.valor - a.valor; });
}

// Computa a meta do mês pelo TOTAL: cartão aprovado no mês + PIX recebido no mês
// linkado a paciente por CPF. PIX não linkado vai pra fila PIX_A_Conferir (não conta
// até aprovarem). comercial = vendas marcadas confirmadas. outros = total - comercial.
function computarMetaMes_(mes) {
  const mesRe = String(mes).slice(0, 7);
  const noMes = function(iso){ return String(iso).slice(0, 7) === mesRe; };

  let cartao = 0;
  lerRedeRecebimentos_().forEach(function(t){ if (noMes(diaToIso_(t.dia))) cartao += t.valor; });

  const shC = getOrCreateSheetGen_(PIX_CONFERIR_SHEET, PIXC_HEADERS);
  const dd = shC.getDataRange().getValues();
  const Hc = {}; PIXC_HEADERS.forEach(function(h, i){ Hc[h] = i; });
  const jaTem = {}; for (let i = 1; i < dd.length; i++) jaTem[String(dd[i][Hc.fitid])] = dd[i];
  const cache = {};
  let pixLinkado = 0, aConfValor = 0, aConfQtd = 0;
  lerPixRecebimentos_(inicioDoMesIso_(mesRe), fimDoMesIso_(mesRe)).forEach(function(r){
    if (!noMes(normData_(r.data))) return;
    if (feegowPacientePorCpf_(r.cpf, cache)) { pixLinkado += r.valor; return; }
    const ex = jaTem[String(r.fitid)];
    const status = ex ? String(ex[Hc.status]) : 'PENDENTE';
    if (!ex) shC.appendRow([r.fitid, r.data, r.valor, r.pagador, r.cpf, '', 'PENDENTE', '']);
    if (status === 'OK') pixLinkado += r.valor;
    else if (status !== 'DESCARTADO') { aConfValor += r.valor; aConfQtd++; }
  });

  const porVendedora = comercialPorVendedora_(mes);
  let comercial = 0; porVendedora.forEach(function(x){ comercial += x.valor; });
  const total = cartao + pixLinkado;
  return {
    mes: mesRe, meta: getMetaMes_(mes), total: total, comercial: comercial,
    outros: Math.max(0, total - comercial), cartao: cartao, pixLinkado: pixLinkado,
    aConferirValor: aConfValor, aConferirQtd: aConfQtd, porVendedora: porVendedora
  };
}

// =========================================================
// ENTRADA MANUAL DE CARD — paciente sem agenda (recepção pede com justificativa)
// O index.html chama get_entradas_manuais / add_entrada_manual / remove_entrada_manual.
// =========================================================
const ENTRADAS_MANUAIS_SHEET = 'Entradas_Manuais';
const EM_HEADERS = ['id','timestamp','nome','justificativa','solicitante','telefone','ativo'];

function handleAddEntradaManual(body) {
  const nome = String(body.nome || '').trim();
  const just = String(body.justificativa || '').trim();
  const quem = String(body.solicitante || '').trim();
  if (!nome || !just || !quem) return jsonErr('nome, justificativa e solicitante são obrigatórios');
  const sh = getOrCreateSheetGen_(ENTRADAS_MANUAIS_SHEET, EM_HEADERS);
  const id = Utilities.getUuid();
  sh.appendRow([id, new Date().toISOString(), nome, just, quem, String(body.telefone || ''), true]);
  return jsonOk({ ok: true, id: id });
}
function handleGetEntradasManuais() {
  const sh = getOrCreateSheetGen_(ENTRADAS_MANUAIS_SHEET, EM_HEADERS);
  const data = sh.getDataRange().getValues();
  const H = {}; EM_HEADERS.forEach(function(h,i){ H[h]=i; });
  const itens = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][H.ativo] === true || String(data[i][H.ativo]).toUpperCase() === 'TRUE') {
      itens.push({ id:String(data[i][H.id]), nome:String(data[i][H.nome]), justificativa:String(data[i][H.justificativa]),
                   solicitante:String(data[i][H.solicitante]), telefone:String(data[i][H.telefone]) });
    }
  }
  return jsonOk({ ok: true, itens: itens });
}
function handleRemoveEntradaManual(body) {
  const id = String(body.id || '');
  if (!id) return jsonErr('id obrigatório');
  const sh = getOrCreateSheetGen_(ENTRADAS_MANUAIS_SHEET, EM_HEADERS);
  const data = sh.getDataRange().getValues();
  const H = {}; EM_HEADERS.forEach(function(h,i){ H[h]=i; });
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][H.id]) === id) { sh.getRange(i+1, H.ativo+1).setValue(false); return jsonOk({ ok: true }); }
  }
  return jsonErr('id não encontrado');
}

// Aprova (OK) ou descarta (DESCARTADO) um PIX da fila. Invalida o cache da meta.
function handleConferirPix(body) {
  const fitid = String(body.fitid || '');
  const decisao = String(body.decisao || '').toUpperCase();
  if (!fitid || (decisao !== 'OK' && decisao !== 'DESCARTADO')) return jsonErr('fitid/decisao inválidos');
  const sh = getOrCreateSheetGen_(PIX_CONFERIR_SHEET, PIXC_HEADERS);
  const dd = sh.getDataRange().getValues();
  const Hc = {}; PIXC_HEADERS.forEach(function(h, i){ Hc[h] = i; });
  for (let i = 1; i < dd.length; i++) {
    if (String(dd[i][Hc.fitid]) === fitid) {
      sh.getRange(i + 1, Hc.status + 1).setValue(decisao);
      sh.getRange(i + 1, Hc.conferido_em + 1).setValue(new Date().toISOString());
      PropertiesService.getScriptProperties().deleteProperty('META_CACHE');
      return jsonOk({ ok: true, fitid: fitid, decisao: decisao });
    }
  }
  return jsonErr('fitid não encontrado');
}

// Setup das credenciais REDE (rodar via curl 1x; não retorna os valores).
function handleSetupRede(body) {
  const p = PropertiesService.getScriptProperties();
  if (body.client_id) p.setProperty('REDE_CLIENT_ID', String(body.client_id));
  if (body.client_secret) p.setProperty('REDE_CLIENT_SECRET', String(body.client_secret));
  if (body.pv) p.setProperty('REDE_PV', String(body.pv));
  if (body.fonte) p.setProperty('REDE_FONTE', String(body.fonte)); // 'arquivo' (CSV) ou 'api'
  return jsonOk({ ok: true, tem_id: !!p.getProperty('REDE_CLIENT_ID'), tem_secret: !!p.getProperty('REDE_CLIENT_SECRET'), pv: p.getProperty('REDE_PV') || '', fonte: p.getProperty('REDE_FONTE') || 'arquivo' });
}

// Apaga os dados de teste das abas (mantém os cabeçalhos). Usado 1x no go-live.
function handleLimparTeste(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const limpas = {};
  ['Vendas_Fechadas', 'Metas'].forEach(function(nome) {
    const sh = ss.getSheetByName(nome);
    if (sh && sh.getLastRow() > 1) { const n = sh.getLastRow() - 1; sh.deleteRows(2, n); limpas[nome] = n; }
  });
  return jsonOk({ ok: true, limpas: limpas });
}

// Conciliação automática (com aviso no Slack) — chamada pelo gatilho horário.
function rodarConciliacaoComSlack() {
  const r = conciliarVendasFechadas_(true); // posta o card quando uma venda é confirmada
  // Resumo do total 2x/dia (12h e 18h), se não postou por venda naquela rodada.
  const h = Number(Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'HH'));
  if ((h === 12 || h === 18) && (!r || !r.confirmadas)) notificarMetaSlack_([]);
  return r;
}

// Cria o gatilho horário. RODAR 1x NO EDITOR (precisa autorizar o escopo de triggers).
function setupTriggerConciliacao() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'rodarConciliacaoComSlack') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('rodarConciliacaoComSlack').timeBased().everyHours(1).create();
  Logger.log('Gatilho horário de conciliação criado.');
  return 'ok';
}

// Fechamento do dia: além do gatilho horário, roda 1x às 19h. Concilia e posta
// um card marcado como "🌙 Fechamento do dia" com o total consolidado.
function rodarFechamentoDia() {
  const r = conciliarVendasFechadas_(false); // concilia sem postar; posto 1 card só, abaixo
  notificarMetaSlack_(r ? r.detalhes : [], true);
  return r;
}

// Cria o gatilho diário das 19h (NÃO remove o gatilho horário). RODAR 1x NO EDITOR.
function setupTriggerFechamento() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'rodarFechamentoDia') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('rodarFechamentoDia').timeBased().everyDays(1).atHour(19).create();
  Logger.log('Gatilho de fechamento das 19h criado.');
  return 'ok';
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

// =========================================================
// WHATSAPP COMERCIAL — monitor de produção das vendedoras
// A instância Z-API do número comercial manda cada mensagem (recebida e
// enviada, com "notificar enviadas por mim" ligado) pro webhook
// ?action=zapi_webhook&wk=..., que grava em
// paraser-extrato-itau.whatsapp.mensagens (BigQuery). Relatório diário às
// 19h no Slack #comercial. Credenciais em Script Properties, nunca no código.
// Atribuição de quem enviou: pelo padrão do messageId (WhatsApp Web começa
// com 3EB0; app do celular não) — validar com mensagens de teste.
// =========================================================
const WPP_BQ_DATASET = 'whatsapp';
const WPP_BQ_TABLE = 'mensagens';
const WPP_BQ_REF = '`' + BQ_PROJECT + '.' + WPP_BQ_DATASET + '.' + WPP_BQ_TABLE + '`';
const WPP_FALHAS_SHEET = 'WhatsApp_Falhas';

function wppProps_() { return PropertiesService.getScriptProperties(); }

// Dispositivo de origem de mensagem enviada pelo número da clínica.
function wppDevice_(messageId, fromMe) {
  if (!fromMe) return '';
  return /^3EB0/i.test(String(messageId || '')) ? 'web' : 'celular';
}

// Recebe o POST do Z-API e grava a mensagem no BigQuery. Responde ok mesmo em
// falha (pro Z-API não desistir do webhook); o erro fica na aba WhatsApp_Falhas.
function handleZapiWebhook(body, params) {
  try {
    const wk = wppProps_().getProperty('WPP_WEBHOOK_KEY') || '';
    if (!wk || String(params.wk || '') !== wk) return jsonErr('wk inválida');
    if (!body || !body.messageId || !body.phone) return jsonOk({ ok: true, skip: 'sem messageId/phone' });
    if (body.isGroup || body.broadcast || body.isStatusReply || body.isNewsletter || body.isEdit) {
      return jsonOk({ ok: true, skip: 'fora do escopo' });
    }
    // Reação/enquete não é mensagem (entortaria "sem resposta" e a 1ª resposta);
    // waitingMessage é placeholder sem conteúdo (o reenvio com conteúdo viria com
    // o MESMO messageId e o dedupe por insertId descartaria a versão boa).
    if (body.reaction || body.pollVote || body.waitingMessage) {
      return jsonOk({ ok: true, skip: 'sem conteúdo de mensagem' });
    }
    const tipo = body.text ? 'texto' : body.audio ? 'audio' : body.image ? 'imagem' :
                 body.video ? 'video' : body.document ? 'documento' : body.sticker ? 'figurinha' : 'outro';
    const texto = body.text ? String(body.text.message || '') :
                  body.image ? String(body.image.caption || '') :
                  body.video ? String(body.video.caption || '') : '';
    const fromMe = body.fromMe === true;
    const row = {
      message_id: String(body.messageId),
      chat_phone: String(body.phone),
      chat_name: String(body.chatName || ''),
      sender_name: String(body.senderName || ''),
      from_me: fromMe,
      momento: new Date(Number(body.momment) || Date.now()).toISOString(),
      tipo: tipo,
      texto: texto,
      device: wppDevice_(body.messageId, fromMe),
      status: String(body.status || ''),
      instance_id: String(body.instanceId || ''),
      raw: JSON.stringify(body),
      ingerido_em: new Date().toISOString()
    };
    const res = BigQuery.Tabledata.insertAll({
      rows: [{ insertId: row.message_id, json: row }]
    }, BQ_PROJECT, WPP_BQ_DATASET, WPP_BQ_TABLE);
    if (res.insertErrors && res.insertErrors.length) throw new Error(JSON.stringify(res.insertErrors));
    return jsonOk({ ok: true });
  } catch (err) {
    try {
      const sh = getOrCreateSheetGen_(WPP_FALHAS_SHEET, ['timestamp', 'erro', 'payload']);
      sh.appendRow([new Date().toISOString(), String(err), JSON.stringify(body).slice(0, 45000)]);
    } catch (e2) { /* sem onde registrar */ }
    return jsonOk({ ok: false });
  }
}

// Chamada à API do Z-API da instância COMERCIAL (credenciais em Script Properties).
function zapiComFetch_(path, method, payload) {
  const p = wppProps_();
  const inst = p.getProperty('ZAPI_COM_INSTANCE'), tok = p.getProperty('ZAPI_COM_TOKEN');
  const ctok = p.getProperty('ZAPI_CLIENT_TOKEN');
  if (!inst || !tok || !ctok) throw new Error('Credenciais Z-API não configuradas (op=set_zapi)');
  const opt = { method: method || 'get', headers: { 'Client-Token': ctok }, muteHttpExceptions: true };
  if (payload) { opt.contentType = 'application/json'; opt.payload = JSON.stringify(payload); }
  const res = UrlFetchApp.fetch('https://api.z-api.io/instances/' + inst + '/token/' + tok + path, opt);
  return { code: res.getResponseCode(), body: res.getContentText() };
}

// Hash SHA-256 da chave admin (a chave em si NUNCA entra no repo, que é público;
// o hash pode: não dá pra reverter uma chave aleatória de 48 hex).
const WPP_ADMIN_KEY_SHA256 = '08172e881cab549931a1e0507f3a8f6ec8aa248267ee2875e3c440fe05d07c42';

function wppHash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(function(b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

// Rotas administrativas do monitor. Autenticação por hash fixo no código
// (sem first-call-wins: não existe janela de captura pós-deploy).
function handleWppAdmin(params) {
  const p = wppProps_();
  const op = String(params.op || '');
  if (wppHash_(String(params.key || '')) !== WPP_ADMIN_KEY_SHA256) return jsonErr('key inválida');
  try {
    if (op === 'init') {
      if (!p.getProperty('WPP_WEBHOOK_KEY')) {
        p.setProperty('WPP_WEBHOOK_KEY', Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''));
      }
      return jsonOk({ ok: true, init: true });
    }
    if (op === 'set_zapi') {
      if (params.instance) p.setProperty('ZAPI_COM_INSTANCE', String(params.instance));
      if (params.token)    p.setProperty('ZAPI_COM_TOKEN', String(params.token));
      if (params.ctoken)   p.setProperty('ZAPI_CLIENT_TOKEN', String(params.ctoken));
      return jsonOk({ ok: true });
    }
    if (op === 'setup_bq') return jsonOk(wppSetupBigQuery_());
    if (op === 'setup_webhook') {
      const wk = p.getProperty('WPP_WEBHOOK_KEY');
      if (!wk) return jsonErr('rode op=init antes');
      // Não apontar o Z-API pra cá sem a tabela existir: o insert falharia em
      // silêncio (aba de falhas) e o período sumiria das métricas.
      try { BigQuery.Tables.get(BQ_PROJECT, WPP_BQ_DATASET, WPP_BQ_TABLE); }
      catch (e) { return jsonErr('rode op=setup_bq antes (tabela não existe): ' + e); }
      const url = ScriptApp.getService().getUrl() + '?action=zapi_webhook&wk=' + wk;
      const r1 = zapiComFetch_('/update-webhook-received', 'put', { value: url });
      const r2 = zapiComFetch_('/update-notify-sent-by-me', 'put', { notifySentByMe: true });
      return jsonOk({ webhook: r1, notifySentByMe: r2 });
    }
    if (op === 'qr')     return jsonOk(zapiComFetch_('/qr-code/image', 'get'));
    if (op === 'status') return jsonOk(zapiComFetch_('/status', 'get'));
    if (op === 'setup_trigger') return jsonOk({ ok: setupTriggerRelatorioWhatsApp() });
    if (op === 'test_report') { rodarRelatorioWhatsApp(); return jsonOk({ ok: true }); }
    if (op === 'diag')    return jsonOk(wppDiag_());
    if (op === 'ultimas') return jsonOk({ itens: wppUltimas_(Number(params.n) || 10) });
    return jsonErr('op desconhecida');
  } catch (err) {
    return jsonErr(String(err));
  }
}

// Cria dataset e tabela (particionada por dia) se não existirem. Mesma
// localização do dataset do extrato, pra permitir join futuro (funil completo).
// Cada passo é nomeado no erro (diagnóstico via rota admin, sem editor).
function wppSetupBigQuery_() {
  const passo = function(nome, fn) {
    try { return fn(); } catch (e) { throw new Error('[' + nome + '] ' + e); }
  };
  let loc = 'US';
  try { loc = BigQuery.Datasets.get(BQ_PROJECT, 'extrato').location || 'US'; } catch (e) {}
  let temDataset = true;
  try {
    BigQuery.Datasets.get(BQ_PROJECT, WPP_BQ_DATASET);
  } catch (e) {
    temDataset = false;
  }
  if (!temDataset) {
    passo('datasets.insert', function() {
      return BigQuery.Datasets.insert({
        datasetReference: { projectId: BQ_PROJECT, datasetId: WPP_BQ_DATASET }, location: loc
      }, BQ_PROJECT);
    });
  }
  try {
    BigQuery.Tables.get(BQ_PROJECT, WPP_BQ_DATASET, WPP_BQ_TABLE);
    return { ok: true, existed: true, location: loc };
  } catch (e) {
    passo('tables.insert', function() {
      return BigQuery.Tables.insert({
      tableReference: { projectId: BQ_PROJECT, datasetId: WPP_BQ_DATASET, tableId: WPP_BQ_TABLE },
      timePartitioning: { type: 'DAY', field: 'momento' },
      schema: { fields: [
        { name: 'message_id',  type: 'STRING' },
        { name: 'chat_phone',  type: 'STRING' },
        { name: 'chat_name',   type: 'STRING' },
        { name: 'sender_name', type: 'STRING' },
        { name: 'from_me',     type: 'BOOLEAN' },
        { name: 'momento',     type: 'TIMESTAMP' },
        { name: 'tipo',        type: 'STRING' },
        { name: 'texto',       type: 'STRING' },
        { name: 'device',      type: 'STRING' },
        { name: 'status',      type: 'STRING' },
        { name: 'instance_id', type: 'STRING' },
        { name: 'raw',         type: 'STRING' },
        { name: 'ingerido_em', type: 'TIMESTAMP' }
      ] }
      }, BQ_PROJECT, WPP_BQ_DATASET);
    });
    return { ok: true, created: true, location: loc };
  }
}

// Roda uma query e devolve as linhas (espera terminar + junta as páginas,
// mesmo tratamento do leitor de PIX).
function wppQuery_(sql, queryParameters) {
  const req = { query: sql, useLegacySql: false, timeoutMs: 30000 };
  if (queryParameters) { req.parameterMode = 'NAMED'; req.queryParameters = queryParameters; }
  let res = BigQuery.Jobs.query(req, BQ_PROJECT);
  const jobId = res.jobReference.jobId, loc = res.jobReference.location;
  let tentativas = 0;
  while (!res.jobComplete) {
    if (++tentativas > 30) throw new Error('BigQuery: consulta WhatsApp não completou a tempo');
    Utilities.sleep(1000);
    res = BigQuery.Jobs.getQueryResults(BQ_PROJECT, jobId, { location: loc, timeoutMs: 30000 });
  }
  let rows = res.rows || [];
  let pageToken = res.pageToken;
  while (pageToken) {
    res = BigQuery.Jobs.getQueryResults(BQ_PROJECT, jobId, { location: loc, pageToken: pageToken });
    rows = rows.concat(res.rows || []);
    pageToken = res.pageToken;
  }
  return rows;
}

// Janela do "dia comercial": das 19h de ontem às 19h de hoje (fuso SP, sem DST).
// Fixa nas 19h em ponto (o gatilho dispara em minuto aleatório dentro da hora):
// nada se perde nem conta duas vezes entre um relatório e o seguinte.
function wppJanelaRelatorio_() {
  const hojeIso = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
  const fim = new Date(hojeIso + 'T19:00:00-03:00');
  const ini = new Date(fim.getTime() - 24 * 3600 * 1000);
  return { ini: ini.toISOString(), fim: fim.toISOString() };
}

const WPP_TS_PARAM_ = function(nome, iso) {
  return { name: nome, parameterType: { type: 'TIMESTAMP' }, parameterValue: { value: iso } };
};

// Mensagens da janela, deduplicadas por message_id, em ordem cronológica (ms).
function wppMensagensJanela_(iniIso, fimIso) {
  const rows = wppQuery_(
    "SELECT ANY_VALUE(chat_phone) chat_phone, ANY_VALUE(chat_name) chat_name, " +
    "ANY_VALUE(from_me) from_me, ANY_VALUE(device) device, UNIX_MILLIS(ANY_VALUE(momento)) ts " +
    "FROM " + WPP_BQ_REF + " WHERE momento >= @ini AND momento < @fim " +
    "GROUP BY message_id ORDER BY ts",
    [WPP_TS_PARAM_('ini', iniIso), WPP_TS_PARAM_('fim', fimIso)]);
  return rows.map(function(r) {
    return {
      chat_phone: r.f[0].v, chat_name: r.f[1].v || '',
      from_me: String(r.f[2].v) === 'true', device: r.f[3].v || '', ts: Number(r.f[4].v)
    };
  });
}

// Quantos números escreveram pela PRIMEIRA vez dentro da janela (novo contato).
// Obs: nas primeiras semanas infla (a tabela nasce vazia, então paciente antiga
// que escreve parece "nova"); se corrige sozinho com o histórico acumulando.
function wppNovosContatosJanela_(iniIso, fimIso) {
  const rows = wppQuery_(
    "SELECT COUNT(*) FROM (SELECT chat_phone, MIN(momento) primeiro " +
    "FROM " + WPP_BQ_REF + " WHERE from_me = FALSE GROUP BY chat_phone) " +
    "WHERE primeiro >= @ini AND primeiro < @fim",
    [WPP_TS_PARAM_('ini', iniIso), WPP_TS_PARAM_('fim', fimIso)]);
  return rows.length ? Number(rows[0].f[0].v) : 0;
}

// Quantas mensagens caíram na aba de falhas dentro da janela (ingestão quebrada
// não pode ser invisível: o relatório avisa que está subcontando).
function wppFalhasJanela_(iniIso, fimIso) {
  const sh = getOrCreateSheetGen_(WPP_FALHAS_SHEET, ['timestamp', 'erro', 'payload']);
  const data = sh.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < data.length; i++) {
    const ts = String(data[i][0]);
    if (ts >= iniIso && ts < fimIso) n++;
  }
  return n;
}

// Duração legível: '18 min', '3h40' (com carry: 3h59m50s vira 4h00, nunca 3h60).
function wppFmtDur_(seg) {
  const min = Math.round(seg / 60);
  if (min < 60) return Math.max(1, min) + ' min';
  return Math.floor(min / 60) + 'h' + ('0' + (min % 60)).slice(-2);
}

// Métricas do dia a partir das mensagens: volumes, 1ª resposta e vácuos.
function wppMetricasDia_(msgs) {
  const chats = {};
  let enviadas = 0, recebidas = 0, web = 0, celular = 0;
  msgs.forEach(function(m) {
    if (m.from_me) { enviadas++; if (m.device === 'web') web++; else if (m.device === 'celular') celular++; }
    else recebidas++;
    (chats[m.chat_phone] = chats[m.chat_phone] || { nome: '', itens: [] }).itens.push(m);
    if (m.chat_name) chats[m.chat_phone].nome = m.chat_name;
  });
  const semResposta = [], esperas = [];
  Object.keys(chats).forEach(function(tel) {
    const c = chats[tel], itens = c.itens;
    const ultima = itens[itens.length - 1];
    if (!ultima.from_me) semResposta.push(c.nome || tel);
    // 1ª resposta: se a conversa do dia abriu com a paciente -> tempo até a 1ª enviada
    const idxIn = itens[0].from_me ? -1 : 0;
    if (idxIn >= 0) {
      for (let j = idxIn + 1; j < itens.length; j++) {
        if (itens[j].from_me) { esperas.push({ seg: (itens[j].ts - itens[idxIn].ts) / 1000, nome: c.nome || tel }); break; }
      }
    }
  });
  esperas.sort(function(a, b) { return a.seg - b.seg; });
  const nE = esperas.length;
  const mediana = !nE ? 0 :
    (nE % 2 ? esperas[(nE - 1) / 2].seg : (esperas[nE / 2 - 1].seg + esperas[nE / 2].seg) / 2);
  const pior = nE ? esperas[nE - 1] : null;
  return {
    totalMsgs: msgs.length, conversas: Object.keys(chats).length,
    enviadas: enviadas, recebidas: recebidas, web: web, celular: celular,
    semResposta: semResposta, mediana: mediana, pior: pior, respostas: esperas.length
  };
}

// Relatório diário no Slack #comercial (gatilho das 19h; também via op=test_report).
// Janela: 19h de ontem às 19h de hoje. Só a query principal é fatal; o resto
// degrada (métrica some do card em vez de derrubar o relatório inteiro).
function rodarRelatorioWhatsApp() {
  const webhookUrl = wppProps_().getProperty('SLACK_COMERCIAL_WEBHOOK');
  if (!webhookUrl) return 'sem SLACK_COMERCIAL_WEBHOOK';

  const j = wppJanelaRelatorio_();
  let msgs;
  try { msgs = wppMensagensJanela_(j.ini, j.fim); }
  catch (e) { return 'BigQuery indisponível (setup pendente?): ' + e; }

  // Status da instância: SEMPRE checar. Desconexão no meio do dia derruba a
  // coleta e, sem isso, o card sairia "normal" com números pela metade.
  let conectado = null;
  try { conectado = /"connected"\s*:\s*true/.test(zapiComFetch_('/status', 'get').body); } catch (e) {}

  if (!msgs.length && conectado === null) return 'Z-API não configurado; nada a reportar';

  const post = function(blocks, resumo) {
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ blocks: blocks, text: resumo })
    });
  };
  const blocks = [{ type: 'header', text: { type: 'plain_text',
    text: '💬 WhatsApp Comercial · ' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM'), emoji: true } }];

  if (conectado === false) {
    blocks.push({ type: 'section', text: { type: 'mrkdwn',
      text: '🔴 *WhatsApp desconectado do monitor!* Os números abaixo podem estar incompletos. Precisa escanear o QR de novo (Felipe sabe como).' } });
  }

  let falhas = 0;
  try { falhas = wppFalhasJanela_(j.ini, j.fim); } catch (e) {}
  if (falhas > 0) {
    blocks.push({ type: 'section', text: { type: 'mrkdwn',
      text: '🛠️ *' + falhas + ' mensagens falharam ao gravar* (aba WhatsApp_Falhas): números subcontados.' } });
  }

  if (!msgs.length) {
    blocks.push({ type: 'section', text: { type: 'mrkdwn', text: 'Nenhuma mensagem registrada na janela (ontem 19h → hoje 19h).' } });
    post(blocks, conectado === false ? '🔴 WhatsApp comercial desconectado' : 'WhatsApp comercial: 0 mensagens');
    return conectado === false ? 'desconectado' : 'zero mensagens';
  }

  const m = wppMetricasDia_(msgs);
  let novos = null;
  try { novos = wppNovosContatosJanela_(j.ini, j.fim); } catch (e) {}
  blocks.push({ type: 'section', text: { type: 'mrkdwn', text:
    '*' + m.conversas + '* conversas' + (novos === null ? '' : ' · *' + novos + '* contatos novos') + '\n' +
    '📤 ' + m.enviadas + ' enviadas · 📥 ' + m.recebidas + ' recebidas\n' +
    '📱 celular ' + m.celular + ' · 💻 web ' + m.web } });
  if (m.respostas) {
    blocks.push({ type: 'section', text: { type: 'mrkdwn', text:
      '⏱️ *1ª resposta:* mediana ' + wppFmtDur_(m.mediana) +
      (m.pior && m.pior.seg > m.mediana ? ' · pior ' + wppFmtDur_(m.pior.seg) + ' (' + m.pior.nome + ')' : '') } });
  }
  if (m.semResposta.length) {
    blocks.push({ type: 'section', text: { type: 'mrkdwn', text:
      '⚠️ *' + m.semResposta.length + ' sem resposta:* ' +
      m.semResposta.slice(0, 6).join(', ') + (m.semResposta.length > 6 ? '…' : '') } });
  } else {
    blocks.push({ type: 'section', text: { type: 'mrkdwn', text: '✅ Nenhuma conversa no vácuo.' } });
  }
  blocks.push({ type: 'context', elements: [{ type: 'mrkdwn',
    text: 'dia comercial: ontem 19h → hoje 19h · atualizado ' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'HH:mm') }] });
  post(blocks, '💬 WhatsApp: ' + m.conversas + ' conversas, ' + m.semResposta.length + ' sem resposta');
  return 'ok: ' + m.totalMsgs + ' mensagens';
}

// Últimas N mensagens (pra validar a atribuição celular/web com testes reais).
function wppUltimas_(n) {
  const rows = wppQuery_(
    "SELECT FORMAT_TIMESTAMP('%d/%m %H:%M', momento, 'America/Sao_Paulo') hora, from_me, device, tipo, " +
    "SUBSTR(texto, 1, 60) texto, chat_name, SUBSTR(message_id, 1, 8) id_prefixo " +
    "FROM " + WPP_BQ_REF + " ORDER BY momento DESC LIMIT " + Math.min(50, Math.max(1, n)));
  return rows.map(function(r) {
    return { hora: r.f[0].v, from_me: String(r.f[1].v) === 'true', device: r.f[2].v || '',
             tipo: r.f[3].v, texto: r.f[4].v || '', chat: r.f[5].v || '', id_prefixo: r.f[6].v };
  });
}

// Diagnóstico: propriedades configuradas + últimos 7 dias + status da instância.
function wppDiag_() {
  const p = wppProps_();
  const out = { props: {
    zapi: !!(p.getProperty('ZAPI_COM_INSTANCE') && p.getProperty('ZAPI_COM_TOKEN') && p.getProperty('ZAPI_CLIENT_TOKEN')),
    webhook_key: !!p.getProperty('WPP_WEBHOOK_KEY')
  } };
  try {
    const rows = wppQuery_(
      "SELECT CAST(DATE(momento, 'America/Sao_Paulo') AS STRING) dia, " +
      "COUNTIF(from_me) enviadas, COUNTIF(NOT from_me) recebidas, " +
      "COUNTIF(from_me AND device='web') web, COUNTIF(from_me AND device='celular') celular, " +
      "COUNT(DISTINCT chat_phone) chats FROM " + WPP_BQ_REF + " GROUP BY dia ORDER BY dia DESC LIMIT 7");
    out.dias = rows.map(function(r) {
      return { dia: r.f[0].v, enviadas: Number(r.f[1].v), recebidas: Number(r.f[2].v),
               web: Number(r.f[3].v), celular: Number(r.f[4].v), chats: Number(r.f[5].v) };
    });
  } catch (e) { out.bq = String(e); }
  try { out.zapi_status = zapiComFetch_('/status', 'get'); } catch (e) { out.zapi_status = String(e); }
  return out;
}

// Cria o gatilho diário das 19h do relatório WhatsApp (não mexe nos outros).
function setupTriggerRelatorioWhatsApp() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'rodarRelatorioWhatsApp') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('rodarRelatorioWhatsApp').timeBased().everyDays(1).atHour(19).create();
  return 'gatilho 19h do relatório WhatsApp criado';
}
