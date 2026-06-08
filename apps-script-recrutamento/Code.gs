// =============================================================================
// Apps Script "Recrutamento Paraser"
// =============================================================================
// Recebe candidatura do quiz LP (app.paraser.com.br/vagas/recepcao):
//   1) Salva resposta no Sheets (linha = candidata)
//   2) Salva CV no Drive (pasta Paraser/RH/2026/Recepcao/CVs/)
//   3) Calcula score automático nas perguntas objetivas (Q1, Q2, Q4, Q5)
//   4) Envia email confirmação pra candidata
//   5) Notifica Slack #atendimento (ou #rh) quando score >= 70
//
// Script Properties necessárias (deixe vazias por enquanto, configure no Cloud):
//   - SHEET_ID         : ID da planilha "Recrutamento Recepcao 2026"
//   - DRIVE_FOLDER_ID  : ID da pasta no Drive onde salvar CVs
//   - SLACK_WEBHOOK_URL: webhook do canal Slack #rh ou #atendimento
//   - NOTIFY_EMAIL     : email do Felipe pra receber cópia das candidaturas top
// =============================================================================

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var props = PropertiesService.getScriptProperties();

    // 1) Score automático
    var score = calcularScore(payload.respostas);

    // 2) Salvar CV no Drive
    var folderId = props.getProperty('DRIVE_FOLDER_ID');
    var cvUrl = '';
    if (folderId && payload.cv && payload.cv.base64) {
      var blob = Utilities.newBlob(
        Utilities.base64Decode(payload.cv.base64),
        payload.cv.tipo || 'application/pdf',
        sanitizarNome(payload.nome) + '_' + payload.cv.nome
      );
      var folder = DriveApp.getFolderById(folderId);
      var file = folder.createFile(blob);
      cvUrl = file.getUrl();
    }

    // 3) Salvar no Sheets
    var sheetId = props.getProperty('SHEET_ID');
    if (sheetId) {
      var ss = SpreadsheetApp.openById(sheetId);
      var sheet = ss.getSheetByName('Candidaturas') || ss.insertSheet('Candidaturas');

      // Cabeçalho se planilha vazia
      if (sheet.getLastRow() === 0) {
        sheet.appendRow([
          'Timestamp', 'Nome', 'Email', 'WhatsApp', 'LinkedIn',
          'Q1_Experiencia', 'Q2_Atendimento_Emocional', 'Q3_Cenario_Resposta',
          'Q4_Ferramentas', 'Q5_Disponibilidade', 'Q6_Motivacao',
          'Score_Automatico', 'Classificacao', 'CV_URL', 'UserAgent'
        ]);
        sheet.setFrozenRows(1);
      }

      var classif = score >= 70 ? 'PRIORIDADE' : 'REVISAR';
      sheet.appendRow([
        payload.timestamp,
        payload.nome,
        payload.email,
        payload.whatsapp,
        payload.linkedin || '',
        payload.respostas.q1 || '',
        payload.respostas.q2 || '',
        payload.respostas.q3 || '',
        payload.respostas.q4 || '',
        payload.respostas.q5 || '',
        payload.respostas.q6 || '',
        score,
        classif,
        cvUrl,
        payload.userAgent || ''
      ]);
    }

    // 4) Email confirmação pra candidata
    enviarEmailConfirmacao(payload.nome, payload.email);

    // 5) Slack notificação (se score alto)
    if (score >= 70) {
      notificarSlack(payload, score, cvUrl);
    }

    // 6) Email pro Felipe (cópia de candidatas top)
    if (score >= 70) {
      var notifyEmail = props.getProperty('NOTIFY_EMAIL');
      if (notifyEmail) {
        enviarEmailFelipe(notifyEmail, payload, score, cvUrl);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, score: score }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Erro: ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Permite chamadas OPTIONS preflight (CORS)
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: 'Recrutamento Paraser webhook ativo' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================================
// SCORE AUTOMÁTICO (0-100)
// Q1 Experiência: 0-25 pts
// Q2 Atendimento emocional: 0-25 pts
// Q4 Ferramentas: 0-25 pts (5 por ferramenta, max 25)
// Q5 Disponibilidade: 0-25 pts
// =============================================================================
function calcularScore(respostas) {
  var score = 0;
  var r = respostas || {};

  // Q1: experiência
  var mapExp = { '0': 5, '1': 10, '2': 18, '3': 22, '4': 25 };
  score += mapExp[r.q1] || 0;

  // Q2: atendimento emocional
  var mapEmo = { 'muito': 25, 'algumas': 18, 'pouco': 10, 'nunca': 8 };
  score += mapEmo[r.q2] || 0;

  // Q4: ferramentas (5 pts por uma, max 25)
  var ferramentas = (r.q4 || '').split(',').filter(function(t){ return t && t !== 'nenhuma'; });
  score += Math.min(25, ferramentas.length * 5);

  // Q5: disponibilidade
  var mapDisp = { 'integral_5': 25, 'integral_6': 25, 'flexivel': 22, 'manha': 12, 'tarde': 12 };
  score += mapDisp[r.q5] || 0;

  return score;
}

// =============================================================================
// EMAIL CONFIRMAÇÃO PRA CANDIDATA
// =============================================================================
function enviarEmailConfirmacao(nome, email) {
  if (!email) return;

  var primeiroNome = (nome || '').split(' ')[0] || 'olá';

  var html =
    '<div style="font-family:Georgia,serif; max-width:540px; margin:0 auto; background:#faf4ea; padding:40px 28px; color:#1a0a26;">' +
      '<div style="background:#1a0626; padding:28px; text-align:center; margin-bottom:24px;">' +
        '<div style="color:#c4a574; font-size:10px; letter-spacing:5px;">— INSTITUTO —</div>' +
        '<div style="color:#fff; font-size:24px; letter-spacing:5px; margin-top:6px;">PARASER</div>' +
      '</div>' +
      '<p style="font-family:Georgia,serif; font-style:italic; color:#3d1652; font-size:24px; line-height:1.3; margin:0 0 20px;">' +
        primeiroNome + ', recebemos sua candidatura.' +
      '</p>' +
      '<p style="font-size:15px; line-height:1.7; color:#1a0a26; margin:0 0 14px;">' +
        'Obrigada por ter dedicado esses minutos pra responder com calma. A gente vai ler cada palavra.' +
      '</p>' +
      '<p style="font-size:15px; line-height:1.7; color:#1a0a26; margin:0 0 14px;">' +
        'Nos próximos dias, nossa equipe entra em contato com as candidatas selecionadas para a próxima etapa.' +
      '</p>' +
      '<p style="font-size:15px; line-height:1.7; color:#5a4868; margin:0; font-style:italic;">' +
        '— Equipe Paraser 💜' +
      '</p>' +
    '</div>';

  GmailApp.sendEmail(email, 'Recebemos sua candidatura — Paraser', '', {
    htmlBody: html,
    name: 'Instituto Paraser'
  });
}

// =============================================================================
// SLACK NOTIFICAÇÃO (só pra candidatas score >= 70)
// =============================================================================
function notificarSlack(payload, score, cvUrl) {
  var webhook = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!webhook) return;

  var msg = '🎯 *Nova candidata PRIORIDADE — Recepção* (score: ' + score + ')\n' +
            '*Nome:* ' + payload.nome + '\n' +
            '*Email:* ' + payload.email + '\n' +
            '*WhatsApp:* ' + payload.whatsapp + '\n' +
            (payload.linkedin ? '*LinkedIn:* ' + payload.linkedin + '\n' : '') +
            '*CV:* ' + (cvUrl || 'erro ao salvar') + '\n\n' +
            '*Por que quer trabalhar:*\n>' + (payload.respostas.q6 || '').substring(0, 300);

  UrlFetchApp.fetch(webhook, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: msg }),
    muteHttpExceptions: true
  });
}

// =============================================================================
// EMAIL PRA FELIPE (candidatas top)
// =============================================================================
function enviarEmailFelipe(email, payload, score, cvUrl) {
  var corpo =
    'Nova candidata PRIORIDADE — score ' + score + '/100\n\n' +
    'Nome: ' + payload.nome + '\n' +
    'Email: ' + payload.email + '\n' +
    'WhatsApp: ' + payload.whatsapp + '\n' +
    (payload.linkedin ? 'LinkedIn: ' + payload.linkedin + '\n' : '') +
    'CV: ' + cvUrl + '\n\n' +
    '--- Cenário (Q3) ---\n' +
    (payload.respostas.q3 || '') + '\n\n' +
    '--- Motivação (Q6) ---\n' +
    (payload.respostas.q6 || '') + '\n\n' +
    'Planilha: https://docs.google.com/spreadsheets/d/' +
      PropertiesService.getScriptProperties().getProperty('SHEET_ID');

  GmailApp.sendEmail(email, '🎯 Candidata PRIORIDADE Paraser: ' + payload.nome + ' (' + score + ')', corpo);
}

// =============================================================================
// HELPERS
// =============================================================================
function sanitizarNome(nome) {
  return (nome || 'candidata')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-zA-Z0-9_-]/g, '_')
    .substring(0, 40);
}

// =============================================================================
// SETUP — preencha as 2 linhas vazias abaixo e rode UMA VEZ
// =============================================================================
// Como rodar:
//   1) No editor do Apps Script, com este arquivo aberto
//   2) Selecione "setupTudo" no menu de funções (em cima, ao lado do botão ▶ Executar)
//   3) Clique em ▶ Executar
//   4) Autorize quando pedir
//   5) Veja no Logs (Ctrl+Enter ou menu "Execuções") que aparece "✅ Tudo configurado"
//
// Depois disso, as Properties ficam salvas — você pode apagar essa função
// ou deixar pra reusar se trocar algum valor.
// =============================================================================
function setupTudo() {
  var props = PropertiesService.getScriptProperties();
  props.setProperties({
    // ✅ JÁ PREENCHIDOS (não precisa editar):
    'SHEET_ID':          '1726GuLXhAHiQbauLaZBievZSRUleFp8_HJc', // ⚠️ se incompleto, pegue o ID completo da planilha
    'NOTIFY_EMAIL':      'felipegrutes@paraser.com.br',

    // ✏️ EDITAR essas 2 linhas antes de rodar:
    'DRIVE_FOLDER_ID':   'COLE_AQUI_O_ID_DA_PASTA_CVS',
    'SLACK_WEBHOOK_URL': 'COLE_AQUI_O_WEBHOOK_DO_SLACK'
  });

  Logger.log('✅ Tudo configurado!');
  Logger.log('SHEET_ID: ' + props.getProperty('SHEET_ID'));
  Logger.log('DRIVE_FOLDER_ID: ' + props.getProperty('DRIVE_FOLDER_ID'));
  Logger.log('SLACK_WEBHOOK_URL: ' + (props.getProperty('SLACK_WEBHOOK_URL').substring(0, 40) + '...'));
  Logger.log('NOTIFY_EMAIL: ' + props.getProperty('NOTIFY_EMAIL'));

  if (props.getProperty('DRIVE_FOLDER_ID') === 'COLE_AQUI_O_ID_DA_PASTA_CVS') {
    Logger.log('⚠️ ATENÇÃO: você ainda não preencheu DRIVE_FOLDER_ID na função setupTudo. Edite o código e rode de novo.');
  }
  if (props.getProperty('SLACK_WEBHOOK_URL') === 'COLE_AQUI_O_WEBHOOK_DO_SLACK') {
    Logger.log('⚠️ ATENÇÃO: você ainda não preencheu SLACK_WEBHOOK_URL na função setupTudo. Edite o código e rode de novo.');
  }
}

// =============================================================================
// TESTE — rode pra confirmar que tudo está configurado
// =============================================================================
function testarSetup() {
  var props = PropertiesService.getScriptProperties();
  Logger.log('SHEET_ID: ' + (props.getProperty('SHEET_ID') || '❌ FALTA'));
  Logger.log('DRIVE_FOLDER_ID: ' + (props.getProperty('DRIVE_FOLDER_ID') || '❌ FALTA'));
  Logger.log('SLACK_WEBHOOK_URL: ' + (props.getProperty('SLACK_WEBHOOK_URL') ? '✓ ok' : '❌ FALTA'));
  Logger.log('NOTIFY_EMAIL: ' + (props.getProperty('NOTIFY_EMAIL') || '❌ FALTA'));
}
