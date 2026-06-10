// ANA Quiz — recebe submissões da LP, grava em ANA_Quiz_Submissoes
// + pré-popula memória do bot em ANA_Conversas pra ANA reconhecer o lead.

function _ss() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('SPREADSHEET_ID não configurado em Script Properties');
  return SpreadsheetApp.openById(id);
}

function doGet(e) {
  var params = (e && e.parameter) || {};
  if (params.action === 'diag' && params.key === 'paraser2026') {
    return ContentService
      .createTextOutput(JSON.stringify(_diag(), null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('ANA Quiz online').setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var r = body.respostas || {};

    if (!body.telefone || !body.nome) {
      return _resp({ ok: false, error: 'nome e telefone obrigatórios' });
    }

    var sh = _ss().getSheetByName('ANA_Quiz_Submissoes');
    if (!sh) return _resp({ ok: false, error: 'aba ANA_Quiz_Submissoes não existe — rode setupAbasAna no ANA Bot primeiro' });

    var telefone = String(body.telefone).replace(/\D/g, '');
    var resumoPerfil = _resumoPerfil(r);

    // Salva submissão crua
    sh.appendRow([
      new Date(),
      telefone,
      body.nome,
      resumoPerfil
    ]);

    // Pré-popula memória do bot pra ANA reconhecer o lead
    _prepararMemoriaInicial(telefone, body.nome, resumoPerfil);

    return _resp({ ok: true });
  } catch (err) {
    Logger.log('quiz doPost erro: ' + err.message);
    return _resp({ ok: false, error: err.message });
  }
}

function _resumoPerfil(r) {
  return 'tempo tentando=' + (r.q1 || '?') +
         ', idade=' + (r.q2 || '?') +
         ', especialista anterior=' + (r.q3 || '?') +
         ', plano=' + (r.q4 || '?') +
         '. Dúvida: "' + (r.q5 || '').substring(0, 300) + '"';
}

function _prepararMemoriaInicial(telefone, nome, resumoPerfil) {
  var sh = _ss().getSheetByName('ANA_Conversas');
  if (!sh) return;

  // O ANA Bot salva memória em LINHAS (não JSON em coluna).
  // Estrutura: [timestamp, telefone, role, mensagem]
  // Pra que a ANA Bot detecte que é lead novo do quiz, salvamos uma mensagem
  // de role='system' (que o bot vai filtrar pra contexto) com [CONTEXTO DO QUIZ].
  sh.appendRow([
    new Date(),
    telefone,
    'system',
    '[CONTEXTO DO QUIZ — Nome: ' + nome + ' · ' + resumoPerfil + ']'
  ]);
}

function _resp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function _diag() {
  var props = PropertiesService.getScriptProperties();
  var out = {
    has_spreadsheet_id: !!props.getProperty('SPREADSHEET_ID'),
    timestamp: new Date().toISOString()
  };
  if (out.has_spreadsheet_id) {
    try {
      var ss = _ss();
      out.planilha_nome = ss.getName();
      out.url = ss.getUrl();
      var abas = ss.getSheets().map(function(s) { return s.getName(); });
      out.abas = abas;
      out.tem_quiz_submissoes = abas.indexOf('ANA_Quiz_Submissoes') !== -1;
      out.tem_conversas = abas.indexOf('ANA_Conversas') !== -1;
      var shq = ss.getSheetByName('ANA_Quiz_Submissoes');
      if (shq) out.total_submissoes = Math.max(0, shq.getLastRow() - 1);
    } catch (e) { out.erro = e.message; }
  }
  return out;
}
