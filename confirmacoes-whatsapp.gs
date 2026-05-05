// ================================================================
// Paraser — Confirmações de Agendamentos via WhatsApp
// ================================================================
// SETUP:
//   1. Preencha as constantes CONFIG abaixo
//   2. Execute debugAgendamentos() uma vez para validar campos da API
//   3. Execute simularEnvio() para ver qual template cada agendamento recebe
//   4. Execute testeEnvio() com um número seu para testar envio real
//   5. Configure trigger: Editar > Acionadores > enviarConfirmacoes
//      Tipo: Por horário, Diário, entre 17h-18h
// ================================================================

// ---- CONFIG Feegow (já configurado no CRM) ----
const CF_FEEGOW_BASE  = 'https://api.feegow.com/v1/api';
const CF_FEEGOW_TOKEN = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJmZWVnb3ciLCJhdWQiOiJwdWJsaWNhcGkiLCJpYXQiOjE3NDM0NzEyNDIsImxpY2Vuc2VJRCI6MTQ0MzR9.oh2VSWT5UPEfYRrPCv34IM1NuP8Aq_ehFYWhE8f5MuU';

// ---- CONFIG Z-API ----
const CF_ZAPI_INSTANCE_ID   = '3F26C64C21E0847CDBB4AA151284BCC4';
const CF_ZAPI_TOKEN         = '3DA7CC26C48283E4C3191295';
const CF_ZAPI_CLIENT_TOKEN  = 'Fd4927cfde3bc4a2799c679493e01f164S';

// ---- CONFIG Slack ----
const CF_SLACK_TOKEN   = 'xoxb-8515459156870-8928400510806-AuvdtpMgOWthr6e8tGAaHafa';
const CF_SLACK_CHANNEL = 'atendimento';

// ---- CONFIG Planilha ----
const CF_SPREADSHEET_ID = '1LO8shBkismiZQvl7kZUXroviFOw1RJqUJxkznVRk-DA';
const CF_CONFIG_SHEET   = 'Config';
const CF_QR_LINK_CELL   = 'B1';
const CF_LOG_SHEET      = 'Confirmacoes_Log';

// ================================================================
// ENDEREÇO (bloco fixo inserido nos templates presenciais)
// ================================================================
const ENDERECO_PARASER =
  'ENDEREÇO:\n' +
  'Ficamos na Rua Prof. Álvaro Rodrigues, 352, 10º andar, Botafogo.\n' +
  'Próximo ao metrô de Botafogo, saída E.';

const DIAS_SEMANA = ['domingo','segunda','terça','quarta','quinta','sexta','sábado'];

// ================================================================
// TEMPLATES
// Variáveis disponíveis: {DATA} {HORA} {DIA_SEMANA} {ENDERECO}
//                        {LINK_QR} {DATA_VISITA}
// ================================================================
const TMPL = {

  MARCELLE_PRESENCIAL:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua consulta PRESENCIAL com a Dra. Marcelle Moura, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'Caso tenha exames peço que envie para o nosso setor de enfermagem, dessa forma iremos anexar ao sistema com mais agilidade. Segue o email: mmenezesmoura@gmail.com\n\n' +
    '⛔ Caso não haja confirmação, a consulta será cancelada. ⛔\n\n' +
    '{ENDERECO}\n\n' +
    'Ressaltamos que exames solicitados durante a consulta, como coleta de preventivo, aplicação de vitaminas, ultrassonografia e demais procedimentos, não estão inclusos no valor da consulta.\n\n' +
    'Podemos confirmar? 💜',

  BRUNA_PRESENCIAL:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua consulta PRESENCIAL com a Dra. Bruna Ortiz, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'Caso tenha exames peço que envie para o e-mail da Dra. Bruna: drabrunaortizguerra@gmail.com.\n' +
    'Desta forma, iremos anexar ao sistema com mais agilidade.\n\n' +
    '⛔ Caso não haja confirmação, a consulta será cancelada. ⛔\n\n' +
    '{ENDERECO}\n\n' +
    'Ressaltamos que exames solicitados durante a consulta, como coleta de preventivo, aplicação de vitaminas, ultrassonografia e demais procedimentos, não estão inclusos no valor da consulta.\n\n' +
    'Podemos confirmar? 💜',

  BRUNA_ONLINE:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua consulta ONLINE com a Dra. Bruna Ortiz, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'Caso tenha exames peço que envie para o e-mail da Dra Bruna: drabrunaortizguerra@gmail.com.\n' +
    'Desta forma, iremos anexar ao sistema com mais agilidade.\n\n' +
    '⛔ Caso não haja confirmação, a consulta será cancelada. ⛔\n\n' +
    'A consulta será realizada por videochamada no WhatsApp\n\n' +
    'Podemos confirmar? 💜',

  MARIO_PRESENCIAL:
    'Olá! Tudo bem?\n' +
    'Estamos entrando em contato para confirmar a CONSULTA PRESENCIAL com o Dr. Mario Barroso, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'Se tiver exames, pedimos a gentileza de enviar para o e-mail do Dr. Mario, para que possamos anexar ao sistema com mais agilidade. Segue e-mail: mariopedro_r@hotmail.com\n\n' +
    '{ENDERECO}\n\n' +
    'Podemos confirmar? 💜',

  MARIO_ONLINE:
    'Olá! Tudo bem?\n' +
    'Estamos entrando em contato para confirmar a CONSULTA ONLINE com o Dr. Mario Barroso, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'Se tiver exames, pedimos a gentileza de enviar para o e-mail do Dr. Mario, para que possamos anexar ao sistema com mais agilidade. Segue e-mail: mariopedro_r@hotmail.com\n\n' +
    'A consulta será realizada através de chamada de vídeo no Whatsapp.\n\n' +
    'Podemos confirmar? 💜',

  JOSELMO_PRESENCIAL:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua consulta PRESENCIAL com o Dr. Joselmo Salvato, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    '⛔ Caso não haja confirmação, a consulta será cancelada. ⛔\n\n' +
    'Ressaltamos que exames solicitados durante a consulta, como coleta de preventivo, aplicação de vitaminas, ultrassonografia e demais procedimentos, não estão inclusos no valor da consulta.\n\n' +
    '{ENDERECO}\n\n' +
    'Podemos confirmar? 💜',

  JOSELMO_ONLINE:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua consulta ONLINE com o Dr. Joselmo Salvato, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'A consulta será realizada através de chamada de vídeo no Whatsapp\n\n' +
    '⛔ Caso não haja confirmação, a consulta será cancelada. ⛔\n\n' +
    'Ressaltamos que exames solicitados durante a consulta, como coleta de preventivo, aplicação de vitaminas, ultrassonografia e demais procedimentos, não estão inclusos no valor da consulta.\n\n' +
    'Podemos confirmar? 💜',

  RODOLFO_PRESENCIAL:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua consulta PRESENCIAL com o Dr. Rodolfo, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'Caso tenha exames peço que envie para o nosso setor de enfermagem, dessa forma iremos anexar ao sistema com mais agilidade. Segue o email: exames@paraser.com.br\n\n' +
    '{ENDERECO}\n\n' +
    'Ressaltamos que exames solicitados durante a consulta, como coleta de preventivo, aplicação de vitaminas, ultrassonografia e demais procedimentos, não estão inclusos no valor da consulta.\n\n' +
    'Podemos confirmar? 💜',

  RODOLFO_ONLINE:
    'Olá! Tudo bem?\n' +
    'Estamos entrando em contato para confirmar a CONSULTA ONLINE com o Dr. Rodolfo Salvato, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'A consulta será realizada por chamada de vídeo no WhatsApp.\n\n' +
    'Caso tenha exames peço que envie para o nosso setor de enfermagem, dessa forma iremos anexar ao sistema com mais agilidade. Segue o email: exames@paraser.com.br\n\n' +
    'Podemos confirmar? 💜',

  HELCE_PRESENCIAL:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar a sua consulta PRESENCIAL com o Dr. Helce Ribeiro, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    '⛔ Caso não haja confirmação, a consulta será cancelada. ⛔\n\n' +
    '{ENDERECO}\n\n' +
    'Podemos confirmar? 💜',

  HELCE_ONLINE:
    'Olá! Tudo bem?\n' +
    'Estamos entrando em contato para confirmar a sua CONSULTA ONLINE com o Dr. Helce Junior, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'Se tiver exames, pedimos a gentileza de enviar para esse contato, para que possamos anexar ao sistema com mais agilidade.\n\n' +
    'A consulta será realizada por chamada de vídeo no WhatsApp.\n\n' +
    'Podemos confirmar? 💜',

  PRISCILA_PRESENCIAL:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua consulta PRESENCIAL com a Dra. Priscila Loyola, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'Caso tenha exames, peço que envie para o e-mail da Dra Priscila: priloyola@gmail.com.\n' +
    'Desta forma, iremos anexar ao sistema com mais agilidade.\n\n' +
    '⛔ Caso não haja confirmação, a consulta será cancelada. ⛔\n\n' +
    '{ENDERECO}\n\n' +
    'Ressaltamos que exames solicitados durante a consulta, como coleta de preventivo, aplicação de vitaminas, ultrassonografia e demais procedimentos, não estão inclusos no valor da consulta.\n\n' +
    'Podemos confirmar? 💜',

  PRISCILA_ONLINE:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua consulta ONLINE com a Dra. Priscila Loyola, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'Caso tenha exames peço que envie para o nosso setor de enfermagem, dessa forma iremos anexar ao sistema com mais agilidade. Segue o email: priloyola@gmail.com\n\n' +
    '⛔ Caso não haja confirmação, a consulta será cancelada. ⛔\n\n' +
    'A consulta será realizada por vídeo chamada no WhatsApp.\n\n' +
    'Podemos confirmar? 💜',

  NUTRI_PRESENCIAL:
    'Olá! Tudo bem?\n\n' +
    'Estamos entrando em contato para confirmar sua CONSULTA PRESENCIAL com a Nutricionista Graziela Siqueira, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    '{ENDERECO}\n\n' +
    'Podemos confirmar? 💜',

  NUTRI_ONLINE:
    'Olá! Tudo bem?\n\n' +
    'Estamos entrando em contato para confirmar sua CONSULTA ONLINE com a Nutricionista Graziela Siqueira, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'A consulta será realizada pelo Google Meet.\n\n' +
    'Podemos confirmar? 💜',

  ACUPUNTURA:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua sessão com a Acupunturista Cristiane, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    '{ENDERECO}\n\n' +
    'Podemos confirmar? 💜',

  ULTRAS_TRATAMENTO:
    'Olá! Tudo bem?\n\n' +
    'Passando para confirmar o seu exame {DIA_SEMANA}, dia {DATA}, às {HORA}.\n\n' +
    'Caso tenha exames anteriores, pedimos que nos envie para o nosso setor de enfermagem. Assim, conseguimos anexá-los ao sistema com mais agilidade. E-mail: exames@paraser.com.br\n\n' +
    '{ENDERECO}\n\n' +
    'Podemos confirmar? 💜',

  ULTRAS_OBSTETRICA:
    'Olá! Tudo bem?\n\n' +
    'Estamos confirmando o seu exame obstétrico {DIA_SEMANA}, dia {DATA}, às {HORA}.\n\n' +
    'Caso possua exames anteriores, pedimos a gentileza de encaminhá-los para o e-mail exames@paraser.com.br, assim conseguimos anexar ao sistema com mais agilidade.\n\n' +
    '💜 Lembre-se de trazer o seu pen drive para a gravação do exame. 💜\n\n' +
    '{ENDERECO}\n\n' +
    'Podemos confirmar? 💜',

  INJURIA:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar seu exame {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'ATENÇÃO !⚠️\n' +
    'HÁ PREPARO PARA O EXAME. CHEGUE 30 MINUTOS ANTES E BEBA 4 COPOS DE ÁGUA! A BEXIGA PRECISA ESTAR CHEIA!\n\n' +
    'Caso tenha exames peço que envie para o nosso setor de enfermagem, dessa forma iremos anexar ao sistema com mais agilidade. Segue o email: exames@paraser.com.br\n\n' +
    '{ENDERECO}\n\n' +
    'Podemos confirmar? 💜',

  QR_CODE:
    'AVISO 🚨:\n' +
    'Para acesso ao prédio, é obrigatório apresentar um QR Code na catraca.\n' +
    'Para gerar o seu QR Code, acesse o link abaixo e preencha seu CPF e nome.\n\n' +
    'ACESSO VISITANTES {DATA_VISITA}: {LINK_QR}\n\n' +
    'Atenção: o check-in para geração do QR Code só pode ser realizado no dia do seu agendamento.',

  ACOMPANHANTES:
    'O link também permite gerar QR Code para acompanhantes.\n\n' +
    'Após gerar o seu QR Code, basta clicar no "+" na parte superior da tela e adicionar o CPF e o nome do visitante. ✨'
};

// ================================================================
// FUNÇÃO PRINCIPAL — configurar trigger diário entre 17h-18h
// ================================================================
function enviarConfirmacoes() {
  var dia = getDiaAlvo();
  var agendamentos = getAgendamentos(dia);

  Logger.log('Agendamentos encontrados para ' + fmtDataBR(dia) + ': ' + agendamentos.length);
  if (!agendamentos.length) return;

  var profMap = carregarProfissionais();

  var qrLink    = getQrLink();
  var dataStr   = fmtDataBR(dia);
  var amanha    = new Date(); amanha.setDate(amanha.getDate() + 1);
  var diaSemana = fmtDataFeegow(dia) === fmtDataFeegow(amanha)
                  ? 'amanhã'
                  : DIAS_SEMANA[dia.getDay()];

  var enviados = 0, semTemplate = 0, semTelefone = 0, erros = 0;

  agendamentos.forEach(function(ag) {
    ag._profNome = profMap[ag.profissional_id] || '';
    ag._procNome = resolveNomeProc(ag);

    var phone, tmplKey;
    try {
      var paciente = getPatientData(ag.paciente_id);
      phone = paciente.phone;
      ag.paciente_nome = paciente.nome;

      if (!phone) {
        logRow(ag, 'SEM_TELEFONE', '', '');
        semTelefone++;
        return;
      }

      tmplKey = resolveTemplateKey(ag);
      if (!tmplKey) {
        logRow(ag, 'SEM_TEMPLATE', phone, '');
        Logger.log('Sem template: prof=' + ag._profNome + ' proc=' + ag._procNome + ' procId=' + ag.procedimento_id + ' tele=' + ag.telemedicina);
        semTemplate++;
        return;
      }

      var hora = formatHora(ag.horario || '');
      var msg  = fillTemplate(TMPL[tmplKey], {
        DATA:        dataStr,
        HORA:        hora,
        DIA_SEMANA:  diaSemana,
        LINK_QR:     qrLink,
        DATA_VISITA: dataStr
      });

      sendWhatsApp(phone, msg);
      Utilities.sleep(2500);

      if (qrLink) {
        var qrMsg = fillTemplate(TMPL.QR_CODE, { DATA_VISITA: dataStr, LINK_QR: qrLink });
        sendWhatsApp(phone, qrMsg);
        Utilities.sleep(1500);
        sendWhatsApp(phone, TMPL.ACOMPANHANTES);
        Utilities.sleep(1500);
      }

      logRow(ag, 'ENVIADO', phone, tmplKey);
      enviados++;

    } catch(e) {
      logRow(ag, 'ERRO: ' + e.message, phone || '', tmplKey || '');
      Logger.log('Erro paciente ' + ag.paciente_id + ': ' + e.message);
      erros++;
    }
  });

  Logger.log(
    'Resultado — Enviados: ' + enviados +
    ' | Sem template: ' + semTemplate +
    ' | Sem telefone: ' + semTelefone +
    ' | Erros: ' + erros
  );
}

// ================================================================
// SIMULAÇÃO — rode para ver qual template cada agendamento recebe
// SEM enviar mensagens reais
// ================================================================
function simularEnvio() {
  var dia = getDiaAlvo();
  var agendamentos = getAgendamentos(dia);

  Logger.log('=== SIMULAÇÃO para ' + fmtDataBR(dia) + ' (' + agendamentos.length + ' agendamentos) ===');

  if (!agendamentos.length) {
    Logger.log('Nenhum agendamento encontrado.');
    return;
  }

  var profMap = carregarProfissionais();

  agendamentos.forEach(function(ag, i) {
    ag._profNome = profMap[ag.profissional_id] || '';
    ag._procNome = resolveNomeProc(ag);

    var tmplKey   = resolveTemplateKey(ag);
    var paciente  = getPatientData(ag.paciente_id);
    Logger.log(
      '[' + (i+1) + '] ' +
      'Paciente: ' + (paciente.nome || 'id=' + ag.paciente_id) + ' | ' +
      'Prof: ' + (ag._profNome || 'profId=' + ag.profissional_id) + ' | ' +
      'Hora: ' + formatHora(ag.horario || '') + ' | ' +
      'ProcID: ' + ag.procedimento_id + ' | ' +
      'LocalID: ' + ag.local_id + ' | ' +
      'Template: ' + (tmplKey || '⚠️ SEM_TEMPLATE')
    );
    Utilities.sleep(300);
  });
}

// ================================================================
// RESOLUÇÃO DO NOME DO PROCEDIMENTO
// O agendamento pode já trazer o nome — tenta vários campos antes
// de usar apenas o ID numérico.
// ================================================================
function resolveNomeProc(ag) {
  return (
    ag.procedimento_nome ||
    ag.procedimento      ||
    ag.descricao_procedimento ||
    ag.nome_procedimento ||
    ''
  );
}

// ================================================================
// FEEGOW — agendamentos do dia
// ================================================================
function getAgendamentos(dia) {
  var ds   = fmtDataFeegow(dia);
  var url  = CF_FEEGOW_BASE + '/appoints/search?data_start=' + ds + '&data_end=' + ds;
  var resp = UrlFetchApp.fetch(url, {
    headers: { 'x-access-token': CF_FEEGOW_TOKEN },
    muteHttpExceptions: true
  });
  var json  = JSON.parse(resp.getContentText());
  var items = Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);

  return items.filter(function(a) {
    var s = (a.status || '').toLowerCase();
    return s !== 'cancelado' && s !== 'bloqueado' && s !== 'bloqueio' && s !== 'desmarcado';
  });
}

// ================================================================
// FEEGOW — mapa de profissionais {id: nome}
// ================================================================
function carregarProfissionais() {
  var map = {};
  try {
    var resp = UrlFetchApp.fetch(CF_FEEGOW_BASE + '/professional/list', {
      headers: { 'x-access-token': CF_FEEGOW_TOKEN }, muteHttpExceptions: true
    });
    var json  = JSON.parse(resp.getContentText());
    var items = Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);
    items.forEach(function(p) {
      var id = p.profissional_id || p.id;
      if (id) map[id] = p.nome || p.name || '';
    });
    Logger.log('Profissionais carregados: ' + Object.keys(map).length);
  } catch(e) { Logger.log('carregarProfissionais erro: ' + e.message); }
  return map;
}

// ================================================================
// FEEGOW — dados do paciente (nome + telefone)
// ================================================================
function getPatientData(pacienteId) {
  if (!pacienteId) return { phone: null, nome: '' };
  var url  = CF_FEEGOW_BASE + '/patient/search?paciente_id=' + pacienteId;
  var resp = UrlFetchApp.fetch(url, {
    headers: { 'x-access-token': CF_FEEGOW_TOKEN },
    muteHttpExceptions: true
  });
  var json = JSON.parse(resp.getContentText());
  var p    = json.content || json;
  var cel  = p.celular || (p.celulares && p.celulares[0]) || p.telefone || p.telefone_celular || '';
  var nome = p.nome || p.paciente_nome || p.name || '';
  return { phone: formatPhone(cel), nome: nome };
}

// ================================================================
// MAPA DE IDs DE PROCEDIMENTOS ESPECIAIS
// A API Feegow não retorna nomes nos agendamentos, só IDs.
// Adicione aqui os IDs conforme forem identificados.
// ================================================================

// procId=69  confirmado: INJÚRIA ENDOMETRIAL (Priscila 29/07/2025 14:45)
// procId=152 confirmado: APLICAÇÃO DE FILGRASTIM (Marcelle 05/08/2025 12:30)
var IDS_INJURIA = [69, 152];

// procId=42  confirmado: USG TRANSLUCÊNCIA NUCAL (Érica 01/10/2025 14:00)
// procId=40  confirmado: USG 1 PÓS BETA (Marcelle 05/05/2026 12:00)
// procId=59  confirmado: USG CONTAGEM DE FOLÍCULOS ANTRAIS (Marcelle + Priscila)
// TODO: USG 2 Pós Beta, USG Morfológica, USG Obstétrica c/ Doppler
//       → rode debugEncontrarProcId() para localizar os IDs e adicione aqui
var IDS_OBSTETRICA = [42, 40, 59];

// procIds de exames de acompanhamento de tratamento (USG Preparo TEC, USG FIV)
// procId=244 confirmado: USG PREPARO TEC 1 - Medicado (Marcelle 05/05/2026 10:00 / 11:00)
// procId=73  confirmado: USG PREPARO TEC 1 - Natural   (Marcelle 05/05/2026 10:40 / 11:20)
// procId=4   confirmado: USG FIV 1ª                   (Marcelle 05/05/2026 11:40)
// procId=203 confirmado: Consulta 1ª Vez (Marcelle) — roteado por profissional, sem ação
// procId=204 confirmado: Consulta 1ª Vez (Marcelle, 12×) — roteado por profissional, sem ação
// procId=281 confirmado: Consulta 1ª Vez (Marcelle 07/04/2026) — roteado por profissional, sem ação
// procId=249 confirmado: Consulta 1ª Vez (Marcelle 11/03/2026) — roteado por profissional, sem ação
// procId=247 confirmado: Consulta 1ª Vez (Marcelle 25/02/2026) — roteado por profissional, sem ação
// procId=12  confirmado: USG COITO PROGRAMADO (Érica 04/05/2026 15:20)
// procId=59  confirmado: USG CONTAGEM DE FOLÍCULOS ANTRAIS → movido para IDS_OBSTETRICA
var IDS_ULTRAS_TRATAMENTO = [244, 73, 4, 12];

// procIds de consultas ONLINE — confirmados via tela do Feegow em 04/05/2026
// procId=252: "CONSULTA 1ª VEZ - DR. RODOLFO SALVATO - Online" (Rodolfo)
// procId=256: "CONSULTA - Online" (usado por Priscila, Rodolfo, Marcelle e outros)
// procId=255: "CONSULTA DE RETORNO - Presencial" → NÃO é online, não incluir aqui
// Execute debugMapearOnlineProcIds() e confira agendamentos dos demais médicos
// (Mario, Joselmo, Helce, Bruna, Graziela) para completar a lista se necessário.
var IDS_ONLINE_PROCS = [252, 256];

// procIds de procedimentos que NÃO recebem confirmação por WhatsApp:
// punção de óvulos, transferência, aspiração de cisto, PRP,
// EMBRION, PRIMORDIA, ORIGEN, SP — e qualquer outro cirúrgico/clínico sem aviso.
// Execute debugProcsSemConfirmacao() para descobrir os IDs e adicione-os aqui.
// procId=89:  Punção De Óvulos / Punção Doadora / 1ª Tec Primórdia (confirmado)
// procId=91:  2ª Tec Embrion (confirmado 21/01/2025)
// procId=120: Aspiração De Cisto (confirmado 09/02/2026 — 11:30 Rodolfo)
// procId=127: PRP (confirmado 27/01/2025)
// procId=139: SP — 1ª/2ª Tec Sp (confirmado 11/04/2026)
// procId=234: 2ª Tec Embrion - Dr. Rodolfo Salvato (variante recente)
// procId=265: Origen — 2ª Tec Laboratório Origen (confirmado 26/03/2026)
// TODO: procId=168 (11× Rodolfo, 08:00-09:00 manhã) — não identificado, monitorar via simularEnvio
var IDS_SEM_CONFIRMACAO = [89, 91, 120, 127, 139, 234, 265];

// ================================================================
// LÓGICA DE TEMPLATE
// Prioridade: procId especial > nome do proc (se vier) > profissional
// ================================================================
function resolveTemplateKey(ag) {
  var procId = ag.procedimento_id;
  var proc   = (ag._procNome || '').toUpperCase();
  var prof   = (ag._profNome || '').toUpperCase();

  // Modal: procId é a fonte mais confiável nesta clínica.
  // (campo telemedicina sempre false; local_id não distingue online/presencial)
  var modal = IDS_ONLINE_PROCS.indexOf(procId) >= 0 ? 'ONLINE' : 'PRESENCIAL';
  // Fallback: se a API um dia passar o nome, aproveita
  if (modal === 'PRESENCIAL' && proc.indexOf('ONLINE') >= 0) modal = 'ONLINE';

  // --- 0. Procedimentos sem confirmação (cirurgias, PRP, EMBRION etc.) ---
  if (IDS_SEM_CONFIRMACAO.indexOf(procId) >= 0) return null;

  // --- 1. IDs especiais hardcoded (mais confiável) ---
  if (IDS_INJURIA.indexOf(procId)           >= 0) return 'INJURIA';
  if (IDS_OBSTETRICA.indexOf(procId)        >= 0) return 'ULTRAS_OBSTETRICA';
  if (IDS_ULTRAS_TRATAMENTO.indexOf(procId) >= 0) return 'ULTRAS_TRATAMENTO';

  // --- 2. Nome do procedimento (fallback, se API passar) ---
  if (proc.includes('INJUR') || proc.includes('FILGRASTIM'))                             return 'INJURIA';
  if (proc.includes('OBSTET') || proc.includes('MORFOL') || proc.includes('TRANSLUC'))           return 'ULTRAS_OBSTETRICA';
  if (proc.includes('POS BETA') || proc.includes('PÓS BETA'))                                    return 'ULTRAS_OBSTETRICA';
  if (proc.includes('FOLICULO') || proc.includes('FOLÍCULO'))                                     return 'ULTRAS_OBSTETRICA';
  if (proc.includes('TRANSVAGINAL') || proc.includes('DOPPLER'))                                  return 'ULTRAS_OBSTETRICA';
  if (proc.includes('PREPARO TEC') || proc.includes('FIV'))                                       return 'ULTRAS_TRATAMENTO';
  if (proc.includes('USG') || proc.includes('ULTRASSOM') || proc.includes('ULTRA'))               return 'ULTRAS_TRATAMENTO';
  if (proc.includes('ACUPUNTURA') || prof.includes('CRISTIANE'))                         return 'ACUPUNTURA';

  // --- 3. Profissional ---
  if (prof.includes('MARCELLE'))                                    return 'MARCELLE_'  + modal;
  if (prof.includes('BRUNA'))                                       return 'BRUNA_'     + modal;
  if (prof.includes('MARIO') || prof.includes('MÁRIO'))             return 'MARIO_'     + modal;
  if (prof.includes('JOSELMO'))                                     return 'JOSELMO_'   + modal;
  if (prof.includes('RODOLFO'))                                     return 'RODOLFO_'   + modal;
  if (prof.includes('HELCE'))                                       return 'HELCE_'     + modal;
  if (prof.includes('PRISCILA') || prof.includes('PRISCILLA'))      return 'PRISCILA_'  + modal;
  if (prof.includes('GRAZIELA') || prof.includes('NUTRI'))          return 'NUTRI_'     + modal;
  // Érica faz só exames; qualquer procId não identificado → tratamento
  if (prof.includes('ERICA') || prof.includes('ÉRICA'))             return 'ULTRAS_TRATAMENTO';

  return null;
}

// ================================================================
// PREENCHIMENTO DE TEMPLATE
// ================================================================
function fillTemplate(tmpl, vars) {
  return tmpl
    .replace(/{DATA}/g,        vars.DATA        || '')
    .replace(/{HORA}/g,        vars.HORA        || '')
    .replace(/{DIA_SEMANA}/g,  vars.DIA_SEMANA  || 'amanhã')
    .replace(/{ENDERECO}/g,    ENDERECO_PARASER)
    .replace(/{LINK_QR}/g,     vars.LINK_QR     || '')
    .replace(/{DATA_VISITA}/g, vars.DATA_VISITA || '');
}

// ================================================================
// Z-API — envio de mensagem
// ================================================================
function sendWhatsApp(phone, message) {
  var url     = 'https://api.z-api.io/instances/' + CF_ZAPI_INSTANCE_ID + '/token/' + CF_ZAPI_TOKEN + '/send-text';
  var payload = JSON.stringify({ phone: phone, message: message });
  var headers = {};
  if (CF_ZAPI_CLIENT_TOKEN) headers['Client-Token'] = CF_ZAPI_CLIENT_TOKEN;
  var resp    = UrlFetchApp.fetch(url, {
    method:      'POST',
    contentType: 'application/json',
    headers:     headers,
    payload:     payload,
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Z-API HTTP ' + code + ': ' + resp.getContentText().substring(0, 200));
  }
  return JSON.parse(resp.getContentText());
}

// ================================================================
// LISTA DE PACIENTES — configura trigger às 8h para esta função
// Popula a aba "Pacientes_Amanha" com nomes formatados (≤40 chars)
// para a recepção usar no Keyaccess antes das 10h30
// ================================================================
function listarPacientesAmanha() {
  var dia          = getDiaAlvo();
  var agendamentos = getAgendamentos(dia);
  var dataStr      = fmtDataBR(dia);

  var ss = SpreadsheetApp.openById(CF_SPREADSHEET_ID);
  var sh = ss.getSheetByName('Pacientes_Amanha');
  if (!sh) {
    sh = ss.insertSheet('Pacientes_Amanha');
  }

  // Limpa conteúdo anterior e escreve cabeçalho
  sh.clearContents();
  sh.appendRow(['PACIENTES PARA ' + dataStr + ' — cole o link Keyaccess em Config!B1 antes das 10h30']);
  sh.appendRow(['#', 'Nome (até 40 caracteres)', 'Hora', 'Profissional']);
  sh.setFrozenRows(2);

  if (!agendamentos.length) {
    sh.appendRow(['', 'Nenhum agendamento encontrado.', '', '']);
    Logger.log('listarPacientesAmanha: nenhum agendamento para ' + dataStr);
    return;
  }

  var profMap = carregarProfissionais();

  // Ordena por horário
  agendamentos.sort(function(a, b) {
    return (a.horario || '').localeCompare(b.horario || '');
  });

  var i = 1;
  var nomesKeyaccess = [];

  agendamentos.forEach(function(ag) {
    try {
      var paciente     = getPatientData(ag.paciente_id);
      var nomeCompleto = (paciente.nome || '').toUpperCase();
      var nome40       = nomeCompleto.substring(0, 40);
      var hora         = formatHora(ag.horario || '');
      var profNome     = profMap[ag.profissional_id] || '';
      sh.appendRow([i, nome40, hora, profNome]);
      if (nome40) nomesKeyaccess.push(nome40);
      i++;
      Utilities.sleep(300);
    } catch(e) {
      sh.appendRow([i, 'ERRO: ' + e.message, '', '']);
      i++;
    }
  });

  Logger.log('listarPacientesAmanha: ' + (i-1) + ' pacientes listados para ' + dataStr);

  // Gera Excel Keyaccess e envia para o Slack
  if (nomesKeyaccess.length > 0) {
    try {
      enviarKeyaccessSlack(nomesKeyaccess, dataStr);
    } catch(e) {
      Logger.log('Slack erro: ' + e.message);
    }
  }
}

// ================================================================
// SLACK — gera Excel no formato Keyaccess e envia para o canal
// ================================================================
function enviarKeyaccessSlack(nomes, dataStr) {
  // Cria planilha temporária no formato Keyaccess
  var tempSs = SpreadsheetApp.create('Keyaccess_Temp_' + dataStr.replace('/', '-'));
  var sh     = tempSs.getActiveSheet();

  sh.appendRow(['Nome completo', 'Empresa', 'Telefone', 'Email']);
  nomes.forEach(function(nome) { sh.appendRow([nome]); });
  SpreadsheetApp.flush();

  // Exporta como .xlsx via URL de exportação do Google
  var ssId      = tempSs.getId();
  var exportUrl = 'https://docs.google.com/spreadsheets/d/' + ssId +
                  '/export?format=xlsx';
  var blob = UrlFetchApp.fetch(exportUrl, {
               headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
             }).getBlob()
               .setName('keyaccess_' + dataStr.replace('/', '-') + '.xlsx');

  var fileName = blob.getName();
  var fileBytes = blob.getBytes();
  var headers = { Authorization: 'Bearer ' + CF_SLACK_TOKEN };

  // Etapa 1 — obtém URL de upload
  var step1 = JSON.parse(UrlFetchApp.fetch(
    'https://slack.com/api/files.getUploadURLExternal?filename=' +
    encodeURIComponent(fileName) + '&length=' + fileBytes.length,
    { headers: headers }
  ).getContentText());
  if (!step1.ok) throw new Error('getUploadURL: ' + step1.error);

  // Etapa 2 — faz upload do arquivo
  UrlFetchApp.fetch(step1.upload_url, {
    method:      'post',
    contentType: 'application/octet-stream',
    payload:     blob.getBytes()
  });

  // Etapa 3 — finaliza e posta no canal
  var channelId = slackGetChannelId(CF_SLACK_CHANNEL);
  var step3 = JSON.parse(UrlFetchApp.fetch(
    'https://slack.com/api/files.completeUploadExternal', {
      method:      'post',
      contentType: 'application/json; charset=utf-8',
      headers:     headers,
      payload:     JSON.stringify({
        files:           [{ id: step1.file_id, title: 'Lista Keyaccess — ' + dataStr }],
        channel_id:      channelId,
        initial_message: '📋 Lista de pacientes para o Keyaccess de *' + dataStr + '*.\nBaixe, suba no Keyaccess e cole o link em *Config!B1* antes das 10h30.'
      })
    }
  ).getContentText());
  if (!step3.ok) throw new Error('completeUpload: ' + step3.error);

  // Remove planilha temporária
  DriveApp.getFileById(ssId).setTrashed(true);

  Logger.log('Slack: arquivo enviado para #' + CF_SLACK_CHANNEL);
}

// ================================================================
// SLACK — retorna o ID do canal a partir do nome
// ================================================================
function slackGetChannelId(channelName) {
  var cursor = '';
  do {
    var url  = 'https://slack.com/api/conversations.list?limit=200&exclude_archived=true&types=public_channel,private_channel' +
               (cursor ? '&cursor=' + encodeURIComponent(cursor) : '');
    var data = JSON.parse(UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + CF_SLACK_TOKEN }
    }).getContentText());
    if (!data.ok) throw new Error('conversations.list: ' + data.error);
    var found = data.channels.filter(function(c) {
      return c.name === channelName || c.name === channelName.replace('#', '');
    });
    if (found.length) return found[0].id;
    cursor = data.response_metadata && data.response_metadata.next_cursor || '';
  } while (cursor);
  throw new Error('Canal Slack não encontrado: ' + channelName);
}

// ================================================================
// PLANILHA — lê link QR Code da célula Config!B1
// ================================================================
function getQrLink() {
  try {
    var ss  = SpreadsheetApp.openById(CF_SPREADSHEET_ID);
    var sh  = ss.getSheetByName(CF_CONFIG_SHEET) || ss.getSheets()[0];
    var val = sh.getRange(CF_QR_LINK_CELL).getValue();
    return (val || '').toString().trim();
  } catch(e) {
    Logger.log('getQrLink erro: ' + e.message);
    return '';
  }
}

// ================================================================
// PLANILHA — log de envios
// ================================================================
function logRow(ag, status, phone, tmplKey) {
  try {
    var ss = SpreadsheetApp.openById(CF_SPREADSHEET_ID);
    var sh = ss.getSheetByName(CF_LOG_SHEET);
    if (!sh) {
      sh = ss.insertSheet(CF_LOG_SHEET);
      sh.appendRow(['Timestamp','Paciente','Profissional','Procedimento','ProcID','Telefone','Template','Status']);
      sh.setFrozenRows(1);
    }
    sh.appendRow([
      new Date(),
      ag.paciente_nome || ag.paciente || '',
      ag._profNome     || ag.profissional_nome || '',
      ag._procNome     || ag.procedimento_nome || ag.procedimento || '',
      ag.procedimento_id || '',
      phone,
      tmplKey,
      status
    ]);
  } catch(e) {
    Logger.log('logRow erro: ' + e.message);
  }
}

// ================================================================
// HELPERS
// ================================================================
// Feriados nacionais + municipais/estaduais do Rio de Janeiro
// Formato: 'DD/MM' (feriados fixos) — atualize anualmente se necessário
var FERIADOS = [
  // Nacionais
  '01/01', // Confraternização Universal
  '20/01', // São Sebastião (municipal RJ)
  '04/03', // Carnaval 2026 (terça)
  '03/03', // Carnaval 2026 (segunda)
  '05/03', // Quarta de Cinzas (meio dia) 2026
  '03/04', // Paixão de Cristo 2026
  '21/04', // Tiradentes
  '01/05', // Dia do Trabalho
  '11/06', // Corpus Christi 2026
  '23/04', // São Jorge (municipal RJ)
  '07/09', // Independência
  '12/10', // Nossa Senhora Aparecida
  '02/11', // Finados
  '15/11', // Proclamação da República
  '20/11', // Consciência Negra
  '24/12', // Véspera de Natal (ponto facultativo RJ)
  '25/12', // Natal
  '31/12', // Véspera de Ano Novo (ponto facultativo RJ)
];

function ehFeriado(d) {
  var chave = pad2(d.getDate()) + '/' + pad2(d.getMonth() + 1);
  return FERIADOS.indexOf(chave) >= 0;
}

function getDiaAlvo() {
  var d = new Date();
  d.setDate(d.getDate() + 1);
  // Avança até encontrar um dia útil (não fim de semana, não feriado)
  while (d.getDay() === 0 || d.getDay() === 6 || ehFeriado(d)) {
    d.setDate(d.getDate() + 1);
  }
  return d;
}

function fmtDataFeegow(d) {
  return pad2(d.getDate()) + '-' + pad2(d.getMonth() + 1) + '-' + d.getFullYear();
}

function fmtDataBR(d) {
  return pad2(d.getDate()) + '/' + pad2(d.getMonth() + 1);
}

function pad2(n) { return n < 10 ? '0' + n : '' + n; }

function formatHora(h) {
  return (h || '').toString().substring(0, 5);
}

function formatPhone(raw) {
  var digits = (raw || '').toString().replace(/\D/g, '');
  if (!digits) return null;
  if (digits.startsWith('55') && digits.length >= 12) return digits;
  if (digits.length >= 10) return '55' + digits;
  return null;
}

// ================================================================
// DEBUG — descobre os procIds de procedimentos sem confirmação.
// Varre 90 dias e lista todos os procIds que aparecem para médicos
// que TÊM template (Rodolfo, Priscila etc.), excluindo os já conhecidos.
// Esses são candidatos a punção, transferência, PRP, etc.
// Para confirmar: abra um agendamento no Feegow com o procedimento
// desejado, anote o ProcID mostrado no simularEnvio() e adicione em
// IDS_SEM_CONFIRMACAO.
// ================================================================
function debugProcsSemConfirmacao() {
  var DIAS = 90;
  var profMap = carregarProfissionais();

  // Profissionais que têm template (pelos nomes usados em resolveTemplateKey)
  var PROF_COM_TEMPLATE = ['MARCELLE','BRUNA','MARIO','MÁRIO','JOSELMO','RODOLFO',
                           'HELCE','PRISCILA','PRISCILLA','GRAZIELA'];

  // ProcIds já mapeados (não precisam aparecer no relatório)
  var CONHECIDOS = IDS_INJURIA.concat(IDS_OBSTETRICA).concat(IDS_ONLINE_PROCS).concat(IDS_SEM_CONFIRMACAO);

  var combos = {}; // procId → { profNomes, count }

  for (var offset = -DIAS; offset <= 7; offset++) {
    var d  = new Date();
    d.setDate(d.getDate() + offset);
    var ds = fmtDataFeegow(d);
    try {
      var resp  = UrlFetchApp.fetch(
        CF_FEEGOW_BASE + '/appoints/search?data_start=' + ds + '&data_end=' + ds,
        { headers: { 'x-access-token': CF_FEEGOW_TOKEN }, muteHttpExceptions: true }
      );
      var json  = JSON.parse(resp.getContentText());
      var items = Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);

      items.forEach(function(ag) {
        var prc     = ag.procedimento_id;
        var profNome = (profMap[ag.profissional_id] || '').toUpperCase();
        var temTemplate = PROF_COM_TEMPLATE.some(function(n) { return profNome.indexOf(n) >= 0; });
        if (!temTemplate) return;
        if (CONHECIDOS.indexOf(prc) >= 0) return;

        if (!combos[prc]) combos[prc] = { profs: {}, count: 0 };
        combos[prc].profs[profNome] = true;
        combos[prc].count++;
      });
    } catch(e) {}
  }

  var lista = Object.keys(combos).map(function(k) {
    return { procId: parseInt(k), count: combos[k].count, profs: Object.keys(combos[k].profs).join(', ') };
  });
  lista.sort(function(a, b) { return b.count - a.count; });

  Logger.log('=== ProcIDs NÃO MAPEADOS em médicos com template (últimos ' + DIAS + ' dias) ===');
  Logger.log('(Candidatos a punção, transferência, PRP, EMBRION, SP etc.)');
  lista.forEach(function(c) {
    Logger.log('ProcID=' + c.procId + '  ocorrências=' + c.count + '  médicos: ' + c.profs);
  });
  Logger.log('Total: ' + lista.length + ' procIds. Abra um agendamento de cada tipo no Feegow para confirmar o nome.');
}

// ================================================================
// DEBUG — mapeia todos os procIds únicos por profissional nos últimos
// 90 dias. Use para descobrir quais procIds são consultas Online:
//   1. Execute esta função
//   2. Para cada profissional, veja quais procIds aparecem
//   3. Compare com a planilha de agendamentos online do Feegow
//      (os procIds das linhas "Online" são os que faltam em IDS_ONLINE_PROCS)
//   4. Adicione os IDs confirmados em IDS_ONLINE_PROCS acima
// ================================================================
function debugMapearOnlineProcIds() {
  var DIAS = 90;
  var profMap = carregarProfissionais();

  // { profId: { procId: { localIds: Set, count: n } } }
  var dados = {};

  for (var offset = -DIAS; offset <= 7; offset++) {
    var d  = new Date();
    d.setDate(d.getDate() + offset);
    var ds = fmtDataFeegow(d);
    try {
      var resp  = UrlFetchApp.fetch(
        CF_FEEGOW_BASE + '/appoints/search?data_start=' + ds + '&data_end=' + ds,
        { headers: { 'x-access-token': CF_FEEGOW_TOKEN }, muteHttpExceptions: true }
      );
      var json  = JSON.parse(resp.getContentText());
      var items = Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);
      items.forEach(function(ag) {
        var pid = ag.profissional_id, prc = ag.procedimento_id, lid = ag.local_id;
        if (!dados[pid]) dados[pid] = {};
        if (!dados[pid][prc]) dados[pid][prc] = { localIds: {}, count: 0 };
        dados[pid][prc].count++;
        dados[pid][prc].localIds[lid] = (dados[pid][prc].localIds[lid] || 0) + 1;
      });
    } catch(e) {}
  }

  Logger.log('=== MAPA PROCIDS POR PROFISSIONAL (últimos ' + DIAS + ' dias) ===');
  Object.keys(dados).sort(function(a, b) { return a - b; }).forEach(function(pid) {
    var nome = profMap[pid] || 'id=' + pid;
    Logger.log('\n--- ' + nome + ' ---');
    var procs = Object.keys(dados[pid]).sort(function(a, b) { return a - b; });
    procs.forEach(function(prc) {
      var info   = dados[pid][prc];
      var lidStr = Object.keys(info.localIds).map(function(l) {
        return 'localId=' + l + '(×' + info.localIds[l] + ')';
      }).join(', ');
      Logger.log('  ProcID=' + prc + '  ocorrências=' + info.count + '  [' + lidStr + ']');
    });
  });
  Logger.log('\nProcIDs JÁ em IDS_ONLINE_PROCS: ' + IDS_ONLINE_PROCS.join(', '));
}

// ================================================================
// DEBUG — tenta buscar o nome de um procedimento pelo ID
// Troque o valor de PROC_ID e execute
// ================================================================
function debugProcedimento() {
  var PROC_ID = 89; // ← troque pelo ID que quer consultar

  var endpoints = [
    '/procedure/detail?procedimento_id=' + PROC_ID,
    '/procedure/detail?id='              + PROC_ID,
    '/procedure/get?procedimento_id='    + PROC_ID,
    '/procedure/get?id='                 + PROC_ID,
    '/procedure/' + PROC_ID,
  ];

  endpoints.forEach(function(ep) {
    var resp = UrlFetchApp.fetch(CF_FEEGOW_BASE + ep, {
      headers: { 'x-access-token': CF_FEEGOW_TOKEN },
      muteHttpExceptions: true
    });
    Logger.log(ep + ' → HTTP ' + resp.getResponseCode() + '\n' + resp.getContentText().substring(0, 300));
  });
}

// ================================================================
// DEBUG — varre N dias e lista todas as combinações únicas
// (profissional_id, procedimento_id) com nome do profissional.
// Use para identificar quais IDs correspondem a Injúria, Filgrastim
// e Translucência Nucal — basta olhar quais IDs aparecem nos dias
// em que esses exames estão agendados.
// ================================================================
function debugScanProcedimentos() {
  var DIAS_ATRAS   = 30;  // quantos dias para trás varrer
  var DIAS_FRENTE  = 30;  // quantos dias para frente varrer

  var profMap = carregarProfissionais();
  var combos  = {};  // chave: "profId|procId"

  for (var offset = -DIAS_ATRAS; offset <= DIAS_FRENTE; offset++) {
    var d  = new Date();
    d.setDate(d.getDate() + offset);
    var ds = fmtDataFeegow(d);

    try {
      var resp = UrlFetchApp.fetch(
        CF_FEEGOW_BASE + '/appoints/search?data_start=' + ds + '&data_end=' + ds,
        { headers: { 'x-access-token': CF_FEEGOW_TOKEN }, muteHttpExceptions: true }
      );
      var json  = JSON.parse(resp.getContentText());
      var items = Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);

      items.forEach(function(ag) {
        var key = ag.profissional_id + '|' + ag.procedimento_id;
        if (!combos[key]) {
          combos[key] = {
            profId:  ag.profissional_id,
            procId:  ag.procedimento_id,
            profNome: profMap[ag.profissional_id] || '?',
            count:   0
          };
        }
        combos[key].count++;
      });
    } catch(e) { /* ignora dias com erro */ }
  }

  // Ordena por profissional e depois por procId
  var lista = Object.values(combos);
  lista.sort(function(a, b) {
    return a.profId - b.profId || a.procId - b.procId;
  });

  Logger.log('=== Combinações únicas (profissional × procedimento) nos últimos/próximos ' + DIAS_ATRAS + ' dias ===');
  lista.forEach(function(c) {
    Logger.log(
      'Prof: ' + c.profNome + ' (id=' + c.profId + ')' +
      '  |  ProcID: ' + c.procId +
      '  |  Ocorrências: ' + c.count
    );
  });
  Logger.log('Total de combinações únicas: ' + lista.length);
}

// ================================================================
// DEBUG — para cada procId desconhecido, retorna a ocorrência mais recente
// (data + hora + profissional) para você abrir no Feegow e confirmar o nome.
// ================================================================
function debugLocalizarProcsDesconhecidos() {
  var IDS_VERIFICAR = [33, 41, 58, 61, 74, 87, 93, 245, 247, 249, 281, 12];
  var DIAS = 90;
  var profMap = carregarProfissionais();

  // { procId: { data, hora, prof } } — guarda só a mais recente
  var achados = {};
  IDS_VERIFICAR.forEach(function(id) { achados[id] = null; });

  for (var offset = -DIAS; offset <= 0; offset++) {
    var d  = new Date();
    d.setDate(d.getDate() + offset);
    var ds = fmtDataFeegow(d);
    try {
      var resp  = UrlFetchApp.fetch(
        CF_FEEGOW_BASE + '/appoints/search?data_start=' + ds + '&data_end=' + ds,
        { headers: { 'x-access-token': CF_FEEGOW_TOKEN }, muteHttpExceptions: true }
      );
      var json  = JSON.parse(resp.getContentText());
      var items = Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);
      items.forEach(function(ag) {
        if (IDS_VERIFICAR.indexOf(ag.procedimento_id) < 0) return;
        // Sobrescreve sempre — como percorremos do mais antigo ao mais novo, fica o mais recente
        achados[ag.procedimento_id] = {
          data: ds,
          hora: formatHora(ag.horario || ''),
          prof: profMap[ag.profissional_id] || 'id=' + ag.profissional_id
        };
      });
    } catch(e) {}
  }

  Logger.log('=== Última ocorrência de cada procId desconhecido ===');
  Logger.log('Abra cada data no Feegow (agenda do profissional) para confirmar o nome.\n');
  IDS_VERIFICAR.forEach(function(id) {
    var a = achados[id];
    if (a) {
      Logger.log('ProcID=' + id + ' → ' + a.data + ' ' + a.hora + ' | ' + a.prof);
    } else {
      Logger.log('ProcID=' + id + ' → ⚠️ não encontrado nos últimos ' + DIAS + ' dias');
    }
  });
}

// ================================================================
// DEBUG — tenta buscar os nomes de vários procIds de uma vez via API Feegow.
// Lista os IDs que obtiveram nome e os que precisam ser conferidos no calendário.
// ================================================================
function debugNomearProcedimentos() {
  // procIDs ainda não identificados da Marcelle (resultado de debugMarcelleUSGProcIds)
  var IDS_VERIFICAR = [33, 41, 58, 61, 74, 87, 93, 245, 247, 249, 281, 12];

  var endpoints = [
    function(id) { return '/procedure/detail?procedimento_id=' + id; },
    function(id) { return '/procedure/get?procedimento_id='    + id; },
    function(id) { return '/procedure/' + id; },
  ];

  Logger.log('=== Buscando nomes de ' + IDS_VERIFICAR.length + ' procIDs via API Feegow ===');

  IDS_VERIFICAR.forEach(function(procId) {
    var nome = null;
    for (var i = 0; i < endpoints.length && !nome; i++) {
      try {
        var resp = UrlFetchApp.fetch(CF_FEEGOW_BASE + endpoints[i](procId), {
          headers: { 'x-access-token': CF_FEEGOW_TOKEN },
          muteHttpExceptions: true
        });
        if (resp.getResponseCode() === 200) {
          var body = resp.getContentText();
          var json = JSON.parse(body);
          var obj  = json.content || json;
          nome = obj.nome || obj.name || obj.descricao || obj.title || null;
          if (!nome && typeof obj === 'string' && obj.length < 200) nome = obj;
        }
      } catch(e) {}
    }
    Logger.log('ProcID=' + procId + ' → ' + (nome || '⚠️ API não retornou nome'));
    Utilities.sleep(200);
  });

  Logger.log('\nPara os IDs sem nome: rode debugEncontrarProcId() trocando PROC_ID e abra a data retornada no Feegow.');
}

// ================================================================
// DEBUG — encontra todas as datas em que um procId específico aparece.
// Útil para localizar no Feegow o nome do procedimento de um ID desconhecido.
// Troque PROC_ID e execute; anote uma das datas retornadas e abra no Feegow.
// ================================================================
function debugEncontrarProcId() {
  var PROC_ID = 204; // ← troque pelo ID que quer identificar
  var DIAS    = 90;

  var profMap = carregarProfissionais();
  var achados = [];

  for (var offset = -DIAS; offset <= 7; offset++) {
    var d  = new Date();
    d.setDate(d.getDate() + offset);
    var ds = fmtDataFeegow(d);
    try {
      var resp  = UrlFetchApp.fetch(
        CF_FEEGOW_BASE + '/appoints/search?data_start=' + ds + '&data_end=' + ds,
        { headers: { 'x-access-token': CF_FEEGOW_TOKEN }, muteHttpExceptions: true }
      );
      var json  = JSON.parse(resp.getContentText());
      var items = Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);
      items.forEach(function(ag) {
        if (ag.procedimento_id !== PROC_ID) return;
        achados.push({
          data:  ds,
          hora:  formatHora(ag.horario || ''),
          prof:  profMap[ag.profissional_id] || 'id=' + ag.profissional_id
        });
      });
    } catch(e) {}
  }

  Logger.log('=== ProcID=' + PROC_ID + ' — ' + achados.length + ' ocorrências (últimos ' + DIAS + ' dias) ===');
  achados.forEach(function(a) {
    Logger.log(a.data + ' ' + a.hora + ' | ' + a.prof);
  });
  Logger.log('Abra uma dessas datas no Feegow para confirmar o nome do procedimento.');
}

// ================================================================
// DEBUG — identifica os procIDs de USG da Dra Marcelle Moura.
// Varre os últimos 90 dias e lista todos os procIDs dela que ainda
// não estão em nenhuma lista conhecida.
// Após rodar: identifique quais são USG Preparo TEC / USG FIV →
//   adicione em IDS_ULTRAS_TRATAMENTO
// Quais são USG Pós Beta / USG Obstétrica / USG Morfológica →
//   adicione em IDS_OBSTETRICA
// ================================================================
function debugMarcelleUSGProcIds() {
  var DIAS = 90;
  var profMap  = carregarProfissionais();

  // Encontra o ID da Marcelle no mapa
  var marcelleId = null;
  Object.keys(profMap).forEach(function(id) {
    if ((profMap[id] || '').toUpperCase().indexOf('MARCELLE') >= 0) marcelleId = id;
  });
  if (!marcelleId) { Logger.log('Profissional Marcelle não encontrada.'); return; }
  Logger.log('Marcelle ID: ' + marcelleId + ' (' + profMap[marcelleId] + ')');

  var CONHECIDOS = IDS_INJURIA.concat(IDS_OBSTETRICA)
                              .concat(IDS_ULTRAS_TRATAMENTO)
                              .concat(IDS_ONLINE_PROCS)
                              .concat(IDS_SEM_CONFIRMACAO);

  var combos = {}; // procId → count

  for (var offset = -DIAS; offset <= 7; offset++) {
    var d  = new Date();
    d.setDate(d.getDate() + offset);
    var ds = fmtDataFeegow(d);
    try {
      var resp  = UrlFetchApp.fetch(
        CF_FEEGOW_BASE + '/appoints/search?data_start=' + ds + '&data_end=' + ds,
        { headers: { 'x-access-token': CF_FEEGOW_TOKEN }, muteHttpExceptions: true }
      );
      var json  = JSON.parse(resp.getContentText());
      var items = Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);
      items.forEach(function(ag) {
        if (String(ag.profissional_id) !== String(marcelleId)) return;
        var prc = ag.procedimento_id;
        if (!combos[prc]) combos[prc] = 0;
        combos[prc]++;
      });
    } catch(e) {}
  }

  var lista = Object.keys(combos).map(function(k) {
    return { procId: parseInt(k), count: combos[k], conhecido: CONHECIDOS.indexOf(parseInt(k)) >= 0 };
  });
  lista.sort(function(a, b) { return b.count - a.count; });

  Logger.log('=== ProcIDs da Marcelle (últimos ' + DIAS + ' dias) ===');
  lista.forEach(function(c) {
    Logger.log('ProcID=' + c.procId + '  ocorrências=' + c.count +
               (c.conhecido ? '  [JÁ MAPEADO]' : '  ← IDENTIFICAR'));
  });
  Logger.log('\nAbra agendamentos da Marcelle no Feegow para cada procId "IDENTIFICAR".');
  Logger.log('USG Preparo TEC / USG FIV → IDS_ULTRAS_TRATAMENTO');
  Logger.log('USG Pós Beta / Morfológica / Obstétrica → IDS_OBSTETRICA');
}

// ================================================================
// DEBUG — lista todos os agendamentos de uma data específica com procIDs.
// Útil para identificar IDs de procedimentos (punção, transferência etc.)
// Altere DATA para a data desejada no formato DD-MM-YYYY e execute.
// ================================================================
function debugDiaEspecifico() {
  var DATA = '05-05-2026'; // ← altere para a data desejada

  var profMap = carregarProfissionais();
  var resp = UrlFetchApp.fetch(
    CF_FEEGOW_BASE + '/appoints/search?data_start=' + DATA + '&data_end=' + DATA,
    { headers: { 'x-access-token': CF_FEEGOW_TOKEN }, muteHttpExceptions: true }
  );
  var json  = JSON.parse(resp.getContentText());
  var items = Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);

  Logger.log('=== ' + DATA + ' — ' + items.length + ' agendamentos ===');
  items.sort(function(a, b) { return (a.horario || '').localeCompare(b.horario || ''); });
  items.forEach(function(ag) {
    Logger.log(
      formatHora(ag.horario || '') + ' | ' +
      (profMap[ag.profissional_id] || 'profId=' + ag.profissional_id) + ' | ' +
      'ProcID=' + ag.procedimento_id + ' | ' +
      'LocalID=' + ag.local_id
    );
  });
}

// ================================================================
// DEBUG — estrutura completa do 1º agendamento de hoje
// Rode para descobrir quais campos a API retorna
// ================================================================
function debugAgendamentos() {
  var hoje = new Date();
  var ds   = fmtDataFeegow(hoje);
  var url  = CF_FEEGOW_BASE + '/appoints/search?data_start=' + ds + '&data_end=' + ds;
  var resp = UrlFetchApp.fetch(url, {
    headers: { 'x-access-token': CF_FEEGOW_TOKEN },
    muteHttpExceptions: true
  });
  var json  = JSON.parse(resp.getContentText());
  var items = Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);
  Logger.log('Total de agendamentos hoje: ' + items.length);
  if (items.length > 0) {
    Logger.log('Campos do 1º agendamento:\n' + JSON.stringify(items[0], null, 2));
    Logger.log('Todos os CAMPOS disponíveis: ' + Object.keys(items[0]).join(', '));
  } else {
    Logger.log('Resposta completa:\n' + JSON.stringify(json, null, 2));
  }
}

// ================================================================
// DEBUG — mapas de profissionais (lista todos)
// ================================================================
function debugMaps() {
  var profMap = carregarProfissionais();
  Logger.log('PROFISSIONAIS:\n' + JSON.stringify(profMap, null, 2));
}

// ================================================================
// TESTE — imprime todas as mensagens nos Registros (sem enviar)
// Para enviar de verdade ao seu número, descomente o bloco inferior
// ================================================================
function testeEnvio() {
  var MEU_NUMERO = '5521999999999'; // ← coloque seu número aqui

  var dataFicta   = '27/04';
  var horaFicta   = '10:00';
  var diaSemana   = 'segunda';
  var qrLinkFicta = 'https://visitante.in/?l=TESTE';

  var keys = Object.keys(TMPL).filter(function(k) { return k !== 'QR_CODE' && k !== 'ACOMPANHANTES'; });

  keys.forEach(function(k) {
    var msg = fillTemplate(TMPL[k], {
      DATA:        dataFicta,
      HORA:        horaFicta,
      DIA_SEMANA:  diaSemana,
      LINK_QR:     qrLinkFicta,
      DATA_VISITA: dataFicta
    });
    Logger.log('=== ' + k + ' ===\n' + msg + '\n');
  });

  Logger.log('Para enviar de verdade ao seu número, descomente o bloco abaixo:');
  /*
  keys.forEach(function(k) {
    var msg = fillTemplate(TMPL[k], { DATA: dataFicta, HORA: horaFicta, DIA_SEMANA: diaSemana,
                                      LINK_QR: qrLinkFicta, DATA_VISITA: dataFicta });
    sendWhatsApp(MEU_NUMERO, '--- TESTE: ' + k + ' ---\n\n' + msg);
    Utilities.sleep(3000);
  });
  */
}
