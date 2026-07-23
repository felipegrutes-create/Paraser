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

// ---- CONFIG — lidos do PropertiesService (nunca hardcode tokens no código) ----
// Para configurar pela primeira vez: execute configurarPropriedades() uma vez.
var _P = PropertiesService.getScriptProperties();

const CF_FEEGOW_BASE       = 'https://api.feegow.com/v1/api';
const CF_FEEGOW_TOKEN      = _P.getProperty('FEEGOW_TOKEN');
const CF_ZAPI_INSTANCE_ID  = _P.getProperty('ZAPI_INSTANCE_ID');
const CF_ZAPI_TOKEN        = _P.getProperty('ZAPI_TOKEN');
const CF_ZAPI_CLIENT_TOKEN = _P.getProperty('ZAPI_CLIENT_TOKEN');
const CF_SLACK_TOKEN       = _P.getProperty('SLACK_TOKEN');
const CF_SLACK_CHANNEL     = 'atendimento';
const CF_SLACK_CHANNEL_REAG = 'reagendamento';  // pedidos de reagendamento vão pra cá (não #atendimento)
// WhatsApp que a paciente recebe ao pedir reagendamento (link ou resposta).
const MSG_REAGENDAMENTO = 'Recebemos seu pedido de reagendamento! 💜 Estamos no aguardo para te oferecer um novo horário. A recepção vai entrar em contato com você em breve.';
const CF_SPREADSHEET_ID    = _P.getProperty('SPREADSHEET_ID');
const CF_CONFIG_SHEET      = 'Config';
const CF_QR_LINK_CELL      = 'B1';
const CF_LOG_SHEET         = 'Confirmacoes_Log';
const CF_PENDENTES_SHEET   = 'Confirmacoes_Pendentes'; // fila de quem recebeu o botão Sim/Não

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

  MARCELLE_ONLINE:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua consulta ONLINE com a Dra. Marcelle Moura, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'Caso tenha exames peço que envie para o nosso setor de enfermagem, dessa forma iremos anexar ao sistema com mais agilidade. Segue o email: mmenezesmoura@gmail.com\n\n' +
    '⛔ Caso não haja confirmação, a consulta será cancelada. ⛔\n\n' +
    'A consulta será realizada por videochamada no WhatsApp.\n\n' +
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

  JOSELMO_RADIOFREQUENCIA:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua sessão de radiofrequência com o Dr. Joselmo Salvato, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    '⛔ Caso não haja confirmação, a consulta será cancelada. ⛔\n\n' +
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

  // Consulta PRESENCIAL do Dr. Rodolfo em SÃO PAULO (endereço próprio, sem QR/link).
  RODOLFO_SP:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua consulta PRESENCIAL com o Dr. Rodolfo, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'Caso tenha exames peço que envie para o nosso setor de enfermagem, dessa forma iremos anexar ao sistema com mais agilidade. Segue o email: exames@paraser.com.br\n\n' +
    'ENDEREÇO:\n' +
    'Av. Indianópolis, 171 - Indianópolis, São Paulo - SP\n\n' +
    'Podemos confirmar?',

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

  MAGALI_PRESENCIAL:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua consulta PRESENCIAL com a Dra. Magali Miranda, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    '⛔ Caso não haja confirmação, a consulta será cancelada. ⛔\n\n' +
    '{ENDERECO}\n\n' +
    'Podemos confirmar? 💜',

  MAGALI_ONLINE:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua consulta ONLINE com a Dra. Magali Miranda, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    '⛔ Caso não haja confirmação, a consulta será cancelada. ⛔\n\n' +
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

  BIANCA_RECEPTORA_PRESENCIAL:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua conversa PRESENCIAL com a Dra. Bianca Salvato, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    '{ENDERECO}\n\n' +
    'Podemos confirmar? 💜',

  BIANCA_RECEPTORA_ONLINE:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua conversa ONLINE com a Dra. Bianca Salvato, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'A consulta será realizada através de chamada de vídeo no Whatsapp\n\n' +
    'Podemos confirmar? 💜',

  SARA_PRESENCIAL:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua consulta PRESENCIAL com a psicóloga Sara, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'Caso não possa comparecer, é importante que nos avise. Nosso serviço de psicologia é oferecido a todas as pacientes em tratamento de Reprodução Assistida, quando uma consulta não é realizada, outra paciente perde a oportunidade de ser atendida.\n\n' +
    'Agradecemos muito pela compreensão e colaboração. Estamos aqui para oferecer todo suporte que vocês precisam!\n\n' +
    '{ENDERECO}\n\n' +
    'Podemos confirmar a sua consulta? 💜',

  SARA_ONLINE:
    'Olá! Tudo bem?\n' +
    'Passando para confirmar sua consulta ONLINE com a psicóloga Sara, {DIA_SEMANA} ({DATA}) às {HORA}.\n\n' +
    'A CONSULTA SERÁ REALIZADA ATRAVÉS DO GOOGLE MEET.\n' +
    'A PSICÓLOGA SARA ENVIARÁ UMA MENSAGEM COM O LINK PARA A CONSULTA.\n\n' +
    'Caso não possa comparecer, é importante que nos avise. Nosso serviço de psicologia é oferecido a todas as pacientes em tratamento de Reprodução Assistida, quando uma consulta não é realizada, outra paciente perde a oportunidade de ser atendida.\n\n' +
    'Agradecemos muito pela compreensão e colaboração. Estamos aqui para oferecer todo suporte que vocês precisam!\n\n' +
    'Podemos confirmar a sua consulta? 💜',

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

  KATIA_PRESENCIAL:
    'Olá! Tudo bem?\n\n' +
    'Passando para confirmar o seu exame de ultrassom com a Dra. Katia Chamorro, {DIA_SEMANA}, dia {DATA}, às {HORA}.\n\n' +
    '{PREPARO}\n\n' +
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
  // NÃO envia confirmações aos sábados e domingos
  // (feriados são permitidos para enviar confirmações do próximo dia útil)
  var hoje = new Date();
  if (hoje.getDay() === 0 || hoje.getDay() === 6) {
    Logger.log('Hoje é ' + DIAS_SEMANA[hoje.getDay()] + ' — nenhuma confirmação será enviada.');
    return;
  }

  var dia = getDiaAlvo();
  var agendamentos;
  try {
    agendamentos = getAgendamentos(dia);
  } catch(e) {
    // Feegow fora do ar / resposta inválida: não dá pra ler a agenda.
    // Avisa no Slack (em vez de morrer em silêncio) e aborta. Reenviar quando voltar.
    Logger.log('🚨 Feegow indisponível — confirmações de ' + fmtDataBR(dia) + ' NÃO enviadas: ' + e.message);
    notificarFeegowFora(fmtDataBR(dia), e.message);
    return;
  }

  Logger.log('Agendamentos encontrados para ' + fmtDataBR(dia) + ': ' + agendamentos.length);
  if (!agendamentos.length) return;

  agendamentos = deduplicarPorPaciente(agendamentos);
  Logger.log('Após deduplicação por paciente: ' + agendamentos.length);

  var profMap = carregarProfissionais();

  var qrLink    = getQrLink();
  var dataStr   = fmtDataBR(dia);
  var amanha    = new Date(); amanha.setDate(amanha.getDate() + 1);
  var diaSemana = fmtDataFeegow(dia) === fmtDataFeegow(amanha)
                  ? 'amanhã'
                  : DIAS_SEMANA[dia.getDay()];

  // Antes de enviar: confirma que a instância Z-API (WhatsApp) está conectada.
  // Sem isso, o Z-API aceita o envio (HTTP 200) mas não entrega — e o Slack
  // reportaria "enviado" falsamente. Se estiver offline, aborta e avisa.
  if (!zapiConectado()) {
    Logger.log('🚨 Z-API DESCONECTADO — nenhuma confirmação enviada para ' + dataStr + '.');
    notificarZapiDesconectado(dataStr, agendamentos.length);
    return;
  }

  var enviados = 0, semTemplate = 0, semTelefone = 0, erros = 0;
  var revisar = [];   // auditoria: exame caindo em template de consulta

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
      // !TMPL[tmplKey]: a chave existe na regra mas o texto não foi cadastrado (ex:
      // MARCELLE_ONLINE faltando 13/07). Sem esta guarda, fillTemplate(undefined) quebra
      // e a paciente cai em "Erro" sem confirmação. Trata como sem template (logado).
      if (!tmplKey || !TMPL[tmplKey]) {
        logRow(ag, 'SEM_TEMPLATE' + (tmplKey ? ' (' + tmplKey + ' não cadastrado)' : ''), phone, tmplKey || '');
        Logger.log('Sem template: key=' + (tmplKey || '(null)') + ' prof=' + ag._profNome + ' proc=' + ag._procNome + ' procId=' + ag.procedimento_id + ' tele=' + ag.telemedicina);
        semTemplate++;
        return;
      }

      var _aud = _auditarRoteamento(ag, tmplKey);
      if (_aud) revisar.push((ag.paciente_nome || ('id ' + ag.paciente_id)) + ' ' + formatHora(ag.horario || '') + ' — ' + _aud);

      var hora = formatHora(ag.horario || '');
      var msg  = fillTemplate(TMPL[tmplKey], {
        DATA:        dataStr,
        HORA:        hora,
        DIA_SEMANA:  diaSemana,
        LINK_QR:     qrLink,
        DATA_VISITA: dataStr,
        PREPARO:     (tmplKey === 'KATIA_PRESENCIAL') ? _preparoKatia(ag) : ''
      });

      sendWhatsApp(phone, msg);
      Utilities.sleep(2500);

      if (qrLink && !tmplKey.endsWith('_ONLINE') && !TMPL_SEM_QR_PREDIO[tmplKey]) {
        var qrMsg = fillTemplate(TMPL.QR_CODE, { DATA_VISITA: dataStr, LINK_QR: qrLink });
        sendWhatsApp(phone, qrMsg);
        Utilities.sleep(1500);
        sendWhatsApp(phone, TMPL.ACOMPANHANTES);
        Utilities.sleep(1500);
      }

      // Confirmação por LINK (os botões do WhatsApp foram bloqueados pela Meta).
      // Só pra quem ainda NÃO confirmou (status != 7). Clicou no link → confirmarFeegow.
      if (Number(ag.status_id) !== 7) {
        var agId = ag.agendamento_id || ag.id || ag.agenda_id;
        if (agId) {
          sendWhatsApp(phone, _msgConfirmacaoLink(agId));
          registrarPendente(phone, ag.paciente_nome || '', agId, dataStr);
          Utilities.sleep(1500);
        } else {
          Logger.log('Sem agendamento_id pro link — campos: ' + Object.keys(ag).join(', '));
        }
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

  notificarSlackConfirmacoes(dataStr, enviados, semTemplate, semTelefone, erros, agendamentos.length, revisar);
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

  agendamentos = deduplicarPorPaciente(agendamentos);
  Logger.log('Após deduplicação: ' + agendamentos.length + ' pacientes únicos');

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
    (carregarNomesProcedimentos()[ag.procedimento_id] || '') ||
    ''
  );
}

// ================================================================
// FEEGOW — agendamentos do dia
// ================================================================
// status_id do Feegow (/appoints/status) que NÃO devem receber confirmação:
//   3 = Atendido | 6 = Não compareceu | 11 = Desmarcado pelo paciente
//   15 = Remarcado | 22 = Cancelado pelo profissional
// Os demais (1 = Marcado-não confirmado, 7 = Marcado-confirmado, etc.) recebem.
var STATUS_NAO_ENVIAR = [3, 6, 11, 15, 22];

function getAgendamentos(dia) {
  var ds   = fmtDataFeegow(dia);
  var url  = CF_FEEGOW_BASE + '/appoints/search?data_start=' + ds + '&data_end=' + ds;
  var resp = UrlFetchApp.fetch(url, {
    headers: { 'x-access-token': CF_FEEGOW_TOKEN },
    muteHttpExceptions: true
  });
  // Detecta Feegow fora do ar: HTTP != 200, resposta não-JSON (página de erro/gateway)
  // ou success=false. Lança erro pra enviarConfirmacoes avisar no Slack, em vez de
  // confundir "Feegow caiu" com "dia sem agenda" (vazio legítimo, segue em silêncio).
  var code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('Feegow fora do ar (HTTP ' + code + ')');
  }
  var json;
  try {
    json = JSON.parse(resp.getContentText());
  } catch(e) {
    throw new Error('Feegow devolveu resposta inválida (não-JSON)');
  }
  if (json && json.success === false) {
    throw new Error('Feegow retornou erro: ' + (json.message || 'success=false'));
  }
  var items = Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);

  return items.filter(function(a) {
    if (a.status_id != null) return STATUS_NAO_ENVIAR.indexOf(Number(a.status_id)) < 0;
    // fallback antigo (caso a API pare de mandar status_id)
    var s = (a.status || '').toLowerCase();
    return s !== 'cancelado' && s !== 'bloqueado' && s !== 'bloqueio' && s !== 'desmarcado';
  });
}

// ================================================================
// Quando uma paciente tem mais de um agendamento no mesmo dia,
// envia confirmação apenas para o PRIMEIRO horário — evita que
// a paciente ignore o horário mais cedo por só ler a última mensagem.
// ================================================================
function deduplicarPorPaciente(agendamentos) {
  agendamentos.sort(function(a, b) {
    return (a.horario || '').localeCompare(b.horario || '');
  });
  var vistos = {};
  return agendamentos.filter(function(ag) {
    if (!ag.paciente_id || vistos[ag.paciente_id]) return false;
    vistos[ag.paciente_id] = true;
    return true;
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
// FEEGOW — mapa de procedimentos {procId: nome}
// A API /appoints/search só devolve o ID do procedimento, não o nome.
// /procedures/list devolve todos os ~289 procedimentos com nome.
// Buscamos uma vez por execução (memoizado em _procNomeMap) — assim a
// lógica de template pode casar pelo NOME e não depende só de IDs chumbados.
// ================================================================
var _procNomeMap = null;
function carregarNomesProcedimentos() {
  if (_procNomeMap) return _procNomeMap;
  _procNomeMap = {};
  try {
    var resp = UrlFetchApp.fetch(CF_FEEGOW_BASE + '/procedures/list', {
      headers: { 'x-access-token': CF_FEEGOW_TOKEN }, muteHttpExceptions: true
    });
    var json  = JSON.parse(resp.getContentText());
    var items = Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);
    items.forEach(function(p) {
      var id = p.procedimento_id || p.id;
      if (id) _procNomeMap[id] = p.nome || p.name || '';
    });
    Logger.log('Nomes de procedimentos carregados: ' + Object.keys(_procNomeMap).length);
  } catch(e) { Logger.log('carregarNomesProcedimentos erro: ' + e.message); }
  return _procNomeMap;
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
// A API /appoints/search só devolve o ID do procedimento, não o nome —
// por isso resolveNomeProc() agora busca o nome em /procedures/list
// (carregarNomesProcedimentos). Estas listas continuam servindo como
// override rápido e como fallback caso a busca de nomes falhe.
// ================================================================

// procId=69  confirmado: INJÚRIA ENDOMETRIAL (Priscila 29/07/2025 14:45)
// procId=152 confirmado: APLICAÇÃO DE FILGRASTIM (Marcelle 05/08/2025 12:30)
// procId=153 USG REAPLICAÇÃO DE FILGRASTIM
// procId=374 confirmado: INJÚRIA ENDOMETRIAL - DR. RODOLFO SALVATO (23/06/2026) — caía no template do Rodolfo
var IDS_INJURIA = [69, 152, 153, 374];

// procId=42  confirmado: USG TRANSLUCÊNCIA NUCAL (Érica 01/10/2025 14:00)
// procId=40  confirmado: USG 1 PÓS BETA (Marcelle 05/05/2026 12:00)
// procId=41  confirmado: USG 2 PÓS BETA (Priscila 28/04/2026 15:20)
// procId=43  confirmado: USG OBSTÉTRICA COMUM (Érica 22/04/2026)
// procId=44  confirmado: USG MORFOLÓGICA (Érica 27/04/2026)
// procId=45  confirmado: USG OBSTÉTRICA COM DOPPLER (Érica 30/04/2026)
// procId=51  confirmado: USG OBSTÉTRICA COM DOPPLER GEMELAR (Érica 26/02/2026)
// procId=171 confirmado: USG PÓS BETA NEG/ABORTO (Rodolfo 28/04/2026)
//
// REMOVIDOS DAQUI (2026-06-10, bug: estavam recebendo template "obstétrica" com pen drive):
//   procId=58, 61 = USG TRANSVAGINAL (não é obstétrica)
//   procId=59     = USG CONTAGEM DE FOLÍCULOS ANTRAIS (avaliação fertilidade, não obstétrica)
// → movidos para IDS_ULTRAS_TRATAMENTO
var IDS_OBSTETRICA = [40, 41, 42, 43, 44, 45, 51, 171];

// procIds de exames de acompanhamento de tratamento (USG Preparo TEC, USG FIV)
// procId=244 confirmado: USG PREPARO TEC 1 - Medicado (Marcelle 05/05/2026 10:00 / 11:00)
// procId=73  confirmado: USG PREPARO TEC 1 - Natural   (Marcelle 05/05/2026 10:40 / 11:20)
// procId=4   confirmado: USG FIV 1ª                   (Marcelle 05/05/2026 11:40)
// procId=203 confirmado: Consulta 1ª Vez (Marcelle) — roteado por profissional, sem ação
// procId=204 confirmado: Consulta 1ª Vez (Marcelle, 12×) — roteado por profissional, sem ação
// procId=281 confirmado: Consulta 1ª Vez (Marcelle 07/04/2026) — roteado por profissional, sem ação
// procId=249 confirmado: Consulta 1ª Vez (Marcelle 11/03/2026) — roteado por profissional, sem ação
// procId=33  confirmado: Consulta de Retorno (Priscila 04/05/2026) — roteado por profissional, sem ação
// procId=156 confirmado: Consulta Presencial (Priscila 07/05/2026 12:00) — roteado por profissional, sem ação
// procId=313 confirmado: Consulta (Magali Miranda 07/05/2026) — roteado por profissional
// procId=247 confirmado: Consulta 1ª Vez (Marcelle 25/02/2026) — roteado por profissional, sem ação
// procId=4   confirmado: USG FIV 1ª (Marcelle 05/05/2026 11:40)
// procId=5   confirmado: USG FIV 2 (Érica Stein 04/05/2026 17:20)
// procId=6   confirmado: USG FIV 3 (Érica Stein 04/05 + Priscila 06/05)
// procId=7   confirmado: USG FIV 4 (Érica Stein 04/05/2026 09:00)
// procId=12  confirmado: USG COITO PROGRAMADO (Érica 04/05/2026 15:20)
// procId=13  confirmado: USG COITO PROGRAMADO (Érica Stein 17/04/2026 11:10)
// procId=73  confirmado: USG PREPARO TEC 1 - Natural (Marcelle 05/05/2026)
// procId=74  confirmado: USG PREPARO TEC (Érica 30/04/2026)
// procId=244 confirmado: USG PREPARO TEC 1 - Medicado (Marcelle 05/05/2026)
// procId=245 confirmado: USG PREPARO TEC (Érica 04/05/2026)
// procId=59  confirmado: USG CONTAGEM DE FOLÍCULOS ANTRAIS → movido para IDS_OBSTETRICA
// procId=8   confirmado: USG FIV 5 (Érica 24/04/2026)
// procId=9   confirmado: USG FIV 6 (Priscila 29/04/2026)
// procId=122 confirmado: USG PREPARO INJÚRIA (Érica 27/04/2026)
// Nomes reais (via /procedures/list):
//   4-9   = USG FIV 1º..6º
//   11-16 = USG COITO PROGRAMADO 1º..6º
//   17-23 = USG INSEMINAÇÃO 1º..6º
//   24-31 = USG MONITORIZAÇÃO OVULAÇÃO 1º..6º (e variantes)
//   65-67 = USG PREPARO ERA 1..3
//   73-75 = USG PREPARO TEC 1..3 - Natural
//   100   = USG 3D PREP ENDOMETRIAL
//   122   = USG PREPARO INJÚRIA
//   167   = USG PREPARO PRP
//   244-246 = USG PREPARO TEC 1..3 - Medicado
var IDS_ULTRAS_TRATAMENTO = [
  4, 5, 6, 7, 8, 9,
  11, 12, 13, 14, 15, 16,
  17, 18, 19, 20, 21, 22, 23,
  24, 25, 26, 27, 28, 29, 30, 31,
  58, 59, 61,   // ← USG TRANSVAGINAL (58, 61) e CONTAGEM FOLÍCULOS (59) — exames de avaliação/fertilidade, NÃO obstétrica
  65, 66, 67,
  73, 74, 75,
  100, 122, 167,
  244, 245, 246
];

// procId de sessão de radiofrequência do Dr. Joselmo Salvato
// procId=307: Radiofrequência (Joselmo 04/05/2026 10:40 / 11:20)
var IDS_JOSELMO_RADIO = [307];

// procIds de consulta da Psicóloga Sara Estruc
// procId=257: Consulta Psicóloga (Sara Estruc 04/05/2026 09:30) — presencial confirmado
// TODO: identificar procId online de Sara se existir
var IDS_SARA = [257];

// procIds de TODOS os procedimentos com "Online" no nome (fonte: /procedures/list).
// Serve como fallback caso a busca de nomes (carregarNomesProcedimentos) falhe —
// senão a regra `proc.indexOf('ONLINE')` no resolveTemplateKey já cobre tudo.
// 33  = CONSULTA DE RETORNO - Online        | 37  = CONSULTA UROLOGISTA DR. HELCE - Online
// 247-254, 266 = CONSULTA 1ª VEZ ... - Online (vários médicos / projetos)
// 256 = CONSULTA - Online                   | 257 = CONSULTA PSICÓLOGA - Online
// 262 = CONSULTA NUTRICIONISTA - Online     | 280,282,283,285 = 1ª VEZ Mario/Bruna - Online
// 309,310,311 = GINECO/UROLOGISTA - Online  | 318,319,349 = JOSELMO - Online
// 322 = CONVERSA RECEPTORA - Online         | 323 = CONVERSA DOADORA - Online
// 327,331,333,334,337,338,341,342,347 = CONSULTA ONLINE - <médico> (e PROJETO ANA)
// (procedimentos "...Presencial" NÃO entram aqui)
var IDS_ONLINE_PROCS = [
  33, 37,
  247, 248, 249, 250, 251, 252, 253, 254, 256, 257, 262, 266,
  280, 282, 283, 285,
  309, 310, 311, 318, 319,
  322, 323,
  327, 331, 333, 334, 337, 338, 341, 342, 347, 349,
  353   // CONSULTA RETORNO ENDOCRINOLOGISTA - DRA. MAGALI (Online) — adicionado 2026-05-13
];

// procIds de procedimentos que NÃO recebem confirmação por WhatsApp.
// (A regra por nome em resolveTemplateKey já barra qualquer "...TEC..." que não
//  seja "USG PREPARO TEC..."; esta lista é o override garantido / fallback.)
//
// PUNÇÃO:
//   89 = PUNÇÃO DE ÓVULOS | 90 = PUNÇÃO DE ÓVULOS - DOADORA
// TEC / TRANSFERÊNCIA DE EMBRIÃO (todas as variantes, fonte /procedures/list):
//   93  = 1ª TEC
//   124 = 2ª TEC - SP                  | 139 = 1° TEC - SP
//   138 = TEC EMBRIÃO DOADO - Rodolfo  | 225,226 = TEC EMBRIÃO DOADO - Priscila
//   143 = TEC (EMBRIÃO FORMADO COM OUTRO MÉDICO) - Rodolfo
//   232 = ... - Priscila               | 233 = ... - Marcelle
//   234 = 2ª TEC EMBRION - Rodolfo     | 235 = ... - Priscila | 236 = ... - Marcelle
//   305 = 2ª TEC EMBRION - Bruna       | 346 = 2ª TEC EMBRION - Joselmo
//   202 = 2ª TEC EMBRION (PROJETO ANA) - Rodolfo | 227 = ... - Priscila
//   228 = ... - Marcelle               | 304 = ... - Mario
//   258 = 1ª TEC - LAB. ORIGEN         | 259 = 1ª TEC - LAB. PRIMÓRDIA
//   264 = 2ª TEC - LAB. PRIMÓRDIA      | 265 = 2ª TEC - LAB. ORIGEN
//   91  = 2ª Tec Embrion (variante antiga)
// OUTROS (cirúrgico/clínico sem aviso):
//   120 = Aspiração De Cisto | 127 = PRP | 147 = PESA | 88 = Colocação DIU
//   87  = Coleta Preventivo  | 267,268 = FOT Receptora
// OBS: procId=75 ("USG PREPARO TEC 3 - Natural") NÃO entra aqui — vai p/ ULTRAS_TRATAMENTO.
// OBS: procId=176 ("Avaliação Doadora" - Bianca) SAIU daqui em 20/07/2026 — passou a
//      receber confirmação (recepção pediu); agora roteado em IDS_BIANCA_RECEPTORA.
var IDS_SEM_CONFIRMACAO = [
  87, 88, 89, 90, 91, 93,
  120, 124, 127, 138, 139, 143, 147,
  202, 225, 226, 227, 228, 232, 233, 234, 235, 236,
  258, 259, 264, 265, 267, 268,
  304, 305, 346
];

// procIds de Conversa com Receptora (Bianca Salvato)
// procId=35:  CONVERSA RECEPTORA - Presencial (Bianca 04/05/2026 10:00)
// procId=36:  CONVERSA DOADORA - Presencial   (Bianca 14/04/2026 11:00)
// procId=322: CONVERSA RECEPTORA - Online     (Bianca 04/05/2026 08:30)
// procId=323: CONVERSA DOADORA - Online       (Bianca 05/05/2026 08:30)
// procId=176: AVALIAÇÃO DOADORA - Presencial   (Bianca) — add 20/07/2026, recepção pediu confirmação
var IDS_BIANCA_RECEPTORA = [35, 36, 176, 322, 323];

// Consultas do Dr. Rodolfo em SÃO PAULO (endereço próprio, Av. Indianópolis).
// Recebem o template RODOLFO_SP e NÃO recebem o QR Code de acesso ao prédio
// (a catraca é da sede do Rio) — pedido do Felipe 23/07/2026. O link de
// confirmação/remarcação continua sendo enviado normalmente.
// 168 = CONSULTA 1° VEZ - DR. RODOLFO (SP) - Presencial
// 396 = CONSULTA DE RETORNO (São Paulo) - DR. RODOLFO SALVATO
// (248 = 1ª VEZ (SP) - Online já cai em RODOLFO_ONLINE, sem endereço/QR)
var IDS_RODOLFO_SP = [168, 396];

// Templates que NÃO recebem o QR Code de acesso ao prédio do Rio (atendimento
// fora da sede). A confirmação/remarcação por link segue normal.
var TMPL_SEM_QR_PREDIO = { RODOLFO_SP: true };

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

  // --- Dra. Katia Chamorro — USG geral/diagnóstica. Roteia pela MÉDICA antes das
  // regras genéricas de USG (senão transvaginal/pélvica cairiam em tratamento/FIV).
  // O preparo certo é montado por _preparoKatia (pelo nome do exame).
  if (prof.includes('KATIA') || prof.includes('CHAMORRO')) return 'KATIA_PRESENCIAL';

  // --- 1. IDs especiais hardcoded (mais confiável) ---
  if (IDS_RODOLFO_SP.indexOf(procId)        >= 0) return 'RODOLFO_SP';   // consulta do Rodolfo em SP
  if (IDS_INJURIA.indexOf(procId)           >= 0) return 'INJURIA';
  if (IDS_OBSTETRICA.indexOf(procId)        >= 0) return 'ULTRAS_OBSTETRICA';
  if (IDS_ULTRAS_TRATAMENTO.indexOf(procId) >= 0) return 'ULTRAS_TRATAMENTO';
  if (IDS_JOSELMO_RADIO.indexOf(procId)     >= 0) return 'JOSELMO_RADIOFREQUENCIA';
  if (IDS_BIANCA_RECEPTORA.indexOf(procId)  >= 0) return 'BIANCA_RECEPTORA_' + modal;
  if (IDS_SARA.indexOf(procId)              >= 0) return 'SARA_' + modal;

  // --- 2. Nome do procedimento (vem de /procedures/list — ver resolveNomeProc) ---
  // Itens de cobrança/honorário/pacote — não são agendamentos de exame
  if (proc.includes('HONORARIO') || proc.includes('HONORÁRIO') || proc.startsWith('PACOTE'))      return null;
  if (proc.includes('INJUR') || proc.includes('INJÚR') || proc.includes('FILGRASTIM'))            return 'INJURIA';
  // TEC (transferência de embrião) / punção — NUNCA recebem confirmação.
  // Qualquer "...TEC..." cai aqui, EXCETO "USG PREPARO TEC..." (contém "PREPARO").
  if (proc.includes('PUNÇÃO') || proc.includes('PUNCAO'))                                         return null;
  if (proc.indexOf('TEC') >= 0 && proc.indexOf('PREPARO') < 0)                                    return null;
  if (proc.includes('EMBRIÃO') || proc.includes('EMBRIAO') || proc.includes('EMBRION'))           return null;
  if (proc.includes('PRIMÓRDIA') || proc.includes('PRIMORDIA') || proc.includes('ORIGEN'))        return null;
  // USG obstétrica (somente exames REALMENTE obstétricos: gestação acompanhada)
  if (proc.includes('OBSTET') || proc.includes('MORFOL') || proc.includes('TRANSLUC'))            return 'ULTRAS_OBSTETRICA';
  if (proc.includes('POS BETA') || proc.includes('PÓS BETA') || proc.includes('PÓS-BETA'))        return 'ULTRAS_OBSTETRICA';
  // ⚠️ NÃO inclua TRANSVAGINAL, DOPPLER ou FOLÍCULO genéricos aqui — são exames de avaliação/fertilidade.
  //    "USG OBSTÉTRICA COM DOPPLER" já é capturado pela regra OBSTET acima.
  //    DOPPLER sozinho (transvaginal/folicular) cai em ULTRAS_TRATAMENTO via regra USG mais abaixo.
  // USG de tratamento (preparo TEC/ERA, FIV, coito programado, inseminação, monitorização...)
  if (proc.includes('PREPARO') || proc.includes('FIV'))                                           return 'ULTRAS_TRATAMENTO';
  if (proc.includes('COITO') || proc.includes('INSEMINA') || proc.includes('MONITORIZA'))         return 'ULTRAS_TRATAMENTO';
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
  if (prof.includes('MAGALI'))                                       return 'MAGALI_'    + modal;
  if (prof.includes('GRAZIELA') || prof.includes('NUTRI'))          return 'NUTRI_'     + modal;
  if (prof.includes('SARA'))                                        return 'SARA_'      + modal;
  // Érica faz só exames; qualquer procId não identificado → tratamento
  if (prof.includes('ERICA') || prof.includes('ÉRICA'))             return 'ULTRAS_TRATAMENTO';

  return null;
}

// ================================================================
// AUDITORIA DE ROTEAMENTO — pega exame/injúria caindo em template de CONSULTA
// (foi o que aconteceu com a INJÚRIA ENDOMETRIAL procId 374 → RODOLFO_PRESENCIAL)
// Devolve um aviso (string) se suspeito, ou null se ok.
// ================================================================
var TMPL_CONSULTA = [
  'RODOLFO_PRESENCIAL','RODOLFO_ONLINE','PRISCILA_PRESENCIAL','PRISCILA_ONLINE',
  'MARCELLE_PRESENCIAL','MARCELLE_ONLINE','BRUNA_PRESENCIAL','BRUNA_ONLINE',
  'MARIO_PRESENCIAL','MARIO_ONLINE','JOSELMO_PRESENCIAL','JOSELMO_ONLINE',
  'HELCE_PRESENCIAL','HELCE_ONLINE','MAGALI_PRESENCIAL','MAGALI_ONLINE',
  'NUTRI_PRESENCIAL','NUTRI_ONLINE'
];
function _auditarRoteamento(ag, tmplKey) {
  var proc = (ag._procNome || '').toUpperCase();
  // palavras que indicam EXAME/procedimento (não consulta)
  var ehExame = /INJ[ÚU]R|FILGRASTIM|\bUSG\b|ULTRASS|PREPARO|\bFIV\b|PUN[ÇC][ÃA]O|EMBRI|DOPPLER|FOL[ÍI]CUL|MONITORIZA/.test(proc);
  if (!ehExame) return null;
  if (TMPL_CONSULTA.indexOf(tmplKey) >= 0) {
    return 'procId ' + ag.procedimento_id + ' "' + (ag._procNome || '?') + '" → ' + tmplKey +
           ' (parece EXAME indo p/ CONSULTA — conferir mapeamento)';
  }
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
    .replace(/{DATA_VISITA}/g, vars.DATA_VISITA || '')
    .replace(/{PREPARO}/g,     vars.PREPARO     || '');
}

// Preparo do exame de USG da Dra. Katia, escolhido pelo NOME do procedimento.
function _preparoKatia(ag) {
  var n = (ag._procNome || '').toUpperCase();
  var EX = 'Trazer exames anteriores.';
  var BEXIGA = 'Bexiga cheia no momento do exame: comece a beber 500 mL de água aos poucos, 30 minutos antes do horário.';
  if (n.indexOf('ABDOME') >= 0 || n.indexOf('ABDÔME') >= 0)
    return '📋 *Preparo do exame:*\n• Jejum de 8 horas\n• Pode beber água ou água de coco\n• Pode tomar seus medicamentos habituais\n• ' + BEXIGA + '\n• ' + EX;
  if (n.indexOf('MAMA') >= 0)
    return '📋 *Preparo do exame:*\n• ' + EX + '\n• Trazer laudos de biópsia, caso já tenha realizado';
  if (n.indexOf('TIRE') >= 0 || n.indexOf('CERVICAL') >= 0)
    return '📋 *Preparo do exame:*\n• ' + EX + '\n• Trazer laudos de punção, caso já tenha realizado\n• Retirar colares e acessórios do pescoço';
  if (n.indexOf('PAREDE') >= 0 || n.indexOf('INGUIN') >= 0 || n.indexOf('ESCROT') >= 0 || n.indexOf('VIRILHA') >= 0 || n.indexOf('ÍNTIMA') >= 0 || n.indexOf('INTIMA') >= 0)
    return '📋 *Preparo do exame:*\n• Em caso de excesso de pelos, recomenda-se depilação prévia, conforme sua preferência\n• ' + EX;
  if (n.indexOf('TRANSVAGINAL') >= 0 || n.indexOf('PELVIC') >= 0 || n.indexOf('PÉLVIC') >= 0 || n.indexOf('URINAR') >= 0 || n.indexOf('URINÁR') >= 0 || n.indexOf('PROSTAT') >= 0 || n.indexOf('PRÓSTAT') >= 0)
    return '📋 *Preparo do exame:*\n• ' + BEXIGA + '\n• ' + EX;
  return '📋 *Preparo do exame:*\n• ' + EX; // genérico, caso entre um exame novo
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
// Z-API — verifica se a instância (WhatsApp do celular) está conectada.
// Retorna true só se o WhatsApp está online e pronto pra enviar.
// Qualquer falha na checagem retorna false (falha segura: não envia às cegas).
// ================================================================
function zapiConectado() {
  var url = 'https://api.z-api.io/instances/' + CF_ZAPI_INSTANCE_ID + '/token/' + CF_ZAPI_TOKEN + '/status';
  var headers = {};
  if (CF_ZAPI_CLIENT_TOKEN) headers['Client-Token'] = CF_ZAPI_CLIENT_TOKEN;
  var resp = UrlFetchApp.fetch(url, {
    method:      'GET',
    headers:     headers,
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    Logger.log('zapiConectado: HTTP ' + code + ' ' + resp.getContentText().substring(0, 200));
    return false;
  }
  var data = JSON.parse(resp.getContentText());
  return data.connected === true;
}

// ================================================================
// LISTA DE PACIENTES — configura trigger às 8h para esta função
// Popula a aba "Pacientes_Amanha" com nomes formatados (≤40 chars)
// para a recepção usar no Keyaccess antes das 10h30
// ================================================================
function listarPacientesAmanha() {
  // NÃO gera lista / posta no Slack aos sábados, domingos e feriados
  var hoje = new Date();
  if (hoje.getDay() === 0 || hoje.getDay() === 6 || ehFeriado(hoje)) {
    Logger.log('Hoje é ' + DIAS_SEMANA[hoje.getDay()] +
               (ehFeriado(hoje) ? ' / feriado' : '') +
               ' — lista de pacientes não será gerada.');
    return;
  }

  var dia          = getDiaAlvo();
  var dataStr      = fmtDataBR(dia);

  // Feegow pode dar um blip às 8h. Tenta 1x, espera 5s e tenta de novo; se ainda assim
  // falhar, AVISA a recepção no Slack (antes ela falhava em silêncio e ninguém sabia).
  var agendamentos;
  try {
    agendamentos = getAgendamentos(dia);
  } catch (e1) {
    Utilities.sleep(5000);
    try {
      agendamentos = getAgendamentos(dia);
    } catch (e2) {
      Logger.log('listarPacientesAmanha: Feegow fora — ' + e2.message);
      try {
        slackPost('🚨 *Lista da recepção (Keyaccess) não saiu.* A Feegow não respondeu na hora de gerar a lista de ' +
                  dataStr + '. Assim que ela voltar, dá pra gerar de novo pra criar o link do Keyaccess.');
      } catch (se) {}
      return;
    }
  }

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
// SLACK — notifica resultado do envio de confirmações
// ================================================================
function notificarSlackConfirmacoes(dataStr, enviados, semTemplate, semTelefone, erros, total, revisar) {
  try {
    var channelId = slackGetChannelId(CF_SLACK_CHANNEL);
    var texto;

    if (erros > 0 && enviados === 0) {
      texto = '🚨 *Confirmações NÃO enviadas para ' + dataStr + '*\n' +
              'Todos os ' + total + ' pacientes falharam.\n' +
              'Erros: ' + erros + ' | Sem template: ' + semTemplate + ' | Sem telefone: ' + semTelefone + '\n' +
              '⚠️ Verifique o Z-API e reenvie manualmente.';
    } else if (erros > 0) {
      texto = '⚠️ *Confirmações enviadas com erros — ' + dataStr + '*\n' +
              '✅ Enviados: ' + enviados + ' | ❌ Erros: ' + erros +
              ' | Sem template: ' + semTemplate + ' | Sem telefone: ' + semTelefone;
    } else {
      texto = '✅ *Confirmações enviadas — ' + dataStr + '*\n' +
              'Enviados: ' + enviados + ' de ' + total + ' pacientes.' +
              (semTemplate > 0 ? ' (' + semTemplate + ' sem template)' : '');
    }

    if (revisar && revisar.length) {
      texto += '\n\n🔎 *REVISAR ROTEAMENTO (' + revisar.length + ')* — pode ter ido no template errado:\n• ' + revisar.join('\n• ');
    }

    UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method:      'post',
      contentType: 'application/json; charset=utf-8',
      headers:     { Authorization: 'Bearer ' + CF_SLACK_TOKEN },
      payload:     JSON.stringify({ channel: channelId, text: texto })
    });
  } catch(e) {
    Logger.log('notificarSlackConfirmacoes erro: ' + e.message);
  }
}

// ================================================================
// SLACK — avisa que o Z-API está desconectado e NADA foi enviado
// ================================================================
function notificarZapiDesconectado(dataStr, total) {
  try {
    var channelId = slackGetChannelId(CF_SLACK_CHANNEL);
    var texto = '🚨 *Z-API DESCONECTADO — confirmações de ' + dataStr + ' NÃO foram enviadas*\n' +
                total + ' paciente(s) ficaram SEM confirmação.\n' +
                'O WhatsApp da clínica está desconectado. Reconecte o QR Code em app.z-api.io ' +
                'e rode as confirmações de novo.';
    UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method:      'post',
      contentType: 'application/json; charset=utf-8',
      headers:     { Authorization: 'Bearer ' + CF_SLACK_TOKEN },
      payload:     JSON.stringify({ channel: channelId, text: texto })
    });
  } catch(e) {
    Logger.log('notificarZapiDesconectado erro: ' + e.message);
  }
}

// ================================================================
// SLACK — avisa que o Feegow está fora do ar e NADA foi enviado
// (a agenda não pôde ser lida, então nenhuma confirmação saiu)
// ================================================================
function notificarFeegowFora(dataStr, motivo) {
  try {
    var channelId = slackGetChannelId(CF_SLACK_CHANNEL);
    var texto = '🚨 *Feegow fora do ar — confirmações de ' + dataStr + ' NÃO foram enviadas*\n' +
                'Não consegui ler a agenda no Feegow (' + motivo + '), então ninguém recebeu confirmação.\n' +
                'Quando o Feegow voltar, rode a função *enviarConfirmacoes* no editor pra mandar as do dia.';
    UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method:      'post',
      contentType: 'application/json; charset=utf-8',
      headers:     { Authorization: 'Bearer ' + CF_SLACK_TOKEN },
      payload:     JSON.stringify({ channel: channelId, text: texto })
    });
  } catch(e) {
    Logger.log('notificarFeegowFora erro: ' + e.message);
  }
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
  '04/06', // Corpus Christi 2026
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
  var IDS_VERIFICAR = [8, 9, 11, 36, 43, 44, 45, 51, 75, 88, 90, 104, 122, 147, 156, 168, 171, 183, 246, 251, 257, 262, 267, 283, 285, 317, 323, 325, 328, 347];
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
// DEBUG — encontra procIDs candidatos a TEC/punção ainda não mapeados.
// Varre 90 dias e lista procIDs que aparecem cedo (08:00-09:30) para
// médicos com template — horário típico de procedimentos de laboratório.
// Adicione os confirmados em IDS_SEM_CONFIRMACAO.
// ================================================================
function debugTECProcIds() {
  var DIAS    = 90;
  var profMap = carregarProfissionais();
  var CONHECIDOS = IDS_INJURIA.concat(IDS_OBSTETRICA)
                              .concat(IDS_ULTRAS_TRATAMENTO)
                              .concat(IDS_BIANCA_RECEPTORA)
                              .concat(IDS_ONLINE_PROCS)
                              .concat(IDS_SEM_CONFIRMACAO);

  var combos = {}; // procId → { profs, count, exemplo }

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
        var hora = formatHora(ag.horario || '');
        if (hora > '10:00') return; // TEC/punção são sempre de manhã cedo
        var prc = ag.procedimento_id;
        if (CONHECIDOS.indexOf(prc) >= 0) return;
        var profNome = profMap[ag.profissional_id] || 'id=' + ag.profissional_id;
        if (!combos[prc]) combos[prc] = { profs: {}, count: 0, data: ds, hora: hora };
        combos[prc].profs[profNome] = true;
        combos[prc].count++;
        combos[prc].data = ds; // mantém a mais recente
        combos[prc].hora = hora;
      });
    } catch(e) {}
  }

  var lista = Object.keys(combos).map(function(k) {
    return {
      procId: parseInt(k),
      count:  combos[k].count,
      profs:  Object.keys(combos[k].profs).join(', '),
      data:   combos[k].data,
      hora:   combos[k].hora
    };
  });
  lista.sort(function(a, b) { return b.count - a.count; });

  Logger.log('=== ProcIDs candidatos a TEC/punção (antes das 10h, últimos ' + DIAS + ' dias) ===');
  lista.forEach(function(c) {
    Logger.log('ProcID=' + c.procId + '  ×' + c.count +
               '  última ocorrência: ' + c.data + ' ' + c.hora +
               '  | ' + c.profs);
  });
  Logger.log('\nAbra as datas no Feegow, confirme o nome e adicione em IDS_SEM_CONFIRMACAO.');
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
  var DATA = '04-05-2026'; // ← altere para a data desejada

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
// SETUP — execute UMA VEZ para salvar os tokens no PropertiesService.
// Após executar, os valores ficam armazenados com segurança no Apps Script
// e não precisam mais estar no código.
// ================================================================
function configurarPropriedades() {
  PropertiesService.getScriptProperties().setProperties({
    'FEEGOW_TOKEN':      '',  // ← cole o token Feegow
    'ZAPI_INSTANCE_ID':  '',  // ← cole o Instance ID do Z-API
    'ZAPI_TOKEN':        '',  // ← cole o Token do Z-API
    'ZAPI_CLIENT_TOKEN': '',  // ← cole o Client-Token do Z-API
    'SLACK_TOKEN':       '',  // ← cole o token Slack (xoxb-...)
    'SPREADSHEET_ID':    '',  // ← cole o ID da planilha Google Sheets
  });
  Logger.log('✅ Propriedades salvas. Pode apagar os valores desta função.');
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

// ================================================================
// ================================================================
// CONFIRMAÇÃO INTERATIVA (Sim/Não) — botão WhatsApp + webhook + Feegow
// ----------------------------------------------------------------
// Fluxo:
//  1) enviarConfirmacoes() chama sendBotaoConfirmacao() e registrarPendente()
//  2) paciente clica Sim/Não (ou responde texto) -> Z-API chama doPost()
//  3) Sim  -> confirmarFeegow() muda status 1->7  + avisa Slack
//     Não  -> só avisa Slack (recepção reagenda manual)
// ================================================================

// ----------------------------------------------------------------
// WEBHOOK — Z-API chama esta URL a cada mensagem recebida.
// Ignora tudo que não for resposta a uma confirmação pendente.
// ----------------------------------------------------------------
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body && body.src === 'manychat') { _gravarNayara(body); return _okText(); } // 📊 coletor das conversas do front da Nayara (Manychat)
    _gravarMsgRecebida(body); // 📊 gravador temporário p/ mapear demandas da recepção (remover após análise)
    try { _repassarMonitor(body); } catch (e) {} // 📡 fan-out pro monitor (recepção lê as msgs); BLINDADO: nunca afeta a confirmação
    if (body.fromMe === true) return _okText();

    var r = interpretarResposta(body);
    if (!r.phone || !r.answer) return _okText(); // não é Sim/Não -> ignora

    processarResposta(r.phone, r.answer);
  } catch (err) {
    Logger.log('doPost erro: ' + err.message);
  }
  return _okText(); // sempre 200 pro Z-API não reenviar em loop
}

// 📡 Fan-out pro monitor WhatsApp: repassa o payload cru da Z-API pro /exec do CRM
// (?action=zapi_webhook&wk=...), que grava no BigQuery. Assim a recepção (número
// das confirmações) é lida SEM trocar o webhook desta instância. A URL (com a wk)
// fica em MONITOR_INGEST_URL (setada 1x via ?action=set_monitor). Sem ela, no-op.
function _repassarMonitor(body) {
  var url = _P.getProperty('MONITOR_INGEST_URL');
  if (!url) return;
  UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
}

// Liga o "notify sent by me" NESTA instância (a das confirmações), pra o monitor
// ver também as mensagens ENVIADAS pela recepção. O doPost já ignora fromMe (linha
// do return), então isto NÃO afeta a confirmação — só faz as enviadas passarem pelo
// doPost e serem repassadas pro monitor. NÃO mexe no webhook-received.
function _ligarNotifySent() {
  try {
    var url = 'https://api.z-api.io/instances/' + CF_ZAPI_INSTANCE_ID + '/token/' + CF_ZAPI_TOKEN + '/update-notify-sent-by-me';
    var r = UrlFetchApp.fetch(url, {
      method: 'put', contentType: 'application/json',
      headers: CF_ZAPI_CLIENT_TOKEN ? { 'Client-Token': CF_ZAPI_CLIENT_TOKEN } : {},
      payload: JSON.stringify({ notifySentByMe: true }), muteHttpExceptions: true
    });
    return { code: r.getResponseCode(), body: r.getContentText().slice(0, 200) };
  } catch (e) { return { error: e.message }; }
}

// 📊 Coletor das conversas do front da Nayara — o Manychat manda (via External Request) uma cópia de cada msg do lead
function _gravarNayara(body) {
  try {
    var ss = SpreadsheetApp.openById(CF_SPREADSHEET_ID);
    var sh = ss.getSheetByName('Msgs_Nayara');
    if (!sh) { sh = ss.insertSheet('Msgs_Nayara'); sh.appendRow(['Data', 'Telefone', 'Nome', 'Mensagem']); }
    var tel  = String(body.tel  || body.phone || '');
    var nome = String(body.nome || body.name  || '');
    var msg  = String(body.msg  || body.text  || '');
    if (!tel && !msg) return;
    sh.appendRow([new Date(), tel, nome, msg]);
  } catch (err) { Logger.log('_gravarNayara erro: ' + err.message); }
}

function doGet(e) {
  var params = (e && e.parameter) || {};
  // Confirmação por link (paciente clica — sem key; o token no link autentica)
  if (params.a === 'c' && params.ag) return _respConf(_confirmarViaLink(params.ag, params.t || ''), params.fmt);
  if (params.a === 'r' && params.ag) return _respConf(_reagendarViaLink(params.ag, params.t || ''), params.fmt);
  if (params.action === 'diag' && params.key === 'paraser2026') {
    return ContentService
      .createTextOutput(JSON.stringify(_diagConfirmacoes(), null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // Liga o fan-out pro monitor: guarda a URL de ingestão + liga o notify-sent.
  // Devolve o instance_id (pra rotular a recepção no monitor). ?url= vazio só liga o notify.
  if (params.action === 'set_monitor' && params.key === 'paraser2026') {
    if (params.url) _P.setProperty('MONITOR_INGEST_URL', String(params.url));
    var _ns = _ligarNotifySent();
    return ContentService.createTextOutput(JSON.stringify({
      ok: true, monitor_url: _P.getProperty('MONITOR_INGEST_URL') || '(vazio)',
      instance_id: CF_ZAPI_INSTANCE_ID, notify_sent: _ns
    }, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'setup-webhook' && params.key === 'paraser2026') {
    return ContentService
      .createTextOutput(JSON.stringify(_configurarWebhookZapi(params.url || ''), null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // Lista os triggers do projeto (diagnóstico: confirmar se o job das 8h existe).
  if (params.action === 'list-triggers' && params.key === 'paraser2026') {
    var _trs = ScriptApp.getProjectTriggers().map(function(t){
      return { fn: t.getHandlerFunction(), tipo: String(t.getEventType()), fonte: String(t.getTriggerSource()) };
    });
    return ContentService.createTextOutput(JSON.stringify(_trs, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  // Reinstala o trigger diário 8h da lista da recepção (idempotente: remove os antigos primeiro).
  if (params.action === 'setup-trigger-lista' && params.key === 'paraser2026') {
    ScriptApp.getProjectTriggers().forEach(function(t){
      if (t.getHandlerFunction() === 'listarPacientesAmanha') ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('listarPacientesAmanha').timeBased().atHour(8).everyDays(1).inTimezone('America/Sao_Paulo').create();
    return ContentService.createTextOutput('{"ok":true,"msg":"trigger listarPacientesAmanha diario 8h (America/Sao_Paulo) reinstalado"}').setMimeType(ContentService.MimeType.JSON);
  }
  // Regera a lista de pacientes da recepção (Keyaccess) sob demanda — desbloqueia
  // quando o job das 8h não rodou/falhou. Escreve na aba Pacientes_Amanha e posta no Slack.
  if (params.action === 'gerar-lista-recepcao' && params.key === 'paraser2026') {
    var _rl;
    try { listarPacientesAmanha(); _rl = { ok: true, msg: 'lista regerada — veja a aba Pacientes_Amanha e o Slack' }; }
    catch (e) { _rl = { ok: false, erro: e.message }; }
    return ContentService.createTextOutput(JSON.stringify(_rl, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'setup-medicos' && params.key === 'paraser2026') {
    setupListaMedicos();
    return ContentService.createTextOutput('{"ok":true,"msg":"aba Medicos_Agenda criada"}').setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'setup-trigger-agenda' && params.key === 'paraser2026') {
    setupTriggerAgendaMedicos();
    return ContentService.createTextOutput('{"ok":true,"msg":"triggers sexta 17h + seg-qui 18h criados"}').setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'preview-agenda' && params.key === 'paraser2026') {
    return ContentService
      .createTextOutput(JSON.stringify(_previewAgendaMedicos(params.para || '', params.dia || ''), null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'enviar-agenda-todos' && params.key === 'paraser2026') {
    enviarAgendaMedicos();
    return ContentService.createTextOutput('{"ok":true,"msg":"dispatch iniciado, veja log + Slack"}').setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'debug-match' && params.key === 'paraser2026') {
    return ContentService
      .createTextOutput(JSON.stringify(_debugMatchAgenda(params.dia || ''), null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'set-webapp-url' && params.key === 'paraser2026') {
    _P.setProperty('WEBAPP_URL', params.url || '');
    return ContentService.createTextOutput(JSON.stringify({ ok: true, url: params.url || '' })).setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'test-slack' && params.key === 'paraser2026') {
    var waT = formatPhone(params.phone || '5521984341020');
    slackPostReag('🧪 [TESTE] 🔄 *Pediu reagendar (link)* — Paciente de Teste (agendamento 99999).\n📲 <https://wa.me/' + waT + '|Chamar a paciente no WhatsApp>');
    return ContentService.createTextOutput(JSON.stringify({ ok: true, wa: waT })).setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'slack-diag' && params.key === 'paraser2026') {
    var diag = { canalAlvo: CF_SLACK_CHANNEL_REAG };
    try {
      var auth = JSON.parse(UrlFetchApp.fetch('https://slack.com/api/auth.test',
        { method: 'post', headers: { Authorization: 'Bearer ' + CF_SLACK_TOKEN }, muteHttpExceptions: true }).getContentText());
      diag.app = auth.user || '?'; diag.appId = auth.user_id || auth.bot_id || '';
    } catch (e) { diag.authErro = e.message; }
    try {
      var lst = JSON.parse(UrlFetchApp.fetch('https://slack.com/api/conversations.list?limit=1000&exclude_archived=true&types=public_channel,private_channel',
        { headers: { Authorization: 'Bearer ' + CF_SLACK_TOKEN } }).getContentText());
      diag.canaisReag = (lst.channels || []).filter(function (c) { return c.name.indexOf('reagend') >= 0; })
        .map(function (c) { return { nome: c.name, appEhMembro: !!c.is_member }; });
      diag.reagExiste = diag.canaisReag.length > 0;
    } catch (e) { diag.listErro = e.message; }
    return ContentService.createTextOutput(JSON.stringify(diag)).setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'primeira-usg-sim' && params.key === 'paraser2026') {
    return ContentService.createTextOutput(JSON.stringify(_simularPrimeiraUsg_(Number(params.limit) || 30))).setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'backfill-usg' && params.key === 'paraser2026') {
    return ContentService.createTextOutput(JSON.stringify(backfillHistoricoUSG(Number(params.meses) || 24))).setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'primeira-usg-run' && params.key === 'paraser2026') {
    notificarPrimeiraUSG();
    return ContentService.createTextOutput(JSON.stringify({ ok: true, msg: 'notificarPrimeiraUSG executada' })).setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'setup-primeira-usg' && params.key === 'paraser2026') {
    ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'notificarPrimeiraUSG') ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('notificarPrimeiraUSG').timeBased().atHour(7).everyDays(1).inTimezone('America/Sao_Paulo').create();
    return ContentService.createTextOutput(JSON.stringify({ ok: true, trigger: 'notificarPrimeiraUSG diário 7h' })).setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'dump-pendentes' && params.key === 'paraser2026') {
    var ps = pendentesSheet_();
    var pout = { total: 0, rows: [] };
    if (ps.getLastRow() > 1) {
      pout.rows = ps.getRange(2, 1, ps.getLastRow() - 1, 6).getValues();
      pout.total = pout.rows.length;
    }
    return ContentService.createTextOutput(JSON.stringify(pout)).setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'dump-msgs' && params.key === 'paraser2026') {
    var dss = SpreadsheetApp.openById(CF_SPREADSHEET_ID);
    var dsh = dss.getSheetByName('Msgs_Recebidas');
    var out = { total: 0, msgs: [], planilha: dss.getName(), planilha_url: dss.getUrl(), aba: 'Msgs_Recebidas' };
    if (dsh && dsh.getLastRow() > 1) {
      out.msgs = dsh.getRange(2, 1, dsh.getLastRow() - 1, 5).getValues();
      out.total = out.msgs.length;
    }
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'dump-nayara' && params.key === 'paraser2026') {
    var nss = SpreadsheetApp.openById(CF_SPREADSHEET_ID);
    var nsh = nss.getSheetByName('Msgs_Nayara');
    var nout = { total: 0, msgs: [], aba: 'Msgs_Nayara' };
    if (nsh && nsh.getLastRow() > 1) {
      nout.msgs = nsh.getRange(2, 1, nsh.getLastRow() - 1, 4).getValues();
      nout.total = nout.msgs.length;
    }
    return ContentService.createTextOutput(JSON.stringify(nout)).setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'test-confirma' && params.key === 'paraser2026') {
    var pNome = params.proc || 'USG DE ABDOMEN TOTAL';
    var msgK = fillTemplate(TMPL.KATIA_PRESENCIAL, {
      DATA: '17/06', HORA: '10:00', DIA_SEMANA: 'amanhã',
      PREPARO: _preparoKatia({ _procNome: pNome })
    });
    if (params.phone) sendWhatsApp(params.phone, msgK);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, proc: pNome, msg: msgK })).setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'test-link' && params.key === 'paraser2026') {
    var agT = params.ag || '000000';
    if (params.phone) sendWhatsApp(params.phone, _msgConfirmacaoLink(agT));
    return ContentService.createTextOutput(JSON.stringify({ ok: true, sent_to: params.phone || '(nao enviado)',
      mensagem: _msgConfirmacaoLink(agT),
      link_confirmar: linkConfirmacao(agT, 'c'),
      link_reagendar: linkConfirmacao(agT, 'r') })).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('Confirmacoes online')
    .setMimeType(ContentService.MimeType.TEXT);
}

function _configurarWebhookZapi(webhookUrl) {
  if (!webhookUrl) return { erro: 'Passe ?url=<URL do Web App>' };

  // Configurar APENAS o webhook de mensagem recebida.
  // Limpar os outros pra evitar bombardeio de eventos irrelevantes.
  var config = [
    { ep: 'update-webhook-received',       valor: webhookUrl }, // SET nosso
    { ep: 'update-webhook-message-status', valor: '' }          // CLEAR (sem ruído)
  ];

  var results = {};
  config.forEach(function(c) {
    try {
      var url = 'https://api.z-api.io/instances/' + CF_ZAPI_INSTANCE_ID +
                '/token/' + CF_ZAPI_TOKEN + '/' + c.ep;
      var headers = { 'Content-Type': 'application/json' };
      if (CF_ZAPI_CLIENT_TOKEN) headers['Client-Token'] = CF_ZAPI_CLIENT_TOKEN;
      var resp = UrlFetchApp.fetch(url, {
        method: 'put',
        headers: headers,
        payload: JSON.stringify({ value: c.valor }),
        muteHttpExceptions: true
      });
      results[c.ep] = { http: resp.getResponseCode(), body: resp.getContentText().substring(0, 200) };
    } catch (e) { results[c.ep] = { erro: e.message }; }
  });

  return { webhook_configurado: webhookUrl, resultados: results };
}

function _diagConfirmacoes() {
  var out = {};

  // 1) Status do Z-API + URL do webhook configurado
  try {
    var url = 'https://api.z-api.io/instances/' + CF_ZAPI_INSTANCE_ID +
              '/token/' + CF_ZAPI_TOKEN + '/webhook-receive-all-messages';
    var headers = {};
    if (CF_ZAPI_CLIENT_TOKEN) headers['Client-Token'] = CF_ZAPI_CLIENT_TOKEN;
    var resp = UrlFetchApp.fetch(url, { method: 'get', headers: headers, muteHttpExceptions: true });
    out.zapi_webhook_receive_all = { http: resp.getResponseCode(), body: resp.getContentText().substring(0, 500) };
  } catch (e1) { out.zapi_webhook_receive_all = { erro: e1.message }; }

  try {
    var url2 = 'https://api.z-api.io/instances/' + CF_ZAPI_INSTANCE_ID +
               '/token/' + CF_ZAPI_TOKEN + '/status';
    var headers2 = {};
    if (CF_ZAPI_CLIENT_TOKEN) headers2['Client-Token'] = CF_ZAPI_CLIENT_TOKEN;
    var resp2 = UrlFetchApp.fetch(url2, { method: 'get', headers: headers2, muteHttpExceptions: true });
    out.zapi_status = { http: resp2.getResponseCode(), body: resp2.getContentText().substring(0, 300) };
  } catch (e2) { out.zapi_status = { erro: e2.message }; }

  // 2) Últimas 20 linhas da fila de pendentes
  try {
    var sh = pendentesSheet_();
    var n  = sh.getLastRow();
    var startRow = Math.max(2, n - 19);
    if (n >= 2) {
      var vals = sh.getRange(startRow, 1, n - startRow + 1, sh.getLastColumn()).getValues();
      out.pendentes_ultimos = vals.map(function(r) {
        return { timestamp: String(r[0]), telefone: r[1], paciente: r[2], agId: r[3], data: String(r[4]), status: r[5] };
      });
    } else {
      out.pendentes_ultimos = [];
    }

    // 3) Contagem por status (últimos 30 dias)
    var contagem = {};
    if (n >= 2) {
      var todasLinhas = sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues();
      var trintaDias = new Date(Date.now() - 30*24*3600*1000);
      todasLinhas.forEach(function(r) {
        var ts = new Date(r[0]);
        if (ts >= trintaDias) {
          var s = r[5] || 'SEM_STATUS';
          contagem[s] = (contagem[s] || 0) + 1;
        }
      });
    }
    out.contagem_30dias = contagem;
  } catch (e3) { out.pendentes_erro = e3.message; }

  // 4) Última execução de enviarConfirmacoes (vê logRow se houver)
  out.timestamp_diag = new Date().toISOString();

  return out;
}

// ----------------------------------------------------------------
// Envia a pergunta de confirmação com botões Sim/Não (clique-only).
// Se o envio do botão falhar, avisa a recepção no Slack — sem fallback de
// texto, já que agora só o CLIQUE no botão confirma.
// ----------------------------------------------------------------
function sendBotaoConfirmacao(phone, agId) {
  var texto = 'Posso confirmar seu agendamento? 💜\n\n' +
              'É só tocar no botão aqui embaixo 👇';

  var url = 'https://api.z-api.io/instances/' + CF_ZAPI_INSTANCE_ID +
            '/token/' + CF_ZAPI_TOKEN + '/send-button-list';
  var headers = {};
  if (CF_ZAPI_CLIENT_TOKEN) headers['Client-Token'] = CF_ZAPI_CLIENT_TOKEN;

  var payload = {
    phone:   phone,
    message: texto,
    buttonList: {
      buttons: [
        { id: 'CONFIRMAR_SIM', label: 'Sim, confirmo' },
        { id: 'CONFIRMAR_NAO', label: 'Não, reagendar' }
      ]
    }
  };

  var resp = UrlFetchApp.fetch(url, {
    method:             'post',
    contentType:        'application/json',
    headers:            headers,
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    Logger.log('send-button-list HTTP ' + code + ': ' +
               resp.getContentText().substring(0, 200));
    // Sem fallback de texto (só o clique confirma). Se o botão não saiu,
    // avisa a recepção pra tratar manualmente.
    slackPost('⚠️ Não consegui enviar o botão de confirmação para ' + phone +
              ' (agendamento ' + agId + '). Confirmar manualmente com a paciente.');
  }
}

// ----------------------------------------------------------------
// Lê o payload do Z-API e devolve {phone, answer:'SIM'|'NAO'|null}.
// SÓ vale clique no botão (ou seleção de lista). Texto digitado é IGNORADO
// de propósito — senão um "Sim" digitado numa conversa de reagendamento com
// a recepção seria lido como confirmação (bug do print 11/06).
// ----------------------------------------------------------------
function interpretarResposta(body) {
  var phone = body.phone || '';

  // Monta o "sinal" APENAS a partir de campos de clique de botão/lista.
  // Nunca lê body.text (mensagem digitada) — esse é o ponto do fix.
  var sig = '';
  if (body.buttonsResponseMessage) {
    sig += ' ' + (body.buttonsResponseMessage.buttonId || '') +
           ' ' + (body.buttonsResponseMessage.message || '');
  }
  if (body.listResponseMessage) {
    sig += ' ' + (body.listResponseMessage.selectedRowId || '') +
           ' ' + (body.listResponseMessage.title || '') +
           ' ' + (body.listResponseMessage.message || '');
  }
  if (body.buttonReply) {
    sig += ' ' + (body.buttonReply.id || '') +
           ' ' + (body.buttonReply.message || '');
  }
  sig = sig.toUpperCase();

  var answer = null;
  if (sig.indexOf('CONFIRMAR_NAO') !== -1 || sig.indexOf('REAGEND') !== -1) {
    answer = 'NAO';
  } else if (sig.indexOf('CONFIRMAR_SIM') !== -1 || sig.indexOf('CONFIRMO') !== -1) {
    answer = 'SIM';
  }
  return { phone: phone, answer: answer };
}

// ----------------------------------------------------------------
// GRAVADOR TEMPORÁRIO — registra mensagens RECEBIDAS (texto) na aba
// Msgs_Recebidas, só pra mapear os tipos de demanda da recepção.
// Ignora mensagens da própria clínica (fromMe) e grupos. Remover após análise.
// ----------------------------------------------------------------
function _gravarMsgRecebida(body) {
  try {
    if (!body || body.isGroup) return; // grava RECEBIDAS e ENVIADAS (p/ medir tempo de resposta); ignora só grupos
    var txt = (body.text && body.text.message) ? body.text.message
            : (typeof body.text === 'string' ? body.text : '');
    if (!txt) return; // só mensagens de texto
    var ss = SpreadsheetApp.openById(CF_SPREADSHEET_ID);
    var sh = ss.getSheetByName('Msgs_Recebidas');
    if (!sh) {
      sh = ss.insertSheet('Msgs_Recebidas');
      sh.appendRow(['Timestamp', 'Telefone', 'Nome', 'Direcao', 'Texto']);
      sh.setFrozenRows(1);
    } else if (sh.getRange(1, 4).getValue() !== 'Direcao') {
      // formato antigo (4 col, só recebidas) — recria p/ medir demanda + tempo de resposta
      sh.clear();
      sh.appendRow(['Timestamp', 'Telefone', 'Nome', 'Direcao', 'Texto']);
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), body.phone || '', body.senderName || body.chatName || '',
                  body.fromMe ? 'ENVIADA' : 'RECEBIDA', txt]);
  } catch (err) {
    Logger.log('gravarMsg erro: ' + err.message);
  }
}

// ----------------------------------------------------------------
// Processa a resposta da paciente.
// ----------------------------------------------------------------
function processarResposta(phone, answer) {
  var pend = buscarPendente(phone);
  if (!pend) return; // sem confirmação pendente pra esse número -> ignora

  if (answer === 'SIM') {
    try {
      confirmarFeegow(pend.agId);
      marcarPendente(pend.row, 'CONFIRMADO');
      sendWhatsApp(phone, 'Perfeito! ✅ Seu agendamento está confirmado. Te esperamos! 💜');
      slackPost('✅ *Confirmado pelo WhatsApp* — ' + (pend.nome || phone) +
                ' (agendamento ' + pend.agId + '). Status mudado pra Confirmado no Feegow.');
    } catch (err) {
      marcarPendente(pend.row, 'ERRO_FEEGOW');
      sendWhatsApp(phone, 'Recebi sua confirmação! ✅ Já avisei a equipe. 💜');
      slackPost('⚠️ ' + (pend.nome || phone) + ' confirmou, mas FALHOU mudar no Feegow ' +
                '(agendamento ' + pend.agId + '): ' + err.message + '. Confirmar manualmente.');
    }
  } else { // NAO
    marcarPendente(pend.row, 'REAGENDAR');
    sendWhatsApp(phone, MSG_REAGENDAMENTO);
    slackPostReag('🔄 *Pediu pra reagendar* — ' + (pend.nome || phone) +
              ' (agendamento ' + pend.agId + '). Recepção: entrar em contato.');
  }
}

// ----------------------------------------------------------------
// Feegow — muda o status do agendamento para 7 (Marcado-confirmado).
// Params via query-string (compatível com a API PHP do Feegow).
// ----------------------------------------------------------------
function confirmarFeegow(agId) {
  var url = CF_FEEGOW_BASE + '/appoints/statusUpdate' +
            '?AgendamentoID=' + encodeURIComponent(agId) +
            '&StatusID=7' +
            '&Obs=' + encodeURIComponent('Confirmado pela paciente via WhatsApp');
  var resp = UrlFetchApp.fetch(url, {
    method:             'post',
    headers:            { 'x-access-token': CF_FEEGOW_TOKEN },
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  var txt  = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Feegow HTTP ' + code + ': ' + txt.substring(0, 200));
  }
  // Feegow às vezes responde 200 com success:false — trata como erro.
  try {
    var j = JSON.parse(txt);
    if (j && j.success === false) {
      throw new Error('Feegow success:false — ' + (j.message || txt.substring(0, 150)));
    }
  } catch (e) { /* resposta não-JSON: confia no HTTP 2xx */ }
}

// ================================================================
// CONFIRMAÇÃO POR LINK
// Os botões interativos do WhatsApp foram bloqueados pela Meta para a
// API não-oficial (Z-API) — passaram a não ser entregues. Em vez de botão,
// a paciente recebe um LINK único por agendamento. Clicou → muda status no
// Feegow (mesma confirmarFeegow do webhook). Token impede confirmar agenda alheia.
// ================================================================
function _linkSecret_() {
  var s = _P.getProperty('LINK_SECRET');
  if (!s) { s = Utilities.getUuid().replace(/-/g, ''); _P.setProperty('LINK_SECRET', s); }
  return s;
}
function _tokenConf_(agId) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(agId) + '|' + _linkSecret_());
  var hex = '';
  for (var i = 0; i < 6; i++) hex += ('0' + (bytes[i] & 0xFF).toString(16)).slice(-2);
  return hex; // 12 chars — inviável de adivinhar
}
function _webAppUrl_() {
  return _P.getProperty('WEBAPP_URL') || ScriptApp.getService().getUrl();
}
function linkConfirmacao(agId, tipo) { // tipo: 'c' confirmar | 'r' reagendar
  // Domínio da clínica (não parece golpe). A página confirmar.html encaminha pro web app.
  var base = _P.getProperty('PUBLIC_LINK_BASE') || 'https://app.paraser.com.br/confirmar.html';
  return base + '?a=' + tipo + '&ag=' + encodeURIComponent(agId) + '&t=' + _tokenConf_(agId);
}
function _msgConfirmacaoLink(agId) {
  return 'É só tocar na opção abaixo 💜\n\n' +
         '✅ *CONFIRMAR CONSULTA*\n' + linkConfirmacao(agId, 'c') + '\n\n' +
         '❌ *PRECISO REMARCAR*\n' + linkConfirmacao(agId, 'r');
}
function _acharPendentePorAg_(agId) {
  var sh = pendentesSheet_();
  if (sh.getLastRow() < 2) return null;
  var dados = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  var qualquer = null;
  for (var i = dados.length - 1; i >= 0; i--) {
    if (String(dados[i][3]) === String(agId)) {
      var p = { row: i + 2, agId: dados[i][3], nome: dados[i][2], tel: dados[i][1], status: String(dados[i][5]) };
      if (p.status === 'PENDENTE') return p;   // prioriza pendente (1ª resposta)
      if (!qualquer) qualquer = p;             // guarda o mais recente em qualquer status (p/ ter nome+telefone)
    }
  }
  return qualquer;
}
function _paginaConf_(emoji, titulo, msg) {
  var html =
    '<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1"><title>Paraser</title></head>' +
    '<body style="margin:0;font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;background:#f3f0f8">' +
    '<div style="max-width:420px;margin:9vh auto;background:#fff;border-radius:18px;padding:38px 26px;' +
    'text-align:center;box-shadow:0 10px 34px rgba(91,58,140,.14)">' +
    '<div style="font-size:56px;line-height:1">' + emoji + '</div>' +
    '<h1 style="color:#5b3a8c;font-size:22px;margin:16px 0 10px">' + titulo + '</h1>' +
    '<p style="color:#555;font-size:16px;line-height:1.5;margin:0">' + msg + '</p>' +
    '<p style="color:#b3a6c9;font-size:13px;margin-top:28px">Clínica Paraser 💜</p>' +
    '</div></body></html>';
  return HtmlService.createHtmlOutput(html)
    .addMetaTag('viewport', 'width=device-width,initial-scale=1');
}
// Retorna {emoji,titulo,msg}. _respConf decide se vira HTML (acesso direto) ou
// JSON (chamado pela página app.paraser.com.br via fetch, com &fmt=json).
function _respConf(res, fmt) {
  if (fmt === 'json') {
    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  }
  return _paginaConf_(res.emoji, res.titulo, res.msg);
}
function _confirmarViaLink(agId, token) {
  if (token !== _tokenConf_(agId)) {
    return { emoji:'⚠️', titulo:'Link inválido', msg:'Esse link de confirmação não é válido. Fale com a recepção pelo WhatsApp. 💜' };
  }
  if (String(agId) === '000000') { // agendamento de teste — não toca no Feegow
    return { emoji:'✅', titulo:'Presença confirmada!', msg:'Tudo certo, te esperamos! 💜' };
  }
  var pend = _acharPendentePorAg_(agId);
  if (pend && pend.status === 'CONFIRMADO') { // já confirmado antes — não repete nada
    return { emoji:'✅', titulo:'Presença confirmada!', msg:'Tudo certo, te esperamos! 💜' };
  }
  try {
    confirmarFeegow(agId);
    if (pend) marcarPendente(pend.row, 'CONFIRMADO');
    if (pend && pend.tel) {
      try { sendWhatsApp(pend.tel, 'Perfeito! ✅ Seu agendamento está confirmado. Te esperamos! 💜'); }
      catch (e) { slackPost('⚠️ Confirmou (ag ' + agId + ') mas falhou ao mandar WhatsApp de agradecimento: ' + e.message); }
    }
    slackPost('✅ *Confirmado via link* — ' + ((pend && pend.nome) || ('ag ' + agId)) +
              ' (agendamento ' + agId + '). Status → Confirmado no Feegow.');
    return { emoji:'✅', titulo:'Presença confirmada!', msg:'Tudo certo, te esperamos! 💜' };
  } catch (err) {
    if (pend) marcarPendente(pend.row, 'ERRO_FEEGOW');
    slackPost('⚠️ Clicou confirmar (ag ' + agId + ') mas FALHOU no Feegow: ' + err.message + '. Confirmar manual.');
    return { emoji:'✅', titulo:'Recebemos sua confirmação!', msg:'Já avisamos a equipe. 💜' };
  }
}
function _reagendarViaLink(agId, token) {
  if (token !== _tokenConf_(agId)) {
    return { emoji:'⚠️', titulo:'Link inválido', msg:'Esse link não é válido. Fale com a recepção pelo WhatsApp. 💜' };
  }
  if (String(agId) === '000000') { // agendamento de teste — não posta no Slack
    return { emoji:'🔄', titulo:'Pedido recebido', msg:'A recepção vai te chamar pra achar um novo horário. 💜' };
  }
  var pend = _acharPendentePorAg_(agId);
  var jaAvisado = pend && pend.status === 'REAGENDAR'; // já pediu antes → não duplica o Slack
  if (pend) marcarPendente(pend.row, 'REAGENDAR');
  // WhatsApp pra paciente: SEMPRE (mesmo se ela clicar de novo) — antes pulava no re-clique.
  if (pend && pend.tel) {
    try { sendWhatsApp(pend.tel, MSG_REAGENDAMENTO); }
    catch (e) { slackPostReag('⚠️ Pedido reagendar (ag ' + agId + ') mas falhou ao mandar WhatsApp de aviso: ' + e.message); }
  }
  // Slack: só na 1ª vez, no canal #reagendamento (não spammar a recepção a cada re-clique).
  if (!jaAvisado) {
    var waReag = (pend && pend.tel) ? formatPhone(pend.tel) : null;
    slackPostReag('🔄 *Pediu reagendar (link)* — ' + ((pend && pend.nome) || ('ag ' + agId)) +
              ' (agendamento ' + agId + ').' +
              (waReag ? '\n📲 <https://wa.me/' + waReag + '|Chamar a paciente no WhatsApp>'
                      : ' Recepção: contatar.'));
  }
  return { emoji:'🔄', titulo:'Pedido recebido', msg:'A recepção vai te chamar pra achar um novo horário. 💜' };
}

// ----------------------------------------------------------------
// Fila de pendentes — aba Confirmacoes_Pendentes
// Colunas: Timestamp | Telefone | Paciente | AgendamentoID | Data | Status
// ----------------------------------------------------------------
function pendentesSheet_() {
  var ss = SpreadsheetApp.openById(CF_SPREADSHEET_ID);
  var sh = ss.getSheetByName(CF_PENDENTES_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CF_PENDENTES_SHEET);
    sh.appendRow(['Timestamp', 'Telefone', 'Paciente', 'AgendamentoID', 'Data', 'Status']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function registrarPendente(phone, nome, agId, dataStr) {
  try {
    pendentesSheet_().appendRow([new Date(), phone, nome, agId, dataStr, 'PENDENTE']);
  } catch (err) {
    Logger.log('registrarPendente erro: ' + err.message);
  }
}

// Acha a confirmação PENDENTE mais recente desse telefone (compara só os dígitos).
function buscarPendente(phone) {
  var sh = pendentesSheet_();
  if (sh.getLastRow() < 2) return null;
  var dados = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  var alvo = soDigitos(phone);
  for (var i = dados.length - 1; i >= 0; i--) { // de baixo pra cima = mais recente
    var tel = soDigitos(String(dados[i][1]));
    var st  = String(dados[i][5]);
    if (st === 'PENDENTE' && telBate(tel, alvo)) {
      return { row: i + 2, agId: dados[i][3], nome: dados[i][2] };
    }
  }
  return null;
}

function marcarPendente(row, novoStatus) {
  try {
    pendentesSheet_().getRange(row, 6).setValue(novoStatus);
  } catch (err) {
    Logger.log('marcarPendente erro: ' + err.message);
  }
}

// ----------------------------------------------------------------
// Slack — post simples no canal de atendimento
// ----------------------------------------------------------------
function slackPost(texto, canal) {
  try {
    if (!CF_SLACK_TOKEN) return false;
    var channelId = slackGetChannelId(canal || CF_SLACK_CHANNEL);
    if (!channelId) return false;
    var resp = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method:             'post',
      contentType:        'application/json; charset=utf-8',
      headers:            { Authorization: 'Bearer ' + CF_SLACK_TOKEN },
      payload:            JSON.stringify({ channel: channelId, text: texto }),
      muteHttpExceptions: true
    });
    try { return JSON.parse(resp.getContentText()).ok === true; } catch (e) { return true; }
  } catch (err) {
    Logger.log('slackPost erro: ' + err.message);
    return false;
  }
}

// Reagendamento → canal #reagendamento. Se ele não existir / o app não estiver
// no canal, o post falha; aí cai no #atendimento com um aviso (nunca perde).
function slackPostReag(texto) {
  if (!slackPost(texto, CF_SLACK_CHANNEL_REAG)) {
    slackPost('⚠️ (não consegui postar no #' + CF_SLACK_CHANNEL_REAG +
              ' — crie o canal e adicione o app da clínica)\n' + texto, CF_SLACK_CHANNEL);
  }
}

// Post no #comercial (mesmo bot). Fallback pro #atendimento se o app não estiver no canal.
function slackPostComercial(texto) {
  if (!slackPost(texto, CF_SLACK_CHANNEL_COMERCIAL)) {
    slackPost('⚠️ (não consegui postar no #' + CF_SLACK_CHANNEL_COMERCIAL +
              ' — adicione o app da clínica ao canal)\n' + texto, CF_SLACK_CHANNEL);
  }
}

// ================================================================
// PRIMEIRA USG — avisa o #comercial no dia em que a paciente faz a
// PRIMEIRA USG dela na clínica (marco de início de tratamento).
// Roda 1x/dia de manhã (trigger notificarPrimeiraUSG). Uma vez por
// paciente na vida (aba de controle Primeira_USG_Avisadas).
// ================================================================
var CF_SLACK_CHANNEL_COMERCIAL = 'comercial';
// Base de pacientes que JÁ fizeram USG (até ontem). A API do Feegow não deixa
// buscar histórico por paciente_id (dá 409), então mantemos a base aqui: um
// backfill inicial varre os meses passados; o diário compara e vai crescendo.
var CF_USG_HIST_SHEET = 'USG_Pacientes_Historico';
// procIds que contam como "fazer USG" (execução de imagem). Exclui PACOTE (cobrança),
// punção/TEC (não é USG). Fonte: /procedures/list (jul/2026).
var USG_PROC_IDS = [4,5,6,7,8,9, 11,12,13,14,15,16, 17,18,19,20,21,22,23,
  25,26,27,28,29,30,31, 40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,
  57,58,59,60,61,62,64,65,66,67, 73,74,75, 100,122,153,167,171,174,
  244,245,246,263, 402,403,404,405,406,407,408,409,410,411,412,413,414,
  50230,50231,50232,50233,50234,50235,50243];
var STATUS_CANCELADO   = [11, 15, 22];       // no histórico, USG cancelada/remarcada não conta
var STATUS_FORA_HOJE   = [6, 11, 15, 22];    // hoje, ignora não-compareceu + cancelados

function _ehUSG_(procId) { return USG_PROC_IDS.indexOf(Number(procId)) >= 0; }
function _dataFeegowISO_(d) { var p = String(d || '').split('-'); return p.length === 3 ? p[2] + '-' + p[1] + '-' + p[0] : String(d); }
// USG de MEIO de ciclo (2º..6º, TEC/ERA 2-3, 2/3 pós beta): implica exame anterior,
// então nunca é "primeira USG" (protege contra dado da base incompleto/errado).
function _procMeioDeCiclo_(nome) {
  var n = String(nome || '').toUpperCase();
  return /\b[2-6]\s*[º°]/.test(n) || /\b[23]\s*P[ÓO]S\s*BETA/.test(n) || /(TEC|ERA)\s*[23]\b/.test(n);
}

// Busca appoints/search por período (dd-mm-yyyy) e opcionalmente paciente.
// Devolve array (pode ser vazio) ou null se falhou de vez (não confundir com vazio).
function _appointsBuscar_(dsFeegow, deFeegow, pacienteId) {
  var url = CF_FEEGOW_BASE + '/appoints/search?data_start=' + dsFeegow + '&data_end=' + deFeegow +
            '&per_page=1000' + (pacienteId ? '&paciente_id=' + pacienteId : '');
  for (var t = 0; t < 3; t++) {
    try {
      var r = UrlFetchApp.fetch(url, { headers: { 'x-access-token': CF_FEEGOW_TOKEN }, muteHttpExceptions: true });
      if (r.getResponseCode() === 200) {
        var j = JSON.parse(r.getContentText());
        return (Array.isArray(j.content) ? j.content : []).filter(function (a) { return a && typeof a === 'object'; });
      }
    } catch (e) {}
    Utilities.sleep(1500);
  }
  return null;
}

// --- Base histórica de quem já fez USG (aba, um paciente_id por linha) ---
function _usgHistSheet_() {
  var ss = SpreadsheetApp.openById(CF_SPREADSHEET_ID);
  var sh = ss.getSheetByName(CF_USG_HIST_SHEET);
  if (!sh) { sh = ss.insertSheet(CF_USG_HIST_SHEET); sh.appendRow(['PacienteID']); sh.setFrozenRows(1); }
  return sh;
}
function _usgHistSet_() {
  var sh = _usgHistSheet_(), set = {};
  if (sh.getLastRow() < 2) return set;
  sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { if (r[0] !== '') set[String(r[0])] = true; });
  return set;
}
function _usgHistAdd_(sh, ids) {
  if (!ids.length) return;
  sh.getRange(sh.getLastRow() + 1, 1, ids.length, 1).setValues(ids.map(function (x) { return [x]; }));
}

// Pacientes DISTINTOS com USG ativa hoje → { paciente_id: agendamento }
function _pacientesUsgHoje_(hojeFeegow) {
  var hoje = _appointsBuscar_(hojeFeegow, hojeFeegow);
  if (hoje === null) return null;
  var pac = {};
  hoje.forEach(function (a) {
    if (!_ehUSG_(a.procedimento_id)) return;
    if (STATUS_FORA_HOJE.indexOf(Number(a.status_id)) >= 0) return;
    if (_procMeioDeCiclo_(resolveNomeProc(a))) return;   // 2º+ implica exame anterior → não é a 1ª
    if (!pac[a.paciente_id]) pac[a.paciente_id] = a;
  });
  return pac;
}

// Backfill: varre os últimos N meses (sem paciente_id) e registra na base todos os
// pacientes que já fizeram USG ATÉ ONTEM. Roda uma vez no setup. Idempotente.
function backfillHistoricoUSG(meses) {
  meses = meses || 36;
  var tz = 'America/Sao_Paulo';
  var ontem = new Date(); ontem.setDate(ontem.getDate() - 1);
  var ontemISO = Utilities.formatDate(ontem, tz, 'yyyy-MM-dd');
  var sh = _usgHistSheet_(), jaTem = _usgHistSet_(), novos = {};
  var d = new Date();
  for (var i = 0; i < meses; i++) {
    var ano = d.getFullYear(), mes = d.getMonth() + 1, ult = new Date(ano, mes, 0).getDate();
    var ags = _appointsBuscar_('01-' + pad2(mes) + '-' + ano, pad2(ult) + '-' + pad2(mes) + '-' + ano);
    if (ags) ags.forEach(function (a) {
      if (_ehUSG_(a.procedimento_id) && STATUS_CANCELADO.indexOf(Number(a.status_id)) < 0 &&
          _dataFeegowISO_(a.data) <= ontemISO && !jaTem[a.paciente_id]) novos[a.paciente_id] = true;
    });
    d.setMonth(d.getMonth() - 1);
    Utilities.sleep(400);
  }
  var ids = Object.keys(novos);
  _usgHistAdd_(sh, ids);
  Logger.log('backfillHistoricoUSG: +' + ids.length + ' pacientes na base (' + meses + ' meses).');
  return { adicionados: ids.length, totalBase: Object.keys(jaTem).length + ids.length };
}

// Diário: avisa o #comercial quem faz a PRIMEIRA USG hoje (não está na base) e
// adiciona esses pacientes à base (pra não reavaliar amanhã).
function notificarPrimeiraUSG() {
  if (!CF_FEEGOW_TOKEN) return;
  var tz = 'America/Sao_Paulo';
  var hojeFeegow = Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy');
  var pac = _pacientesUsgHoje_(hojeFeegow);
  if (pac === null) { Logger.log('notificarPrimeiraUSG: Feegow não respondeu.'); return; }
  var base = _usgHistSet_(), sh = _usgHistSheet_(), novos = [];
  Object.keys(pac).forEach(function (pid) {
    if (base[pid]) return;                          // já fez USG antes → não é a primeira
    var ag = pac[pid], p = getPatientData(pid);
    var proc = resolveNomeProc(ag) || ('procId ' + ag.procedimento_id);
    var prof = carregarProfissionais()[ag.profissional_id] || '';
    slackPostComercial('🔬 *Primeira USG* — ' + (p.nome || ('paciente ' + pid)) +
      ' faz a *primeira USG* dela hoje (' + proc + (prof ? ' · ' + prof : '') +
      ' às ' + formatHora(ag.horario || '') + ').\nMarco de início de tratamento — vale acompanhar. 💜');
    novos.push(pid);                                // primeira USG → entra na base
  });
  _usgHistAdd_(sh, novos);
  Logger.log('notificarPrimeiraUSG: ' + novos.length + ' primeira(s) USG de ' + Object.keys(pac).length + ' com USG hoje.');
}

// DRY-RUN: quem receberia HOJE, sem postar nem gravar. Rápido (1 chamada + base).
function _simularPrimeiraUsg_(limit) {
  var tz = 'America/Sao_Paulo';
  var hojeISO = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var hojeFeegow = Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy');
  var pac = _pacientesUsgHoje_(hojeFeegow);
  if (pac === null) return { erro: 'Feegow não respondeu' };
  var base = _usgHistSet_(), ids = Object.keys(pac), out = [];
  ids.slice(0, limit || 30).forEach(function (pid) {
    var p = getPatientData(pid);
    out.push({ paciente: p.nome || pid, proc: resolveNomeProc(pac[pid]), veredito: base[pid] ? 'ja fez antes' : 'PRIMEIRA USG' });
  });
  return { data: hojeISO, totalPacientesUsgHoje: ids.length, baseTamanho: Object.keys(base).length, amostra: out };
}

// ----------------------------------------------------------------
// Helpers
// ----------------------------------------------------------------
function soDigitos(s) { return String(s).replace(/\D/g, ''); }

// Telefones batem se os últimos 8 dígitos coincidem (ignora 55/DDD/9 extra).
function telBate(a, b) {
  if (!a || !b) return false;
  var n = 8;
  return a.slice(-n) === b.slice(-n);
}

function _okText() {
  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}

// ----------------------------------------------------------------
// SETUP — rode UMA VEZ pra apontar o webhook "ao receber" do Z-API
// pra este Web App. Depois disso, as respostas Sim/Não chegam aqui.
// Se você reimplantar o Web App (nova URL), atualize WEBHOOK_URL e rode de novo.
// ----------------------------------------------------------------
function configurarWebhookZapi() {
  var WEBHOOK_URL = 'https://script.google.com/macros/s/AKfycbwCDFfEx0JEn0m9vwbaLEt6jqoYXR_Mjy4Z9Gu6i6srkozNQG_sJCbR9fON5aybvi-5dw/exec';

  var url = 'https://api.z-api.io/instances/' + CF_ZAPI_INSTANCE_ID +
            '/token/' + CF_ZAPI_TOKEN + '/update-webhook-received';
  var headers = {};
  if (CF_ZAPI_CLIENT_TOKEN) headers['Client-Token'] = CF_ZAPI_CLIENT_TOKEN;

  var resp = UrlFetchApp.fetch(url, {
    method:             'put',
    contentType:        'application/json',
    headers:            headers,
    payload:            JSON.stringify({ value: WEBHOOK_URL }),
    muteHttpExceptions: true
  });
  Logger.log('Z-API update-webhook-received: HTTP ' + resp.getResponseCode() +
             ' — ' + resp.getContentText().substring(0, 300));
}

// ================================================================
// AGENDA DO MÉDICO — envia agenda do dia seguinte pra cada médico
// configurado na aba "Medicos_Agenda" da planilha. Trigger Seg-Sex 18h.
// ================================================================

var CF_MEDICOS_SHEET = 'Medicos_Agenda';

// Roda 1x pra criar a aba e popular com a lista inicial.
function setupListaMedicos() {
  var ss = SpreadsheetApp.openById(CF_SPREADSHEET_ID);
  var sh = ss.getSheetByName(CF_MEDICOS_SHEET);
  if (!sh) sh = ss.insertSheet(CF_MEDICOS_SHEET);
  sh.clear();
  sh.appendRow(['Nome', 'Telefone', 'Ativo']);
  sh.setFrozenRows(1);

  var lista = [
    ['Dra. Bianca Salvato',    '5521966723841', 'SIM'],
    ['Dra. Bruna Ortiz',       '5521967360998', 'SIM'],
    ['Cristiane Chagas',       '5521983569930', 'SIM'],
    ['Dra. Érica Stein',       '5521997876716', 'SIM'],
    ['Graziela Siqueira',      '5521994550819', 'SIM'],
    ['Dr. Helce Ribeiro',      '5521992436355', 'SIM'],
    ['Dr. Joselmo Salvato',    '5521999899118', 'SIM'],
    ['Dra. Kátia Chamorro',    '5521988277790', 'SIM'],
    ['Dra. Mabel Iglesias',    '5521997424582', 'SIM'],
    ['Dra. Magali Miranda',    '5532990210011', 'SIM'],
    ['Dra. Marcelle Moura',    '5521997399888', 'SIM'],
    ['Dr. Mario Barroso',      '5521983293133', 'SIM'],
    ['Dra. Priscila Loyola',   '5521991510502', 'SIM'],
    ['Dr. Rodolfo Salvato',    '5521993612289', 'SIM'],
    ['Sara Salvato',           '5521996777207', 'SIM'],
    ['Dra. Verônica Pintor',   '5521965206906', 'SIM']
  ];
  sh.getRange(2, 1, lista.length, 3).setValues(lista);
  sh.setColumnWidths(1, 3, 220);
  Logger.log('✅ Aba "' + CF_MEDICOS_SHEET + '" criada com ' + lista.length + ' médicos. Edite ativo/telefone direto na planilha.');
}

function lerListaMedicos_() {
  var ss = SpreadsheetApp.openById(CF_SPREADSHEET_ID);
  var sh = ss.getSheetByName(CF_MEDICOS_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  return vals
    .map(function(r) { return { nome: String(r[0] || '').trim(), telefone: String(r[1] || '').replace(/\D/g, ''), ativo: String(r[2] || '').trim().toUpperCase() }; })
    .filter(function(m) { return m.nome && m.telefone && m.ativo === 'SIM'; });
}

// Match: confronta um agendamento com a lista, retorna o telefone do médico
// pelo primeiro nome único (Bianca, Bruna, Rodolfo, etc). Case-insensitive,
// sem acentos.
function _normalizarNome(s) {
  return String(s || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '') // remove acentos
    .toUpperCase()
    .replace(/\bDRA?\.?\s*/g, '') // remove "Dr." / "Dra." + ponto + espaço
    .replace(/[^A-Z0-9\s]/g, ' ') // qualquer não-alfanumérico vira espaço
    .replace(/\s+/g, ' ')
    .trim();
}

function _matchMedico(profNomeFeegow, listaMedicos) {
  if (!profNomeFeegow) return null;
  var alvo = _normalizarNome(profNomeFeegow);
  if (!alvo) return null;
  // tenta match por primeiro nome
  var primeiroAlvo = alvo.split(' ')[0];
  for (var i = 0; i < listaMedicos.length; i++) {
    var nomeListado = _normalizarNome(listaMedicos[i].nome);
    var primeiroListado = nomeListado.split(' ')[0];
    if (primeiroAlvo && primeiroAlvo === primeiroListado) return listaMedicos[i];
  }
  return null;
}

// Formata um horário "HH:MM" a partir do campo Hora do Feegow ("14:30" ou "14:30:00")
function _formatHora(h) {
  var s = String(h || '');
  var m = s.match(/^(\d{1,2}):(\d{2})/);
  if (!m) return s;
  return (m[1].length === 1 ? '0' + m[1] : m[1]) + 'h' + m[2];
}

function _statusEmoji(statusId) {
  // 7 = Marcado-Confirmado (já confirmou pelo WhatsApp ou manual)
  // 1 = Marcado (não confirmado ainda)
  // 3 = Atendido
  if (Number(statusId) === 7 || Number(statusId) === 3) return '✅';
  return '⏳';
}

function _statusTexto(statusId) {
  if (Number(statusId) === 7) return 'Confirmado';
  if (Number(statusId) === 3) return 'Atendido';
  return 'Não confirmado';
}

function _diaSemanaPorExt(d) {
  var dias = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'];
  return dias[d.getDay()];
}

function _titleCaseNome(s) {
  var minus = ['da', 'de', 'do', 'das', 'dos', 'e', 'di', 'du'];
  return String(s || '').toLowerCase().trim().split(/\s+/).map(function(w, i) {
    if (i > 0 && minus.indexOf(w) >= 0) return w;
    if (!w.length) return w;
    return w.charAt(0).toUpperCase() + w.slice(1);
  }).join(' ');
}

// Preenche ag.paciente_nome em todos os agendamentos (faz 1 call por paciente único).
function _carregarNomesPacientes(agendamentos) {
  var cache = {};
  agendamentos.forEach(function(ag) {
    if (!ag.paciente_id) return;
    if (cache[ag.paciente_id] === undefined) {
      try {
        var p = getPatientData(ag.paciente_id);
        cache[ag.paciente_id] = (p && p.nome) || '';
      } catch (e) {
        cache[ag.paciente_id] = '';
      }
    }
    ag.paciente_nome = cache[ag.paciente_id];
  });
}

function _tituloMedico(nomeCompleto) {
  // "Dra. Bianca Salvato" → "Dra. Bianca"
  // "Sara Salvato" → "Sara"
  // "Cristiane Chagas" → "Cristiane"
  var m = String(nomeCompleto || '').match(/^(Dra?\.\s*)(\S+)/i);
  if (m) return m[1] + m[2];
  return String(nomeCompleto || '').split(' ')[0];
}

function _formatarMensagemAgenda(nomeMedico, dataAlvo, agendamentos) {
  var titulo = _tituloMedico(nomeMedico);
  var dataStr = Utilities.formatDate(dataAlvo, 'GMT-3', 'dd/MM');
  var diaSem  = _diaSemanaPorExt(dataAlvo);

  var saudacao = 'Oi, ' + titulo + '! 💜\n\n';
  if (!agendamentos.length) {
    return saudacao + 'Você não tem agendamentos pra amanhã (' + diaSem + ', ' + dataStr + ').\n\nBom descanso!\n\nEquipe Paraser';
  }

  agendamentos.sort(function(a, b) { return String(a.hora || '').localeCompare(String(b.hora || '')); });

  // Detecta gaps maiores que 60min entre consultas e marca como "livre"
  var linhas = [];
  var prevFimMin = null;
  agendamentos.forEach(function(ag) {
    var horaIni = _formatHora(ag.hora);
    var match = String(ag.hora || '').match(/^(\d{1,2}):(\d{2})/);
    if (match) {
      var minIni = (+match[1]) * 60 + (+match[2]);
      if (prevFimMin !== null && minIni - prevFimMin >= 60) {
        var gapH = Math.floor(prevFimMin / 60), gapM = prevFimMin % 60;
        var gapStr = (gapH < 10 ? '0' + gapH : gapH) + 'h' + (gapM < 10 ? '0' + gapM : gapM);
        linhas.push('🕓 ' + gapStr + ' — _livre_');
      }
      var dur = Number(ag.tempo) || Number(ag.duracao) || 30;
      prevFimMin = minIni + dur;
    }
    var emoji = _statusEmoji(ag.status_id);
    var nomePac = _titleCaseNome(ag.paciente_nome || 'paciente');
    var procName = _titleCaseNome((ag._procNome || 'procedimento').replace(/\s*-\s*/g, ' - '));
    var statusTxt = _statusTexto(ag.status_id);
    linhas.push(
      emoji + ' *' + horaIni + '* · *' + nomePac + '*\n' +
      '     ' + procName + ' · _' + statusTxt + '_'
    );
  });

  var total = agendamentos.length;
  var confirmados = agendamentos.filter(function(a) { return Number(a.status_id) === 7 || Number(a.status_id) === 3; }).length;
  var rodape = '\n\nTotal: *' + total + ' paciente' + (total > 1 ? 's' : '') + '*';
  if (confirmados) rodape += ' · ' + confirmados + ' confirmado' + (confirmados > 1 ? 's' : '') + ' ✅';
  rodape += '.';

  return saudacao +
         'Sua agenda de amanhã (' + diaSem + ', ' + dataStr + '):\n\n' +
         linhas.join('\n\n') +
         rodape +
         '\n\nEquipe Paraser';
}

// Bloco de pendências "Em atendimento" anexado na agenda do médico.
// Só é chamado quando o médico tem pelo menos um preso.
function _blocoPresosMedico(arr, procMap) {
  arr.sort(function(x, y) { return cfDataKey(x.data).localeCompare(cfDataKey(y.data)); });
  var n = arr.length;
  var linhas = ['🔴 *Pendente no Feegow:* você tem *' + n + '* exame' + (n === 1 ? '' : 's') +
                ' ainda em "Em atendimento". Quando puder, marque como "Atendido" (senão somem dos relatórios):'];
  arr.slice(0, 8).forEach(function(a) {
    var proc = procMap[a.procedimento_id] || '';
    var pac  = '';
    try { pac = getPatientData(a.paciente_id).nome || ''; } catch (e) {}
    if (!pac) pac = '#' + (a.paciente_id || '?');
    linhas.push('• ' + cfDataBRFeegow(a.data) + ' ' + _titleCaseNome(pac) + (proc ? ' · ' + _titleCaseNome(proc) : ''));
  });
  if (n > 8) linhas.push('_+' + (n - 8) + ' mais_');
  return linhas.join('\n');
}

// FUNÇÃO PRINCIPAL — chamada pelo trigger Seg-Sex 18h
function enviarAgendaMedicos() {
  var hoje = new Date();
  var dow = hoje.getDay(); // 0=dom, 6=sáb
  if (dow === 0 || dow === 6) {
    Logger.log('Fim de semana — não envia agenda hoje.');
    return;
  }

  // Z-API conectado?
  if (typeof zapiConectado === 'function' && !zapiConectado()) {
    Logger.log('🚨 Z-API desconectado — agendas NÃO enviadas.');
    if (typeof slackPost === 'function') slackPost('🚨 *Agenda médicos* — Z-API desconectado, agendas de amanhã NÃO foram enviadas.');
    return;
  }

  // Trava anti-duplicação: se já enviou hoje (ex.: disparo manual + trigger no mesmo dia), não reenvia.
  var _propsAg = PropertiesService.getScriptProperties();
  var _hojeStr = Utilities.formatDate(hoje, 'America/Sao_Paulo', 'yyyy-MM-dd');
  if (_propsAg.getProperty('AGENDA_MED_ENVIADA_EM') === _hojeStr) {
    Logger.log('Agenda médicos já enviada hoje (' + _hojeStr + ') — pulando pra não duplicar.');
    return;
  }

  // Dia alvo = amanhã (se hoje é sexta, alvo é segunda)
  var alvo = new Date(hoje);
  alvo.setDate(alvo.getDate() + 1);
  while (alvo.getDay() === 0 || alvo.getDay() === 6) alvo.setDate(alvo.getDate() + 1);

  // Pega agenda + carrega nomes
  var agendamentos;
  try {
    agendamentos = getAgendamentos(alvo);
  } catch(e) {
    // Feegow fora do ar: não dá pra montar a agenda. Avisa no Slack e aborta.
    Logger.log('🚨 Feegow indisponível — agenda dos médicos NÃO enviada: ' + e.message);
    if (typeof slackPost === 'function') slackPost('🚨 *Agenda médicos* — Feegow fora do ar (' + e.message + '), agendas de amanhã NÃO foram enviadas.');
    return;
  }
  var profMap = carregarProfissionais();
  var procMap = (typeof carregarNomesProcedimentos === 'function') ? carregarNomesProcedimentos() : {};
  agendamentos.forEach(function(ag) {
    ag._profNome = profMap[ag.profissional_id] || '';
    if (!ag._procNome && procMap[ag.procedimento_id]) ag._procNome = procMap[ag.procedimento_id];
    if (!ag._procNome) ag._procNome = (typeof resolveNomeProc === 'function') ? resolveNomeProc(ag) : '';
    if (!ag.hora && ag.horario) ag.hora = ag.horario;
  });

  var medicos = lerListaMedicos_();
  if (!medicos.length) {
    Logger.log('Sem médicos ativos na aba ' + CF_MEDICOS_SHEET);
    return;
  }

  // Filtra só agendamentos de médicos da lista pra evitar chamar getPatientData de pacientes que ninguém vai ver
  var idsMedicosLista = {};
  medicos.forEach(function(med) {
    agendamentos.forEach(function(ag) {
      if (_matchMedico(ag._profNome, [med])) idsMedicosLista[ag.agendamento_id] = true;
    });
  });
  var relevantes = agendamentos.filter(function(ag) { return idsMedicosLista[ag.agendamento_id]; });
  _carregarNomesPacientes(relevantes);

  // Pendências: agendamentos presos em "Em atendimento" (do 1º do mês passado até
  // ontem). Anexadas na msg de cada médico (só quem tem pendência vê). Se falhar,
  // segue sem — a agenda é o que importa, a pendência é bônus.
  var presosEmAtend = [];
  var presosOk = false;
  try {
    var _hj    = new Date();
    var _ontemP = new Date(_hj.getTime() - 24 * 60 * 60 * 1000);
    var _iniP  = new Date(_hj.getFullYear(), _hj.getMonth() - 1, 1);
    presosEmAtend = cfBuscarEmAtendimento(_iniP, _ontemP);
    presosOk = true;
  } catch (e) {
    Logger.log('Pendências "Em atendimento" não carregadas (agenda segue normal): ' + e.message);
  }

  var enviados = 0, falhas = 0, semAgenda = 0;
  medicos.forEach(function(med) {
    try {
      var agDoMedico = agendamentos.filter(function(ag) { return _matchMedico(ag._profNome, [med]); });

      // Sem grade no dia seguinte -> NÃO envia (só registra). Médico sem
      // agendamento amanhã não precisa receber a mensagem.
      if (!agDoMedico.length) {
        semAgenda++;
        logRowAgenda_(med, alvo, 0, 'SEM_AGENDA (nao enviado)');
        return;
      }

      var msg = _formatarMensagemAgenda(med.nome, alvo, agDoMedico);
      var presosDoMed = presosEmAtend.filter(function(a) {
        return _matchMedico(profMap[a.profissional_id] || '', [med]);
      });
      if (presosDoMed.length) {
        msg = msg.replace('\n\nEquipe Paraser', '\n\n' + _blocoPresosMedico(presosDoMed, procMap) + '\n\nEquipe Paraser');
      }
      sendWhatsApp(med.telefone, msg);
      Utilities.sleep(1200);
      enviados++;
      logRowAgenda_(med, alvo, agDoMedico.length, 'ENVIADO');
    } catch (err) {
      falhas++;
      Logger.log('Falha enviando agenda pra ' + med.nome + ': ' + err.message);
      logRowAgenda_(med, alvo, 0, 'ERRO: ' + err.message);
    }
  });

  var resumo = '📅 *Agenda médicos enviada* — ' +
               Utilities.formatDate(alvo, 'GMT-3', 'dd/MM') +
               ' · enviadas: ' + enviados + ' / falhas: ' + falhas + ' / sem agenda: ' + semAgenda;
  Logger.log(resumo);
  if (typeof slackPost === 'function') slackPost(resumo);

  // Toda segunda, posta no Slack o panorama dos presos em "Em atendimento" (mesma
  // info que foi pra cada médico). Aproveita este trigger, sem gatilho separado.
  if (dow === 1 && presosOk && typeof slackPost === 'function') {
    try { slackPost(_resumoPresosSlack(presosEmAtend, profMap, procMap)); }
    catch (e) { Logger.log('Resumo presos (Slack) falhou: ' + e.message); }
  }

  _propsAg.setProperty('AGENDA_MED_ENVIADA_EM', _hojeStr);
}

function logRowAgenda_(med, dataAlvo, totalAg, status) {
  var ss = SpreadsheetApp.openById(CF_SPREADSHEET_ID);
  var sh = ss.getSheetByName('Agenda_Medicos_Log');
  if (!sh) {
    sh = ss.insertSheet('Agenda_Medicos_Log');
    sh.appendRow(['Timestamp', 'Medico', 'Telefone', 'Data Alvo', 'Total Agend', 'Status']);
    sh.setFrozenRows(1);
  }
  sh.appendRow([new Date(), med.nome, med.telefone, Utilities.formatDate(dataAlvo, 'GMT-3', 'dd/MM/yyyy'), totalAg, status]);
}

function _debugMatchAgenda(diaForcado) {
  var alvo;
  if (diaForcado) {
    var p = diaForcado.split('/');
    alvo = new Date(+p[2], +p[1] - 1, +p[0]);
  } else {
    alvo = new Date();
    alvo.setDate(alvo.getDate() + 1);
  }

  var agendamentos = getAgendamentos(alvo);
  var profMap = carregarProfissionais();
  var amostras = agendamentos.slice(0, 10).map(function(ag) {
    var nomeFromMap = profMap[ag.profissional_id] || '';
    return {
      paciente: ag.paciente_nome,
      hora: ag.hora || ag.horario,
      profissional_id: ag.profissional_id,
      _profNome_do_map: nomeFromMap,
      campos_disponiveis: Object.keys(ag).filter(function(k){ return !k.startsWith('_'); }).slice(0, 25)
    };
  });

  var medicos = lerListaMedicos_();
  agendamentos.forEach(function(ag) { ag._profNome = profMap[ag.profissional_id] || ''; });

  var tentativas = medicos.map(function(med) {
    var matchados = agendamentos.filter(function(ag) { return _matchMedico(ag._profNome, [med]); });
    return {
      medico_lista: med.nome,
      primeiro_normalizado: _normalizarNome(med.nome).split(' ')[0],
      qtd_matchados: matchados.length
    };
  });

  return {
    dia: Utilities.formatDate(alvo, 'GMT-3', 'dd/MM/yyyy'),
    total_agendamentos: agendamentos.length,
    prof_map_size: Object.keys(profMap).length,
    amostras_brutas: amostras,
    tentativas_match: tentativas
  };
}

// PREVIEW — envia pro número do Felipe (ou outro) TODAS as mensagens
// que sairiam pra cada médico, marcadas como "[PREVIEW Dr. X]" pra ele aprovar.
// Não toca nos médicos reais.
function _previewAgendaMedicos(paraNumero, diaForcado) {
  if (!paraNumero) return { erro: 'passe ?para=<telefone com DDI>' };

  var alvo;
  if (diaForcado) {
    var p = diaForcado.split('/');
    if (p.length === 3) alvo = new Date(+p[2], +p[1] - 1, +p[0]);
    else return { erro: 'dia inválido. Formato: dd/MM/yyyy' };
  } else {
    alvo = new Date();
    alvo.setDate(alvo.getDate() + 1);
    while (alvo.getDay() === 0 || alvo.getDay() === 6) alvo.setDate(alvo.getDate() + 1);
  }

  var agendamentos = getAgendamentos(alvo);
  var profMap = carregarProfissionais();
  var procMap = (typeof carregarNomesProcedimentos === 'function') ? carregarNomesProcedimentos() : {};
  agendamentos.forEach(function(ag) {
    ag._profNome = profMap[ag.profissional_id] || '';
    if (!ag._procNome && procMap[ag.procedimento_id]) ag._procNome = procMap[ag.procedimento_id];
    if (!ag._procNome) ag._procNome = (typeof resolveNomeProc === 'function') ? resolveNomeProc(ag) : '';
    if (!ag.hora && ag.horario) ag.hora = ag.horario;
  });

  var medicos = lerListaMedicos_();

  // Carrega nomes só dos pacientes que vão aparecer
  var idsLista = {};
  medicos.forEach(function(med) {
    agendamentos.forEach(function(ag) {
      if (_matchMedico(ag._profNome, [med])) idsLista[ag.agendamento_id] = true;
    });
  });
  var relevantes = agendamentos.filter(function(ag) { return idsLista[ag.agendamento_id]; });
  _carregarNomesPacientes(relevantes);

  var enviados = 0, semAgenda = 0, resumo = [];

  medicos.forEach(function(med) {
    var agDoMedico = agendamentos.filter(function(ag) { return _matchMedico(ag._profNome, [med]); });
    // Espelha a produção: quem não tem grade no dia seguinte não recebe.
    if (!agDoMedico.length) {
      semAgenda++;
      resumo.push({ medico: med.nome, agendamentos: 0, enviado: false });
      return;
    }
    var msg = '[PREVIEW » ' + med.nome + ']\n\n' + _formatarMensagemAgenda(med.nome, alvo, agDoMedico);
    sendWhatsApp(paraNumero, msg);
    Utilities.sleep(1500);
    enviados++;
    resumo.push({ medico: med.nome, agendamentos: agDoMedico.length, enviado: true });
  });

  return {
    ok: true,
    para: paraNumero,
    dia_alvo: Utilities.formatDate(alvo, 'GMT-3', 'dd/MM/yyyy'),
    total_medicos: medicos.length,
    sem_agenda: semAgenda,
    enviados: enviados,
    resumo: resumo
  };
}

// Setup trigger Seg-Sex 18h (chamada manual 1x)
// Wrappers dos triggers: SEXTA envia às 17h (agenda de segunda); SEG–QUI às 18h.
// Cada dia só UM wrapper dispara, então não há envio em duplicidade.
function enviarAgendaMedicosSexta() {
  if (new Date().getDay() === 5) enviarAgendaMedicos();
}
function enviarAgendaMedicosSemana() {
  var d = new Date().getDay();
  if (d >= 1 && d <= 4) enviarAgendaMedicos();
}

function setupTriggerAgendaMedicos() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'enviarAgendaMedicos' || fn === 'enviarAgendaMedicosSexta' || fn === 'enviarAgendaMedicosSemana') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Sexta-feira: 17h
  ScriptApp.newTrigger('enviarAgendaMedicosSexta')
    .timeBased().everyDays(1).atHour(17).inTimezone('America/Sao_Paulo').create();
  // Segunda a quinta: 18h
  ScriptApp.newTrigger('enviarAgendaMedicosSemana')
    .timeBased().everyDays(1).atHour(18).inTimezone('America/Sao_Paulo').create();
  Logger.log('✅ Triggers: sexta 17h + seg-qui 18h (cada wrapper checa o dia; trava anti-duplo por dia)');
}

// ================================================================
// ALERTA — agendamentos presos em "Em atendimento" no Feegow
// ----------------------------------------------------------------
// Só o MÉDICO consegue marcar "Atendido". Quando ele esquece de
// fechar, o agendamento fica preso em "Em atendimento" e some dos
// relatórios que filtram só "Atendido" (Médicos, produção). O
// repasse já conta "Em atendimento" (não depende disto), mas as
// outras telas não. Este alerta lista os presos por médico e posta
// no Slack toda segunda 8h, pra alguém pedir ao médico pra fechar.
// Lê o Feegow DIRETO (não a planilha), então cada item some do
// alerta assim que o status é fechado. Janela: 1º dia do mês
// passado até ontem (cobre o ciclo de repasse recente).
// ================================================================
var _cfStatusMap = null;
function cfCarregarStatusMap() {
  if (_cfStatusMap) return _cfStatusMap;
  _cfStatusMap = {};
  try {
    var resp = UrlFetchApp.fetch(CF_FEEGOW_BASE + '/appoints/status', {
      headers: { 'x-access-token': CF_FEEGOW_TOKEN }, muteHttpExceptions: true
    });
    var json  = JSON.parse(resp.getContentText());
    var items = Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);
    items.forEach(function(s) {
      var id = s.status_id || s.id;
      if (id != null) _cfStatusMap[id] = String(s.nome || s.name || s.status || '').toLowerCase().trim();
    });
  } catch (e) { Logger.log('cfCarregarStatusMap erro: ' + e.message); }
  return _cfStatusMap;
}

// dd/MM/yyyy a partir de um Date
function cfDataBR(d) {
  return Utilities.formatDate(d, 'America/Sao_Paulo', 'dd/MM/yyyy');
}
// dd/MM/yyyy a partir da string do Feegow (aceita yyyy-mm-dd, dd-mm-yyyy, dd/mm/yyyy)
function cfDataBRFeegow(s) {
  s = String(s || '').slice(0, 10);
  var m;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/))) return m[3] + '/' + m[2] + '/' + m[1];
  if ((m = s.match(/^(\d{2})-(\d{2})-(\d{4})$/))) return m[1] + '/' + m[2] + '/' + m[3];
  return s;
}
// chave yyyymmdd pra ordenar cronologicamente
function cfDataKey(s) {
  s = String(s || '').slice(0, 10);
  var m;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/))) return m[1] + m[2] + m[3];
  if ((m = s.match(/^(\d{2})[-\/](\d{2})[-\/](\d{4})$/))) return m[3] + m[2] + m[1];
  return s;
}

// Busca no Feegow os agendamentos "Em atendimento" entre duas datas (Date).
function cfBuscarEmAtendimento(ini, fim) {
  var url  = CF_FEEGOW_BASE + '/appoints/search?data_start=' + fmtDataFeegow(ini) + '&data_end=' + fmtDataFeegow(fim);
  var resp = UrlFetchApp.fetch(url, { headers: { 'x-access-token': CF_FEEGOW_TOKEN }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('Feegow fora do ar (HTTP ' + resp.getResponseCode() + ')');
  var json = JSON.parse(resp.getContentText());
  if (json && json.success === false) throw new Error('Feegow: ' + (json.message || 'success=false'));
  var items = Array.isArray(json.content) ? json.content : (Array.isArray(json) ? json : []);
  var smap  = cfCarregarStatusMap();
  return items.filter(function(a) {
    var nome = (smap[a.status_id] || a.status || '').toLowerCase().trim();
    return nome.indexOf('em atendimento') === 0;
  });
}

// Monta o texto do resumo de presos pro Slack (agrupado por médico). Reaproveitado
// pelo alertarEmAtendimento (manual) e pela agenda de segunda (automático).
function _resumoPresosSlack(presos, profMap, procMap) {
  var hoje  = new Date();
  var ontem = new Date(hoje.getTime() - 24 * 60 * 60 * 1000);
  var ini   = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
  var periodo = cfDataBR(ini) + ' a ' + cfDataBR(ontem);
  if (!presos || !presos.length) {
    return '🩺 *Presos em "Em atendimento"* (' + periodo + ')\n✅ Nenhum preso. Feegow em dia.';
  }
  var grupos = {};
  presos.forEach(function(a) {
    var pnome = profMap[a.profissional_id] || ('Profissional ' + (a.profissional_id || '?'));
    (grupos[pnome] = grupos[pnome] || []).push(a);
  });
  var nomes  = Object.keys(grupos).sort(function(x, y) { return grupos[y].length - grupos[x].length; });
  var linhas = ['🩺 *Agendamentos presos em "Em atendimento"* (' + periodo + ')',
                '_Só o médico marca "Atendido" no Feegow. Peça pra fecharem os itens abaixo (senão somem dos relatórios)._', ''];
  var total = 0;
  nomes.forEach(function(nome) {
    var arr = grupos[nome].sort(function(x, y) { return cfDataKey(x.data).localeCompare(cfDataKey(y.data)); });
    total += arr.length;
    linhas.push('• *' + nome + '* · ' + arr.length);
    arr.slice(0, 6).forEach(function(a) {
      var proc = procMap[a.procedimento_id] || '';
      var pac  = '';
      try { pac = getPatientData(a.paciente_id).nome || ''; } catch (e) {}
      if (!pac) pac = '#' + (a.paciente_id || '?');
      linhas.push('     ' + cfDataBRFeegow(a.data) + '  ' + _titleCaseNome(pac) + (proc ? ' · ' + _titleCaseNome(proc) : ''));
    });
    if (arr.length > 6) linhas.push('     _+' + (arr.length - 6) + ' mais_');
  });
  linhas.push('');
  linhas.push('Total: *' + total + '* preso' + (total === 1 ? '' : 's') + '.');
  return linhas.join('\n');
}

// Resumo dos presos no Slack. Roda automático toda segunda (dentro da agenda dos
// médicos, sem gatilho separado); também pode ser chamado na mão pra testar.
function alertarEmAtendimento() {
  var itens;
  try {
    var hoje  = new Date();
    var ontem = new Date(hoje.getTime() - 24 * 60 * 60 * 1000);
    var ini   = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
    itens = cfBuscarEmAtendimento(ini, ontem);
  } catch (e) {
    slackPost('🚨 Alerta "Em atendimento" não rodou: ' + e.message);
    return;
  }
  slackPost(_resumoPresosSlack(itens, carregarProfissionais(), carregarNomesProcedimentos()));
}
