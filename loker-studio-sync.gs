/****************************************************************************************
 * *
 * FEEGOW → GOOGLE SHEETS: SOLUÇÃO HÍBRIDA E DEFINITIVA (V5)          *
 * *
 * - CORREÇÃO DE ORDENAÇÃO: A planilha agora é ordenada pela "Data de Criação"        *
 * (coluna Q) ao final da carga inicial.                                           *
 * - FASE 1: Carga Inicial Massiva (baseada na DATA DE CRIAÇÃO).                       *
 * - FASE 2: Sincronização Horária Automática (baseada na DATA DO AGENDAMENTO).        *
 * *
 ****************************************************************************************/

// =================================================================================
// CONFIGURAÇÕES GLOBAIS
// =================================================================================
const API_URL = "https://api.feegow.com/v1/api/reports/generate";
const API_TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJmZWVnb3ciLCJhdWQiOiJwdWJsaWNhcGkiLCJpYXQiOjE3NDM0NzEyNDIsImxpY2Vuc2VJRCI6MTQ0MzR9.oh2VSWT5UPEfYRrPCv34IM1NuP8Aq_ehFYWhE8f5MuU";
const SHEET_NAME = "Dados Diários";
const HEADERS = [
    "ID Agendamento","Data","Horário","Paciente","Celular","CPF","E-mail","Endereço",
    "Faturado","Idade","Local","Profissional","Procedimento",
    "Status","Valor","Usuário","Data de Criação","Mês"
];
// ⚠️ A API IGNORA offset e limit no report schedule-appointments: ela devolve a faixa
// de datas inteira, sempre o mesmo conjunto. LIMITE_API só sobrou na Fase 1 (carga
// inicial, já concluída, que pede 1 dia por vez e nunca chega perto de 1000).
const LIMITE_API = 1000;
const DATA_INICIO_GERAL = "01/01/2025";


// =================================================================================
// FUNÇÃO PRINCIPAL - A SER CHAMADA PELO ACIONADOR DE HORA EM HORA
// =================================================================================
function sincronizarFeegow() {
  const properties = PropertiesService.getUserProperties();
  const cargaInicialConcluida = properties.getProperty('CARGA_INICIAL_CONCLUIDA');

  // 18/08/2026 — BATIMENTO CARDÍACO. O acionador parou de disparar em 07/08 e
  // ninguém viu por 11 dias: a aba congelou e as consultas de 1ª vez (as que são
  // marcadas em cima da hora) sumiram do dashboard. Sem log de execução acessível
  // por API, não dava pra provar se o acionador estava vivo. Agora toda rodada
  // carimba a hora aqui; `?acao=status` responde e o dashboard avisa se envelhecer.
  try {
    if (cargaInicialConcluida !== 'true') {
      Logger.log('FASE 1: Executando lote da Carga Inicial Massiva (baseada na Data de Criação).');
      _executarLoteCargaInicial();
    } else {
      Logger.log('FASE 2: Executando Sincronização Horária de Rotina.');
      _executarSincronizacaoHoraria();
    }
    properties.setProperty('ULTIMO_SYNC_OK', new Date().toISOString());
    properties.deleteProperty('ULTIMO_SYNC_ERRO');
  } catch (e) {
    properties.setProperty('ULTIMO_SYNC_ERRO', new Date().toISOString() + ' | ' + e.message);
    throw e;
  }
}


// =================================================================================
// ACIONADOR — recriar do zero (18/08/2026)
// =================================================================================
// `ScriptApp.getProjectTriggers()` lista o acionador mesmo depois que o Google o
// desativa por falhas repetidas, então "aparece na lista" NÃO é prova de que ele
// dispara. Apagar e criar de novo é o único jeito de garantir um acionador vivo.
function recriarAcionador() {
  let apagados = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sincronizarFeegow') { ScriptApp.deleteTrigger(t); apagados++; }
  });
  ScriptApp.newTrigger('sincronizarFeegow').timeBased().everyHours(1).create();
  return { apagados: apagados, criados: 1 };
}


// =================================================================================
// FASE 1: LÓGICA DA CARGA INICIAL MASSIVA
// =================================================================================
function _executarLoteCargaInicial() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(300000)) {
    Logger.log('CARGA INICIAL: Execução simultânea evitada.');
    return;
  }

  const properties = PropertiesService.getUserProperties();
  const cursorDataStr = properties.getProperty('cursorData') || DATA_INICIO_GERAL;
  let cursorOffset = parseInt(properties.getProperty('cursorOffset') || '0', 10);
  
  const hoje = new Date(); 
  const dataFimBusca = addDays(hoje, 730); 
  let cursorData = strBRToDate(cursorDataStr);

  Logger.log(`CARGA INICIAL: Buscando agendamentos do dia ${formatBR(cursorData)}, offset ${cursorOffset}.`);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME);
        sheet.appendRow(HEADERS);
    }
    const idxPorId = new Map(sheet.getDataRange().getValues().slice(1).map((row, i) => [String(row[0]), i + 2]));

    const tempoInicio = new Date();
    const limiteExecucao = 5 * 60 * 1000;

    while (cursorData <= dataFimBusca && (new Date() - tempoInicio < limiteExecucao)) {
      const dI = formatBR(cursorData);
      const payload = { report: "schedule-appointments", DATA_INICIO: dI, DATA_FIM: dI, offset: cursorOffset, limit: LIMITE_API };
      const registrosApi = requisicaoFeegow(API_URL, API_TOKEN, payload);

      if (registrosApi === null) {
          Logger.log(`Falha na API para o dia ${dI}. Tentando novamente no próximo lote.`);
          break; 
      }
      
      if (registrosApi.length > 0) {
        const registrosFiltrados = registrosApi.filter(ag => {
            const dataCriStr = extrairDataCriacao(ag);
            if (!dataCriStr) return false;
            const dataCri = strBRToDate(dataCriStr);
            return dataCri <= hoje;
        });

        if (registrosFiltrados.length > 0) {
            Logger.log(`Processando ${registrosFiltrados.length} de ${registrosApi.length} registros encontrados para ${dI} (criados até hoje)`);
            processarDados(sheet, registrosFiltrados, idxPorId);
        }
      }

      if (registrosApi.length < LIMITE_API) {
        cursorData.setDate(cursorData.getDate() + 1);
        cursorOffset = 0;
      } else {
        cursorOffset += LIMITE_API;
      }
    }

    if (cursorData > dataFimBusca) {
      Logger.log("🎯🎉 CARGA INICIAL CONCLUÍDA! Ordenando dados...");
      properties.setProperty('CARGA_INICIAL_CONCLUIDA', 'true');
      properties.deleteProperty('cursorData');
      properties.deleteProperty('cursorOffset');
      
      const ultimaLinha = sheet.getLastRow();
      if(ultimaLinha > 1) {
        // ==> CORREÇÃO DE ORDENAÇÃO APLICADA AQUI <==
        // Ordena pela coluna 17 (Data de Criação) em ordem crescente.
        sheet.getRange(2, 1, ultimaLinha - 1, HEADERS.length).sort({ column: 17, ascending: true });
      }
      Logger.log("Ordenação finalizada.");

    } else {
      properties.setProperties({ 'cursorData': formatBR(cursorData), 'cursorOffset': String(cursorOffset) });
      Logger.log(`CARGA INICIAL: Lote finalizado. Próxima busca de agendamentos começará de: ${formatBR(cursorData)}, offset: ${cursorOffset}`);
    }
  } finally {
    lock.releaseLock();
  }
}


// =================================================================================
// FASE 2: LÓGICA DA SINCRONIZAÇÃO HORÁRIA DE ROTINA
// =================================================================================
function _executarSincronizacaoHoraria() {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(300000)) {
        Logger.log('SINCRONIZAÇÃO HORÁRIA: Execução simultânea evitada.');
        return;
    }

    try {
        const hoje = new Date();
        const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
        const dataInicio = inicioMes;
        const dataFim = addDays(hoje, 180);

        Logger.log(`SINCRONIZAÇÃO HORÁRIA: Verificando período de ${formatBR(dataInicio)} a ${formatBR(dataFim)}`);

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEET_NAME);
        if (!sheet) return;
        
        const idxPorId = new Map(sheet.getDataRange().getValues().slice(1).map((row, i) => [String(row[0]), i + 2]));

        // 30/07/2026 — Um pedido por MÊS, não um por dia.
        // Antes: 1 chamada por dia, do 1º do mês até hoje+180 = ~210 chamadas POR RODADA,
        // de hora em hora (~5 mil por dia). Agora: 7 chamadas.
        // O relatório devolve a faixa inteira numa chamada só (testado: outubro/2025
        // devolveu 1166 registros de uma vez). ⚠️ NÃO paginar por offset aqui: a API
        // IGNORA offset e limit, devolve sempre o mesmo conjunto, então o laço antigo
        // (`while(true)` até vir menos que LIMITE_API) viraria LOOP INFINITO em qualquer
        // bloco com 1000 registros ou mais. Por isso o bloco é mensal e sem paginação.
        let chamadas = 0, processados = 0;
        let blocoIni = new Date(dataInicio);
        while (blocoIni <= dataFim) {
            let blocoFim = new Date(blocoIni.getFullYear(), blocoIni.getMonth() + 1, 0); // último dia do mês
            if (blocoFim > dataFim) blocoFim = new Date(dataFim);

            const payload = { report: "schedule-appointments", DATA_INICIO: formatBR(blocoIni), DATA_FIM: formatBR(blocoFim) };
            const registros = requisicaoFeegow(API_URL, API_TOKEN, payload);
            chamadas++;

            if (registros === null) {
                Logger.log(`Falha na API no bloco ${formatBR(blocoIni)}–${formatBR(blocoFim)}. Segue pro próximo bloco.`);
            } else {
                if (registros.length > 0) {
                    processarDados(sheet, registros, idxPorId);
                    processados += registros.length;
                }
                if (registros.length >= 5000) {
                    Logger.log(`⚠️ Bloco ${formatBR(blocoIni)} devolveu ${registros.length} registros. Se a API começar a truncar, quebrar o bloco em quinzenas.`);
                }
            }

            blocoIni = new Date(blocoFim.getFullYear(), blocoFim.getMonth(), blocoFim.getDate() + 1);
        }
        Logger.log(`SINCRONIZAÇÃO HORÁRIA: Verificação concluída. ${chamadas} chamadas, ${processados} registros.`);
    } finally {
        lock.releaseLock();
    }
}


// =================================================================================
// FUNÇÕES DE APOIO (PROCESSAMENTO E HELPERS)
// =================================================================================
function processarDados(sheet, dadosApi, idxPorId) {
    const appends = [];
    const porLinha = new Map(); // nº da linha na planilha -> valores da linha
    dadosApi.forEach(ag => {
        const id = String(ag.AgendamentoID);
        const dataCri = extrairDataCriacao(ag);
        const linha = [
            id, ag.Data, ag.Hora, ag.NomePaciente, ag.Cel1 || "", ag.CPF || "", ag.Email1 || "",
            `${ag.Endereco || ""}, ${ag.Numero || ""} - ${ag.Bairro || ""}, ${ag.Cidade || ""}`,
            ag.Faturado || "Não", ag.Idade || "", ag.NomeLocal || "",
            ag.NomeProfissional || ag.Profissional || "", ag.NomeProcedimento || "",
            ag.StaConsulta || "", ag.Valor || "0,00", ag.Nome || "",
            dataCri || "",
            obterMesPorExtenso(dataCri || "")
        ];
        const rowIndex = idxPorId.get(id);
        if (rowIndex) {
            porLinha.set(rowIndex, linha);
        } else {
            appends.push(linha);
        }
    });

    try {
        // 30/07/2026 — escreve em BLOCOS CONTÍGUOS, não uma chamada por linha.
        // Antes: um setValues por linha (~1.171 escritas por rodada). Agora linhas
        // vizinhas viram um bloco só (~91 escritas, medido na aba real).
        // De propósito NÃO escrevo a faixa inteira de uma vez (linha 7530 até 41146):
        // isso reescreveria 32 mil linhas alheias e poderia apagar o carimbo
        // "Excluído no Feegow" que a faxina do CRM grava nessa mesma aba.
        let blocos = 0;
        const linhas = Array.from(porLinha.keys()).sort((a, b) => a - b);
        let i = 0;
        while (i < linhas.length) {
            let j = i;
            while (j + 1 < linhas.length && linhas[j + 1] === linhas[j] + 1) j++;
            const valores = [];
            for (let k = i; k <= j; k++) valores.push(porLinha.get(linhas[k]));
            sheet.getRange(linhas[i], 1, valores.length, HEADERS.length).setValues(valores);
            blocos++;
            i = j + 1;
        }
        if (linhas.length) Logger.log(`  ${linhas.length} linhas atualizadas em ${blocos} blocos.`);
        if (appends.length > 0) {
            const newRowStart = sheet.getLastRow() + 1;
            sheet.getRange(newRowStart, 1, appends.length, HEADERS.length).setValues(appends);
            appends.forEach((linha, i) => idxPorId.set(linha[0], newRowStart + i));
        }
    } catch (e) {
        Logger.log(`ERRO ao escrever na planilha: ${e.message}.`);
    }
}

// ... As funções helpers (requisicaoFeegow, etc.) permanecem as mesmas ...
function requisicaoFeegow(API_URL, API_TOKEN, payload) { try { const resp = UrlFetchApp.fetch(API_URL, { method: "post", headers: { "Content-Type": "application/json", "x-access-token": API_TOKEN }, payload: JSON.stringify(payload), muteHttpExceptions: true }); const code = resp.getResponseCode(); if (code !== 200) { Logger.log(`HTTP ${code} | payload=${JSON.stringify(payload)}`); return null; } const body = JSON.parse(resp.getContentText()); return Array.isArray(body.data) ? body.data : Array.isArray(body.result) ? body.result : Array.isArray(body.results) ? body.results : []; } catch (e) { Logger.log(`Erro HTTP: ${e.message}`); return null; } }
function extrairDataCriacao(ag) { const chaves = ["DataCriacao","Criacao","CriacaoAgendamento","Data_Criacao","DataDeCriacao","created_at","CreatedAt","Criação do agendamento","Criacao do agendamento"]; for (let k of chaves) { if (ag && ag[k]) { const raw = String(ag[k]).trim(); if (/^\d{4}-\d{2}-\d{2}/.test(raw)) { const [datePart] = raw.split(/[T ]/); const [y,m,d] = datePart.split("-"); return `${d}/${m}/${y}`; } const m = raw.match(/(\d{2}\/\d{2}\/\d{4})/); if (m) return m[1]; } } return ""; }
function obterMesPorExtenso(dataString) { if (!dataString) return ""; const mm = parseInt(dataString.split("/")[1], 10); const meses = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"]; return meses[mm-1] || ""; }
function strBRToDate(s){ const [d,m,y]=s.split("/").map(n=>parseInt(n,10)); return new Date(y,m-1,d); }
function formatBR(d){ return Utilities.formatDate(d,"GMT-3","dd/MM/yyyy"); }
function addDays(d,n){ const x=new Date(d.getTime()); x.setDate(x.getDate()+n); return x; }
// =================================================================================
// GATILHO MANUAL (30/07/2026) — só pela URL /dev, que exige login de uma conta com
// acesso de edição a esta planilha. Serve pra rodar/conferir o sync na hora, sem
// esperar o acionador de hora em hora. Não cria nada público.
// =================================================================================
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.key !== 'paraser2026') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (p.acao === 'sincronizar') {
    const t0 = new Date();
    sincronizarFeegow();
    return ContentService.createTextOutput(JSON.stringify({
      ok: true, segundos: Math.round((new Date() - t0) / 1000)
    })).setMimeType(ContentService.MimeType.JSON);
  }
  if (p.acao === 'triggers') {
    return ContentService.createTextOutput(JSON.stringify({
      triggers: ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction() + ' (' + t.getEventType() + ')'),
      cargaInicialConcluida: PropertiesService.getUserProperties().getProperty('CARGA_INICIAL_CONCLUIDA')
    })).setMimeType(ContentService.MimeType.JSON);
  }
  if (p.acao === 'recriar') {
    return ContentService.createTextOutput(JSON.stringify(recriarAcionador()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (p.acao === 'status') {
    const pr = PropertiesService.getUserProperties();
    return ContentService.createTextOutput(JSON.stringify({
      ultimoSyncOk: pr.getProperty('ULTIMO_SYNC_OK') || null,
      ultimoSyncErro: pr.getProperty('ULTIMO_SYNC_ERRO') || null,
      triggers: ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction())
    })).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: false }))
    .setMimeType(ContentService.MimeType.JSON);
}
