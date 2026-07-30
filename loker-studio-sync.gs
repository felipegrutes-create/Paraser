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
const LIMITE_API = 1000;
const DATA_INICIO_GERAL = "01/01/2025";


// =================================================================================
// FUNÇÃO PRINCIPAL - A SER CHAMADA PELO ACIONADOR DE HORA EM HORA
// =================================================================================
function sincronizarFeegow() {
  const properties = PropertiesService.getUserProperties();
  const cargaInicialConcluida = properties.getProperty('CARGA_INICIAL_CONCLUIDA');

  if (cargaInicialConcluida !== 'true') {
    Logger.log('FASE 1: Executando lote da Carga Inicial Massiva (baseada na Data de Criação).');
    _executarLoteCargaInicial();
  } else {
    Logger.log('FASE 2: Executando Sincronização Horária de Rotina.');
    _executarSincronizacaoHoraria();
  }
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

        for (let diaAtual = new Date(dataInicio); diaAtual <= dataFim; diaAtual.setDate(diaAtual.getDate() + 1)) {
            let offset = 0;
            while (true) {
                const dataFormatada = formatBR(diaAtual);
                const payload = { report: "schedule-appointments", DATA_INICIO: dataFormatada, DATA_FIM: dataFormatada, offset: offset, limit: LIMITE_API };
                const registros = requisicaoFeegow(API_URL, API_TOKEN, payload);

                if (registros === null) break;
                if (registros.length > 0) processarDados(sheet, registros, idxPorId);
                if (registros.length < LIMITE_API) break;
                offset += LIMITE_API;
            }
        }
        Logger.log('SINCRONIZAÇÃO HORÁRIA: Verificação concluída.');
    } finally {
        lock.releaseLock();
    }
}


// =================================================================================
// FUNÇÕES DE APOIO (PROCESSAMENTO E HELPERS)
// =================================================================================
function processarDados(sheet, dadosApi, idxPorId) {
    const appends = [];
    const updates = [];
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
            updates.push({ rangeA1: `A${rowIndex}`, values: [linha] });
        } else {
            appends.push(linha);
        }
    });

    try {
        if (updates.length > 0) updates.forEach(u => sheet.getRange(u.rangeA1).offset(0, 0, 1, HEADERS.length).setValues(u.values));
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