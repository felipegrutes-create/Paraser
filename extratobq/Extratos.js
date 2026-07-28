/**
 * === CONFIGURAÇÕES GERAIS ===
 */
const PASTA_PARASER_ID = "1hCmfZGUE8LAMdE-89v9s3bdM_bXh7OIl";
const PASTA_INSTITUTO_ID = "1WbX5eUW3jWklniZnF3nJoEDaGm1RdyGf";
const PLANILHA_LANC_PARASER = "LANÇAMENTOS";
const PLANILHA_LANC_INSTITUTO = "LANÇAMENTOS_I";
const PLANILHA_PAGAMENTOS = "PAGAMENTOS";

const TZ_FIXO = "America/Sao_Paulo";

/**
 * === MENU PERSONALIZADO ===
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💰 Atualizar Extratos")
    .addItem("🔄 Puxar do banco (ParaSer + Instituto)", "atualizarExtratosBQ")
    .addItem("🧪 Testar classificação (aba TESTE_BQ)", "testarClassificacaoBQ")
    .addItem("▶️ Ligar automático (hora em hora)", "ligarExtratoBQ")
    .addItem("⏸️ Desligar automático", "desligarExtratoBQ")
    .addSeparator()
    .addItem("Atualizar Extrato ParaSer (pasta)", "atualizarExtratoParaser")
    .addItem("Atualizar Extrato Instituto (pasta)", "atualizarExtratoInstituto")
    .addItem("Conciliar saídas com PAGAMENTOS", "conciliarSaidasExistentesComPagamentos")
    .addToUi();
}

/**
 * === FUNÇÕES DE ATALHO ===
 */
function atualizarExtratoParaser() {
  processarExtratos(PASTA_PARASER_ID, PLANILHA_LANC_PARASER, "PARASER");
}

function atualizarExtratoInstituto() {
  processarExtratos(PASTA_INSTITUTO_ID, PLANILHA_LANC_INSTITUTO, "INSTITUTO");
}

/**
 * === FUNÇÃO PRINCIPAL ===
 */
function processarExtratos(pastaId, abaLancamentos, empresaNome) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let pasta;
  try {
    pasta = DriveApp.getFolderById(pastaId);
  } catch (e) {
    ui.alert(
      `❌ Erro: não foi possível acessar a pasta do ${empresaNome}.\n` +
      `Verifique o ID da pasta e permissões.`
    );
    return;
  }

  const aba = ss.getSheetByName(abaLancamentos);
  const pagamentos = ss.getSheetByName(PLANILHA_PAGAMENTOS);

  if (!aba) {
    ui.alert(`❌ Erro: a aba "${abaLancamentos}" não foi encontrada.`);
    return;
  }

  if (!pagamentos) {
    ui.alert(`❌ Erro: a aba "${PLANILHA_PAGAMENTOS}" não foi encontrada.`);
    return;
  }

  const dadosPagamentos = pagamentos.getDataRange().getValues();

  let qtdComData = aba.getRange("A:A").getValues().filter(String).length;
  if (qtdComData < 2) qtdComData = 1;
  let ultimaComData = qtdComData;

  const dadosExistentes = ultimaComData > 1
    ? aba.getRange(2, 1, ultimaComData - 1, 14).getValues()
    : [];

  // ✅ Considera somente ontem e hoje
  const dataMinima = obterDataDiasAtrasSP(1);

  // 🔑 Mapa de lançamentos existentes: DATA + VALOR + DESCRIÇÃO normalizada
  const mapaExistentes = new Set();

  dadosExistentes.forEach(l => {
    const dataKey = dataParaChave(l[0]);
    const valorNorm = normalizarValor(l[3]);
    const descNorm = normalizarDescricao(l[1]);

    if (!dataKey || valorNorm === null) return;

    mapaExistentes.add(`${dataKey}|${valorNorm}|${descNorm}`);
  });

  let novosLancamentos = 0;
  let pendentes = 0;

  const cadastroSheet = ss.getSheetByName("Cadastros");
  let fornecedoresValidos = [];
  let planosValidos = [];
  let classificacoesValidas = [];

  if (cadastroSheet) {
    fornecedoresValidos = cadastroSheet.getRange("H2:H").getValues().flat().filter(String);

    const receitas = cadastroSheet.getRange("A2:A").getValues().flat().filter(String);
    const despesas = cadastroSheet.getRange("B2:B").getValues().flat().filter(String);
    planosValidos = [...receitas, ...despesas];

    const regras = aba.getRange("C2:C").getDataValidations();

    if (regras && regras[0] && regras[0][0]) {
      const criterio = regras[0][0].getCriteriaValues();
      if (criterio && criterio[0]) classificacoesValidas = criterio[0];
    }
  }

  const nomeSubpasta = "Processados";
  const pastas = pasta.getFoldersByName(nomeSubpasta);
  const subpasta = pastas.hasNext() ? pastas.next() : pasta.createFolder(nomeSubpasta);

  const arquivos = pasta.getFiles();

  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();

    const mime = arquivo.getMimeType();

    if (mime === MimeType.SHORTCUT) {
      Logger.log(`⚠️ Atalho ignorado: ${arquivo.getName()} (${arquivo.getId()})`);
      continue;
    }

    if (mime !== MimeType.GOOGLE_SHEETS) {
      Logger.log(`⚠️ Arquivo ignorado, não é Google Sheets: ${arquivo.getName()} [${mime}]`);
      continue;
    }

    let arquivoPlanilha;

    try {
      arquivoPlanilha = SpreadsheetApp.openById(arquivo.getId());
    } catch (e) {
      Logger.log(`❌ Falha ao abrir: ${arquivo.getName()} — ID ${arquivo.getId()} — ${e}`);
      continue;
    }

    const primeiraAba = arquivoPlanilha.getSheets()[0];

    if (!primeiraAba) {
      Logger.log(`⚠️ Planilha sem abas: ${arquivo.getName()}`);
      continue;
    }

    const todasLinhas = primeiraAba.getDataRange().getValues();

    let linhaInicio = -1;

    for (let i = 0; i < todasLinhas.length; i++) {
      const linhaTxt = normalizarDescricao(todasLinhas[i].join(" "));
      const temData = linhaTxt.includes("DATA");
      const temLanc =
        linhaTxt.includes("LANÇ") ||
        linhaTxt.includes("LANC") ||
        linhaTxt.includes("HIST") ||
        linhaTxt.includes("DESCR");

      if (temData && temLanc) {
        linhaInicio = i;
        break;
      }
    }

    if (linhaInicio === -1) {
      Logger.log(`⚠️ Não foi possível localizar cabeçalho em ${arquivo.getName()}, ignorado.`);
      continue;
    }

    const cabecalho = todasLinhas[linhaInicio];

    let idxData = -1;
    let idxLanc = -1;
    let idxValor = -1;
    let idxSaldo = -1;

    for (let c = 0; c < cabecalho.length; c++) {
      const txt = normalizarDescricao(cabecalho[c] || "");

      if (txt.includes("DATA")) idxData = c;

      if (
        idxLanc === -1 &&
        (
          txt.includes("LANÇ") ||
          txt.includes("LANC") ||
          txt.includes("HIST") ||
          txt.includes("DESCR")
        )
      ) {
        idxLanc = c;
      }

      if (txt.includes("VALOR")) idxValor = c;
      if (txt.includes("SALDO")) idxSaldo = c;
    }

    if (idxData === -1) idxData = 0;
    if (idxLanc === -1) idxLanc = 1;
    if (idxValor === -1) idxValor = 3;
    if (idxSaldo === -1) idxSaldo = 4;

    const dados = todasLinhas.slice(linhaInicio + 1);
    let houveImportacao = false;

    dados.forEach(linha => {
      const dataBruta = linha[idxData];
      const lancamento = linha[idxLanc];
      const valorRaw = linha[idxValor];

      if (
        !dataBruta ||
        valorRaw === null ||
        typeof valorRaw === "undefined" ||
        valorRaw === ""
      ) {
        return;
      }

      const dataConvertida = parseDataBR(dataBruta);
      if (!dataConvertida) return;

      // ✅ Considera apenas ontem e hoje
      if (dataConvertida.getTime() < dataMinima.getTime()) return;

      const chaveData = Utilities.formatDate(dataConvertida, TZ_FIXO, "yyyy-MM-dd");

      const valorNum = normalizarValor(valorRaw);
      if (valorNum === null) return;

      const descOriginal = lancamento ? lancamento.toString() : "";
      const descNorm = normalizarDescricao(descOriginal);

      const chaveCompleta = `${chaveData}|${valorNum}|${descNorm}`;

      // ✅ Impede duplicidade
      if (mapaExistentes.has(chaveCompleta)) return;

      let classificacao = "";
      let fornecedor = "";
      let planoContas = "";

      if (descNorm.includes("REDE")) {
        classificacao = "Clientes/REDE";
        fornecedor = "REDE";
        planoContas = "Receita Produtos e Serviços";
      } else if (descNorm.includes("PIX")) {
        classificacao = "Clientes/PIX";
        fornecedor = "PACIENTES - GERAL";
        planoContas = "Receita Produtos e Serviços";
      }

      if (valorNum < 0) {
        const valorAbs = Math.abs(valorNum);

        for (let i = 1; i < dadosPagamentos.length; i++) {
          const linhaP = dadosPagamentos[i];

          const dataPgto = linhaP[0];
          const planoPgto = linhaP[1];
          const fornecedorPgto = linhaP[2];
          const status = linhaP[4];
          const valorPgto = Number(linhaP[5]);
          const contaPgto = linhaP[9];

          if (!dataPgto || isNaN(valorPgto) || !fornecedorPgto) continue;
          if (status !== "PAGO") continue;
          if (contaPgto !== "CC ITAÚ") continue;

          const dataPgtoKey = dataParaChave(dataPgto);
          if (dataPgtoKey !== chaveData) continue;
          if (Math.abs(valorPgto) !== valorAbs) continue;

          const fornecedorPgtoNorm = normalizarDescricao(fornecedorPgto);
          const fornecedorPrefixo = fornecedorPgtoNorm.substring(0, 15);

          const fornecedorBate =
            fornecedorPgtoNorm.length > 0 &&
            (
              descNorm.includes(fornecedorPrefixo) ||
              fornecedorPgtoNorm.includes(descNorm.substring(0, Math.min(15, descNorm.length)))
            );

          if (!fornecedorBate) continue;

          fornecedor = fornecedorPgto;
          planoContas = planoPgto;

          if (!classificacao) classificacao = "Pagamentos";

          break;
        }
      }

      if (descNorm.includes("HOSPINOVA") || descNorm.includes("ONCO PROD")) {
        classificacao = "Medicamentos";
      }

      fornecedor = normalizarDropdown(fornecedor, fornecedoresValidos);
      planoContas = normalizarDropdown(planoContas, planosValidos);

      let classificacaoValida = false;

      if (classificacao && classificacoesValidas.length > 0) {
        classificacaoValida = classificacoesValidas
          .map(v => v.toString().trim().toUpperCase())
          .includes(classificacao.toString().trim().toUpperCase());
      }

      if (!classificacaoValida) classificacao = "";

      ultimaComData++;
      const novaLinhaIndex = ultimaComData;

      aba.insertRowAfter(novaLinhaIndex - 1);

      gravarDataPura_(aba.getRange(novaLinhaIndex, 1), dataConvertida); // dia puro, sem hora

      aba.getRange(novaLinhaIndex, 2).setValue(descOriginal || "");
      aba.getRange(novaLinhaIndex, 3).setValue(classificacao);
      aba.getRange(novaLinhaIndex, 4).setValue(valorNum || "");
      aba.getRange(novaLinhaIndex, 10).setValue(valorNum > 0 ? "Entrada" : "Saída");
      aba.getRange(novaLinhaIndex, 11).setValue(obterMes(dataConvertida));
      aba.getRange(novaLinhaIndex, 12).setValue(empresaNome);
      aba.getRange(novaLinhaIndex, 13).setValue(fornecedor || "");
      aba.getRange(novaLinhaIndex, 14).setValue(planoContas || "");

      if (novaLinhaIndex > 2) {
        aba.getRange(novaLinhaIndex - 1, 6, 1, 4)
          .copyTo(
            aba.getRange(novaLinhaIndex, 6, 1, 4),
            SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
            false
          );
      }

      if (novaLinhaIndex > 2) {
        const dvFornecedor = aba.getRange(novaLinhaIndex - 1, 13).getDataValidation();

        if (dvFornecedor) {
          aba.getRange(novaLinhaIndex, 13).setDataValidation(dvFornecedor);
        }
      }

      if (!classificacao) {
        aba.getRange(novaLinhaIndex, 1, 1, 14).setBackground("#FFF9C4");
        pendentes++;
      } else {
        aba.getRange(novaLinhaIndex, 1, 1, 14).setBackground(null);
      }

      novosLancamentos++;
      houveImportacao = true;

      mapaExistentes.add(chaveCompleta);
    });

    if (houveImportacao) {
      try {
        arquivo.moveTo(subpasta);
      } catch (e) {
        Logger.log(`⚠️ Não foi possível mover para Processados: ${arquivo.getName()} — ${e}`);
      }
    }
  }

  const msgPeriodo =
    `\n(Período considerado: últimos 2 dias — desde ${Utilities.formatDate(dataMinima, TZ_FIXO, "dd/MM/yyyy")})`;

  if (novosLancamentos === 0) {
    ui.alert(`Nenhum novo lançamento encontrado para ${empresaNome}.${msgPeriodo}`);
  } else {
    ui.alert(
      `✅ Extratos da empresa ${empresaNome} atualizados com sucesso!\n\n` +
      `Foram importados ${novosLancamentos} novos lançamentos.` +
      (pendentes > 0 ? `\n⚠️ ${pendentes} linhas precisam de categorização manual.` : "") +
      msgPeriodo
    );
  }
}

/**
 * === CONCILIAÇÃO PARA SAÍDAS JÁ EXISTENTES ===
 */
function conciliarSaidasExistentesComPagamentos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pagamentos = ss.getSheetByName(PLANILHA_PAGAMENTOS);

  if (!pagamentos) {
    SpreadsheetApp.getUi().alert(`❌ Erro: a aba "${PLANILHA_PAGAMENTOS}" não foi encontrada.`);
    return;
  }

  const dadosPagamentos = pagamentos.getDataRange().getValues();

  const cadastroSheet = ss.getSheetByName("Cadastros");
  let fornecedoresValidos = [];
  let planosValidos = [];
  let classificacoesValidas = [];

  if (cadastroSheet) {
    fornecedoresValidos = cadastroSheet.getRange("H2:H").getValues().flat().filter(String);

    const receitas = cadastroSheet.getRange("A2:A").getValues().flat().filter(String);
    const despesas = cadastroSheet.getRange("B2:B").getValues().flat().filter(String);
    planosValidos = [...receitas, ...despesas];

    const abaParaser = ss.getSheetByName(PLANILHA_LANC_PARASER);

    if (abaParaser) {
      const regrasParaser = abaParaser.getRange("C2:C").getDataValidations();

      if (regrasParaser && regrasParaser[0] && regrasParaser[0][0]) {
        const criterio = regrasParaser[0][0].getCriteriaValues();
        if (criterio && criterio[0]) classificacoesValidas = criterio[0];
      }
    }
  }

  const abasLanc = [PLANILHA_LANC_PARASER, PLANILHA_LANC_INSTITUTO];

  abasLanc.forEach(nomeAba => {
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) return;

    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha <= 2) return;

    const range = aba.getRange(2, 1, ultimaLinha - 1, 14);
    const valores = range.getValues();

    for (let i = 0; i < valores.length; i++) {
      const linha = valores[i];

      const dataCell = linha[0];
      const descOriginal = linha[1];
      let classificacao = linha[2];
      const valorCell = linha[3];
      const tipo = linha[9];
      let fornecedor = linha[12];
      let planoContas = linha[13];

      const dataKey = dataParaChave(dataCell);
      const valorNum = normalizarValor(valorCell);

      if (!dataKey || valorNum === null) continue;
      if (tipo !== "Saída") continue;

      if (fornecedor && planoContas) continue;

      const descNorm = normalizarDescricao(descOriginal);
      const valorAbs = Math.abs(valorNum);

      for (let j = 1; j < dadosPagamentos.length; j++) {
        const linhaP = dadosPagamentos[j];

        const dataPgto = linhaP[0];
        const planoPgto = linhaP[1];
        const fornecedorPgto = linhaP[2];
        const status = linhaP[4];
        const valorPgto = Number(linhaP[5]);
        const contaPgto = linhaP[9];

        if (!dataPgto || isNaN(valorPgto) || !fornecedorPgto) continue;
        if (status !== "PAGO") continue;
        if (contaPgto !== "CC ITAÚ") continue;

        const dataPgtoKey = dataParaChave(dataPgto);
        if (dataPgtoKey !== dataKey) continue;
        if (Math.abs(valorPgto) !== valorAbs) continue;

        const fornecedorPgtoNorm = normalizarDescricao(fornecedorPgto);
        const fornecedorPrefixo = fornecedorPgtoNorm.substring(0, 15);

        const fornecedorBate =
          fornecedorPgtoNorm.length > 0 &&
          (
            descNorm.includes(fornecedorPrefixo) ||
            fornecedorPgtoNorm.includes(descNorm.substring(0, Math.min(15, descNorm.length)))
          );

        if (!fornecedorBate) continue;

        fornecedor = fornecedorPgto;
        planoContas = planoPgto;

        if (!classificacao) classificacao = "Pagamentos";

        break;
      }

      fornecedor = normalizarDropdown(fornecedor, fornecedoresValidos);
      planoContas = normalizarDropdown(planoContas, planosValidos);

      if (classificacao && classificacoesValidas.length > 0) {
        const ok = classificacoesValidas
          .map(v => v.toString().trim().toUpperCase())
          .includes(classificacao.toString().trim().toUpperCase());

        if (!ok) classificacao = linha[2];
      }

      valores[i][2] = classificacao || linha[2];
      valores[i][12] = fornecedor || linha[12];
      valores[i][13] = planoContas || linha[13];
    }

    range.setValues(valores);
  });

  SpreadsheetApp.getUi().alert("✅ Conciliação de saídas com PAGAMENTOS concluída.");
}

/**
 * === FUNÇÕES AUXILIARES ===
 */
function normalizarValor(valor) {
  if (valor === null || valor === "" || typeof valor === "undefined") return null;
  if (typeof valor === "number") return valor;

  let texto = valor.toString().trim();
  if (texto === "") return null;

  let negativo = false;

  if (texto.startsWith("(") && texto.endsWith(")")) {
    negativo = true;
    texto = texto.slice(1, -1).trim();
  }

  texto = texto.replace(/[^\d,.\-]/g, "");
  texto = texto.replace(/\.(?=\d{3}(,|\.|$))/g, "");
  texto = texto.replace(",", ".");

  const num = parseFloat(texto);

  if (isNaN(num)) return null;

  return negativo ? -Math.abs(num) : num;
}

function normalizarDescricao(desc) {
  if (!desc) return "";

  return desc
    .toString()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

function normalizarDropdown(valor, lista) {
  if (!valor) return "";

  const item = lista.find(
    v => v.toString().trim().toUpperCase() === valor.toString().trim().toUpperCase()
  );

  return item || valor;
}

function dataParaChave(dataCell) {
  const d = parseDataBR(dataCell);

  if (!d) return null;

  return Utilities.formatDate(d, TZ_FIXO, "yyyy-MM-dd");
}

function criarDataSPMeioDia(ano, mes, dia) {
  const mm = String(mes).padStart(2, "0");
  const dd = String(dia).padStart(2, "0");

  return new Date(`${ano}-${mm}-${dd}T12:00:00-03:00`);
}

/**
 * Grava a data na celula como DIA PURO, nunca como data-com-hora.
 *
 * Por que existe: parseDataBR devolve meio-dia de Sao Paulo, e o FUSO DESTA
 * PLANILHA e America/Los_Angeles. O mesmo instante, lido no fuso dela, aparece
 * como 07:00 ou 08:00 (muda com o horario de verao) — e a celula deixa de ser
 * um numero inteiro. Formula de outra aba que compara por data nao casa com
 * 45925,3333, so com 45925. Foi isso que escondeu 25 lancamentos das outras
 * abas ate 27/07/2026.
 *
 * A solucao grava o NUMERO DE SERIE do dia, por aritmetica pura, sem fuso
 * nenhum no meio. Usar em TODO lugar que escreve data na coluna A.
 */
function gravarDataPura_(celula, data) {
  const s = Utilities.formatDate(data, TZ_FIXO, "yyyy-MM-dd").split("-");
  const serial = Math.round(
    (Date.UTC(Number(s[0]), Number(s[1]) - 1, Number(s[2])) - Date.UTC(1899, 11, 30)) / 86400000
  );
  celula.setValue(serial);
  celula.setNumberFormat("dd/MM/yyyy");
}

function parseDataBR(dataBruta) {
  if (!dataBruta) return null;

  if (Object.prototype.toString.call(dataBruta) === "[object Date]" && !isNaN(dataBruta)) {
    const txt = Utilities.formatDate(dataBruta, TZ_FIXO, "dd/MM/yyyy");
    return parseDataBR(txt);
  }

  if (typeof dataBruta === "string") {
    let txt = dataBruta.trim();

    if (txt.indexOf(" ") !== -1) txt = txt.split(" ")[0];
    if (/^\d{4}-\d{2}-\d{2}T/.test(txt)) txt = txt.substring(0, 10);

    if (/^\d{4}-\d{2}-\d{2}$/.test(txt)) {
      const [ano, mes, dia] = txt.split("-").map(Number);

      if (!ano || !mes || !dia) return null;

      return criarDataSPMeioDia(ano, mes, dia);
    }

    if (txt.includes("/")) {
      const partes = txt.split("/");

      if (partes.length !== 3) return null;

      const dia = Number(partes[0]);
      const mes = Number(partes[1]);
      let anoStr = partes[2];

      let ano;

      if (anoStr.length === 2) {
        const num = Number(anoStr);
        ano = num >= 50 ? 1900 + num : 2000 + num;
      } else {
        ano = Number(anoStr);
      }

      if (!dia || !mes || !ano) return null;

      return criarDataSPMeioDia(ano, mes, dia);
    }
  }

  return null;
}

function obterDataDiasAtrasSP(diasAtras) {
  const hojeTxt = Utilities.formatDate(new Date(), TZ_FIXO, "dd/MM/yyyy");
  const hoje = parseDataBR(hojeTxt);

  hoje.setDate(hoje.getDate() - diasAtras);

  const dataTxt = Utilities.formatDate(hoje, TZ_FIXO, "dd/MM/yyyy");

  return parseDataBR(dataTxt);
}

function obterMes(data) {
  if (!data) return "";

  const meses = [
    "JANEIRO",
    "FEVEREIRO",
    "MARÇO",
    "ABRIL",
    "MAIO",
    "JUNHO",
    "JULHO",
    "AGOSTO",
    "SETEMBRO",
    "OUTUBRO",
    "NOVEMBRO",
    "DEZEMBRO"
  ];

  const d = parseDataBR(data);

  if (!d) return "";

  const mesIdx = Number(Utilities.formatDate(d, TZ_FIXO, "M")) - 1;

  return meses[mesIdx];
}
