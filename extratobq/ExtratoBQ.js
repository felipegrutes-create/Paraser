/**
 * === EXTRATO VIA BIGQUERY (automático) ===
 *
 * Puxa os lançamentos das 2 contas direto do BigQuery
 * (paraser-extrato-itau.extrato.lancamentos), que o coletor Cloud Run abastece
 * de hora em hora. Substitui o passo manual de baixar OFX e salvar na pasta.
 *
 * Classifica por DE-PARA (o que já foi classificado na aba, à mão):
 *   1) por assinatura da descrição (BOLETO PAGO X, SISPAG, ...)
 *   2) pelo favorecido/contraparte (PIX ENVIADO <fornecedor/funcionário/sócio>)
 * + REGRAS dos padrões fixos do banco. O que não casar fica em branco (amarelo).
 *
 * ⚠️ As descrições do banco NÃO batem com as que a analista digitou (o arquivo que
 * ela baixava formata diferente: "RECEBIMENTO REDE" vs "REDE", "PIX RECEBIDO" vs
 * "PIX TRANSF"). Por isso o dedupe NÃO pode ser por descrição. A escrita usa um
 * MARCO de data (Script Property BQ_CORTE): o sistema só insere lançamentos DEPOIS
 * do que a analista já fechou; daí pra frente deduplica entre as próprias linhas.
 *
 * Reusa helpers de Extratos.js: normalizarValor, normalizarDescricao,
 * normalizarDropdown, parseDataBR, dataParaChave, obterMes, TZ_FIXO.
 */

const BQ_PROJECT_FLUXO = "paraser-extrato-itau";
const BQ_TABELA_FLUXO = "`paraser-extrato-itau.extrato.lancamentos`";
const DRYRUN_KEY = "paraser-dryrun-2026";

const EMPRESAS_BQ = [
  { tag: "PARASER",   conta: "071500088498", aba: "LANÇAMENTOS" },
  { tag: "INSTITUTO", conta: "071500997912", aba: "LANÇAMENTOS_I" }
];

function lerExtratoBQ_(conta, dias) {
  const hoje = new Date();
  const ini = new Date(hoje.getTime() - dias * 86400000);
  const fmt = d => Utilities.formatDate(d, TZ_FIXO, "yyyy-MM-dd");

  const sql =
    "SELECT id, FORMAT_DATE('%d/%m/%Y', data_contabil) AS data, operation, valor, " +
    "descricao, descricao_completa, contraparte_nome " +
    "FROM " + BQ_TABELA_FLUXO + " " +
    "WHERE conta = @conta AND data_contabil BETWEEN @ini AND @fim " +
    "ORDER BY data_contabil, id";

  const req = {
    query: sql, useLegacySql: false, parameterMode: "NAMED", timeoutMs: 30000,
    queryParameters: [
      { name: "conta", parameterType: { type: "STRING" }, parameterValue: { value: conta } },
      { name: "ini",   parameterType: { type: "DATE" },   parameterValue: { value: fmt(ini) } },
      { name: "fim",   parameterType: { type: "DATE" },   parameterValue: { value: fmt(hoje) } }
    ]
  };

  let res = BigQuery.Jobs.query(req, BQ_PROJECT_FLUXO);
  const jobId = res.jobReference.jobId;
  const loc = res.jobReference.location;
  let t = 0;
  while (!res.jobComplete) {
    if (++t > 30) throw new Error("BigQuery: consulta não completou a tempo");
    Utilities.sleep(1000);
    res = BigQuery.Jobs.getQueryResults(BQ_PROJECT_FLUXO, jobId, { location: loc, timeoutMs: 30000 });
  }
  let rows = res.rows || [];
  let pageToken = res.pageToken;
  while (pageToken) {
    res = BigQuery.Jobs.getQueryResults(BQ_PROJECT_FLUXO, jobId, { location: loc, pageToken: pageToken });
    rows = rows.concat(res.rows || []);
    pageToken = res.pageToken;
  }
  return rows.map(r => ({
    id: r.f[0].v,
    data: r.f[1].v,
    operation: r.f[2].v,
    valor: Number(r.f[3].v),
    descricao: r.f[4].v || "",
    descricaoCompleta: r.f[5].v || "",
    contraparte: r.f[6].v || ""
  }));
}

function assinaturaLanc_(texto) {
  let s = normalizarDescricao(texto);
  s = s.replace(/\bAT\d+|\bDB\d+/g, " ");
  s = s.replace(/\d{2}\/\d{2}(\/\d{2,4})?/g, " ");
  s = s.replace(/\d{3}\.\d{3}\.\d{3}-\d{2}/g, " ");
  s = s.replace(/\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/g, " ");
  s = s.replace(/\d{4,}/g, " ");
  s = s.replace(/\s+/g, " ").trim();
  return s;
}

function ehTransferenciaNominal_(descNorm) {
  return /\bPIX\b|\bTED\b|\bDOC\b|\bTRANSF/.test(descNorm);
}

function classificacoesDaAba_(aba) {
  const regras = aba.getRange("C2:C").getDataValidations();
  if (!regras || !regras[0] || !regras[0][0]) return [];
  const crit = regras[0][0].getCriteriaValues();
  if (!crit || !crit[0]) return [];
  const v = crit[0];
  if (Array.isArray(v)) return v.filter(String);
  if (v && typeof v.getValues === "function") return v.getValues().flat().filter(String);
  return [];
}

const _top = o => { let mx = -1, r = ""; for (const k in o) if (o[k] > mx) { mx = o[k]; r = k; } return r; };
const _inc = (o, k) => { const v = String(k || "").trim(); if (v) o[v] = (o[v] || 0) + 1; };

function construirDeParaHistorico_(aba) {
  const ultima = aba.getRange("A:A").getValues().filter(String).length;
  if (ultima < 2) return {};
  const vals = aba.getRange(2, 1, ultima - 1, 14).getValues();
  const mapa = {};
  vals.forEach(l => {
    const desc = l[1], classif = l[2], forn = l[12], plano = l[13];
    if (!desc || !classif) return;
    const chave = assinaturaLanc_(desc);
    if (!chave) return;
    if (!mapa[chave]) mapa[chave] = { c: {}, f: {}, p: {} };
    _inc(mapa[chave].c, classif); _inc(mapa[chave].f, forn); _inc(mapa[chave].p, plano);
  });
  const out = {};
  for (const k in mapa) out[k] = { classificacao: _top(mapa[k].c), fornecedor: _top(mapa[k].f), plano: _top(mapa[k].p) };
  return out;
}

function construirDeParaFornecedor_(aba) {
  const ultima = aba.getRange("A:A").getValues().filter(String).length;
  if (ultima < 2) return [];
  const vals = aba.getRange(2, 1, ultima - 1, 14).getValues();
  const mapa = {};
  vals.forEach(l => {
    const classif = l[2], forn = l[12], plano = l[13];
    if (!classif || !forn) return;
    const key = normalizarDescricao(forn);
    if (key.length < 5) return;
    if (!mapa[key]) mapa[key] = { c: {}, p: {}, forn: forn };
    _inc(mapa[key].c, classif); _inc(mapa[key].p, plano);
  });
  return Object.keys(mapa)
    .map(k => ({ key: k, classificacao: _top(mapa[k].c), fornecedor: mapa[k].forn, plano: _top(mapa[k].p) }))
    .sort((a, b) => b.key.length - a.key.length);
}

function matchFornecedor_(contraparteNorm, listaForn) {
  if (!contraparteNorm || contraparteNorm.length < 5) return null;
  for (let i = 0; i < listaForn.length; i++) {
    const it = listaForn[i];
    if (contraparteNorm.indexOf(it.key) >= 0 || it.key.indexOf(contraparteNorm) >= 0) return it;
  }
  return null;
}

function regrasBanco_(descNorm, operation) {
  const ent = operation === "C";
  if (/^REDE\b/.test(descNorm))                        return { classificacao: "Clientes/REDE", fornecedor: "REDE", plano: "Receita Produtos e Serviços" };
  if (/REND PAGO APLIC/.test(descNorm))                return { classificacao: "Rendimento", fornecedor: "ITAÚ", plano: "Receita Rendimentos" };
  if (/INSTITUTO PARASER|PARASER SERV/.test(descNorm)) return { classificacao: "Parte Relacionada", fornecedor: "", plano: "Empréstimo" };
  if (/SISPAG TRIB|TRIBUTOS|\bDARF\b|\bISS\b|\bDAS\b|\bGPS\b|\bFGTS\b|\bISSQN\b|\bIOF\b/.test(descNorm)) return { classificacao: "Tributos", fornecedor: "", plano: "" };
  if (ent && /\bPIX\b/.test(descNorm))                 return { classificacao: "Clientes/PIX", fornecedor: "PACIENTES - GERAL", plano: "Receita Produtos e Serviços" };
  return null;
}

function classificarBQ_(lanc, dePara, listaForn, dropClass, dropForn, dropPlano) {
  const descNorm = normalizarDescricao(lanc.descricaoCompleta || lanc.descricao);
  const chave1 = assinaturaLanc_(lanc.descricao || lanc.descricaoCompleta);
  const chave2 = assinaturaLanc_(lanc.descricaoCompleta);
  let hit = dePara[chave1] || (chave2 && dePara[chave2]) || null;
  let origem = hit ? "histórico" : "";

  if (!hit && ehTransferenciaNominal_(descNorm)) {
    const m = matchFornecedor_(normalizarDescricao(lanc.contraparte), listaForn);
    if (m) { hit = { classificacao: m.classificacao, fornecedor: m.fornecedor, plano: m.plano }; origem = "favorecido"; }
  }

  let classificacao = "", fornecedor = "", plano = "";
  if (hit && hit.classificacao) {
    classificacao = hit.classificacao; fornecedor = hit.fornecedor || ""; plano = hit.plano || "";
  } else {
    const r = regrasBanco_(descNorm, lanc.operation);
    if (r) { classificacao = r.classificacao; fornecedor = r.fornecedor; plano = r.plano; origem = "regra"; }
  }

  fornecedor = normalizarDropdown(fornecedor, dropForn);
  plano = normalizarDropdown(plano, dropPlano);
  const classifOk = classificacao && (dropClass || [])
    .some(v => v.toString().trim().toUpperCase() === classificacao.toString().trim().toUpperCase());
  if (!classifOk) classificacao = "";
  if (!classificacao) origem = "pendente";

  return { classificacao, fornecedor, plano, origem };
}

// Lê os dropdowns e monta os de-para de uma aba (usado no teste e na escrita).
function _contextoAba_(ss, emp) {
  const cad = ss.getSheetByName("Cadastros");
  const dropForn = cad ? cad.getRange("H2:H").getValues().flat().filter(String) : [];
  const dropPlano = cad
    ? [...cad.getRange("A2:A").getValues().flat().filter(String), ...cad.getRange("B2:B").getValues().flat().filter(String)]
    : [];
  const aba = ss.getSheetByName(emp.aba);
  return {
    aba: aba, dropForn: dropForn, dropPlano: dropPlano,
    dropClass: aba ? classificacoesDaAba_(aba) : [],
    dePara: aba ? construirDeParaHistorico_(aba) : {},
    listaForn: aba ? construirDeParaFornecedor_(aba) : []
  };
}

function dryRunClassificar_(dias) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const res = [];
  EMPRESAS_BQ.forEach(emp => {
    const ctx = _contextoAba_(ss, emp);
    if (!ctx.aba) return;
    lerExtratoBQ_(emp.conta, dias).forEach(l => {
      const c = classificarBQ_(l, ctx.dePara, ctx.listaForn, ctx.dropClass, ctx.dropForn, ctx.dropPlano);
      res.push({
        empresa: emp.tag, data: l.data, desc: (l.descricaoCompleta || l.descricao),
        contraparte: l.contraparte, valor: l.valor, tipo: l.operation === "C" ? "Entrada" : "Saída",
        classificacao: c.classificacao, fornecedor: c.fornecedor, plano: c.plano, origem: c.origem
      });
    });
  });
  return res;
}

/**
 * === TESTE SEGURO (DRY-RUN) === Escreve o que classificaria numa aba "TESTE_BQ". NÃO toca nas abas reais.
 */
function testarClassificacaoBQ() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const res = dryRunClassificar_(25);
  let out = ss.getSheetByName("TESTE_BQ");
  if (out) out.clear(); else out = ss.insertSheet("TESTE_BQ");
  const linhas = [["Empresa", "Data", "Descrição (banco)", "Contraparte", "Valor", "Tipo", "Classificação", "Fornecedor", "Plano de Contas", "Origem"]];
  let classificados = 0;
  res.forEach(r => {
    linhas.push([r.empresa, r.data, r.desc.substring(0, 50), (r.contraparte || "").substring(0, 30),
      r.valor, r.tipo, r.classificacao, r.fornecedor, r.plano, r.origem]);
    if (r.classificacao) classificados++;
  });
  const pct = res.length ? Math.round(classificados / res.length * 100) : 0;
  linhas.push([]);
  linhas.push(["RESUMO", res.length + " lançamentos", classificados + " classificados (" + pct + "%)", (res.length - classificados) + " pendentes"]);
  out.getRange(1, 1, linhas.length, 10).setValues(linhas.map(r => { const x = r.slice(); while (x.length < 10) x.push(""); return x; }));
  out.getRange(1, 1, 1, 10).setFontWeight("bold");
  out.setFrozenRows(1);
  const msg = "TESTE_BQ: " + classificados + "/" + res.length + " (" + pct + "%) classificados.";
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert("✅ " + msg); } catch (e) {}
}

/**
 * === PRODUÇÃO: puxa do BQ e escreve nas abas reais ===
 * Só insere lançamentos com data DEPOIS do marco BQ_CORTE (o que a analista já fechou
 * fica intocado). Dedupe entre as linhas do próprio sistema (data+valor+descrição do
 * banco, que é consistente). Fórmulas de saldo copiadas da linha anterior; pendente = amarelo.
 */
function atualizarExtratosBQ() {
  const corteStr = PropertiesService.getScriptProperties().getProperty("BQ_CORTE");
  // Sem marco definido, o corte é HOJE: nunca insere retroativo, então mesmo rodando
  // "Puxar do banco" antes de ligar, é impossível duplicar o que a analista já lançou.
  const corte = corteStr ? new Date(corteStr + "T12:00:00-03:00") : new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let totalNovos = 0, totalPend = 0;
  const resumoEmp = [];

  EMPRESAS_BQ.forEach(emp => {
    const ctx = _contextoAba_(ss, emp);
    const aba = ctx.aba;
    if (!aba) return;

    let ultimaComData = aba.getRange("A:A").getValues().filter(String).length;
    if (ultimaComData < 2) ultimaComData = 1;
    const existentes = ultimaComData > 1 ? aba.getRange(2, 1, ultimaComData - 1, 14).getValues() : [];

    // dedupe só entre linhas DEPOIS do corte (as do próprio sistema). As da analista
    // (data <= corte) têm descrição diferente e ficam de fora do corte e do dedupe.
    const mapaEx = new Set();
    existentes.forEach(l => {
      const d = parseDataBR(l[0]);
      if (corte && d && d.getTime() <= corte.getTime()) return;
      const dk = dataParaChave(l[0]), vn = normalizarValor(l[3]), dn = normalizarDescricao(l[1]);
      if (dk && vn !== null) mapaEx.add(dk + "|" + vn + "|" + dn);
    });

    let novos = 0, pend = 0;
    lerExtratoBQ_(emp.conta, 40).forEach(l => {
      const dataConv = parseDataBR(l.data);
      if (!dataConv) return;
      if (corte && dataConv.getTime() <= corte.getTime()) return;

      const valorNum = normalizarValor(l.valor);
      if (valorNum === null) return;
      const chaveData = Utilities.formatDate(dataConv, TZ_FIXO, "yyyy-MM-dd");
      const descOriginal = l.descricaoCompleta || l.descricao || "";
      const chave = chaveData + "|" + valorNum + "|" + normalizarDescricao(descOriginal);
      if (mapaEx.has(chave)) return;

      const c = classificarBQ_(l, ctx.dePara, ctx.listaForn, ctx.dropClass, ctx.dropForn, ctx.dropPlano);

      ultimaComData++;
      const idx = ultimaComData;
      aba.insertRowAfter(idx - 1);
      gravarDataPura_(aba.getRange(idx, 1), dataConv); // dia puro: ver Extratos.js
      aba.getRange(idx, 2).setValue(descOriginal);
      aba.getRange(idx, 3).setValue(c.classificacao);
      aba.getRange(idx, 4).setValue(valorNum);
      aba.getRange(idx, 10).setValue(valorNum > 0 ? "Entrada" : "Saída");
      aba.getRange(idx, 11).setValue(obterMes(dataConv));
      aba.getRange(idx, 12).setValue(emp.tag);
      aba.getRange(idx, 13).setValue(c.fornecedor);
      aba.getRange(idx, 14).setValue(c.plano);
      if (idx > 2) {
        aba.getRange(idx - 1, 6, 1, 4).copyTo(aba.getRange(idx, 6, 1, 4), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
        const dvF = aba.getRange(idx - 1, 13).getDataValidation();
        if (dvF) aba.getRange(idx, 13).setDataValidation(dvF);
      }
      if (!c.classificacao) { aba.getRange(idx, 1, 1, 14).setBackground("#FFF9C4"); pend++; }
      else aba.getRange(idx, 1, 1, 14).setBackground(null);

      mapaEx.add(chave);
      novos++;
    });

    // Religa a cadeia do Saldo Acumulado. Cada insertRowAfter empurra pra baixo a 1ª linha
    // de buffer (vazia, logo abaixo do último lançamento) e, no empurrão, a referência
    // "I(linha de cima)" dela fica presa na linha ANTIGA: pula os lançamentos recém-inseridos
    // e trava um saldo defasado (era o que aparecia — o buffer mostrando o saldo de 2
    // lançamentos atrás). As linhas de buffer abaixo encadeiam nessa 1ª, então basta recopiar
    // a fórmula F:I do último lançamento pra ela que a coluna inteira volta a bater. Roda
    // sempre (mesmo com 0 novos) pra também corrigir um estado já defasado.
    const primBuffer = ultimaComData + 1;
    if (ultimaComData > 1 && primBuffer <= aba.getMaxRows() && aba.getRange(primBuffer, 1).getValue() === "") {
      aba.getRange(ultimaComData, 6, 1, 4).copyTo(aba.getRange(primBuffer, 6, 1, 4), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
    }

    totalNovos += novos; totalPend += pend;
    resumoEmp.push(emp.tag + ": " + novos + " novos" + (pend ? " (" + pend + " a conferir)" : ""));
  });

  const msg = totalNovos + " novos do banco. " + resumoEmp.join(" · ") + (corteStr ? " [após " + corteStr + "]" : "");
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert("✅ " + msg); } catch (e) {}
  return msg;
}

/**
 * Liga a automação: define o MARCO (só insere lançamentos depois de hoje) e cria o
 * gatilho de hora em hora. Rodar 1x. A analista fecha até hoje na mão; o sistema assume amanhã.
 */
function ligarExtratoBQ() {
  const props = PropertiesService.getScriptProperties();
  const hojeStr = Utilities.formatDate(new Date(), TZ_FIXO, "yyyy-MM-dd");
  props.setProperty("BQ_CORTE", hojeStr);
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "atualizarExtratosBQ") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("atualizarExtratosBQ").timeBased().everyHours(1).create();
  const msg = "Automação ligada. Marco = " + hojeStr + " (insere a partir de amanhã). Roda de hora em hora.";
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert("✅ " + msg); } catch (e) {}
  return msg;
}

function desligarExtratoBQ() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "atualizarExtratosBQ") { ScriptApp.deleteTrigger(t); n++; } });
  const msg = "Automação desligada (" + n + " gatilho[s] removido[s]).";
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert("✅ " + msg); } catch (e) {}
  return msg;
}

/**
 * Endpoint só de LEITURA pro refino (protegido por chave). Resumo + pendentes (sem valores).
 */
function doGet(e) {
  const out = ContentService.createTextOutput().setMimeType(ContentService.MimeType.JSON);
  if (!e || !e.parameter || e.parameter.key !== DRYRUN_KEY) return out.setContent(JSON.stringify({ error: "unauthorized" }));
  try {
    if (e.parameter.cmp) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const r = {};
      EMPRESAS_BQ.forEach(emp => {
        const aba = ss.getSheetByName(emp.aba);
        if (!aba) return;
        const ultima = aba.getRange("A:A").getValues().filter(String).length;
        const n = Math.min(8, ultima - 1);
        const linhas = n > 0 ? aba.getRange(ultima - n + 1, 1, n, 4).getValues()
          .map(l => dataParaChave(l[0]) + "|" + normalizarValor(l[3]) + "|" + normalizarDescricao(l[1])) : [];
        const bq = lerExtratoBQ_(emp.conta, 3)
          .map(l => { const d = parseDataBR(l.data); return (d ? Utilities.formatDate(d, TZ_FIXO, "yyyy-MM-dd") : "?") + "|" + normalizarValor(l.valor) + "|" + normalizarDescricao(l.descricaoCompleta || l.descricao); });
        r[emp.tag] = { abaUltimas: linhas, bqRecentes: bq };
      });
      return out.setContent(JSON.stringify(r));
    }
    const res = dryRunClassificar_(Number(e.parameter.dias) || 25);
    const porOrigem = {};
    res.forEach(r => { porOrigem[r.origem] = (porOrigem[r.origem] || 0) + 1; });
    const classificados = res.filter(r => r.classificacao).length;
    const pend = res.filter(r => !r.classificacao)
      .map(r => ({ e: r.empresa, t: r.tipo, d: r.desc.substring(0, 40), c: (r.contraparte || "").substring(0, 28) }));
    return out.setContent(JSON.stringify({
      tot: res.length, classificados: classificados, pct: res.length ? Math.round(classificados / res.length * 100) : 0,
      porOrigem: porOrigem, pendentes: pend.slice(0, 60)
    }));
  } catch (err) {
    return out.setContent(JSON.stringify({ error: String(err) }));
  }
}
