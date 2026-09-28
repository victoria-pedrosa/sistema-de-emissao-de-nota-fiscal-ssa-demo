// =========================================================================
// CONFIGURAÇÕES DE IDS E PERMISSÕES
// =========================================================================
const ID_PLANILHA_SISTEMA = "ID_EXEMPLO";
const ID_PLANILHA_EXEMPLO = "ID_EXEMPLO";
const ID_PLANILHA_SENHAS = PropertiesService.getScriptProperties().getProperty('ID_PLANILHA_SENHAS');
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const GEMINI_MODEL_NAME = "gemini-3.6-flash";

const ID_PASTA_RAIZ_ROBO = "ID_EXEMPLO";
const ID_PASTA_NOTAS_EMITIDAS = "ID_EXEMPLO";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Exemplo NFS-e - Painel Corporativo')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function rotinaAutomata5Minutos() {
  sincronizarRetornosDoPython();
  exportarFilaParaExcel();
}

function sincronizacaoRapida() {
  sincronizarRetornosDoPython();
}

function configurarAcionadorSincronizacaoRapida() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sincronizacaoRapida') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('sincronizacaoRapida')
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log('Acionador de sincronizacaoRapida configurado para rodar a cada 1 minuto.');
}
function extrairNumeroNotaComIA(arquivo) {
  try {
    if (!GEMINI_API_KEY || GEMINI_API_KEY.includes("COLE_SUA_CHAVE")) return null;
    const blob = arquivo.getBlob();
    const mimeType = blob.getContentType();
    const dadosBase64 = Utilities.base64Encode(blob.getBytes());

    const instrucao = "Analise o PDF de nota fiscal de servico (NFS-e) anexado e identifique o numero da nota (tambem chamado de numero da NFS-e ou numero do RPS convertido), o numero sequencial oficial impresso no documento emitido pela prefeitura/Sefaz. NAO utilize horario, data, codigo de verificacao ou protocolo como numero da nota.\n" +
      "Responda SOMENTE com um objeto JSON valido (sem markdown, sem texto extra), no formato {\"numero_nota\": \"...\"}. Use string vazia se nao conseguir identificar com confianca. Nao invente numeros.";

    const payload = {
      contents: [{ role: "user", parts: [
        { inlineData: { mimeType: mimeType, data: dadosBase64 } },
        { text: instrucao }
      ] }],
      generationConfig: { responseMimeType: "application/json", temperature: 0.1 }
    };

    const url = "https://generativelanguage.googleapis.com/v1beta/models/" + GEMINI_MODEL_NAME + ":generateContent?key=" + GEMINI_API_KEY;
    const resp = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const codigoStatus = resp.getResponseCode();
    const respTexto = resp.getContentText();
    if (codigoStatus < 200 || codigoStatus >= 300) {
      Logger.log("extrairNumeroNotaComIA: erro na API do Gemini (" + codigoStatus + "): " + respTexto.substring(0, 300));
      return null;
    }

    const jsonResp = JSON.parse(respTexto);
    const cand = jsonResp.candidates && jsonResp.candidates[0];
    const textoResposta = cand && cand.content && cand.content.parts && cand.content.parts[0] ? cand.content.parts[0].text : "";
    if (!textoResposta) return null;

    let dadosIA;
    try {
      dadosIA = JSON.parse(textoResposta);
    } catch (parseErr) {
      const mm = textoResposta.match(/\{[\s\S]*\}/);
      if (!mm) return null;
      dadosIA = JSON.parse(mm[0]);
    }

    const numero = String(dadosIA.numero_nota || "").trim();
    return numero !== "" ? numero : null;
  } catch (e) {
    Logger.log("extrairNumeroNotaComIA: falha ao extrair numero da nota via IA: " + e.toString());
    return null;
  }
}


function construirMapaPrestadores() {
  try {
    const sheetPrest = SpreadsheetApp.openById(ID_PLANILHA_SISTEMA).getSheetByName("Prestadores");
    if (!sheetPrest) return {};
    const dadosPrest = sheetPrest.getDataRange().getValues();
    const mapa = {};
    for (let r = 1; r < dadosPrest.length; r++) {
      const cnpjNorm = String(dadosPrest[r][2] || "").replace(/[^0-9]/g, "");
      const razaoSocial = String(dadosPrest[r][1] || "").trim();
      if (cnpjNorm && razaoSocial) mapa[cnpjNorm] = razaoSocial;
    }
    return mapa;
  } catch (eMapa) {
    Logger.log("construirMapaPrestadores: falha ao carregar aba Prestadores: " + eMapa.toString());
    return {};
  }
}

function nomeArquivoConferePrestador(nomeArquivo, linhaDados, idxPrestCnpj, mapaPrestadores) {
  if (idxPrestCnpj === -1) return true;
  const cnpjNorm = String(linhaDados[idxPrestCnpj] || "").replace(/[^0-9]/g, "");
  if (!cnpjNorm) return true;
  const razaoSocial = mapaPrestadores ? mapaPrestadores[cnpjNorm] : null;
  if (!razaoSocial) return true;
  const normalizar = function(s) { return String(s || "").toUpperCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^A-Z0-9]/g, ""); };
  const nomeNorm = normalizar(nomeArquivo);
  const razaoNorm = normalizar(razaoSocial);
  return razaoNorm.length > 0 && nomeNorm.indexOf(razaoNorm) !== -1;
}


function sincronizarRetornosDoPython() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log("sincronizarRetornosDoPython: lock ocupado, execucao ignorada.");
    return { sucesso: false, erro: "Execucao concorrente em andamento." };
  }
  try {
  try {
    const pasta = DriveApp.getFolderById(ID_PASTA_NOTAS_EMITIDAS);
    const arquivos = pasta.getFiles();
    const sheet = SpreadsheetApp.openById(ID_PLANILHA_SISTEMA).getSheetByName("Solicitacoes");
    const dados = sheet.getDataRange().getValues();
    const cabecalhos = dados[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, '_'));

    const idxId = cabecalhos.indexOf("id");
    const idxStatus = cabecalhos.indexOf("status");
    let idxLink = cabecalhos.findIndex(h => h === "link_pdf" || h === "url_pdf" || h === "pdf_url");
    let idxEmitidaEm = cabecalhos.indexOf("emitida_em");
    let idxNumeroNota = cabecalhos.indexOf("numero_nota");
    let idxSocio = cabecalhos.indexOf("socio_solicitante");
    let idxPrestCnpj = cabecalhos.indexOf("prestador_cnpj");
    let idxPrestName = cabecalhos.indexOf("prestador_name");
    let idxCriadaPor = cabecalhos.indexOf("criada_por");
    let idxDiscr = cabecalhos.indexOf("discriminacao");
    let idxVlrTotal = cabecalhos.indexOf("valor_total");
  let idxObsErro = cabecalhos.indexOf("obs_erro");
  if (idxObsErro === -1) {
    const novaCol = cabecalhos.length + 1;
    sheet.getRange(1, novaCol).setValue("obs_erro");
    idxObsErro = novaCol - 1;
  }

    if (idxId === -1 || idxStatus === -1) return { sucesso: false, erro: "Colunas principais ausentes." };

    let notasAtualizadas = 0;
  const mapaPrestadores = construirMapaPrestadores();

    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      const nome = arquivo.getName();
      const jaProcessado = nome.startsWith("[PROCESSADO]");

    const matchP = nome.match(/-\s*P(\d+)\.pdf\s*$/i);
    const matchNF = nome.match(/NF[\s_-]*(?:GERADA[\s_-]*)?(\d+)/i);
    const matchSucesso = matchP || matchNF;
    // CORRECAO (27/07/2026): "-P<id>.pdf" traz o ID interno da fila (confiavel).
    // "NF <numero>" e o numero OFICIAL da nota (prefeitura/Sefaz) - numeracao independente do
    // ID interno, que pode coincidir por acaso com o ID de outra solicitacao pendente (foi o que
    // causou o vinculo errado na nota ID 17). Por isso, quando o numero vier somente do padrao
    // "NF <numero>", exigimos confirmacao extra pelo nome do prestador.
    const idVeioDoPadraoConfiavel = !!matchP;
      if (matchSucesso) {
        const idNota = matchSucesso[1];
        let precisaProcessar = false;
        for (let iChk = 1; iChk < dados.length; iChk++) {
          if (String(dados[iChk][idxId]).trim() === String(idNota)) {
          if (!idVeioDoPadraoConfiavel && !nomeArquivoConferePrestador(nome, dados[iChk], idxPrestCnpj, mapaPrestadores)) continue;
            const statusChk = String(dados[iChk][idxStatus]).trim().toUpperCase();
            if (statusChk === "EMITIR" || statusChk === "PROCESSANDO" || statusChk.includes("ERRO")) precisaProcessar = true;
            break;
          }
        }
        if (!(jaProcessado && !precisaProcessar)) {
        try { arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (eShare) { Logger.log("Aviso: nao foi possivel compartilhar o arquivo " + nome + ": " + eShare.toString()); }
        const urlPDF = arquivo.getUrl();

        for (let i = 1; i < dados.length; i++) {
          if (String(dados[i][idxId]).trim() === String(idNota)) {
        if (!idVeioDoPadraoConfiavel && !nomeArquivoConferePrestador(nome, dados[i], idxPrestCnpj, mapaPrestadores)) {
                Logger.log("sincronizarRetornosDoPython: arquivo '" + nome + "' corresponde ao ID " + idNota + " mas o nome do prestador nao confere; ignorando para evitar vinculo incorreto (possivel colisao com o numero oficial da nota).");
                break;
              }
            const statusAtual = String(dados[i][idxStatus]).trim().toUpperCase();
            if (statusAtual === "EMITIR" || statusAtual === "PROCESSANDO" || statusAtual.includes("ERRO")) {

              const dataAtualFormatada = Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd HH:mm");

              const numeroNotaReal = extrairNumeroNotaComIA(arquivo) || "Gerada Sefaz";
              sheet.getRange(i + 1, idxStatus + 1).setValue("EMITIDA");
              if (idxLink !== -1) sheet.getRange(i + 1, idxLink + 1).setValue(urlPDF);
              if (idxEmitidaEm !== -1) sheet.getRange(i + 1, idxEmitidaEm + 1).setValue(dataAtualFormatada);
              if (idxNumeroNota !== -1) sheet.getRange(i + 1, idxNumeroNota + 1).setValue(numeroNotaReal);
      if (idxObsErro !== -1) sheet.getRange(i + 1, idxObsErro + 1).setValue("");

              try {
                const ssEscr = SpreadsheetApp.openById(ID_PLANILHA_EXEMPLO);
                const sheetBaseEscr = ssEscr.getSheetByName("base");

                const pName = idxPrestName !== -1 ? dados[i][idxPrestName] : "";
                const socio = idxSocio !== -1 ? dados[i][idxSocio] : "";
                const pCnpj = idxPrestCnpj !== -1 ? dados[i][idxPrestCnpj] : "";
                const vlr = idxVlrTotal !== -1 ? dados[i][idxVlrTotal] : 0;
                const user = idxCriadaPor !== -1 ? dados[i][idxCriadaPor] : "";
                const discr = idxDiscr !== -1 ? dados[i][idxDiscr] : "";

                sheetBaseEscr.appendRow([idNota, dataAtualFormatada, pName, socio, pCnpj, numeroNotaReal, vlr, user, discr, urlPDF]);
              } catch(eEscr) {
                Logger.log("Erro ao alimentar aba base da planilha Exemplo: " + eEscr.toString());
              }

              notasAtualizadas++;
            }
            break;
          }
        }
        if (!nome.startsWith("[PROCESSADO]")) arquivo.setName("[PROCESSADO] " + nome);
        }
      }
      const matchErro = nome.match(/^erro_(\d+)\.txt$/i);
      if (matchErro) {
        const idNota = matchErro[1];
        const msgErro = arquivo.getBlob().getDataAsString();

        for (let i = 1; i < dados.length; i++) {
          if (String(dados[i][idxId]).trim() === String(idNota)) {
          const statusAtualErro = String(dados[i][idxStatus]).trim().toUpperCase();
          if (statusAtualErro !== "EMITIDA") {
            Logger.log("Robo reportou erro na nota ID " + idNota + ": " + msgErro.slice(0, 300));
            sheet.getRange(i + 1, idxStatus + 1).setValue("EMITIR");
            if (idxObsErro !== -1) sheet.getRange(i + 1, idxObsErro + 1).setValue(String(msgErro).slice(0, 500));
            notasAtualizadas++;
          } else {
            Logger.log("Arquivo de erro obsoleto ignorado para nota ID " + idNota + " (nota ja emitida com sucesso).");
          }
          break;
        }
        }
        arquivo.setTrashed(true);
      }
    }

    if (notasAtualizadas > 0) {
      exportarFilaParaExcel();
    }
    return { sucesso: true, atualizadas: notasAtualizadas };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
    }
  } finally {
    lock.releaseLock();
  }
}

function sincronizarPdfsDrive() {
  return sincronizarRetornosDoPython();
}

function exportarFilaParaExcel() {
  try {
    const ssSistema = SpreadsheetApp.openById(ID_PLANILHA_SISTEMA);
    const abaSol = ssSistema.getSheetByName("Solicitacoes");
    const dados = abaSol.getDataRange().getValues();

    let abaTemp = ssSistema.getSheetByName("TEMP_FILA_ROBO");
    if (!abaTemp) abaTemp = ssSistema.insertSheet("TEMP_FILA_ROBO");
    abaTemp.clear();

    const cabecalhos = dados[0];
    const idxStatus = cabecalhos.map(h => String(h).trim().toLowerCase()).indexOf("status");
    const idxPrestCnpj = cabecalhos.map(h => String(h).trim().toLowerCase()).indexOf("prestador_cnpj");

    const cabecalhosComSenha = cabecalhos.concat(["prestador_senha"]);
    const dadosFiltrados = [cabecalhosComSenha];

    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][idxStatus]).trim().toUpperCase() === "EMITIR") {
        const senha = idxPrestCnpj > -1 ? buscarSenhaNaPlanilhaDeSenhas(dados[i][idxPrestCnpj]) : "";
        dadosFiltrados.push(dados[i].concat([senha]));
      }
    }

    const pastaRaiz = DriveApp.getFolderById(ID_PASTA_RAIZ_ROBO);

    if (dadosFiltrados.length <= 1) {
      abaTemp.getRange(1, 1, 1, cabecalhosComSenha.length).setValues([cabecalhosComSenha]);
    } else {
      abaTemp.getRange(1, 1, dadosFiltrados.length, dadosFiltrados[0].length).setValues(dadosFiltrados);
    }

    SpreadsheetApp.flush();
    Utilities.sleep(3000);

    const urlExport = "https://docs.google.com/spreadsheets/d/" + ID_PLANILHA_SISTEMA + "/export?format=xlsx&gid=" + abaTemp.getSheetId() + "&time=" + new Date().getTime();
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(urlExport, { headers: { 'Authorization': 'Bearer ' + token } });

    const arquivosAntigos = pastaRaiz.getFilesByName("fila_emissao.xlsx");
    while(arquivosAntigos.hasNext()) { arquivosAntigos.next().setTrashed(true); }

    pastaRaiz.createFile(response.getBlob()).setName("fila_emissao.xlsx");

  } catch(e) { Logger.log("Erro Gerar XLSX: " + e.toString()); }
}

function buscarSenhaNaPlanilhaDeSenhas(cnpjAlvo) {
  try {
    const dados = SpreadsheetApp.openById(ID_PLANILHA_SENHAS).getSheets()[0].getDataRange().getValues();
    const cnpjLimpo = String(cnpjAlvo).replace(/\D/g, '');
    for (let i = 1; i < dados.length; i++) {
      for (let c = 0; c < dados[i].length; c++) {
        if (String(dados[i][c]).replace(/\D/g, '') === cnpjLimpo) return String(dados[i][c+1] || dados[i][c+2] || dados[i][c+3] || "").trim();
      }
    }
  } catch(e) {} return "";
}

function parseNumber(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === 'number') return val;
  let s = String(val).replace(/[R$\s]/g, '');
  if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
  return parseFloat(s) || 0;
}

function inicializarDadosDoSistema() {
  try {
    const ssSistema = SpreadsheetApp.openById(ID_PLANILHA_SISTEMA);
    const ssEscr = SpreadsheetApp.openById(ID_PLANILHA_EXEMPLO);

    const abaPrest = ssEscr.getSheetByName("bancodedados");
    const dadosPrest = abaPrest.getDataRange().getValues();
    dadosPrest.shift();
    const listaPrestadores = dadosPrest.map((linha, index) => ({
      id: index + 1, razao_social: linha[1] ? String(linha[1]).trim() : "", cnpj: inlineObs(linha[2]),
              observacao: inlineObs(linha[3]), dados_bancarios: linha[5] ? String(linha[5]).trim() : "", descricao_padrao: linha[6] ? String(linha[6]).trim() : "",
      cnae_padrao: inlineObs(linha[8]), aliquota_padrao: parseFloat(linha[9]) || 0,
      regime_tributario: linha[11] ? String(linha[11]).trim() : "", banco: linha[12] ? String(linha[12]).trim() : "",
      agencia: linha[13] ? String(linha[13]).trim() : "", conta: linha[14] ? String(linha[14]).trim() : "", chave_pix: linha[15] ? String(linha[15]).trim() : ""
    })).filter(p => p.razao_social !== "");

    const abaTom = ssSistema.getSheetByName("Tomadores");
    const dadosTom = abaTom.getDataRange().getValues();
    dadosTom.shift();
    const listaTomadores = dadosTom.map((linha, index) => ({
      id: index + 1, razao_social: linha[1] ? String(linha[1]).trim() : "", cnpj: inlineObs(linha[2]),
      email: inlineObs(linha[3]), regime_tributario: linha[4] ? String(linha[4]).trim() : "",
      observacao: inlineObs(linha[5]), banco: linha[6] ? String(linha[6]).trim() : "",
      agencia: linha[7] ? String(linha[7]).trim() : "", conta: linha[8] ? String(linha[8]).trim() : "", chave_pix: linha[9] ? String(linha[9]).trim() : ""
    })).filter(t => t.razao_social !== "");

    const abaUser = ssSistema.getSheetByName("Usuarios");
    const dadosUser = abaUser.getDataRange().getValues();
    dadosUser.shift();
    const listaUsuarios = dadosUser.map((linha, index) => ({
      id: index + 1, email: linha[0] ? String(linha[0]).trim().toLowerCase() : "", nome: linha[1] ? String(linha[1]).trim() : "",
      perfil: inlineObs(linha[2]), ativo: (linha[3] == 1 || String(linha[3]).trim() === "1") ? 1 : 0
    })).filter(u => u.email !== "");

    const abaSol = ssSistema.getSheetByName("Solicitacoes");
    const dadosSol = abaSol.getDataRange().getValues();
    const cabecalhosSol = dadosSol.shift();
    const cabecalhosNorm = cabecalhosSol.map(h => String(h).trim().toLowerCase().replace(/\s+/g, '_').replace(/[áàãâ]/g, 'a').replace(/[í]/g, 'i').replace(/[éê]/g, 'e').replace(/[óôõ]/g, 'o'));

    const listaSolicitacoes = dadosSol.map((linha) => {
      let obj = {};
      for(let i=0; i<cabecalhosSol.length; i++) { obj[cabecalhosSol[i]] = inlineObs(linha[i]); obj[cabecalhosNorm[i]] = linha[i]; }
      let dataStr = (obj["criada_em"] instanceof Date) ? Utilities.formatDate(obj["criada_em"], "GMT-3", "yyyy-MM-dd HH:mm") : (obj["criada_em"] || "");

      return {
        id: obj["id"] ? String(obj["id"]) : "",
        prestador_cnpj: obj["prestador_cnpj"] || obj["cnpj_prestador"] || "",
        cnpj_tomador: obj["tomador_cnpj"] || obj["cnpj_tomador"] || "",
        socio_solicitante: obj["socio_solicitante"] || obj["socio"] || "",
        valor_total: parseNumber(obj["valor_total"]),
        status: obj["status"] ? String(obj["status"]).trim().toUpperCase() : "PENDENTE",
        criada_em: dataStr,
        criada_por: obj["criada_por"] || obj["criada_por_nome"] || "",
        cnae: obj["cnae"] ? String(obj["cnae"]) : "",
        aliquota: obj["aliquota"] ? String(obj["aliquota"]) : "",
        iss_retido: obj["iss_retido"] ? String(obj["iss_retido"]) : "",
      pis_cofins_retido: obj["pis_cofins_retido"] ? String(obj["pis_cofins_retido"]) : "",
        data_competencia: obj["data_competencia"] ? String(obj["data_competencia"]) : "",
        codigo_tributario_nacional: obj["codigo_tributario_nacional"] ? String(obj["codigo_tributario_nacional"]) : "",
        ctiss: obj["ctiss"] ? String(obj["ctiss"]) : "",
        nbs: obj["nbs"] ? String(obj["nbs"]) : "",
        cclasstrib: obj["cclasstrib"] ? String(obj["cclasstrib"]) : "",
        cindop: obj["cindop"] ? String(obj["cindop"]) : "",
        discriminacao: obj["discriminacao"] ? String(obj["discriminacao"]) : "",
        inss: parseNumber(obj["inss"]),
        irrf: parseNumber(obj["irrf"]),
        csll: parseNumber(obj["csll"]),
        cofins: parseNumber(obj["cofins"]),
        pis: parseNumber(obj["pis"]),
        outras_retencoes: parseNumber(obj["outras_retencoes"]),
        valor_liquido: parseNumber(obj["valor_liquido"]),
        link_pdf: obj["link_pdf"] || obj["url_pdf"] || obj["pdf_url"] || ""
      };
    }).filter(s => s.id !== "");

    return { sucesso: true, prestadores: listaPrestadores, tomadores: listaTomadores, usuarios: listaUsuarios, solicitacoes: listaSolicitacoes };
  } catch (erro) { return { sucesso: false, erro: erro.toString() }; }
}

function inlineObs(val) { return val ? String(val).trim() : ""; }

function salvarNotaNaPlanilha(dados) {
  try {
    const ssSistema = SpreadsheetApp.openById(ID_PLANILHA_SISTEMA);
    const sheetSolSistema = ssSistema.getSheetByName("Solicitacoes");
    const timestamp = Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd HH:mm:ss");

    const dataSol = sheetSolSistema.getDataRange().getValues();
    const cabecalhos = dataSol[0];

    const atualizacoes = {
      "prestador_cnpj": dados.prestador_cnpj, "tomador_cnpj": dados.tomador_cnpj, "socio_solicitante": dados.socio_solicitante, "valor_total": dados.valor_total,
      "aliquota": dados.aliquota, "cnae": dados.cnae, "codigo_tributario_nacional": dados.codigo_tributario_nacional, "ctiss": dados.ctiss,
      "nbs": dados.nbs, "cclasstrib": dados.cclasstrib, "cindop": dados.cindop, "discriminacao": dados.discriminacao, "iss_retido": dados.iss_retido, "pis_cofins_retido": dados.pis_cofins_retido, "data_competencia": dados.data_competencia,
      "inss": dados.inss, "irrf": dados.irrf, "csll": dados.csll, "cofins": dados.cofins, "pis": dados.pis, "outras_retencoes": dados.outras_retencoes,
      "valor_liquido": dados.valor_liquido, "criada_por": dados.criada_por_nome, "criada_em": timestamp, "prestador_name": dados.prestador_name || ""
    };

    if (dados.id) {
      let rowIndex = -1;
      for (let i = 1; i < dataSol.length; i++) { if (String(dataSol[i][cabecalhos.indexOf("id")]) === String(dados.id)) { rowIndex = i + 1; break; } }
      if (rowIndex > -1) {
        atualizacoes["status"] = "AJUSTADA";
        cabecalhos.forEach((colNome, index) => {
          let colNorm = String(colNome).trim().toLowerCase().replace(/\s+/g, '_');
          if (atualizacoes[colNorm] !== undefined) sheetSolSistema.getRange(rowIndex, index + 1).setValue(atualizacoes[colNorm]);
        });
        return { sucesso: true, mensagem: "Solicitação AJUSTADA!", novoId: dados.id };
      } else { throw new Error("Nota não encontrada."); }
    } else {
      const novoId = sheetSolSistema.getLastRow() + 1;
      atualizacoes["id"] = novoId; atualizacoes["status"] = "PENDENTE";
      const novaLinha = cabecalhos.map(colNome => atualizacoes[String(colNome).trim().toLowerCase().replace(/\s+/g, '_')] !== undefined ? atualizacoes[String(colNome).trim().toLowerCase().replace(/\s+/g, '_')] : "");
      sheetSolSistema.appendRow(novaLinha);

      return { sucesso: true, mensagem: "Solicitação salva!", novoId: novoId };
    }
  } catch (erro) { return { sucesso: false, erro: erro.toString() }; }
}

function alterarStatusNota(idNota, novoStatus) {
  try {
    const abaSol = SpreadsheetApp.openById(ID_PLANILHA_SISTEMA).getSheetByName("Solicitacoes");
    const dados = abaSol.getDataRange().getValues();
    const idxId = dados[0].indexOf("id"); const idxStatus = dados[0].indexOf("status");

    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][idxId]) === String(idNota)) {
        abaSol.getRange(i + 1, idxStatus + 1).setValue(novoStatus);
        return { sucesso: true };
      }
    }
    return { sucesso: false, erro: "Nota não encontrada." };
  } catch (e) { return { sucesso: false, erro: e.toString() }; }
}

function excluirNotaDaPlanilha(idNota, emailUsuarioLogado) {
  try {
    const abaSol = SpreadsheetApp.openById(ID_PLANILHA_SISTEMA).getSheetByName("Solicitacoes");
    const dadosSol = abaSol.getDataRange().getValues();
    const idxId = dadosSol[0].indexOf("id");

    for (let i = 1; i < dadosSol.length; i++) {
      if (String(dadosSol[i][idxId]) === String(idNota)) {
        abaSol.deleteRow(i + 1);
        SpreadsheetApp.flush();
        Utilities.sleep(1000);
        exportarFilaParaExcel();
        return { sucesso: true, mensagem: "Excluída!" };
      }
    }
    return { sucesso: false, erro: "Não encontrada." };
  } catch (e) { return { sucesso: false, erro: e.toString() }; }
}

function salvarPrestadorNaPlanilha(p) {
  try {
    const sheet = SpreadsheetApp.openById(ID_PLANILHA_EXEMPLO).getSheetByName("bancodedados");
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for(let i=1; i<data.length; i++) { if(data[i][2] === p.cnpj) { rowIndex = i + 1; break; } }
    if(rowIndex > -1) {
      sheet.getRange(rowIndex, 2).setValue(p.razao_social);
      sheet.getRange(rowIndex, 7).setValue(p.descricao_padrao);
      sheet.getRange(rowIndex, 9).setValue(p.cnae_padrao);
      sheet.getRange(rowIndex, 10).setValue(p.aliquota_padrao);
      sheet.getRange(rowIndex, 12).setValue(p.regime_tributario);
      sheet.getRange(rowIndex, 13).setValue(p.banco);
      sheet.getRange(rowIndex, 14).setValue(p.agencia);
      sheet.getRange(rowIndex, 15).setValue(p.conta);
      sheet.getRange(rowIndex, 16).setValue(p.chave_pix);
    } else {
      sheet.appendRow(["ext_" + new Date().getTime(), p.razao_social, p.cnpj, "", "", "", p.descricao_padrao, "", p.cnae_padrao, p.aliquota_padrao, "", p.regime_tributario, p.banco, p.agencia, p.conta, p.chave_pix]);
    }
    return { sucesso: true };
  } catch(e) { return { sucesso: false, erro: e.toString() }; }
}

function sha256Hex(texto) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8);
  let hex = "";
  for (let i = 0; i < bytes.length; i++) { let b = bytes[i]; if (b < 0) b += 256; const h = b.toString(16); hex += (h.length === 1 ? "0" + h : h); }
  return hex;
}

function autenticar(email, senha) {
  email = String(email || "").trim().toLowerCase(); senha = String(senha || "");
  const dados = SpreadsheetApp.openById(ID_PLANILHA_SISTEMA).getSheetByName("Usuarios").getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).trim().toLowerCase() === email) {
      if (!(dados[i][3] == 1 || String(dados[i][3]).trim() === "1")) return { status: "inativo" };
      if (!dados[i][4]) return { status: "primeiro_acesso", nome: String(dados[i][1]).trim() };
      if (sha256Hex(senha) === String(dados[i][4]).trim()) return { status: "sucesso", nome: String(dados[i][1]).trim(), perfil: String(dados[i][2] || "suporte").trim().toLowerCase() };
      return { status: "senhaincorreta" };
    }
  } return { status: "nao_encontrado" };
}

function definirSenha(email, novaSenha, nomeUsuario) {
  email = String(email || "").trim().toLowerCase(); novaSenha = String(novaSenha || "");
  if (novaSenha.length < 6) throw new Error("Mínimo 6 caracteres.");
  const aba = SpreadsheetApp.openById(ID_PLANILHA_SISTEMA).getSheetByName("Usuarios");
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).trim().toLowerCase() === email) {
      aba.getRange(i + 1, 5).setValue(sha256Hex(novaSenha));
      if (String(nomeUsuario || "").trim() !== "") aba.getRange(i + 1, 2).setValue(String(nomeUsuario).trim());
      return { status: "ok" };
    }
  } throw new Error("Usuario não encontrado.");
}

function processarPedidoComIA(textoPedido, arquivoBase64, mimeType, nomeLogado, emailLogado) {
  try {
    if (!GEMINI_API_KEY || GEMINI_API_KEY.includes("COLE_SUA_CHAVE")) { throw new Error("A chave do Gemini não foi configurada."); }

    const parts = [];
    if (arquivoBase64 && mimeType) {
      parts.push({ inlineData: { mimeType: mimeType, data: arquivoBase64 } });
    }

    const instrucao = "Voce e um assistente que ajuda a preencher uma solicitacao de nota fiscal de servico (NFS-e) para uma empresa brasileira.\n" +
      "Analise o texto do pedido abaixo e, se houver um arquivo anexado (imagem, PDF, print de conversa, etc.), analise tambem o conteudo do arquivo para extrair as informacoes.\n" +
      "Responda SOMENTE com um objeto JSON valido (sem markdown, sem texto extra, sem \`\`\`) com as chaves abaixo. Use string vazia \"\" para qualquer campo que voce nao conseguir identificar com confianca. Nao invente dados.\n" +
      "Campos: prestador_name, prestador_cnpj, tomador_cnpj, socio_solicitante, valor_total (numero com ponto decimal, sem simbolo de moeda), discriminacao, data_competencia (yyyy-MM-dd), cnae, codigo_tributario_nacional, ctiss, nbs, cclasstrib, cindop, iss_retido, pis_cofins_retido, inss, irrf, csll, cofins, pis, outras_retencoes, observacoes.\n\n" +
      "Texto do pedido:\n" + String(textoPedido || "");

    parts.push({ text: instrucao });

    const payload = {
      contents: [{ role: "user", parts: parts }],
      generationConfig: { responseMimeType: "application/json", temperature: 0.2 }
    };

    const url = "https://generativelanguage.googleapis.com/v1beta/models/" + GEMINI_MODEL_NAME + ":generateContent?key=" + GEMINI_API_KEY;
    const resp = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const codigoStatus = resp.getResponseCode();
    const respTexto = resp.getContentText();
    if (codigoStatus < 200 || codigoStatus >= 300) {
      throw new Error("Erro na API do Gemini (" + codigoStatus + "): " + respTexto.substring(0, 300));
    }

    const jsonResp = JSON.parse(respTexto);
    const cand = jsonResp.candidates && jsonResp.candidates[0];
    const textoResposta = cand && cand.content && cand.content.parts && cand.content.parts[0] ? cand.content.parts[0].text : "";
    if (!textoResposta) throw new Error("A IA nao retornou nenhum dado. Tente detalhar melhor o pedido ou verifique o anexo.");

    let dadosIA;
    try {
      dadosIA = JSON.parse(textoResposta);
    } catch (parseErr) {
      const m = textoResposta.match(/\{[\s\S]*\}/);
      if (!m) throw new Error("Nao foi possivel interpretar a resposta da IA.");
      dadosIA = JSON.parse(m[0]);
    }

    const sheet = SpreadsheetApp.openById(ID_PLANILHA_SISTEMA).getSheetByName("Solicitacoes");
    const dados = sheet.getDataRange().getValues();
    const cabecalhos = dados[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, "_"));
    const idx = {};
    cabecalhos.forEach((h, i) => idx[h] = i);

    const novaLinha = new Array(cabecalhos.length).fill("");
    const setIf = (campo, valor) => { if (idx[campo] !== undefined && valor !== undefined && valor !== null && valor !== "") novaLinha[idx[campo]] = valor; };

    const novoId = calcularProximoId(dados, idx);
    const timestamp = Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd HH:mm:ss");

    setIf("id", novoId);
    setIf("prestador_name", dadosIA.prestador_name);
    setIf("prestador_cnpj", dadosIA.prestador_cnpj);
    setIf("tomador_cnpj", dadosIA.tomador_cnpj);
    setIf("socio_solicitante", dadosIA.socio_solicitante || nomeLogado);
    setIf("valor_total", dadosIA.valor_total);
    setIf("discriminacao", dadosIA.discriminacao);
    setIf("data_competencia", dadosIA.data_competencia);
    setIf("cnae", dadosIA.cnae);
    setIf("codigo_tributario_nacional", dadosIA.codigo_tributario_nacional);
    setIf("ctiss", dadosIA.ctiss);
    setIf("nbs", dadosIA.nbs);
    setIf("cclasstrib", dadosIA.cclasstrib);
    setIf("cindop", dadosIA.cindop);
    setIf("iss_retido", dadosIA.iss_retido);
    setIf("pis_cofins_retido", dadosIA.pis_cofins_retido);
    setIf("inss", dadosIA.inss);
    setIf("irrf", dadosIA.irrf);
    setIf("csll", dadosIA.csll);
    setIf("cofins", dadosIA.cofins);
    setIf("pis", dadosIA.pis);
    setIf("outras_retencoes", dadosIA.outras_retencoes);
    setIf("status", "PENDENTE_REVISAO");
    setIf("criada_por", "IA (" + (nomeLogado || emailLogado || "painel") + ")");
    setIf("criada_em", timestamp);
    setIf("prioridade", "Normal");
    setIf("log_emissao", "Rascunho criado pelo assistente de IA (" + GEMINI_MODEL_NAME + ") a partir do pedido descrito. Revise todos os campos antes de salvar." + (dadosIA.observacoes ? (" Obs. da IA: " + dadosIA.observacoes) : ""));

    sheet.appendRow(novaLinha);

    return { sucesso: true, novoId: novoId };
  } catch (e) { return { sucesso: false, erro: e.toString() }; }
}

// =========================================================================
// O SEGREDO DEFINITIVO: EXTRAÇÃO DE DADOS DE PDF DA NOTA SALVADOR
// =========================================================================
function processarPDFNotaSalvador(base64Data, fileName) {
  try {
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', fileName);
    var resource = { title: 'Temp_Extracao_Salvador_' + fileName };
    var tempFile = Drive.Files.insert(resource, blob, {ocr: true});
    var doc = DocumentApp.openById(tempFile.id);
    var texto = doc.getBody().getText();
    Drive.Files.remove(tempFile.id);

    var dados = {};

    dados.regimePrestador = "Lucro Presumido / Real";
    if (/optante[\s\S]{1,50}simples nacional/i.test(texto) && !/n[ãa]o optante[\s\S]{1,50}simples nacional/i.test(texto)) {
        dados.regimePrestador = "Simples Nacional";
    }

    // REGEX ATUALIZADO: Captura tanto CNPJ (14 dígitos) quanto CPF (11 dígitos)
    var regexDoc = /(?:\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2})|(?:\d{3}\.\d{3}\.\d{3}\-\d{2})/g;
    var docsEncontrados = texto.match(regexDoc) || [];
    dados.prestadorCNPJ = docsEncontrados[0] || "";
    dados.tomadorCNPJ = docsEncontrados.length > 1 ? docsEncontrados[1] : (docsEncontrados[0] || "");

    var matchValor = texto.match(/VALOR TOTAL DA NOTA[\s\S]{0,50}?R\$\s*([\d\.,]+)/i);
    if (matchValor) { dados.valorTotal = matchValor[1]; }

    var matchDiscr = texto.match(/DISCRIMINAÇÃO DOS SERVIÇOS([\s\S]*?)(?:VALOR TOTAL|CÓDIGO DO SERVIÇO|VALOR LÍQUIDO|INSCRIÇÃO MUNICIPAL|TRIBUTAÇÃO|RETENÇÕES)/i);
    if (matchDiscr) { dados.discriminacao = matchDiscr[1].replace(/^\s+|\s+$/g, '').trim(); }

    dados.aliquota = "";
    var matchAliq = texto.match(/Al[íi]quota[\s\S]{0,30}?([\d]{1,2}(?:[\.,][\d]{1,4})?)\s*%/i);
    if (matchAliq) {
        dados.aliquota = matchAliq[1].replace('.', ',');
    } else {
        var idxAliq = texto.toLowerCase().indexOf("alíquota");
        if (idxAliq === -1) idxAliq = texto.toLowerCase().indexOf("aliquota");
        if (idxAliq > -1) {
            var substr = texto.substring(idxAliq, idxAliq + 50);
            var numMatch = substr.match(/([\d]{1,2}(?:[\.,][\d]{1,2})?)/);
            if (numMatch) {
                var v = parseFloat(numMatch[1].replace(',', '.'));
                if (v > 0 && v <= 10) dados.aliquota = numMatch[1].replace('.', ',');
            }
        }
    }
    if (dados.aliquota && !dados.aliquota.includes(',')) dados.aliquota += ',00';

    var matchCNAE = texto.match(/CNAE[\s\S]{0,40}?([\d\.\-\/]+)/i);
    if (matchCNAE) { dados.cnae = matchCNAE[1]; }

    var textoLimpo = texto.replace(/2003/g, '    ');
    var matchItem = textoLimpo.match(/(?:Item da LC|C[óo]digo do Servi[çc]o)[\s\S]{0,30}?([\d]{1,2}[\.\-][\d]{1,2})/i);
    if (matchItem) { dados.itemLC = matchItem[1].replace('-', '.'); }

    dados.issRetido = "Não";
    var matchIssRet = texto.match(/Valor do ISS Retido[\s\S]{0,30}?([\d\.,]+)/i);
    if (matchIssRet && parseFloat(matchIssRet[1].replace(/\./g, '').replace(',', '.')) > 0) dados.issRetido = "Sim";
    if (/Exigibilidade do ISS[\s\S]*?Retido/i.test(texto)) dados.issRetido = "Sim";

  dados.pisCofinsRetido = "Não";
  var matchPisCofinsRet = texto.match(/Valor (?:do |de )?PIS\s*\/?\s*COFINS[\s\S]{0,10}?Retido[\s\S]{0,30}?([\d\.,]+)/i);
  if (matchPisCofinsRet && parseFloat(matchPisCofinsRet[1].replace(/\./g, '').replace(',', '.')) > 0) dados.pisCofinsRetido = "Sim";
  if (/PIS\s*\/?\s*COFINS[\s\S]{0,20}?Retido/i.test(texto)) dados.pisCofinsRetido = "Sim";

    // 9. RETENÇÕES FEDERAIS - A PROVA REAL MATEMÁTICA
    dados.pis = "0,00"; dados.cofins = "0,00"; dados.inss = "0,00"; dados.irrf = "0,00"; dados.csll = "0,00";
    var baseCalculo = dados.valorTotal ? parseFloat(dados.valorTotal.replace(/\./g, '').replace(',', '.')) : 0;

    if (baseCalculo > 0) {
        var allNums = texto.substring(texto.length / 2).match(/[\d]{1,3}(?:\.[\d]{3})*\,[\d]{2}/g) || [];
        for (var i = 0; i < allNums.length; i++) {
            var valStr = allNums[i];
            var valNum = parseFloat(valStr.replace(/\./g, '').replace(',', '.'));

            // Trava Suprema: O número do imposto NUNCA vai ser o valor da nota
            if (valNum === 0 || valNum >= baseCalculo) continue;

            var perc = valNum / baseCalculo;

            if (Math.abs(perc - 0.0065) <= 0.0001) dados.pis = valStr; // 0.65%
            else if (Math.abs(perc - 0.03) <= 0.0001) dados.cofins = valStr; // 3%
            else if (Math.abs(perc - 0.01) <= 0.0001) dados.csll = valStr; // 1%
            else if (Math.abs(perc - 0.015) <= 0.0001 || Math.abs(perc - 0.012) <= 0.0001) dados.irrf = valStr; // 1.5% ou 1.2%
            else if (Math.abs(perc - 0.11) <= 0.0001) dados.inss = valStr; // 11%
        }
    }

    return dados;

  } catch (e) {
    return { erro: e.toString() };
  }
}

/**
 * ============================================================
 * INTEGRACAO ZENDESK -> SOLICITACAO DE NOTA FISCAL (via IA)
 * ============================================================
 * Recebe POST do webhook do Zendesk (trigger com tag "nf_automatica_ia")
 * e cria uma nova linha na aba "Solicitacoes" com status "PENDENTE_REVISAO".
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ticketId = body.ticket_id || "";
    const requesterName = (body.requester_name || "").trim();
    const requesterEmail = (body.requester_email || "").trim();
    const organizationName = (body.organization_name || "").trim();
    const noteText = stripHtml(body.note_text || "");

    if (!noteText) {
      return respostaJson({ ok: false, erro: "note_text vazio - nada para processar." });
    }

    const sheet = SpreadsheetApp.openById(ID_PLANILHA_SISTEMA).getSheetByName("Solicitacoes");
    const dados = sheet.getDataRange().getValues();
    const cabecalhos = dados[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, '_'));

    const idx = {};
    cabecalhos.forEach((h, i) => idx[h] = i);

    const prestadorInfo = buscarPrestadorPorNomeOrganizacao(organizationName);
    const organizationCnpj = prestadorInfo ? prestadorInfo.cnpj : "";
    const validacaoPrestador = validarPrestador(requesterEmail, organizationCnpj);
    const prestadorCnpj = validacaoPrestador.cnpjFinal;

    const tomadorCnpj = extrairCNPJ(noteText);
    const valorTotal = extrairValorTotal(noteText);
    const socioSolicitante = extrairAssinatura(noteText) || requesterName || requesterEmail;

    const notaAnterior = (tomadorCnpj && prestadorCnpj) ? buscarUltimaNotaEmitida(dados, idx, prestadorCnpj, tomadorCnpj) : null;

    const novaLinha = new Array(cabecalhos.length).fill("");
    const setIf = (campo, valor) => { if (idx[campo] !== undefined) novaLinha[idx[campo]] = valor; };
    const novoId = calcularProximoId(dados, idx);

    setIf("id", novoId);
    setIf("prestador_cnpj", prestadorCnpj);
    setIf("tomador_cnpj", tomadorCnpj);
    setIf("socio_solicitante", socioSolicitante);
    setIf("valor_total", valorTotal);
    setIf("status", "PENDENTE_REVISAO");
    setIf("criada_por", "Automacao Zendesk (IA)");
    setIf("criada_em", new Date());
    setIf("prioridade", "Normal");

    let logMsg;
    if (notaAnterior) {
      ["cnae", "codigo_tributario_nacional", "ctiss", "nbs", "cclasstrib", "cindop", "discriminacao", "iss_retido", "data_competencia", "inss", "irrf", "csll",
       "cofins", "pis", "outras_retencoes", "desc_condicionado", "desc_incondicionado"
      ].forEach(campo => {
        if (idx[campo] !== undefined) novaLinha[idx[campo]] = notaAnterior[idx[campo]];
      });
      logMsg = "Criada via ticket Zendesk #" + ticketId + ". Dados reaproveitados da solicitacao id=" + notaAnterior[idx["id"]] + " (mesmo prestador e tomador, ultima nota EMITIDA). Aliquota precisa ser revisada manualmente.";
    } else {
      setIf("cnae", prestadorInfo ? prestadorInfo.cnae_padrao : "");
      setIf("discriminacao", montarDiscriminacao(noteText));
      setIf("iss_retido", extrairIssRetido(noteText) || (prestadorInfo ? prestadorInfo.iss_retido : ""));
      setIf("inss", extrairRetencao(noteText, "INSS"));
      setIf("irrf", extrairRetencao(noteText, "IRRF"));
      setIf("csll", extrairRetencao(noteText, "CSLL"));
      setIf("cofins", extrairRetencao(noteText, "COFINS"));
      setIf("pis", extrairRetencao(noteText, "PIS"));
      logMsg = "Criada via ticket Zendesk #" + ticketId + ". Sem nota anterior para este tomador - dados extraidos automaticamente do texto do ticket (CNAE/ISS padrao do prestador aplicado quando disponivel). Aliquota e retencoes precisam ser revisados manualmente.";
    }
    if (validacaoPrestador.aviso) {
      logMsg += " " + validacaoPrestador.aviso;
    }
    setIf("log_emissao", logMsg);

    sheet.appendRow(novaLinha);
    return respostaJson({ ok: true, id: novoId, tinhaNotaAnterior: !!notaAnterior, avisoPrestador: validacaoPrestador.aviso || null });

  } catch (err) {
    return respostaJson({ ok: false, erro: String(err) });
  }
}

function stripHtml(html) {
  return String(html)
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, ' ')
    .replace(/<(br|\/p|\/div|\/li)[^>]*>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/[ \t]+\n/g, '\n')
    .trim();
}

function respostaJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function limparCnpj(txt) {
  const digits = String(txt).replace(/\D/g, "");
  return digits.length === 14 ? digits : String(txt).trim();
}

function extrairCNPJ(texto) {
  const re = /\b\d{2}\.?\d{3}\.?\d{3}\/?\d{4}-?\d{2}\b/g;
  const encontrados = texto.match(re);
  if (!encontrados) return "";
  return encontrados[0].replace(/\D/g, "");
}

function extrairValorTotal(texto) {
  const linhas = texto.split(/\n/);
  for (const linha of linhas) {
    if (/valor.*total|total.*valor|valor\s+da\s+nota|valor\s+bruto/i.test(linha)) {
      const v = extrairPrimeiroValorMonetario(linha);
      if (v !== null) return v;
    }
  }
  return extrairPrimeiroValorMonetario(texto);
}

function extrairPrimeiroValorMonetario(texto) {
  const re = /R\$\s*([\d.,]+)/;
  const m = texto.match(re);
  if (!m) return null;
  return converterMoedaBrParaNumero(m[1]);
}

function converterMoedaBrParaNumero(str) {
  let s = str.trim();
  if (s.indexOf(',') > -1) s = s.replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function extrairAssinatura(texto) {
  const linhas = texto.split(/\n/).map(l => l.trim()).filter(Boolean);
  for (let i = linhas.length - 1; i >= 0 && i >= linhas.length - 5; i--) {
    const m = linhas[i].match(/(?:atenciosamente|abracos|cordialmente)[,:\s-]+(.+)/i);
    if (m && m[1]) return m[1].trim();
  }
  return "";
}

function extrairIssRetido(texto) {
  if (/iss\s+retido[:\s]*sim|reter[aa]?\s+iss|com\s+retencao\s+de\s+iss/i.test(texto)) return "Sim";
  if (/iss\s+retido[:\s]*n[aa]o|sem\s+retencao\s+de\s+iss/i.test(texto)) return "Nao";
  return "";
}

function extrairRetencao(texto, sigla) {
  const re = new RegExp(sigla + "[:\\s]+([\\d.,]+)\\s*%?", "i");
  const m = texto.match(re);
  if (!m) return "";
  const n = converterMoedaBrParaNumero(m[1]);
  return n === null ? "" : n;
}

function montarDiscriminacao(texto) {
  let base = texto.trim();
  if (base.length > 1500) base = base.substring(0, 1500) + "...";
  return base;
}

function buscarUltimaNotaEmitida(dados, idx, prestadorCnpj, tomadorCnpj) {
  const idxTomador = idx["tomador_cnpj"];
  const idxPrestador = idx["prestador_cnpj"];
  const idxStatus = idx["status"];
  if (idxTomador === undefined || idxPrestador === undefined) return null;
  const tomNorm = String(tomadorCnpj || "").replace(/\D/g, "");
  const prestNorm = String(prestadorCnpj || "").replace(/\D/g, "");
  if (!tomNorm || !prestNorm) return null;
  for (let i = dados.length - 1; i >= 1; i--) {
    const cnpjTomLinha = String(dados[i][idxTomador] || "").replace(/\D/g, "");
    const cnpjPrestLinha = String(dados[i][idxPrestador] || "").replace(/\D/g, "");
    const statusLinha = idxStatus !== undefined ? String(dados[i][idxStatus] || "").trim().toUpperCase() : "";
    if (cnpjTomLinha === tomNorm && cnpjPrestLinha === prestNorm && statusLinha === "EMITIDA") return dados[i];
  }
  return null;
}

function calcularProximoId(dados, idx) {
  const idxId = idx["id"];
  let max = 0;
  for (let i = 1; i < dados.length; i++) {
    const v = Number(dados[i][idxId]);
    if (!isNaN(v) && v > max) max = v;
  }
  return max + 1;
}

function buscarPrestadorPorNomeOrganizacao(organizationName) {
  const nomeNorm = String(organizationName).trim().toLowerCase();
  if (!nomeNorm) return null;
  const sheetPrestadores = SpreadsheetApp.openById(ID_PLANILHA_SISTEMA).getSheetByName("Prestadores");
  const dados = sheetPrestadores.getDataRange().getValues();
  const cabecalhos = dados[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, '_'));
  const idxRazao = cabecalhos.indexOf("razao_social");
  const idxCnpj = cabecalhos.indexOf("cnpj");
  const idxCnae = cabecalhos.indexOf("cnae_padrao");
  const idxIss = cabecalhos.indexOf("iss_retido");
  const idxDescPadrao = cabecalhos.indexOf("descricao_da_nota_padrao");
  if (idxRazao === -1) return null;

  for (let i = 1; i < dados.length; i++) {
    const razao = String(dados[i][idxRazao] || "").trim().toLowerCase();
    if (razao && (razao === nomeNorm || razao.includes(nomeNorm) || nomeNorm.includes(razao))) {
      return {
        cnpj: idxCnpj > -1 ? limparCnpj(dados[i][idxCnpj]) : "",
        cnae_padrao: idxCnae > -1 ? dados[i][idxCnae] : "",
        iss_retido: idxIss > -1 ? dados[i][idxIss] : "",
        descricao_padrao: idxDescPadrao > -1 ? dados[i][idxDescPadrao] : ""
      };
    }
  }
  return null;
}

function validarPrestador(requesterEmail, prestadorCnpjZendesk) {
  const emailNorm = String(requesterEmail).trim().toLowerCase();
  const sheetSocios = SpreadsheetApp.openById(ID_PLANILHA_SISTEMA).getSheetByName("Socios_Prestador");
  const dados = sheetSocios.getDataRange().getValues();
  const cabecalhos = dados[0].map(h => String(h).trim().toLowerCase());
  const idxEmail = cabecalhos.indexOf("email");
  const idxCnpj = cabecalhos.indexOf("prestador_cnpj");

  if (idxEmail === -1 || idxCnpj === -1 || !emailNorm) {
    return { cnpjFinal: prestadorCnpjZendesk, aviso: "" };
  }

  const cnpjsDoSocio = [];
  for (let i = 1; i < dados.length; i++) {
    const email = String(dados[i][idxEmail] || "").trim().toLowerCase();
    if (email && email === emailNorm) {
      const cnpj = limparCnpj(dados[i][idxCnpj]);
      if (cnpj && cnpjsDoSocio.indexOf(cnpj) === -1) cnpjsDoSocio.push(cnpj);
    }
  }

  let aviso = "";
  if (cnpjsDoSocio.length === 0) {
    aviso = "Solicitante nao encontrado na base Socios_Prestador. Prestador definido apenas pela organizacao do ticket no Zendesk - confirme manualmente.";
  } else if (cnpjsDoSocio.length > 1) {
    aviso = "ATENCAO: este solicitante esta vinculado a MAIS DE UMA empresa prestadora (" + cnpjsDoSocio.join(", ") + "). Empresa identificada automaticamente via organizacao do ticket: " + (prestadorCnpjZendesk || "(nenhuma)") + ". CONFIRME que e a empresa correta antes de aprovar a emissao.";
  } else if (prestadorCnpjZendesk && cnpjsDoSocio[0] !== prestadorCnpjZendesk) {
    aviso = "ATENCAO: o CNPJ da organizacao do ticket no Zendesk (" + prestadorCnpjZendesk + ") nao coincide com o CNPJ cadastrado para este socio (" + cnpjsDoSocio[0] + "). Revisar manualmente qual esta correto.";
  }

  const cnpjFinal = prestadorCnpjZendesk || cnpjsDoSocio[0] || "";
  return { cnpjFinal, aviso };
}


// ===== Integracao via Google Formulario (bypass de bloqueio de Web App publico) =====
const ID_FORMULARIO_INTAKE = 'ID_EXEMPLO';

function onFormSubmit(e) {
  try {
    const nv = {}; if (e && e.response && typeof e.response.getItemResponses === 'function') { e.response.getItemResponses().forEach(function(ir){ nv[ir.getItem().getTitle()] = ir.getResponse(); }); } else if (e && e.namedValues) { Object.keys(e.namedValues).forEach(function(k){ nv[k] = (e.namedValues[k] && e.namedValues[k][0]) ? e.namedValues[k][0] : ''; }); } const pick = function(nome) { return nv[nome] || ''; };
    const payload = {
      ticket_id: pick('ticket_id'),
      requester_name: pick('requester_name'),
      requester_email: pick('requester_email'),
      organization_name: pick('organization_name'),
      note_text: pick('note_text')
    };
    const fakeE = { postData: { contents: JSON.stringify(payload) } };
    doPost(fakeE);
  } catch (err) {
    Logger.log('Erro em onFormSubmit: ' + err);
  }
}

function setupFormTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(t);
    }
  });
  const form = FormApp.openById(ID_FORMULARIO_INTAKE);
  ScriptApp.newTrigger('onFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();
  Logger.log('Gatilho onFormSubmit criado com sucesso para o formulario: ' + form.getTitle());
}


function debugTestFormSubmit() {
  var fakeE = { postData: { contents: JSON.stringify({ticket_id:'85145', requester_name:'Cliente Exemplo', requester_email:'exemplo@cliente.com.br', organization_name:'', note_text:'teste debug conteudo texto'}) } };
  var result = doPost(fakeE);
  Logger.log(result.getContent());
}


function testeDiagnosticoOnFormSubmit() {
  var payload = {ticket_id:'85145', requester_name:'Cliente Exemplo', requester_email:'exemplo@cliente.com.br', organization_name:'', note_text:'Teste diagnostico automacao NF - verificar gravacao.'};
  var fakeE = { postData: { contents: JSON.stringify(payload) } };
  var res = doPost(fakeE);
  Logger.log(res.getContent());
}