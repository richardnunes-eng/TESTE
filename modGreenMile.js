/**
 * ==============================================================================
 * MÓDULO: modGreenMile.gs (VERSÃO CORRIGIDA - PROTEÇÃO + OTIMIZAÇÃO)
 * ==============================================================================
 * ✅ Filtro de data: Só processa rotas de dezembro/2025 em diante
 * ✅ Proteção contra perda de dados
 * ✅ Colunas extras para sequência e valor
 * ✅ Logs detalhados para debug
 * ==============================================================================
 */

const TAMANHO_LOTE = 120;

// ✅ DATA MÍNIMA - 1 de Dezembro de 2025
const DATA_MINIMA_GM = new Date("2025-12-01");

// ✅ COLUNAS COMPLETAS (adicionei plannedSequenceNum e plannedSize3)
const COLUNAS_GM = [
  "route.key", 
  "stop.plannedSequenceNum",      // ✅ Sequência da parada
  "stop.actualArrival", 
  "stop.actualDeparture", 
  "stop.hasSignature",                
  "stop.undeliverableCode.description", 
  "stop.deliveryStatus", 
  "stop.location.description", 
  "stop.location.addressLine1",
  "stop.location.district",
  "stop.actualSize1", 
  "stop.plannedSize1", 
  "stop.baseLineSize1", 
  "stop.actualSize2", 
  "stop.plannedSize2",
  "stop.plannedSize3",            // ✅ Valor da entrega
  "stop.stopType.type", 
  "stop.location.city",
  "stop.location.key",
];

function parseNumeroSeguro(valor) {
  if (valor === null || valor === undefined) return 0;
  if (typeof valor === "number") return valor;
  const normalizado = String(valor)
    .replace(/[^\d,.-]/g, "")
    .replace(/\.(?=.*\.)/g, "")
    .replace(",", ".");
  const parsed = parseFloat(normalizado);
  return Number.isNaN(parsed) ? 0 : parsed;
}

function sincronizarGreenMileStable() {
  console.time("⏱️ Tempo Total GreenMile");
  console.log("🚀 Iniciando Sincronização GreenMile (Modo Seguro)...");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsMain = ss.getSheetByName(SHEET_MAIN);
  let wsOut = ss.getSheetByName(SHEET_GM);

  if (!wsMain) { 
    console.error(`❌ ERRO CRÍTICO: Aba '${SHEET_MAIN}' não encontrada.`); 
    return; 
  }
  
  if (!wsOut) { 
    wsOut = ss.insertSheet(SHEET_GM); 
    console.log(`📋 Aba '${SHEET_GM}' criada.`);
  }

  // === 1. CARREGAR HISTÓRICO ATUAL ===
  let dadosHistorico = [];
  let rotasNoBanco = new Set();           
  let rotasPendentesNoBanco = new Set();
  let linhasOriginais = 0;

  const rangeGM = wsOut.getDataRange();
  const valuesGM = rangeGM.getValues();
  
  if (valuesGM.length > 1) {
    linhasOriginais = valuesGM.length - 1;
    const headersGM = valuesGM[0];
    const idxKeyGM = headersGM.findIndex(h => String(h).trim().toLowerCase() === "route.key");
    const idxDepGM = headersGM.findIndex(h => String(h).trim().toLowerCase() === "stop.actualdeparture");
    
    if (idxKeyGM !== -1) {
      dadosHistorico = valoresParaObjetos(valuesGM);
      console.log(`📚 Histórico carregado: ${dadosHistorico.length} registros`);
      
      if (idxDepGM !== -1) {
        for (let i = 1; i < valuesGM.length; i++) {
          let rKey = String(valuesGM[i][idxKeyGM]).trim();
          let departure = valuesGM[i][idxDepGM]; 

          if(rKey) rotasNoBanco.add(rKey);

          // Se não tem saída, está pendente
          if (rKey && (!departure || departure === "" || String(departure).trim() === "")) {
            rotasPendentesNoBanco.add(rKey);
          }
        }
      }
      console.log(`📊 Rotas no banco: ${rotasNoBanco.size} | Pendentes: ${rotasPendentesNoBanco.size}`);
    }
  }

  // === 2. CRUZAMENTO COM CLICKUP (ENTREGAS) ===
  const dataMain = wsMain.getDataRange().getValues();
  const headersMain = dataMain[0];
  
  const colNomeIdx = headersMain.findIndex(h => { 
    let t = String(h).toUpperCase().trim(); 
    return t === "NOME" || t === "PLANO" || t === "ROTA"; 
  });
  const colTipoIdx = headersMain.findIndex(h => String(h).toUpperCase().trim() === "TIPO DE TAREFA");
  
  // ✅ Coluna de data para filtrar
  const colDataIdx = headersMain.findIndex(h => { 
    let t = String(h).toUpperCase().trim(); 
    return t === "DATA DE CRIAÇÃO" || t === "DATA DE SAÍDA" || t === "DATA DE SAIDA"; 
  });

  if (colNomeIdx === -1) { 
    console.error("❌ ERRO: Coluna de Rota não encontrada."); 
    return; 
  }

  let rotasParaBaixar = new Set();
  let rotasIgnoradasPorData = 0;

  for (let i = 1; i < dataMain.length; i++) {
    // Filtro por tipo
    if (colTipoIdx !== -1) {
      let tipo = String(dataMain[i][colTipoIdx] || "").trim().toLowerCase();
      if (tipo !== "tarefa principal") continue; 
    }

    // ✅ FILTRO DE DATA - Ignora rotas anteriores a dezembro/2025
    if (colDataIdx !== -1) {
      let dataRota = dataMain[i][colDataIdx];
      if (dataRota instanceof Date) {
        if (dataRota < DATA_MINIMA_GM) {
          rotasIgnoradasPorData++;
          continue;
        }
      }
    }

    let val = String(dataMain[i][colNomeIdx] || "");
    if (val.length < 2) continue;
    
    let rotaKey = val.includes("-") ? val.split("-")[0].trim() : val;
    
    // FILTRO 610: Só processa se começar com 610
    if (!String(rotaKey).startsWith("610")) continue; 

    // LÓGICA DE DECISÃO
    if (!rotasNoBanco.has(rotaKey)) {
      rotasParaBaixar.add(rotaKey); // Nova
    } 
    else if (rotasPendentesNoBanco.has(rotaKey)) {
      rotasParaBaixar.add(rotaKey); // Atualização (Pendente)
    }
    // Se existe e não está pendente, IGNORA (Preserva histórico)
  }

  if (rotasIgnoradasPorData > 0) {
    console.log(`📅 Rotas ignoradas por data (antes de dez/2025): ${rotasIgnoradasPorData}`);
  }

  const listaDownload = Array.from(rotasParaBaixar);
  console.log(`📥 Rotas para baixar/atualizar: ${listaDownload.length}`);

  // === 3. DOWNLOAD ===
  let novosRegistros = [];
  let rotasBaixadasComSucesso = new Set();
  let errosDownload = 0;

  if (listaDownload.length > 0) {
    const token = obterTokenJWT();
    
    if (!token) {
      console.error("❌ ERRO: Falha ao obter token JWT. Abortando download.");
    } else {
      console.log("🔑 Token JWT obtido com sucesso.");
      
      for (let i = 0; i < listaDownload.length; i += TAMANHO_LOTE) {
        const loteAtual = listaDownload.slice(i, i + TAMANHO_LOTE);
        const requests = loteAtual.map(rota => prepararRequest(token, rota, COLUNAS_GM));

        try {
          const responses = UrlFetchApp.fetchAll(requests);
          
          responses.forEach((res, index) => {
            const rotaReferencia = loteAtual[index];
            
            if (res.getResponseCode() === 200) {
              try {
                let json = JSON.parse(res.getContentText());
                let items = json.content || json.rows || json.items || json;

                if (Array.isArray(items) && items.length > 0) {
                  rotasBaixadasComSucesso.add(rotaReferencia); 
                  
                  items.forEach(item => {
                    if (!item["route.key"]) item["route.key"] = rotaReferencia;
                    let flatItem = flattenObject(item, "");
                    
                    // Tratamento de ordersInfo
                    let rawInfo = flatItem["stop.ordersInfo"];
                    if (rawInfo) {
                      flatItem["stop.orders.number"] = String(rawInfo).replace(/[\[\]\"]/g, '');
                    } else {
                      flatItem["stop.orders.number"] = "";
                    }
                    
                    // Garantir que valores numéricos sejam números
                    const plannedSize3Raw = flatItem["stop.plannedSize3"] 
                      ?? flatItem["stop.plannedSize3.value"]
                      ?? flatItem["stop.plannedSize3.amount"];
                    flatItem["stop.plannedSize3"] = parseNumeroSeguro(plannedSize3Raw);
                    flatItem["stop.plannedSize1"] = parseNumeroSeguro(flatItem["stop.plannedSize1"] || 0);
                    flatItem["stop.plannedSequenceNum"] = parseInt(flatItem["stop.plannedSequenceNum"] || 0);
                    
                    novosRegistros.push(flatItem);
                  });
                }
              } catch (parseError) {
                console.warn(`⚠️ Erro ao parsear resposta da rota ${rotaReferencia}: ${parseError.message}`);
                errosDownload++;
              }
            } else {
              console.warn(`⚠️ HTTP ${res.getResponseCode()} para rota ${rotaReferencia}`);
              errosDownload++;
            }
          });
        } catch (fetchError) {
          console.error(`❌ Erro no lote ${i}: ${fetchError.message}`);
          errosDownload++;
        }
        
        // Delay entre lotes
        if (i + TAMANHO_LOTE < listaDownload.length) {
          Utilities.sleep(100);
        }
      }
    }
  }

  console.log(`✅ Download concluído: ${rotasBaixadasComSucesso.size} rotas | ${novosRegistros.length} registros | ${errosDownload} erros`);

  // === 4. MERGE ===
  let dadosFinais = dadosHistorico.filter(d => {
    let rKey = String(d["route.key"]).trim();
    // Mantém histórico SE a rota NÃO foi baixada com sucesso agora
    return !rotasBaixadasComSucesso.has(rKey);
  });

  dadosFinais = dadosFinais.concat(novosRegistros);

  // ✅ TRAVA DE SEGURANÇA 1: Lista final vazia
  if (dadosFinais.length === 0) {
    if (linhasOriginais > 0) {
      console.error("❌ CRÍTICO: Lista final vazia mas havia dados. Abortando para proteger dados.");
      return;
    }
    console.log("📭 Nenhum dado para salvar.");
    return;
  }

  // ✅ TRAVA DE SEGURANÇA 2: Lista diminuiu muito
  if (linhasOriginais > 0 && dadosFinais.length < linhasOriginais * 0.8) {
    console.error(`❌ CRÍTICO: Lista final (${dadosFinais.length}) é muito menor que original (${linhasOriginais}). Abortando.`);
    return;
  }

  // === 5. PREPARAR MATRIZ ===
  let headersSet = new Set();
  COLUNAS_GM.forEach(c => headersSet.add(c));
  dadosFinais.forEach(r => Object.keys(r).forEach(k => { 
    if(k !== "stop.ordersInfo") headersSet.add(k); 
  }));
  
  let headerArr = Array.from(headersSet);
  let matriz = [headerArr];

  dadosFinais.forEach(item => {
    let row = headerArr.map(h => {
      let val = item[h];
      // Converter strings ISO para Date
      if (typeof val === 'string' && val.match(/^\d{4}-\d{2}-\d{2}T/)) {
        return new Date(val);
      }
      return (val === undefined || val === null) ? "" : val;
    });
    matriz.push(row);
  });

  // === 6. ESCRITA SEGURA ===
  console.log(`💾 Salvando ${matriz.length - 1} registros...`);
  
  const numCols = matriz[0].length;
  const numRows = matriz.length;
  
  // Primeiro, escreve os dados novos
  wsOut.getRange(1, 1, numRows, numCols).setValues(matriz);

  // Depois, limpa o excesso (se houver) - COM PROTEÇÃO
  const totalLinhas = wsOut.getMaxRows();
  const linhasExcedentes = totalLinhas - numRows;

  if (linhasExcedentes > 0) {
    // ✅ Só limpa se for menos de 20% do total
    const percentualExcedente = linhasExcedentes / Math.max(linhasOriginais, 1);
    
    if (percentualExcedente <= 0.2 || linhasOriginais === 0) {
      try {
        wsOut.getRange(numRows + 1, 1, linhasExcedentes, wsOut.getMaxColumns()).clearContent();
        console.log(`🧹 Limpou ${linhasExcedentes} linhas excedentes.`);
      } catch(e) {
        // Ignora erro se não tiver o que limpar
      }
    } else {
      console.warn(`⚠️ Muitas linhas excedentes (${linhasExcedentes}). Não limpando por segurança.`);
    }
  }

  // Formatar cabeçalho
  wsOut.getRange(1, 1, 1, numCols).setFontWeight("bold").setBackground("#f3f3f3");
  wsOut.setFrozenRows(1);

  console.log(`✅ Sincronização GreenMile concluída! Total: ${matriz.length - 1} registros`);
  console.timeEnd("⏱️ Tempo Total GreenMile");
}

// ============================================================================
// HELPERS
// ============================================================================
function valoresParaObjetos(values) {
  let headers = values[0];
  let result = [];
  
  for (let i = 1; i < values.length; i++) {
    let obj = {};
    let hasData = false;
    
    for (let j = 0; j < headers.length; j++) {
      let h = String(headers[j]).trim();
      if(h) { 
        obj[h] = values[i][j]; 
        if(values[i][j] !== "") hasData = true; 
      }
    }
    
    if(hasData && obj["route.key"]) {
      result.push(obj);
    }
  }
  return result;
}

function prepararRequest(token, rotaKey, colunasDesejadas) {
  const criteriaUrlObj = { 
    "filters": colunasDesejadas, 
    "viewType": "STOP", 
    "firstResult": 0, 
    "maxResults": 1000 
  };
  
  const baseUrl = (typeof GM_URL_BASE !== 'undefined') 
    ? GM_URL_BASE 
    : "https://3coracoes.greenmile.com/greenmile-connect/rest/api/v1/fulfillment/routes";
    
  const urlFinal = `${baseUrl}?criteria=${encodeURIComponent(JSON.stringify(criteriaUrlObj))}`;
  
  const payloadBody = { 
    "criteriaChain": [{ 
      "and": [{ 
        "attr": "route.key", 
        "eq": rotaKey, 
        "matchMode": "EXACT" 
      }] 
    }], 
    "sort": [{ 
      "attr": "stop.plannedSequenceNum", 
      "type": "ASC" 
    }] 
  };
  
  return { 
    "url": urlFinal, 
    "method": "POST", 
    "contentType": "application/json;charset=UTF-8", 
    "headers": { 
      "Authorization": "Bearer " + token, 
      "Greenmile-Module": "LIVE", 
      "Accept": "application/json" 
    }, 
    "payload": JSON.stringify(payloadBody), 
    "muteHttpExceptions": true 
  };
}

function obterTokenJWT() {
  try {
    const user = (typeof GM_USERNAME !== 'undefined') ? GM_USERNAME : "";
    const pass = (typeof GM_PASSWORD !== 'undefined') ? GM_PASSWORD : "";
    
    if(!user || !pass) {
      console.error("❌ Credenciais GreenMile não configuradas (GM_USERNAME/GM_PASSWORD)");
      return null;
    }
    
    const resp = UrlFetchApp.fetch("https://3coracoes.greenmile.com/login", { 
      "method": "post", 
      "payload": `j_username=${encodeURIComponent(user)}&j_password=${encodeURIComponent(pass)}`, 
      "headers": {
        "Accept": "application/json", 
        "Greenmile-Module": "LIVE"
      }, 
      "muteHttpExceptions": true 
    });
    
    if (resp.getResponseCode() === 200) { 
      let j = JSON.parse(resp.getContentText()); 
      if (j.analyticsToken && j.analyticsToken.access_token) {
        return j.analyticsToken.access_token; 
      }
      return j.access_token; 
    } else {
      console.error(`❌ Erro login GreenMile: HTTP ${resp.getResponseCode()}`);
    }
  } catch (e) { 
    console.error("❌ Erro Auth GreenMile: " + e.message); 
  }
  return null;
}

function flattenObject(ob, prefix) {
  let toReturn = {};
  
  for (let i in ob) {
    if (!ob.hasOwnProperty(i)) continue;
    
    let key = prefix ? prefix + "." + i : i;
    
    if ((typeof ob[i]) === 'object' && ob[i] !== null) {
      if (ob[i] instanceof Date) {
        toReturn[key] = ob[i].toISOString();
      } else { 
        let f = flattenObject(ob[i], key); 
        for (let x in f) toReturn[x] = f[x]; 
      }
    } else { 
      toReturn[key] = ob[i]; 
    }
  }
  return toReturn;
}

// ============================================================================
// FUNÇÃO AUXILIAR - Forçar re-download de uma rota específica
// ============================================================================
function forcarDownloadRota(rotaKey) {
  console.log(`🔄 Forçando download da rota: ${rotaKey}`);
  
  const token = obterTokenJWT();
  if (!token) {
    console.error("❌ Falha ao obter token");
    return null;
  }
  
  const request = prepararRequest(token, rotaKey, COLUNAS_GM);
  
  try {
    const response = UrlFetchApp.fetch(request.url, {
      method: request.method,
      contentType: request.contentType,
      headers: request.headers,
      payload: request.payload,
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      const items = json.content || json.rows || json.items || json;
      console.log(`✅ Encontrados ${items.length} registros para rota ${rotaKey}`);
      return items;
    } else {
      console.error(`❌ HTTP ${response.getResponseCode()}`);
    }
  } catch (e) {
    console.error(`❌ Erro: ${e.message}`);
  }
  return null;
}
