/**
 * ==============================================================================
 * MÓDULO: modClickUp.gs (VERSÃO 2.0 - RÁPIDO E SEGURO)
 * ==============================================================================
 * ✅ Sem filtro de unidade - puxa TUDO
 * ✅ Ignora apenas: SINISTRO, CANCELADO
 * ✅ Busca paralela otimizada
 * ✅ Proteção contra perda de dados
 * ✅ Data mínima: 1 de dezembro de 2025
 * ==============================================================================
 */

// === CONFIGURAÇÕES ===
const CLICKUP_TOKEN = "pk_87986690_9X1MC60UE18B1X9PEJFRMEFTT6GNHHFS"; 
const BASE_URL = "https://api.clickup.com/api/v2/list/";

// ✅ DATA MÍNIMA - 1 de Dezembro de 2025
const DATA_MINIMA_CLICKUP = new Date("2025-12-01T00:00:00").getTime();

// ✅ STATUS IGNORADOS (não puxa esses)
const STATUS_IGNORADOS = ["sinistro", "cancelado"];

// ✅ LISTAS DO CLICKUP
const CONFIG_LISTAS = {
  ENTREGAS: { 
    id: "901314444197", 
    nomeAba: "ENTREGAS"
  },
  MOTORISTAS: { 
    id: "901310964393", 
    nomeAba: "MOTORISTAS"
  }
};

const HEADER_PADRAO = [
  "ID", "Nome", "Status", "Status Cor", "URL", "Data de Criação", 
  "Data de Fechamento", "Data de Atualização", 
  "Prioridade", "Tempo Estimado (h)", "Tempo Gasto (h)", 
  "Tipo de Tarefa", "ID do Pai", "Checklists"
];

const REGEX_EMOJI = /([\u2700-\u27BF]|[\uE000-\uF8FF]|\uD83C[\uDC00-\uDFFF]|\uD83D[\uDC00-\uDFFF]|[\u2011-\u26FF]|\uD83E[\uDD10-\uDDFF]|\u200D|\uFE0F)/g;

// ============================================================================
// FUNÇÃO PRINCIPAL - SYNC RÁPIDO
// ============================================================================
function ExecutarIntegracaoMestre() {
  console.time("⏱️ TOTAL ClickUp");
  console.log("🚀 INICIANDO SYNC CLICKUP (v2.0 - Sem filtros de unidade)");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProps = PropertiesService.getScriptProperties();

  for (let chave in CONFIG_LISTAS) {
    const config = CONFIG_LISTAS[chave];
    console.time(`⏱️ ${config.nomeAba}`);
    
    try {
      sincronizarLista(ss, scriptProps, config);
    } catch (erro) {
      console.error(`❌ ERRO em ${config.nomeAba}: ${erro.message}`);
    }
    
    console.timeEnd(`⏱️ ${config.nomeAba}`);
  }

  console.timeEnd("⏱️ TOTAL ClickUp");
}

// ============================================================================
// SYNC DE UMA LISTA
// ============================================================================
function sincronizarLista(ss, scriptProps, config) {
  const { id: listId, nomeAba } = config;
  const lastTimeKey = `LAST_TIME_${nomeAba}`;
  
  // Timestamp de início
  let lastTime = scriptProps.getProperty(lastTimeKey);
  let timeStart = lastTime ? parseInt(lastTime) : DATA_MINIMA_CLICKUP;
  
  // Garantir que não busque antes de dezembro/2025
  if (timeStart < DATA_MINIMA_CLICKUP) {
    timeStart = DATA_MINIMA_CLICKUP;
  }
  
  const timeNow = Date.now();
  
  console.log(`📋 [${nomeAba}] ${nomeAba === "MOTORISTAS" ? "Buscando TUDO (sem filtro de data)" : "Buscando desde: " + new Date(timeStart).toLocaleDateString('pt-BR')}`);

  // 1. BUSCAR DADOS DA API
  const novosDados = buscarTarefasClickUp(listId, timeStart, nomeAba);
  
  if (novosDados.tarefas.length === 0) {
    console.log(`[${nomeAba}] ✅ Nenhuma alteração.`);
    scriptProps.setProperty(lastTimeKey, timeNow.toString());
    return;
  }
  
  console.log(`[${nomeAba}] 📥 ${novosDados.tarefas.length} tarefas encontradas`);

  // 2. CARREGAR HISTÓRICO DA PLANILHA
  let ws = ss.getSheetByName(nomeAba);
  let dadosAtuais = [];
  let linhasOriginais = 0;
  
  if (ws && ws.getLastRow() > 1) {
    const valores = ws.getDataRange().getValues();
    linhasOriginais = valores.length - 1;
    dadosAtuais = converterParaObjetos(valores);
    console.log(`[${nomeAba}] 📚 Histórico: ${dadosAtuais.length} registros`);
  }

  // 3. TRAVA DE SEGURANÇA - Leitura falhou
  if (linhasOriginais > 0 && dadosAtuais.length === 0) {
    throw new Error(`CRÍTICO: Falha ao ler planilha ${nomeAba}. Abortando.`);
  }

  // 4. MERGE (Histórico + Novos)
  const mapaGeral = new Map();
  
  // Adiciona histórico
  dadosAtuais.forEach(d => {
    if (d["ID"]) mapaGeral.set(String(d["ID"]), d);
  });
  
  // Adiciona/sobrescreve com novos
  novosDados.tarefas.forEach(d => {
    if (d["ID"]) mapaGeral.set(String(d["ID"]), d);
  });
  
  const listaFinal = Array.from(mapaGeral.values());
  
  console.log(`[${nomeAba}] 📊 Total após merge: ${listaFinal.length}`);

  // 5. TRAVA DE SEGURANÇA - Lista diminuiu muito
  if (linhasOriginais > 10 && listaFinal.length < linhasOriginais * 0.8) {
    throw new Error(`CRÍTICO: Lista diminuiu de ${linhasOriginais} para ${listaFinal.length}. Abortando.`);
  }

  // 6. SALVAR NA PLANILHA
  salvarNaPlanilha(ss, nomeAba, listaFinal, novosDados.campos, linhasOriginais);
  
  // 7. ATUALIZAR TIMESTAMP
  scriptProps.setProperty(lastTimeKey, timeNow.toString());
  
  console.log(`[${nomeAba}] ✅ Concluído! Total: ${listaFinal.length}`);
}

// ============================================================================
// BUSCA OTIMIZADA NA API
// ============================================================================
function buscarTarefasClickUp(listId, timeGT, nomeAba) {
  const tarefas = [];
  const campos = new Map();
  
  let page = 0;
  let temMais = true;
  
  // ✅ MOTORISTAS: Busca TUDO (sem filtro de data na URL)
  // ✅ ENTREGAS: Busca só alterações recentes
  const urlBase = `${BASE_URL}${listId}/task?archived=false&subtasks=true&include_closed=true`;
  
  while (temMais) {
    let url;
    if (nomeAba === "MOTORISTAS") {
      // Sem filtro de data - puxa TUDO
      url = `${urlBase}&page=${page}`;
    } else {
      // Com filtro de data
      url = `${urlBase}&page=${page}&date_updated_gt=${timeGT}`;
    }
    
    try {
      const response = UrlFetchApp.fetch(url, {
        headers: { "Authorization": CLICKUP_TOKEN },
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() !== 200) {
        console.warn(`⚠️ HTTP ${response.getResponseCode()} na página ${page}`);
        break;
      }
      
      const json = JSON.parse(response.getContentText());
      const tasks = json.tasks || [];
      
      if (tasks.length === 0) {
        temMais = false;
        break;
      }
      
      // Processar tarefas
      tasks.forEach(task => {
        const tarefa = processarTarefa(task, campos, nomeAba);
        if (tarefa) tarefas.push(tarefa);
      });
      
      page++;
      
      // Se veio menos de 100, é a última página
      if (tasks.length < 100) {
        temMais = false;
      }
      
      // Pequeno delay para não sobrecarregar a API
      if (temMais) Utilities.sleep(50);
      
    } catch (e) {
      console.error(`❌ Erro página ${page}: ${e.message}`);
      temMais = false;
    }
  }
  
  console.log(`   📄 Páginas processadas: ${page + 1}`);
  return { tarefas, campos };
}

// ============================================================================
// PROCESSAR UMA TAREFA
// ============================================================================
function processarTarefa(task, camposMap, nomeAba) {
  const statusAtual = task.status ? task.status.status.toLowerCase().trim() : "";
  const dataCriacao = task.date_created ? parseInt(task.date_created) : 0;
  
  // ✅ MOTORISTAS: SEM NENHUM FILTRO - puxa TUDO
  if (nomeAba === "MOTORISTAS") {
    // Não filtra nada, passa direto
  } 
  // ✅ ENTREGAS: Filtra apenas status e data
  else {
    // Filtro de status ignorados (só ENTREGAS)
    if (STATUS_IGNORADOS.includes(statusAtual)) {
      return null;
    }
    
    // Filtro de data de criação (só ENTREGAS)
    if (dataCriacao < DATA_MINIMA_CLICKUP) {
      return null;
    }
  }
  
  // Montar objeto
  const rowData = {
    "ID": task.id,
    "Nome": task.name,
    "Status": task.status ? task.status.status : "",
    "Status Cor": task.status ? task.status.color : "",
    "URL": task.url,
    "Data de Criação": msToDate(task.date_created),
    "Data de Fechamento": msToDate(task.date_closed),
    "Data de Atualização": msToDate(task.date_updated),
    "Prioridade": (task.priority && task.priority.priority) ? task.priority.priority : "N/A",
    "Tempo Estimado (h)": task.time_estimate ? (parseFloat(task.time_estimate) / 3600000) : "",
    "Tempo Gasto (h)": task.time_spent ? (parseFloat(task.time_spent) / 3600000) : "",
    "Tipo de Tarefa": task.parent ? "Subtask" : "Tarefa Principal",
    "ID do Pai": task.parent ? task.parent : "-"
  };
  
  // Checklists
  let strChecklist = "";
  if (task.checklists && task.checklists.length > 0) {
    const items = [];
    task.checklists.forEach(c => {
      if (c.items) c.items.forEach(i => items.push(i.name));
    });
    strChecklist = items.join(", ");
  }
  rowData["Checklists"] = strChecklist;
  
  // Custom Fields
  if (task.custom_fields) {
    task.custom_fields.forEach(cf => {
      if (cf.value !== undefined && cf.value !== null) {
        // Mapear nome do campo
        if (!camposMap.has(cf.id)) {
          let nomeLimpo = removerEmojis(cf.name);
          const u = nomeLimpo.toUpperCase();
          
          // Padronizar nomes comuns
          if (u.includes("MOTORISTA") && !u.includes("CPF") && !u.includes("AJUDANTE")) {
            nomeLimpo = "MOTORISTA";
          }
          if (u.includes("CONTATO") || (u.includes("CELULAR") && u.includes("MOTORISTA"))) {
            nomeLimpo = "CONTATO MOTORISTA";
          }
          if (u.includes("PLACA")) {
            nomeLimpo = "PLACA";
          }
          
          camposMap.set(cf.id, { name: nomeLimpo, type: cf.type });
        }
        
        rowData[camposMap.get(cf.id).name] = resolverCustomField(cf);
      }
    });
  }
  
  return rowData;
}

// ============================================================================
// SALVAR NA PLANILHA (MODO SEGURO)
// ============================================================================
function salvarNaPlanilha(ss, nomeAba, listaFinal, camposNovos, linhasOriginais) {
  let ws = ss.getSheetByName(nomeAba);
  
  // Criar aba se não existir
  if (!ws) {
    ws = ss.insertSheet(nomeAba);
    ws.getRange(1, 1, 1, HEADER_PADRAO.length).setValues([HEADER_PADRAO]).setFontWeight("bold");
  }
  
  // Montar headers
  let headers = [];
  if (ws.getLastColumn() > 0) {
    headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  } else {
    headers = [...HEADER_PADRAO];
    ws.getRange(1, 1, 1, HEADER_PADRAO.length).setValues([HEADER_PADRAO]);
  }
  
  const headerMap = {};
  headers.forEach((h, i) => headerMap[String(h).trim()] = i);
  
  // Adicionar colunas novas dos custom fields
  const colunasNovas = [];
  camposNovos.forEach((info, id) => {
    if (!headerMap.hasOwnProperty(info.name)) {
      colunasNovas.push(info.name);
      headerMap[info.name] = headers.length + colunasNovas.length - 1;
    }
  });
  
  // Adicionar colunas de dados que não estão no header
  if (listaFinal.length > 0) {
    Object.keys(listaFinal[0]).forEach(k => {
      if (!headerMap.hasOwnProperty(k) && !colunasNovas.includes(k)) {
        colunasNovas.push(k);
        headerMap[k] = headers.length + colunasNovas.length - 1;
      }
    });
  }
  
  // Escrever novas colunas no header
  if (colunasNovas.length > 0) {
    ws.getRange(1, headers.length + 1, 1, colunasNovas.length)
      .setValues([colunasNovas])
      .setFontWeight("bold");
    headers = [...headers, ...colunasNovas];
  }
  
  // Montar matriz de dados
  if (listaFinal.length === 0) {
    console.warn("Lista vazia - nada a salvar");
    return;
  }
  
  const matriz = listaFinal.map(item => {
    const row = new Array(headers.length).fill("");
    headers.forEach((h, i) => {
      const val = item[h];
      if (val !== undefined && val !== null) row[i] = val;
    });
    return row;
  });
  
  // Escrever dados
  const numCols = headers.length;
  const numRows = matriz.length;
  
  ws.getRange(2, 1, numRows, numCols).setValues(matriz);
  
  // Limpar linhas excedentes (com proteção)
  const totalLinhas = ws.getMaxRows();
  const linhasExcedentes = totalLinhas - numRows - 1;
  
  if (linhasExcedentes > 0) {
    const percentual = linhasExcedentes / Math.max(linhasOriginais, 1);
    
    if (percentual <= 0.15 || linhasOriginais === 0) {
      try {
        ws.getRange(numRows + 2, 1, linhasExcedentes, ws.getMaxColumns()).clearContent();
      } catch(e) {}
    } else {
      console.warn(`⚠️ ${linhasExcedentes} linhas órfãs (${Math.round(percentual*100)}%) - não limpando`);
    }
  }
  
  // Formatação
  ws.getRange(1, 1, 1, numCols).setFontWeight("bold").setBackground("#f3f3f3");
  ws.setFrozenRows(1);
  
  SpreadsheetApp.flush();
}

// ============================================================================
// HELPERS
// ============================================================================
function converterParaObjetos(values) {
  const headers = values[0];
  const result = [];
  
  // Encontrar coluna ID
  let idIndex = -1;
  for (let k = 0; k < headers.length; k++) {
    const h = String(headers[k]).trim().toUpperCase();
    if (h === "ID" || h === "TASK ID") {
      idIndex = k;
      break;
    }
  }
  
  for (let i = 1; i < values.length; i++) {
    const obj = {};
    let hasData = false;
    
    for (let j = 0; j < headers.length; j++) {
      const h = String(headers[j]).trim();
      if (h) {
        obj[h] = values[i][j];
        if (values[i][j] !== "") hasData = true;
      }
    }
    
    // Garantir que tem ID
    if (!obj["ID"] && idIndex > -1 && values[i][idIndex]) {
      obj["ID"] = values[i][idIndex];
    }
    
    if (hasData && obj["ID"]) {
      result.push(obj);
    }
  }
  
  return result;
}

function removerEmojis(texto) {
  if (!texto) return "";
  return texto.toString().replace(REGEX_EMOJI, '').replace(/\s+/g, ' ').trim();
}

function resolverCustomField(cf) {
  // Dropdown
  if (cf.type === "drop_down" || (cf.type_config && cf.type_config.options)) {
    const found = cf.type_config.options.find(o => o.orderindex == cf.value || o.id == cf.value);
    return found ? found.name : cf.value;
  }
  
  // Date
  if (cf.type === "date") {
    return msToDate(cf.value);
  }
  
  // Labels
  if (cf.type === "labels" && cf.value && cf.type_config) {
    const arr = [];
    cf.value.forEach(v => {
      const f = cf.type_config.options.find(o => o.id == v);
      if (f) arr.push(f.label);
    });
    return arr.join(", ");
  }
  
  // Numbers
  if (["currency", "number", "formula"].includes(cf.type)) {
    if (cf.value === null || cf.value === "") return "";
    const n = parseFloat(cf.value);
    return isNaN(n) ? cf.value : n;
  }
  
  return cf.value;
}

function msToDate(ms) {
  if (!ms || ms == 0) return "";
  return new Date(Number(ms));
}

// ============================================================================
// FUNÇÕES UTILITÁRIAS
// ============================================================================

/**
 * Força reset do timestamp para baixar tudo desde dezembro/2025
 */
function FORCAR_RESET_CLICKUP() {
  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.setProperty("LAST_TIME_ENTREGAS", DATA_MINIMA_CLICKUP.toString());
  scriptProps.setProperty("LAST_TIME_MOTORISTAS", DATA_MINIMA_CLICKUP.toString());
  console.log("✅ Reset concluído! Execute ExecutarIntegracaoMestre() agora.");
}

/**
 * Sync apenas ENTREGAS
 */
function SyncEntregas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProps = PropertiesService.getScriptProperties();
  sincronizarLista(ss, scriptProps, CONFIG_LISTAS.ENTREGAS);
}

/**
 * Sync apenas MOTORISTAS
 */
function SyncMotoristas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProps = PropertiesService.getScriptProperties();
  sincronizarLista(ss, scriptProps, CONFIG_LISTAS.MOTORISTAS);
}

/**
 * Envia status para ClickUp
 */
function enviarStatusParaClickup(taskId, novoStatus) {
  if (!taskId) return false;
  
  const url = `https://api.clickup.com/api/v2/task/${taskId}`;
  const payload = { "status": novoStatus };
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: "put",
      headers: { 
        "Authorization": CLICKUP_TOKEN, 
        "Content-Type": "application/json" 
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) return true;
    console.error("Erro ClickUp: " + response.getContentText());
    return false;
  } catch (e) {
    console.error("Erro Req: " + e.toString());
    return false;
  }
}

/**
 * Debug: Ver quantas tarefas tem na lista
 */
function DEBUG_ContarTarefas() {
  for (let chave in CONFIG_LISTAS) {
    const config = CONFIG_LISTAS[chave];
    console.log(`\n=== ${config.nomeAba} ===`);
    
    let total = 0;
    let page = 0;
    let temMais = true;
    
    while (temMais) {
      const url = `${BASE_URL}${config.id}/task?archived=false&subtasks=true&include_closed=true&page=${page}`;
      
      try {
        const response = UrlFetchApp.fetch(url, {
          headers: { "Authorization": CLICKUP_TOKEN },
          muteHttpExceptions: true
        });
        
        if (response.getResponseCode() !== 200) break;
        
        const json = JSON.parse(response.getContentText());
        const tasks = json.tasks || [];
        
        total += tasks.length;
        
        if (tasks.length < 100) temMais = false;
        else page++;
        
        Utilities.sleep(50);
      } catch (e) {
        temMais = false;
      }
    }
    
    console.log(`Total de tarefas: ${total}`);
  }
}
