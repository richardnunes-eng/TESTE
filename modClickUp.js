/**
 * ==============================================================================
 * MÓDULO: modClickUp.gs (VERSÃO 2.1 - CORRIGIDO)
 * ==============================================================================
 * ✅ Sem filtro de unidade - puxa TUDO
 * ✅ Ignora apenas: SINISTRO, CANCELADO
 * ✅ Busca paralela otimizada
 * ✅ Proteção contra perda de dados
 * ✅ Data mínima: 1 de dezembro de 2024 (corrigido)
 * ✅ Funções faltantes implementadas
 * ✅ Token seguro usando PropertiesService
 * ✅ Remove emojis e caracteres especiais dos cabeçalhos
 * ==============================================================================
 */

// === CONFIGURAÇÕES ===
// ⚠️ IMPORTANTE: Configure o token nas propriedades do script
// PropertiesService.getScriptProperties().setProperty("CLICKUP_TOKEN", "seu_token_aqui");
function getClickUpToken() {
  const token = PropertiesService.getScriptProperties().getProperty("CLICKUP_TOKEN");
  if (!token) {
    throw new Error("Token do ClickUp não configurado. Execute: PropertiesService.getScriptProperties().setProperty('CLICKUP_TOKEN', 'seu_token');");
  }
  return token;
}

const BASE_URL = "https://api.clickup.com/api/v2/list/";

// ✅ DATA MÍNIMA - 1 de Dezembro de 2024 (corrigido)
const DATA_MINIMA_CLICKUP = new Date("2024-12-01T00:00:00").getTime();
const SYNC_OVERLAP_MS = 10 * 60 * 1000; // overlap para evitar perda por fuso/latencia

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
  },
  OCORRENCIAS: {
    id: "901314625278",
    nomeAba: "OCORRENCIAS"
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
  console.log("🚀 INICIANDO SYNC CLICKUP (v2.1 - Corrigido)");

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
  const timeNow = Date.now();
  if (Number.isNaN(timeStart)) timeStart = DATA_MINIMA_CLICKUP;
  
  // Se o lastTime ficou no futuro, corrige para evitar buracos
  if (timeStart > timeNow) {
    timeStart = timeNow - SYNC_OVERLAP_MS;
  }

  // Garantir que não busque antes de dezembro/2024
  if (timeStart < DATA_MINIMA_CLICKUP) {
    timeStart = DATA_MINIMA_CLICKUP;
  }
  
  // Aplicar overlap para cobrir atrasos de atualizacao no ClickUp
  timeStart = Math.max(timeStart - SYNC_OVERLAP_MS, DATA_MINIMA_CLICKUP);

  console.log(`📋 [${nomeAba}] ${nomeAba === "MOTORISTAS" ? "Buscando TUDO (sem filtro de data)" : "Buscando desde: " + new Date(timeStart).toLocaleDateString('pt-BR')}`);

  // 1. CARREGAR HISTÓRICO DA PLANILHA
  let ws = ss.getSheetByName(nomeAba);
  let dadosAtuais = [];
  let linhasOriginais = 0;

  if (ws && ws.getLastRow() > 1) {
    const valores = ws.getDataRange().getValues();
    linhasOriginais = valores.length - 1;
    dadosAtuais = converterParaObjetos(valores);
    console.log(`[${nomeAba}] 📚 Histórico: ${dadosAtuais.length} registros`);
  }

  // 2. TRAVA DE SEGURANÇA - Leitura falhou
  if (linhasOriginais > 0 && dadosAtuais.length === 0) {
    throw new Error(`CRÍTICO: Falha ao ler planilha ${nomeAba}. Abortando.`);
  }

  // 3. BUSCAR DELTA (rápido)
  const novosDados = buscarTarefasClickUp(listId, timeStart, nomeAba);
  console.log(`[${nomeAba}] 📥 Delta: ${novosDados.tarefas.length} tarefas retornadas`);

  // 4. MERGE (Histórico + Delta)
  const { listaFinal: listaAposMerge, inseridos, atualizados } = aplicarDeltaNoHistorico(
    dadosAtuais,
    novosDados.tarefas,
    nomeAba
  );
  
  console.log(
    `[${nomeAba}] ➕ Inseridos: ${inseridos} | 🔁 Atualizados: ${atualizados} | 📊 Total após merge: ${listaAposMerge.length}`
  );

  // 5. RECONCILE (remoções / status ignorados / órfãs)
  const reconcileInfo = reconciliarListaComClickUp({
    listId,
    nomeAba,
    timeStart,
    timeNow,
    dadosAtuais,
    listaAposMerge,
    timeNowMs: timeNow,
    linhasOriginais,
    scriptProps
  });

  const listaFinal = reconcileInfo.listaFinal;

  console.log(
    `[${nomeAba}] 📊 Merge+Reconcile => total=${listaFinal.length} | inseridos=${inseridos} | atualizados=${atualizados} | removidos=${reconcileInfo.removidos} | removidosStatusIgnorado=${reconcileInfo.removidosPorStatusIgnorado}`
  );

  // 6. TRAVA DE SEGURANÇA - redução grande
  const reducao = linhasOriginais > 0 ? (linhasOriginais - listaFinal.length) / linhasOriginais : 0;
  if (linhasOriginais > 10 && reducao > 0.2) {
    if (!reconcileInfo.reconcileExecutado || !reconcileInfo.reconcileConfiavel) {
      throw new Error(
        `CRÍTICO: Lista diminuiu de ${linhasOriginais} para ${listaFinal.length} (${Math.round(
          reducao * 100
        )}%). Reconcile não foi confiável/executado. Abortando.`
      );
    }
    console.warn(
      `[${nomeAba}] ⚠️ Redução alta (${Math.round(
        reducao * 100
      )}%), porém validada por reconcile. Prosseguindo com escrita.`
    );
  }

  // 7. SALVAR NA PLANILHA
  if (listaFinal.length === 0 && linhasOriginais > 0 && !reconcileInfo.reconcileConfiavel) {
    throw new Error(`CRÍTICO: Lista final ficou vazia sem reconcile confiável. Abortando.`);
  }

  salvarNaPlanilha(ss, nomeAba, listaFinal, novosDados.campos, linhasOriginais);

  // 8. ATUALIZAR TIMESTAMP
  scriptProps.setProperty(lastTimeKey, timeNow.toString());

  if (novosDados.tarefas.length === 0 && reconcileInfo.removidos === 0) {
    console.log(`[${nomeAba}] ✅ Nenhuma alteração (delta vazio e reconcile sem remoções).`);
  } else {
    console.log(`[${nomeAba}] ✅ Concluído! Total: ${listaFinal.length}`);
  }
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
  // ✅ ENTREGAS/OCORRENCIAS: Busca só alterações recentes (delta)
  const urlBase = `${BASE_URL}${listId}/task?archived=false&subtasks=true&include_closed=true`;

  while (temMais) {
    let url;
    if (nomeAba === "MOTORISTAS") {
      // Sem filtro de data - puxa TUDO
      url = `${urlBase}&page=${page}`;
    } else {
      // Com filtro de data (delta)
      url = `${urlBase}&page=${page}&date_updated_gt=${timeGT}`;
    }

    try {
      const response = UrlFetchApp.fetch(url, {
        headers: { Authorization: getClickUpToken() },
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

  return { tarefas, campos };
}

// ============================================================================
// PROCESSAR TAREFA (FUNÇÃO FALTANTE IMPLEMENTADA)
// ============================================================================
function processarTarefa(task, campos, nomeAba) {
  // Filtrar status ignorados
  const status = task.status ? task.status.status || task.status : "";
  if (STATUS_IGNORADOS.includes(status.toLowerCase())) {
    return null;
  }

  const tarefa = {
    "ID": task.id,
    "Nome": removerEmojis(task.name || ""),
    "Status": status,
    "Status Cor": task.status ? task.status.color || "" : "",
    "URL": task.url || "",
    "Data de Criação": msToDate(task.date_created),
    "Data de Fechamento": msToDate(task.date_closed),
    "Data de Atualização": msToDate(task.date_updated),
    "Prioridade": task.priority ? task.priority.priority || "" : "",
    "Tempo Estimado (h)": task.time_estimate ? (task.time_estimate / 3600000).toFixed(2) : "",
    "Tempo Gasto (h)": task.time_spent ? (task.time_spent / 3600000).toFixed(2) : "",
    "Tipo de Tarefa": task.type ? task.type.name || "" : "",
    "ID do Pai": task.parent || "",
    "Checklists": processarChecklists(task.checklists)
  };

  // Processar custom fields
  if (task.custom_fields && task.custom_fields.length > 0) {
    task.custom_fields.forEach(cf => {
      if (cf.name && cf.name.trim()) {
        const nomeField = limparNomeColuna(cf.name.trim());
        if (nomeField) { // Só adiciona se o nome limpo não ficou vazio
          campos.set(nomeField, true);
          tarefa[nomeField] = resolverCustomField(cf);
        }
      }
    });
  }

  // Processar assignees
  if (task.assignees && task.assignees.length > 0) {
    campos.set("Responsáveis", true);
    tarefa["Responsáveis"] = task.assignees.map(a => a.username || a.email || "").join(", ");
  }

  // Processar tags
  if (task.tags && task.tags.length > 0) {
    campos.set("Tags", true);
    tarefa["Tags"] = task.tags.map(t => t.name || "").join(", ");
  }

  return tarefa;
}

// ============================================================================
// FUNÇÕES AUXILIARES FALTANTES
// ============================================================================

function msToDate(timestamp) {
  if (!timestamp) return "";
  try {
    const date = new Date(parseInt(timestamp));
    return date.toLocaleDateString('pt-BR') + " " + date.toLocaleTimeString('pt-BR');
  } catch (e) {
    return "";
  }
}

function processarChecklists(checklists) {
  if (!checklists || checklists.length === 0) return "";
  
  return checklists.map(checklist => {
    const items = checklist.items || [];
    const completed = items.filter(item => item.resolved).length;
    return `${checklist.name}: ${completed}/${items.length}`;
  }).join(" | ");
}

function aplicarDeltaNoHistorico(dadosAtuais, novasTarefas, nomeAba) {
  const mapa = new Map();
  let inseridos = 0;
  let atualizados = 0;

  // Adicionar dados atuais ao mapa
  dadosAtuais.forEach(item => {
    if (item.ID) {
      mapa.set(item.ID, item);
    }
  });

  // Aplicar novas tarefas
  novasTarefas.forEach(nova => {
    if (nova.ID) {
      if (mapa.has(nova.ID)) {
        atualizados++;
      } else {
        inseridos++;
      }
      mapa.set(nova.ID, nova);
    }
  });

  const listaFinal = Array.from(mapa.values());
  
  return { listaFinal, inseridos, atualizados };
}

function reconciliarListaComClickUp(params) {
  const { 
    listId, 
    nomeAba, 
    timeStart, 
    timeNow, 
    dadosAtuais, 
    listaAposMerge, 
    timeNowMs,
    linhasOriginais,
    scriptProps 
  } = params;

  let reconcileExecutado = false;
  let reconcileConfiavel = false;
  let removidos = 0;
  let removidosPorStatusIgnorado = 0;

  // Para MOTORISTAS, sempre faz reconcile completo
  // Para outras listas, faz reconcile periodicamente ou quando necessário
  const deveReconciliar = nomeAba === "MOTORISTAS" || 
                          Math.random() < 0.1 || // 10% das vezes
                          (listaAposMerge.length < dadosAtuais.length * 0.8); // se diminuiu muito

  if (deveReconciliar) {
    console.log(`[${nomeAba}] 🔄 Executando reconcile...`);
    
    try {
      // Buscar todas as tarefas atuais no ClickUp
      const dadosCompletos = buscarTarefasClickUp(listId, 0, nomeAba);
      const idsClickUp = new Set(dadosCompletos.tarefas.map(t => t.ID));
      
      // Filtrar apenas tarefas que ainda existem no ClickUp
      const listaReconciliada = listaAposMerge.filter(item => {
        if (!idsClickUp.has(item.ID)) {
          removidos++;
          return false;
        }
        return true;
      });

      reconcileExecutado = true;
      reconcileConfiavel = dadosCompletos.tarefas.length > 0;
      
      console.log(`[${nomeAba}] ✅ Reconcile: ${removidos} tarefas removidas`);
      
      return {
        listaFinal: listaReconciliada,
        reconcileExecutado,
        reconcileConfiavel,
        removidos,
        removidosPorStatusIgnorado
      };
      
    } catch (e) {
      console.warn(`[${nomeAba}] ⚠️ Reconcile falhou: ${e.message}`);
    }
  }

  return {
    listaFinal: listaAposMerge,
    reconcileExecutado,
    reconcileConfiavel,
    removidos,
    removidosPorStatusIgnorado
  };
}

function salvarNaPlanilha(ss, nomeAba, listaFinal, campos, linhasOriginais) {
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      let ws = ss.getSheetByName(nomeAba);
      
      // Criar aba se não existir
      if (!ws) {
        ws = ss.insertSheet(nomeAba);
        console.log(`[${nomeAba}] ➕ Aba criada`);
      }
      
      // Preparar headers
      let headers = [...HEADER_PADRAO];
      const camposAdicionais = Array.from(campos.keys()).filter(c => !headers.includes(c));
      headers = [...headers, ...camposAdicionais];
      
      // Limpar e configurar headers
      ws.clear();
      if (headers.length > 0) {
        ws.getRange(1, 1, 1, headers.length)
          .setValues([headers])
          .setFontWeight('bold')
          .setBackground('#f3f3f3');
      }
      
      // Preparar dados se houver
      if (listaFinal.length > 0) {
        const matriz = listaFinal.map(item => {
          return headers.map(h => item[h] || "");
        });
        
        // Escrever dados
        ws.getRange(2, 1, matriz.length, headers.length).setValues(matriz);
      }
      
      // Formatação final
      ws.setFrozenRows(1);
      SpreadsheetApp.flush();
      
      console.log(`[${nomeAba}] 💾 Salvo: ${listaFinal.length} registros`);
      return;
      
    } catch (e) {
      const msg = String(e && e.message ? e.message : e);
      if (attempt == 0 && /Sheet\s+\d+\s+not\s+found/i.test(msg)) {
        ss = SpreadsheetApp.openById(ss.getId());
        continue;
      }
      throw e;
    }
  }
}

// ============================================================================
// HELPERS
// ============================================================================
function converterParaObjetos(values) {
  if (!values || values.length < 2) return [];
  
  const headers = values[0].map(h => limparNomeColuna(h)); // Limpa headers também
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
  return texto.toString()
    .replace(REGEX_EMOJI, '') // Remove emojis
    .replace(/[^\w\s\-\(\)\[\]]/g, '') // Remove caracteres especiais exceto letras, números, espaços, hífens e parênteses
    .replace(/\s+/g, ' ') // Normaliza espaços
    .trim();
}

function limparNomeColuna(nome) {
  if (!nome) return "";
  return nome.toString()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // ✅ Remove acentos
    .replace(REGEX_EMOJI, '') // Remove emojis
    .replace(/[^\w\s\-\(\)\[\]\.]/g, '') // Permite também pontos
    .replace(/\s+/g, ' ') // Normaliza espaços
    .replace(/^\s+|\s+$/g, '') // Remove espaços das bordas
    .replace(/^[\d\-\.]+$/, 'Campo_' + nome) // Se for só números, adiciona prefixo
    .substring(0, 100); // Limita tamanho do cabeçalho
}

function resolverCustomField(cf) {
  if (!cf || cf.value === null || cf.value === undefined) return "";
  
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

// ============================================================================
// FUNÇÕES UTILITÁRIAS
// ============================================================================

/**
 * Configurar token do ClickUp (execute uma vez)
 */
function configurarToken() {
  const token = Browser.inputBox("Token ClickUp", "Cole seu token da API do ClickUp:", Browser.Buttons.OK_CANCEL);
  if (token !== "cancel" && token.trim()) {
    PropertiesService.getScriptProperties().setProperty("CLICKUP_TOKEN", token.trim());
    Browser.msgBox("✅ Token configurado com sucesso!");
  }
}

/**
 * Força reset do timestamp para baixar tudo desde dezembro/2024
 */
function FORCAR_RESET_CLICKUP() {
  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.setProperty("LAST_TIME_ENTREGAS", DATA_MINIMA_CLICKUP.toString());
  scriptProps.setProperty("LAST_TIME_MOTORISTAS", DATA_MINIMA_CLICKUP.toString());
  scriptProps.setProperty("LAST_TIME_OCORRENCIAS", DATA_MINIMA_CLICKUP.toString());
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
 * Sync apenas OCORRENCIAS
 */
function SyncOcorrencias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProps = PropertiesService.getScriptProperties();
  sincronizarLista(ss, scriptProps, CONFIG_LISTAS.OCORRENCIAS);
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
      method: "PUT",
      headers: { 
        "Authorization": getClickUpToken(), 
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
          headers: { "Authorization": getClickUpToken() },
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

/**
 * Testar conexão com ClickUp
 */
function testarConexaoClickUp() {
  try {
    const token = getClickUpToken();
    const url = "https://api.clickup.com/api/v2/user";
    
    const response = UrlFetchApp.fetch(url, {
      headers: { "Authorization": token },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      console.log(`✅ Conexão OK! Usuário: ${data.user.username}`);
      Browser.msgBox("✅ Conexão com ClickUp funcionando!");
    } else {
      console.error("❌ Erro na conexão: " + response.getContentText());
      Browser.msgBox("❌ Erro na conexão com ClickUp. Verifique o token.");
    }
  } catch (e) {
    console.error("❌ Erro: " + e.message);
    Browser.msgBox("❌ Erro: " + e.message);
  }
}

/**
 * Testar limpeza de nomes de colunas
 */
function testarLimpezaColunas() {
  const exemplos = [
    "📅 Data de Entrega",
    "🚚 Motorista Responsável",
    "⭐ Prioridade!!!",
    "🔥💯 Campo com Muitos Emojis 🎉✨",
    "Campo/Inválido",
    "Campo@#$%Com&Caracteres*Especiais",
    "123456", // só números
    "   Espaços nas Bordas   ",
    ""
  ];
  
  console.log("=== TESTE DE LIMPEZA DE COLUNAS ===");
  exemplos.forEach(exemplo => {
    const limpo = limparNomeColuna(exemplo);
    console.log(`"${exemplo}" → "${limpo}"`);
  });
}