/**
 * ==============================================================================
 * ARQUIVO: modDashboard.gs (VERSÃO CORRIGIDA - NFE E VALOR SIMPLIFICADOS)
 * ==============================================================================
 * ✅ CORREÇÃO: Chave de NFe = PLANO + CÓDIGO DO CLIENTE (simples)
 * ✅ CORREÇÃO: Valor = stop.plannedSize3 direto do GreenMile
 * ==============================================================================
 *
 * ==============================================================================
 * IMPORTANTE: AUTENTICAÇÃO E DEPLOY
 * ==============================================================================
 *
 * ⚠️ APPS SCRIPT WEB APP NÃO É EMBUTÍVEL COM OAUTH EM IFRAME ⚠️
 *
 * O Google Apps Script bloqueia OAuth em iframes por política de segurança
 * (frame-ancestors, X-Frame-Options). Isso NÃO pode ser removido via código.
 *
 * SOLUÇÃO IMPLEMENTADA:
 * 1. Detecção automática de iframe no front-end (JS-Logica.html)
 * 2. Tentativa de escape automático redirecionando window.top
 * 3. Fallback: abertura automática em nova aba
 * 4. Tela de erro com botão manual apenas se tudo falhar
 *
 * CONFIGURAÇÃO DE DEPLOY (appsscript.json):
 * {
 *   "webapp": {
 *     "executeAs": "USER_DEPLOYING",     // ✅ Executa como quem fez o deploy
 *     "access": "ANYONE_ANONYMOUS"        // ✅ Permite acesso sem login Google
 *   }
 * }
 *
 * NOTAS DE DEPLOY:
 * - Use "clasp push" para enviar código
 * - Deploy via: Apps Script Editor > Deploy > New Deployment > Web App
 * - Escolha: "Execute as: Me" e "Who has access: Anyone"
 * - Copie a URL do Web App e distribua aos usuários
 * - URL formato: https://script.google.com/macros/s/{SCRIPT_ID}/exec
 *
 * ALTERNATIVA PARA EMBED EM IFRAME:
 * - Hospedar front-end separado (Firebase/Vercel/Cloudflare Pages)
 * - Apps Script funciona apenas como API (doGet/doPost retornando JSON)
 * - Usar Google Identity Services (GIS) no front-end para autenticação
 * - Enviar token ao Apps Script para validação
 *
 * ==============================================================================
 */

// === CONFIGURAÇÃO DAS ABAS ===
const SHEET_MAIN = "ENTREGAS";
const SHEET_GM = "GreenMile";
const SHEET_MOT = "MOTORISTAS";

/**
 * Wrapper padrão para todas as respostas da API
 * Garante estrutura consistente: { ok, data, error, ts }
 */
function apiResponse(ok, data, error) {
  return {
    ok: Boolean(ok),
    data: data || null,
    error: error || null,
    ts: new Date().toISOString()
  };
}

function isAuthError(err) {
  const msg = String((err && err.message) || err || "");
  return /authorization|perm(iss|iss)?(ion)?|not authorized|insufficient|access denied|autoriz/i.test(msg);
}

function safeExecute(fn) {
  try {
    return { ok: true, value: fn() };
  } catch (err) {
    const needsAuth = isAuthError(err);
    const message = needsAuth
      ? "Autorize o app para continuar."
      : (err && err.message ? err.message : String(err));
    return { ok: false, needsAuth, message, error: err };
  }
}

function checkAuth() {
  const result = safeExecute(() => {
    PropertiesService.getScriptProperties().getProperty("AUTH_CHECK");
    const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();

    // Tentar obter email do usuário (pode falhar se política da organização bloquear)
    let userEmail = "authenticated";
    try {
      const user = Session.getEffectiveUser();
      if (user && user.getEmail) {
        userEmail = user.getEmail() || "authenticated";
      }
    } catch (e) {
      console.warn('Nao foi possivel obter email do usuario:', e);
    }

    return {
      ssId: ssId,
      userEmail: userEmail,
      env: 'production',
      version: '2.0'
    };
  });

  if (result.ok) {
    return apiResponse(true, result.value, null);
  }

  return apiResponse(false, null, {
    needsAuth: result.needsAuth,
    message: result.message
  });
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Dashboard');
  return template.evaluate()
      .setTitle('THX LOG | Centro de Comando')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// === RELATORIOS EM LOOP (TRIGGER A CADA 1 MINUTO) ===
const RELATORIOS_LOOP_HANDLER = "runRelatoriosLoop";

function startRelatoriosLoop() {
  stopRelatoriosLoop();
  ScriptApp.newTrigger(RELATORIOS_LOOP_HANDLER)
    .timeBased()
    .everyMinutes(1)
    .create();
  return apiResponse(true, { started: true }, null);
}

function stopRelatoriosLoop() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === RELATORIOS_LOOP_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  return apiResponse(true, { stopped: true }, null);
}

function runRelatoriosLoop() {
  const tasks = [
    { name: "dashboard", run: () => getDashboardData("force") }
  ];

  tasks.forEach((task) => {
    const result = safeExecute(task.run);
    if (!result.ok) {
      console.warn(`Relatorio ${task.name} falhou: ${result.message}`);
    }
  });
}

/**
 * FUNÇÃO PRINCIPAL - RETORNA DADOS DO DASHBOARD
 */
function getDashboardData(modo) {
  const authProbe = safeExecute(() => SpreadsheetApp.getActiveSpreadsheet().getId());
  if (!authProbe.ok) {
    return apiResponse(false, {
      drivers: [],
      stats: { total: 0, emRota: 0, finalizados: 0, criticos: 0 }
    }, {
      message: authProbe.message,
      needsAuth: authProbe.needsAuth
    });
  }

  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = "payload_dashboard_v6";  // ✅ Nova versão do cache
    const CACHE_DURATION = 600;
    
    // Fast path com cache
    if (modo !== 'force') {
      const cachedData = cache.get(cacheKey);
      if (cachedData) {
        console.log("⚡ Cache hit");
        return JSON.parse(cachedData);
      }
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();

    // Sincronização GreenMile (apenas em force ou check)
    if (modo === 'force' || modo === 'check') {
      try {
        const lastSync = parseInt(props.getProperty("LAST_AUTO_SYNC") || "0");
        const now = new Date().getTime();
        const INTERVALO = 5 * 60 * 1000;

        if (modo === 'force' || (now - lastSync > INTERVALO)) {
          if (typeof sincronizarGreenMileStable === 'function') {
            sincronizarGreenMileStable();
            props.setProperty("LAST_AUTO_SYNC", now.toString());
          }
        }
      } catch (e) { 
        console.warn("Erro Sync GreenMile: " + e.message); 
      }
    }

    // Leitura das abas
    const wsMain = ss.getSheetByName(SHEET_MAIN);
    const wsGM = ss.getSheetByName(SHEET_GM);
    const wsMot = ss.getSheetByName(SHEET_MOT);

    if (!wsMain) {
      return apiResponse(false, { drivers: [], stats: {} }, "Aba '" + SHEET_MAIN + "' não encontrada.");
    }
    if (!wsGM) {
      return apiResponse(false, { drivers: [], stats: {} }, "Aba '" + SHEET_GM + "' não encontrada.");
    }
    if (!wsMot) {
      return apiResponse(false, { drivers: [], stats: {} }, "Aba '" + SHEET_MOT + "' não encontrada.");
    }

    // Leitura em bulk
    const dataMainRaw = wsMain.getDataRange().getValues();
    const dataMainDisplay = wsMain.getDataRange().getDisplayValues();
    const dataGM = wsGM.getDataRange().getValues();
    const dataMot = wsMot.getDataRange().getValues();

    // Verificar se tem dados
    if (dataMainRaw.length < 2) {
      return apiResponse(false, { drivers: [], stats: {} }, "Aba ENTREGAS está vazia.");
    }

    const colMain = mapDashboardCols(dataMainRaw[0], 'MAIN');
    const colGM = mapDashboardCols(dataGM[0], 'GM');
    const colMot = mapDashboardCols(dataMot[0], 'MOT');

    if (colMain.PLANO === -1) {
      return apiResponse(false, { drivers: [], stats: {} }, "Coluna PLANO não encontrada em ENTREGAS.");
    }

    // ========================================================================
    // ✅ MAPEAMENTO NFE SIMPLIFICADO: PLANO|CLIENTE => NFE
    // ========================================================================
    const mapNfe = new Map();

    if (colMain.PLANO !== -1 && colMain.ID_GM_LOC !== -1 && colMain.CHECKLISTS !== -1) {
      for (let i = 1; i < dataMainDisplay.length; i++) {
        // Pegar o plano (route.key) - normalizado
        const planoFull = String(dataMainDisplay[i][colMain.PLANO] || "").trim();
        const planoChave = planoFull.split('-')[0].replace(/[^a-zA-Z0-9]/g, "").toUpperCase();
        
        // Código do cliente (location.key)
        const clientId = String(dataMainDisplay[i][colMain.ID_GM_LOC] || "").trim();
        
        // Número da NFe
        const nfe = String(dataMainDisplay[i][colMain.CHECKLISTS] || "").trim();
        
        // ✅ CHAVE SIMPLES: PLANO|CLIENTE
        if (planoChave && clientId && nfe && nfe !== "" && nfe !== "---" && nfe !== "-") {
          const chave = `${planoChave}|${clientId}`;
          
          // Se já existe, concatena
          const existente = mapNfe.get(chave);
          if (existente) {
            if (!existente.includes(nfe)) {
              mapNfe.set(chave, `${existente}, ${nfe}`);
            }
          } else {
            mapNfe.set(chave, nfe);
          }
        }
      }
    }
    
    console.log(`📝 Total de NFes mapeadas: ${mapNfe.size}`);

    // ========================================================================
    // ✅ SUBTASKS (REENTREGAS) - AGRUPAR POR TAREFA PAI
    // ========================================================================
    const mapSubtasks = new Map();

    if (colMain.TIPO !== -1 && colMain.ID_PAI !== -1) {
      for (let i = 1; i < dataMainRaw.length; i++) {
        const tipo = String(dataMainRaw[i][colMain.TIPO] || "").toLowerCase();
        if (!tipo.includes("subtask")) continue;

        const parentId = String(dataMainRaw[i][colMain.ID_PAI] || "").trim();
        if (!parentId) continue;

        const nomeSubtask = String(
          (colMain.NOME !== -1 ? dataMainDisplay[i][colMain.NOME] : dataMainDisplay[i][colMain.PLANO]) ||
          dataMainRaw[i][colMain.PLANO] ||
          ""
        ).trim();

        const statusClickup = String(
          (colMain.STATUS_CLICKUP !== -1 ? dataMainDisplay[i][colMain.STATUS_CLICKUP] : "") ||
          dataMainRaw[i][colMain.STATUS_CLICKUP] ||
          ""
        ).trim();
        const statusClickupColor = String(
          (colMain.STATUS_CLICKUP_COLOR !== -1 ? dataMainDisplay[i][colMain.STATUS_CLICKUP_COLOR] : "") ||
          dataMainRaw[i][colMain.STATUS_CLICKUP_COLOR] ||
          ""
        ).trim();
        const statusLower = statusClickup.toLowerCase();
        const isDone = statusLower.includes("finaliz") || statusLower.includes("conclu") || statusLower.includes("entreg") || statusLower.includes("fechad");
        const isDev = statusLower.includes("devol") || statusLower.includes("retorn") || statusLower.includes("return");

        const nfeSubtask = colMain.CHECKLISTS !== -1
          ? String(dataMainDisplay[i][colMain.CHECKLISTS] || "").trim()
          : "";

        const valorSubtask = colMain.VALOR_MAIN !== -1
          ? parseNumeroSeguro(dataMainRaw[i][colMain.VALOR_MAIN] || 0)
          : 0;

        const pesoSubtask = colMain.PESO_MAIN !== -1
          ? parseNumeroSeguro(dataMainRaw[i][colMain.PESO_MAIN] || 0)
          : 0;

        const dataOrdem = parseDateSafe(
          (colMain.DATA_ATUALIZACAO !== -1 ? dataMainRaw[i][colMain.DATA_ATUALIZACAO] : null) ||
          (colMain.DATA_CRIACAO !== -1 ? dataMainRaw[i][colMain.DATA_CRIACAO] : null)
        );

        const stop = {
          seq: 0,
          cliente: nomeSubtask || "Reentrega",
          clienteCodigo: "---",
          status: statusClickup || "Reentrega",
          statusColor: statusClickupColor || "",
          hora: "--:--",
          saida: "--:--",
          isDev: isDev,
          isDone: isDone,
          permanencia: null,
          nfe: nfeSubtask || "---",
          valor: valorSubtask,
          peso: pesoSubtask,
          enderecoCompleto: "",
          _ordem: dataOrdem ? dataOrdem.getTime() : null
        };

        if (!mapSubtasks.has(parentId)) {
          mapSubtasks.set(parentId, []);
        }
        mapSubtasks.get(parentId).push(stop);
      }
    }

    const mapSubtaskInfo = new Map();
    mapSubtasks.forEach((stops, parentId) => {
      stops.sort((a, b) => {
        if (a._ordem !== null && b._ordem !== null) return a._ordem - b._ordem;
        if (a._ordem !== null) return -1;
        if (b._ordem !== null) return 1;
        return String(a.cliente || "").localeCompare(String(b.cliente || ""));
      });
      stops.forEach((stop, idx) => {
        stop.seq = idx + 1;
        delete stop._ordem;
      });

      const totais = stops.reduce((acc, stop) => {
        acc.total += 1;
        if (stop.isDone) acc.feitos += 1;
        if (stop.isDev) acc.dev += 1;
        acc.peso += parseNumeroSeguro(stop.peso || 0);
        acc.valor += parseNumeroSeguro(stop.valor || 0);
        return acc;
      }, { total: 0, feitos: 0, dev: 0, peso: 0, valor: 0 });

      mapSubtaskInfo.set(parentId, { ...totais, stops });
    });

    // Mapeamento Motoristas
    const mapMotoristasAux = new Map();
    if (colMot.PLACA !== -1) {
      for (let i = 1; i < dataMot.length; i++) {
        const chave = normalizarChave(dataMot[i][colMot.PLACA]);
        if (chave) {
          mapMotoristasAux.set(chave, {
            nome: dataMot[i][colMot.MOTORISTA] || "",
            contato: dataMot[i][colMot.CONTATO] || "",
            modelo: dataMot[i][colMot.MODELO] || "",
            perfil: dataMot[i][colMot.PERFIL] || ""
          });
        }
      }
    }

    // ========================================================================
    // ✅ PROCESSAMENTO GREENMILE - VALOR DIRETO DO stop.plannedSize3
    // ========================================================================
    const mapGMData = new Map();
    
    if (colGM.ROUTE_KEY !== -1) {
      for (let i = 1; i < dataGM.length; i++) {
        const rKey = normalizarChave(dataGM[i][colGM.ROUTE_KEY]);
        if (!rKey) continue;

        if (!mapGMData.has(rKey)) {
          mapGMData.set(rKey, { total: 0, feitos: 0, dev: 0, peso: 0, valor: 0, pArrival: null, uDep: null, stops: [] });
        }
        const rota = mapGMData.get(rKey);
        const row = dataGM[i];

        const dArr = parseDateSafe(row[colGM.ARR]);
        const dDep = parseDateSafe(row[colGM.DEP]);

        if (dArr && (!rota.pArrival || dArr < rota.pArrival)) rota.pArrival = dArr;
        if (dDep && (!rota.uDep || dDep > rota.uDep)) rota.uDep = dDep;

        const statusGM = String(row[colGM.STATUS] || "").toLowerCase();
        const isDev = (row[colGM.DEV_CODE] && String(row[colGM.DEV_CODE]).trim() !== "") || statusGM.includes("return") || statusGM.includes("devolu");
        const isDone = (row[colGM.DEP] && String(row[colGM.DEP]).trim() !== "") || statusGM.includes("complete") || statusGM.includes("deliver");

        rota.total++;
        if (isDone) rota.feitos++;
        if (isDev) rota.dev++;
        
        // ✅ PESO: plannedSize1 ou actualSize1
        const pesoStop = parseNumeroSeguro(row[colGM.PESO_P] || row[colGM.PESO_A] || 0);
        rota.peso += pesoStop;
        
        // ✅ VALOR: DIRETO do stop.plannedSize3 (coluna 15)
        const valorStop = parseNumeroSeguro(row[colGM.VALOR] || 0);
        rota.valor += valorStop;

        // Permanência (tempo no cliente em minutos)
        let permanencia = null;
        if (dArr && dDep) {
          permanencia = Math.round((dDep.getTime() - dArr.getTime()) / 60000);
          if (permanencia < 0) permanencia = null;
        }

        const enderecoCompleto = montarEnderecoCompleto(
          row[colGM.LOC_DESC],
          row[colGM.LOC_ADDRESS],
          row[colGM.LOC_DISTRICT],
          row[colGM.LOC_CITY]
        );

        // ✅ BUSCA NFE: CHAVE SIMPLES = PLANO|CLIENTE
        const clientId = String(row[colGM.LOC_KEY] || "").trim();
        const chaveBusca = `${rKey}|${clientId}`;
        const nfeEncontrada = mapNfe.get(chaveBusca) || "---";

        rota.stops.push({
          seq: parseInt(row[colGM.SEQ] || 0),
          cliente: String(row[colGM.CLIENTE] || "").substring(0, 25),
          clienteCodigo: clientId || "---",
          status: isDev ? "Devolução" : (isDone ? "Realizado" : "Pendente"),
          hora: formatarHora(dArr),
          saida: formatarHora(dDep),
          isDev: isDev,
          isDone: isDone,
          permanencia: permanencia,
          nfe: nfeEncontrada,           // ✅ NFe da chave simples
          valor: valorStop,              // ✅ Valor do plannedSize3
          peso: pesoStop,
          enderecoCompleto: enderecoCompleto
        });
      }
    }

    // Payload Final
    const payload = {
      drivers: [],
      stats: { total: 0, emRota: 0, finalizados: 0, criticos: 0 },
      lastUpdate: Utilities.formatDate(new Date(), "GMT-3", "HH:mm")
    };

    const processados = new Set();
    const nowHour = new Date().getHours();

    for (let i = 1; i < dataMainRaw.length; i++) {
      const tipo = String(dataMainRaw[i][colMain.TIPO] || "").toLowerCase();
      if (!tipo.includes("principal")) continue;

      const planoFull = String(dataMainRaw[i][colMain.PLANO] || "");
      const planoChave = normalizarChave(planoFull.split('-')[0]);
      if (!planoChave || processados.has(planoChave)) continue;
      processados.add(planoChave);

      const clickupId = String(dataMainRaw[i][colMain.ID_CLICKUP] || "").trim();
      const subtaskInfo = clickupId ? mapSubtaskInfo.get(clickupId) : null;
      const placaChave = normalizarChave(dataMainRaw[i][colMain.PLACA]);
      const motoristaInfo = mapMotoristasAux.get(placaChave) || {};
      const dadosGM = mapGMData.get(planoChave);
      
      let status = "OUTROS", label = "Apoio / Outros", sClass = "secondary", prog = 0;
      let devPct = 0;
      let entreguePct = 0;

      if (dadosGM && dadosGM.total > 0) {
        prog = Math.round((dadosGM.feitos / dadosGM.total) * 100) || 0;
        devPct = Math.round((dadosGM.dev / dadosGM.total) * 100) || 0;
        entreguePct = Math.max(prog - devPct, 0);
        status = "EM_ROTA"; label = "Em Trânsito"; sClass = "info";

        if (prog >= 100) {
          if (dadosGM.dev > 0) { status = "RESSALVA"; label = "Com Devolução"; sClass = "warning"; }
          else { status = "FINALIZADO"; label = "Finalizado"; sClass = "success"; }
        } else if (nowHour >= 14 && prog < 50) {
          status = "CRITICO"; label = "Atrasado"; sClass = "danger";
        }
      }

      // Tempo em Rota
      let tempoRota = "--";
      if (dadosGM && dadosGM.pArrival) {
        const fimTempo = dadosGM.uDep || new Date();
        const diffMs = fimTempo.getTime() - dadosGM.pArrival.getTime();
        const diffMins = Math.round(diffMs / 60000);
        if (diffMins > 0 && diffMins < 1440) {
          const horas = Math.floor(diffMins / 60);
          const mins = diffMins % 60;
          tempoRota = horas > 0 ? `${horas}h ${String(mins).padStart(2, '0')}m` : `${mins}m`;
        }
      }

      // Data de Saída
      let dataSaida = "--/--";
      let dataSaidaIso = "";
      const rawDataSaida = dataMainRaw[i][colMain.DATA_SAIDA];
      if (rawDataSaida instanceof Date) {
        dataSaida = Utilities.formatDate(rawDataSaida, Session.getScriptTimeZone(), "dd/MM");
        dataSaidaIso = Utilities.formatDate(rawDataSaida, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else if (rawDataSaida) {
        const parsed = parseDateSafe(rawDataSaida);
        if (parsed) {
          dataSaida = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "dd/MM");
          dataSaidaIso = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
      }

      // Contagem de Stats
      payload.stats.total++;
      if (sClass === 'info') payload.stats.emRota++;
      if (sClass === 'success' || sClass === 'warning') payload.stats.finalizados++;
      if (sClass === 'danger') payload.stats.criticos++;

      const totaisFallback = subtaskInfo || { total: 0, feitos: 0, dev: 0, peso: 0, valor: 0, stops: [] };
      const detalhes = (dadosGM && dadosGM.stops && dadosGM.stops.length > 0)
        ? dadosGM.stops
        : totaisFallback.stops;

      const entregasTexto = (dadosGM && dadosGM.total > 0)
        ? `${dadosGM.feitos}/${dadosGM.total}`
        : `${totaisFallback.feitos}/${totaisFallback.total}`;

      const pesoTotal = (dadosGM && dadosGM.total > 0)
        ? dadosGM.peso
        : totaisFallback.peso;

      const valorTotal = (dadosGM && dadosGM.total > 0)
        ? dadosGM.valor
        : totaisFallback.valor;

      if ((!dadosGM || dadosGM.total === 0) && totaisFallback.total > 0) {
        prog = Math.round((totaisFallback.feitos / totaisFallback.total) * 100) || 0;
        devPct = Math.round((totaisFallback.dev / totaisFallback.total) * 100) || 0;
        entreguePct = Math.max(prog - devPct, 0);
      }

      const statusClickupColor = String(
        (colMain.STATUS_CLICKUP_COLOR !== -1 ? dataMainDisplay[i][colMain.STATUS_CLICKUP_COLOR] : "") ||
        dataMainRaw[i][colMain.STATUS_CLICKUP_COLOR] ||
        ""
      ).trim();

      payload.drivers.push({
        id: planoChave,
        placa: String(dataMainRaw[i][colMain.PLACA] || "").toUpperCase(),
        motorista: toTitleCase(String(dataMainRaw[i][colMain.MOTORISTA] || motoristaInfo.nome || "Indefinido")),
        tel: String(motoristaInfo.contato || "").replace(/[^0-9]/g, ""),
        veiculo: toTitleCase(motoristaInfo.modelo || "Veículo"),
        plano: planoFull,
        progresso: prog,
        devolucaoPct: Math.min(Math.max(devPct, 0), 100),
        entreguePct: Math.min(Math.max(entreguePct, 0), 100),
        status: status,
        statusLabel: label,
        statusClass: sClass,
        entregas: entregasTexto,
        peso: (pesoTotal || 0).toFixed(0),
        valorTotal: valorTotal || 0,  // ✅ Soma dos plannedSize3
        unidade: dataMainRaw[i][colMain.UNIDADE] || "---",
        clickupId: dataMainRaw[i][colMain.ID_CLICKUP] || "",
        statusClickup: dataMainRaw[i][colMain.STATUS_CLICKUP] || "",
        statusClickupColor: statusClickupColor,
        dataSaida: dataSaida,
        dataSaidaIso: dataSaidaIso,
        tempoRota: tempoRota,
        detalhes: detalhes
      });
    }

    // Ordenação: Críticos primeiro
    const pesoStatus = { "CRITICO": 4, "EM_ROTA": 3, "RESSALVA": 2, "FINALIZADO": 1, "OUTROS": 0 };
    payload.drivers.sort((a,b) => pesoStatus[b.status] - pesoStatus[a.status]);

    // Envolve payload com apiResponse antes de cachear
    const response = apiResponse(true, payload, null);

    // Salva no cache
    try {
      cache.put(cacheKey, JSON.stringify(response), CACHE_DURATION);
    } catch(e) {
      console.warn("Cache overflow: " + e.message);
    }

    console.log("✅ getDashboardData: " + payload.drivers.length + " motoristas retornados");
    return response;

  } catch (erro) {
    console.error("❌ ERRO FATAL em getDashboardData: " + erro.message);
    console.error("Stack: " + erro.stack);
    const needsAuth = isAuthError(erro);
    return apiResponse(false, {
      drivers: [],
      stats: { total: 0, emRota: 0, finalizados: 0, criticos: 0 }
    }, {
      message: needsAuth ? "Autorize o app para continuar." : "Erro ao carregar dados: " + erro.message,
      needsAuth: needsAuth
    });
  }
}

/**
 * ============================================================================
 * EXPORTAÇÃO DE PLANILHA - VERSÃO COMPLETA E OTIMIZADA
 * ============================================================================
 */

function exportDashboardCsv() {
    const authProbe = safeExecute(() => SpreadsheetApp.getActiveSpreadsheet().getId());
    if (!authProbe.ok) {
        return apiResponse(false, null, {
            message: authProbe.message,
            needsAuth: authProbe.needsAuth
        });
    }
    try {
        Logger.log("📊 Iniciando exportação de planilha...");
        
        // 1️⃣ Obtém os dados do dashboard
        const dashboardData = getDashboardData();
        
        if (!dashboardData || !dashboardData.ok || !dashboardData.data || !dashboardData.data.drivers || dashboardData.data.drivers.length === 0) {
            Logger.log("⚠️ Nenhum dado encontrado para exportar");
            return apiResponse(false, null, "Nenhum dado disponível para exportação.");
        }

        const drivers = dashboardData.data.drivers;
        Logger.log(`✅ ${drivers.length} rotas encontradas`);
        
        // 2️⃣ Cria nova planilha
        const agora = new Date();
        const dataStr = Utilities.formatDate(agora, Session.getScriptTimeZone(), "dd-MM-yyyy_HH-mm");
        const nomePlanilha = `Dashboard_THX_${dataStr}`;
        
        const ss = SpreadsheetApp.create(nomePlanilha);
        const sheet = ss.getActiveSheet();
        sheet.setName("Rotas");
        
        Logger.log(`✅ Planilha criada: ${nomePlanilha}`);
        
        // 3️⃣ Define cabeçalhos
        const headers = [
            "Status",
            "Motorista",
            "Telefone",
            "Veículo",
            "Placa",
            "Unidade",
            "Plano/Rota",
            "Data Saída",
            "Progresso (%)",
            "Entregue (%)",
            "Devolução (%)",
            "Entregas",
            "Peso (kg)",
            "Valor Total (R$)",
            "Tempo em Rota",
            "Status ClickUp",
            "ID ClickUp"
        ];
        
        // 4️⃣ Escreve cabeçalhos
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        
        // 5️⃣ Formata cabeçalhos
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setBackground("#1a73e8");
        headerRange.setFontColor("#FFFFFF");
        headerRange.setFontWeight("bold");
        headerRange.setHorizontalAlignment("center");
        headerRange.setVerticalAlignment("middle");
        
        // 6️⃣ Prepara dados
        const rows = [];
        
        const normalizePercent = value => {
            const numeric = parseFloat(value);
            if (Number.isNaN(numeric)) return 0;
            return numeric > 1 ? numeric / 100 : numeric;
        };

        drivers.forEach(driver => {
            const row = [
                driver.statusLabel || "-",
                driver.motorista || "-",
                driver.tel || "-",
                driver.veiculo || "-",
                driver.placa || "-",
                driver.unidade || "-",
                driver.plano || "-",
                driver.dataSaida || "-",
                normalizePercent(driver.progresso),
                normalizePercent(driver.entreguePct),
                normalizePercent(driver.devolucaoPct),
                parseInt(driver.entregas) || 0,
                parseFloat(driver.peso) || 0,
                parseFloat(driver.valorTotal) || 0,
                driver.tempoRota || "-",
                driver.statusClickup || "-",
                driver.clickupId || "-"
            ];
            rows.push(row);
        });
        
        // 7️⃣ Escreve dados
        if (rows.length > 0) {
            sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
            Logger.log(`✅ ${rows.length} linhas escritas`);
        }
        
        // 8️⃣ Formata colunas numéricas
        const lastRow = rows.length + 1;
        
        // Progresso, Entregue, Devolução (%)
        sheet.getRange(2, 9, rows.length, 3).setNumberFormat("0.00%");
        
        // Valor Total (R$)
        sheet.getRange(2, 14, rows.length, 1).setNumberFormat("R$ #,##0.00");
        
        // Peso (kg)
        sheet.getRange(2, 13, rows.length, 1).setNumberFormat("#,##0.00");
        
        // Centraliza algumas colunas
        sheet.getRange(2, 8, rows.length, 1).setHorizontalAlignment("center"); // Data
        sheet.getRange(2, 9, rows.length, 4).setHorizontalAlignment("center"); // Progresso/Entregas
        
        // 9️⃣ Ajusta largura das colunas
        sheet.autoResizeColumns(1, headers.length);
        
        // Garante largura mínima
        for (let i = 1; i <= headers.length; i++) {
            const currentWidth = sheet.getColumnWidth(i);
            if (currentWidth < 100) {
                sheet.setColumnWidth(i, 100);
            }
            if (currentWidth > 300) {
                sheet.setColumnWidth(i, 300);
            }
        }
        
        // 🔟 Formatação condicional por status
        applyConditionalFormatting(sheet, rows.length);
        
        // 1️⃣1️⃣ Congela cabeçalho
        sheet.setFrozenRows(1);
        
        // 1️⃣2️⃣ Adiciona filtros
        sheet.getRange(1, 1, lastRow, headers.length).createFilter();
        
        // 1️⃣3️⃣ Move para o início
        ss.setActiveSheet(sheet);
        
        // 1️⃣4️⃣ Compartilha com o usuário
        const userEmail = Session.getActiveUser().getEmail();
        if (userEmail) {
            ss.addEditor(userEmail);
            Logger.log(`✅ Planilha compartilhada com: ${userEmail}`);
        }
        
        // 1️⃣5️⃣ Obtém URL e retorna
        const url = ss.getUrl();
        const downloadUrl = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx`;
        Logger.log(`✅ Exportação concluída: ${url}`);

        return apiResponse(true, {
            url: url,
            downloadUrl: downloadUrl,
            fileName: nomePlanilha,
            rowCount: rows.length
        }, null);

    } catch (error) {
        Logger.log(`❌ Erro na exportação: ${error.toString()}`);
        Logger.log(`Stack: ${error.stack}`);

        const needsAuth = isAuthError(error);

        return apiResponse(false, null, {
            message: needsAuth ? "Autorize o app para continuar." : `Erro ao gerar planilha: ${error.toString()}`,
            needsAuth: needsAuth
        });
    }
}

/**
 * ============================================================================
 * FORMATAÇÃO CONDICIONAL
 * ============================================================================
 */
function applyConditionalFormatting(sheet, rowCount) {
    try {
        const statusRange = sheet.getRange(2, 1, rowCount, 1);
        
        // Verde para "FINALIZADO"
        const ruleFinalizado = SpreadsheetApp.newConditionalFormatRule()
            .whenTextContains("FINALIZADO")
            .setBackground("#D4EDDA")
            .setFontColor("#155724")
            .setRanges([statusRange])
            .build();
        
        // Amarelo para "EM ROTA"
        const ruleEmRota = SpreadsheetApp.newConditionalFormatRule()
            .whenTextContains("EM ROTA")
            .setBackground("#FFF3CD")
            .setFontColor("#856404")
            .setRanges([statusRange])
            .build();
        
        // Vermelho para "CRÍTICO"
        const ruleCritico = SpreadsheetApp.newConditionalFormatRule()
            .whenTextContains("CRÍTICO")
            .setBackground("#F8D7DA")
            .setFontColor("#721C24")
            .setRanges([statusRange])
            .build();
        
        // Laranja para "RESSALVA"
        const ruleRessalva = SpreadsheetApp.newConditionalFormatRule()
            .whenTextContains("RESSALVA")
            .setBackground("#FFE5CC")
            .setFontColor("#CC5500")
            .setRanges([statusRange])
            .build();
        
        const rules = sheet.getConditionalFormatRules();
        rules.push(ruleFinalizado, ruleEmRota, ruleCritico, ruleRessalva);
        sheet.setConditionalFormatRules(rules);
        
        Logger.log("✅ Formatação condicional aplicada");
        
    } catch (error) {
        Logger.log(`⚠️ Erro ao aplicar formatação condicional: ${error}`);
    }
}

/**
 * ============================================================================
 * ABA DE ESTATÍSTICAS
 * ============================================================================
 */
function addStatisticsSheet(ss, drivers) {
    try {
        const statsSheet = ss.insertSheet("Estatísticas");
        
        // Calcula estatísticas
        const total = drivers.length;
        const emRota = drivers.filter(d => d.status === "EM_ROTA").length;
        const criticos = drivers.filter(d => d.status === "CRITICO").length;
        const finalizados = drivers.filter(d => d.status === "FINALIZADO").length;
        const ressalvas = drivers.filter(d => d.status === "RESSALVA").length;
        
        const totalEntregas = drivers.reduce((sum, d) => sum + (parseInt(d.entregas) || 0), 0);
        const totalPeso = drivers.reduce((sum, d) => sum + (parseFloat(d.peso) || 0), 0);
        const totalValor = drivers.reduce((sum, d) => sum + (parseFloat(d.valorTotal) || 0), 0);
        
        const progressoMedio = drivers.reduce((sum, d) => sum + (parseFloat(d.progresso) || 0), 0) / total;
        
        // Agrupa por unidade
        const unidades = {};
        drivers.forEach(d => {
            const unidade = d.unidade || "Sem Unidade";
            if (!unidades[unidade]) {
                unidades[unidade] = 0;
            }
            unidades[unidade]++;
        });
        
        // Monta dados da aba
        const statsData = [
            ["📊 ESTATÍSTICAS DO DASHBOARD", ""],
            [""],
            ["📅 Data/Hora da Exportação:", new Date().toLocaleString('pt-BR')],
            [""],
            ["═══════════════════════════", "═══════════════"],
            ["🚛 RESUMO GERAL", ""],
            ["═══════════════════════════", "═══════════════"],
            ["Total de Rotas:", total],
            ["Em Rota:", emRota],
            ["Críticos:", criticos],
            ["Finalizados:", finalizados],
            ["Ressalvas:", ressalvas],
            [""],
            ["═══════════════════════════", "═══════════════"],
            ["📦 TOTAIS", ""],
            ["═══════════════════════════", "═══════════════"],
            ["Total de Entregas:", totalEntregas],
            ["Peso Total (kg):", totalPeso],
            ["Valor Total:", totalValor],
            ["Progresso Médio (%):", progressoMedio],
            [""],
            ["═══════════════════════════", "═══════════════"],
            ["🏢 POR UNIDADE", "Quantidade"],
            ["═══════════════════════════", "═══════════════"]
        ];
        
        // Adiciona dados por unidade
        Object.keys(unidades).sort().forEach(unidade => {
            statsData.push([unidade, unidades[unidade]]);
        });
        
        // Escreve dados
        statsSheet.getRange(1, 1, statsData.length, 2).setValues(statsData);
        
        // Formata título
        statsSheet.getRange(1, 1, 1, 2).merge();
        statsSheet.getRange(1, 1).setBackground("#1a73e8");
        statsSheet.getRange(1, 1).setFontColor("#FFFFFF");
        statsSheet.getRange(1, 1).setFontWeight("bold");
        statsSheet.getRange(1, 1).setFontSize(14);
        statsSheet.getRange(1, 1).setHorizontalAlignment("center");
        
        // Formata seções
        const sectionRows = [6, 14, 22];
        sectionRows.forEach(row => {
            statsSheet.getRange(row, 1, 1, 2).setBackground("#E8F0FE");
            statsSheet.getRange(row, 1, 1, 2).setFontWeight("bold");
        });
        
        // Formata valores
        statsSheet.getRange(18, 2).setNumberFormat("#,##0.00");
        statsSheet.getRange(19, 2).setNumberFormat("R$ #,##0.00");
        statsSheet.getRange(20, 2).setNumberFormat("0.00%");
        
        // Ajusta largura
        statsSheet.setColumnWidth(1, 300);
        statsSheet.setColumnWidth(2, 150);
        
        Logger.log("✅ Aba de estatísticas criada");
        
    } catch (error) {
        Logger.log(`⚠️ Erro ao criar aba de estatísticas: ${error}`);
    }
}

// Compatibilidade com chamadas antigas
function exportDashboardXlsx() {
  return exportDashboardCsv();
}

// ============================================================================
// CSV HELPERS
// ============================================================================
function toCsv(rows) {
  return rows.map(row => row.map(csvCell).join(",")).join("\r\n");
}

function csvCell(value) {
  if (value === null || value === undefined) return "";
  const s = String(value);
  if (s.includes("\"") || s.includes(",") || s.includes("\n") || s.includes("\r")) {
    return "\"" + s.replace(/\"/g, "\"\"") + "\"";
  }
  return s;
}

// ============================================================================
// HELPERS
// ============================================================================
function parseDateSafe(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  try {
    let d = new Date(val);
    if (!isNaN(d.getTime())) return d;
  } catch(e) {}
  return null;
}

function normalizarChave(val) {
  if (!val) return "";
  return String(val).replace(/[^a-zA-Z0-9]/g, "").toUpperCase().trim();
}

function formatarHora(val) {
  if (!val) return "--:--";
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
  let s = String(val);
  if (s.includes("T")) {
    let timePart = s.split("T")[1];
    return timePart ? timePart.substring(0,5) : "--:--";
  }
  if (s.length >= 5) return s.substring(0,5);
  return s;
}

function toTitleCase(str) {
  if (!str) return "";
  return str.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase()); 
}

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

function montarEnderecoCompleto(desc, addressLine1, district, city) {
  const vistos = new Set();
  const partes = [desc, addressLine1, district, city]
    .map(item => String(item || "").trim())
    .filter(Boolean)
    .filter(item => {
      const chave = item.toLowerCase();
      if (vistos.has(chave)) return false;
      vistos.add(chave);
      return true;
    });
  return partes.join(" - ");
}

function mapDashboardCols(headers, type) {
  const map = {
    PLANO:-1, PLACA:-1, MOTORISTA:-1, ROUTE_KEY:-1,
    ARR:-1, DEP:-1, STATUS:-1,
    PESO_P:-1, PESO_A:-1,
    ID_GM_LOC:-1, CHECKLISTS:-1, TIPO:-1, ID_PAI:-1,
    VALOR:-1, LOC_KEY:-1,
    DATA_SAIDA:-1, ID_CLICKUP:-1,
    UNIDADE:-1, CONTATO:-1, MODELO:-1, PERFIL:-1,
    DEV_CODE:-1, CLIENTE:-1, SEQ:-1, CLIENT_ID:-1, PARENT_ID:-1,
    LOC_DESC:-1, LOC_ADDRESS:-1, LOC_CITY:-1, LOC_DISTRICT:-1,
    STATUS_CLICKUP:-1, STATUS_CLICKUP_COLOR:-1, NOME:-1, DATA_CRIACAO:-1, DATA_ATUALIZACAO:-1,
    DATA_FECHAMENTO:-1, VALOR_MAIN:-1, PESO_MAIN:-1
  };

  headers.forEach((h, i) => {
    const t = String(h || "").trim().toUpperCase();

    // MAIN (Aba ENTREGAS)
    if (type === 'MAIN') {
      if (t === 'PLANO' || t === 'ROTA' || t === 'NOME') map.PLANO = i;
      if (t === 'NOME') map.NOME = i;
      if (t === 'PLACA') map.PLACA = i;
      if (t === 'MOTORISTA') map.MOTORISTA = i;
      if (t === 'UNIDADE' || t === 'BASE' || t === 'CD' || t === 'FILIAL') map.UNIDADE = i;
      if (t === 'ID GM LOCALIZAÇÃO' || t === 'ID GM LOCALIZACAO' || t === 'ID GM LOCALIZAÇÃO') map.ID_GM_LOC = i;
      if (t === 'CHECKLISTS' || t === 'CHECKLIST' || t === 'NFE' || t === 'NOTAS') map.CHECKLISTS = i;
      if (t === 'TIPO DE TAREFA' || t === 'TIPO') map.TIPO = i;
      if (t === 'DATA DE SAÍDA' || t === 'DATA DE SAIDA') map.DATA_SAIDA = i;
      if (t === 'DATA DE CRIAÇÃO' || t === 'DATA DE CRIACAO') map.DATA_CRIACAO = i;
      if (t === 'DATA DE ATUALIZAÇÃO' || t === 'DATA DE ATUALIZACAO') map.DATA_ATUALIZACAO = i;
      if (t === 'DATA DE FECHAMENTO') map.DATA_FECHAMENTO = i;
      if (t === 'ID' || t === 'TASK ID' || t === 'CLICKUP ID') map.ID_CLICKUP = i;
      if (t === 'ID DO PAI' || t === 'PARENT ID') map.ID_PAI = i;
      if (t === 'STATUS' || t === 'STATUS CLICKUP' || t === 'SITUACAO') map.STATUS_CLICKUP = i;
      if (t === 'STATUS COR' || t === 'STATUS COLOR' || t === 'STATUS CLICKUP COR' || t === 'COR STATUS') map.STATUS_CLICKUP_COLOR = i;
      if (t === 'CLIENTE' || t === 'NOME DO CLIENTE') map.CLIENTE = i;
      if (t.includes('VALOR')) map.VALOR_MAIN = i;
      if (t.includes('PESO')) map.PESO_MAIN = i;
    }

    // MOT (Aba MOTORISTAS)
    if (type === 'MOT') {
      if (t === 'PLACA') map.PLACA = i;
      if (t === 'MOTORISTA' || t === 'NOME') map.MOTORISTA = i;
      if (t.includes('CONTATO')) map.CONTATO = i;
      if (t === 'MODELO') map.MODELO = i;
      if (t.includes('PERFIL')) map.PERFIL = i;
    }

    // GM (Aba GreenMile)
    if (type === 'GM') {
      if (t === 'ROUTE.KEY' || t === 'ROUTE KEY') map.ROUTE_KEY = i;
      if (t === 'STOP.ACTUALARRIVAL' || t.includes('ACTUALARRIVAL')) map.ARR = i;
      if (t === 'STOP.ACTUALDEPARTURE' || t.includes('ACTUALDEPARTURE')) map.DEP = i;
      if (t.includes('UNDELIVERABLECODE') || t.includes('DEVOLUCAO')) map.DEV_CODE = i;
      if (t === 'STOP.DELIVERYSTATUS' || t.includes('DELIVERYSTATUS')) map.STATUS = i;
      if (t === 'STOP.LOCATION.DESCRIPTION' || t.includes('LOCATION.DESCRIPTION')) {
        map.CLIENTE = i;
        map.LOC_DESC = i;
      }
      if (t === 'STOP.LOCATION.ADDRESSLINE1' || t.includes('LOCATION.ADDRESSLINE1')) map.LOC_ADDRESS = i;
      if (t === 'STOP.LOCATION.CITY' || t.includes('LOCATION.CITY')) map.LOC_CITY = i;
      if (t === 'STOP.LOCATION.DISTRICT' || t.includes('LOCATION.DISTRICT')) map.LOC_DISTRICT = i;
      if (t === 'STOP.PLANNEDSEQUENCENUM' || t.includes('PLANNEDSEQUENCENUM')) map.SEQ = i;
      if (t === 'STOP.PLANNEDSIZE1' || t.includes('PLANNEDSIZE1')) map.PESO_P = i;
      if (t === 'STOP.ACTUALSIZE1' || t.includes('ACTUALSIZE1')) map.PESO_A = i;
      
      // ✅ VALOR: stop.plannedSize3 (coluna 15 conforme debug)
      if (t === 'STOP.PLANNEDSIZE3' || t === 'PLANNEDSIZE3') map.VALOR = i;
      
      // ✅ LOCATION.KEY: stop.location.key (coluna 18 conforme debug)
      if (t === 'STOP.LOCATION.KEY' || t === 'LOCATION.KEY') map.LOC_KEY = i;
    }
  });

  return map;
}

// ============================================================================
// SALVAR OCORRÊNCIA
// ============================================================================
function salvarOcorrencia(formData) {
  const authProbe = safeExecute(() => SpreadsheetApp.getActiveSpreadsheet().getId());
  if (!authProbe.ok) {
    return apiResponse(false, null, {
      message: authProbe.message,
      needsAuth: authProbe.needsAuth
    });
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let ws = ss.getSheetByName("Ocorrencias");
    if (!ws) {
      ws = ss.insertSheet("Ocorrencias");
      ws.appendRow(["DATA", "MOTORISTA", "ROTA", "CLIENTE", "DATA SAIDA", "NFE", "MOTIVO", "CAUSADOR", "VALOR", "DESCRICAO"]);
    }
    ws.appendRow([new Date(), formData.motorista, formData.rota, formData.cliente, formData.dataSaida, formData.nfe, formData.motivo, formData.causador, formData.valor, formData.descricao]);
    return apiResponse(true, { saved: true }, null);
  } catch (e) {
    const needsAuth = isAuthError(e);
    return apiResponse(false, null, {
      message: needsAuth ? "Autorize o app para continuar." : e.toString(),
      needsAuth: needsAuth
    });
  }
}

// ============================================================================
// FINALIZAR TAREFA NO CLICKUP
// ============================================================================
function finalizarTarefaBackend(clickupId, rotaId) {
  const authProbe = safeExecute(() => SpreadsheetApp.getActiveSpreadsheet().getId());
  if (!authProbe.ok) {
    return apiResponse(false, null, {
      message: authProbe.message,
      needsAuth: authProbe.needsAuth
    });
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ws = ss.getSheetByName(SHEET_MAIN);
    
    // Envia para o ClickUp
    const sucessoApi = enviarStatusParaClickup(clickupId, "Finalizada");

    if (!sucessoApi) {
      return apiResponse(false, null, "Erro na API do ClickUp");
    }

    // Atualiza na planilha
    if (ws && clickupId) {
      const data = ws.getDataRange().getValues();
      const headers = data[0];
      
      let colIdClickup = -1;
      let colStatus = -1;
      let colStatusColor = -1;
      
      headers.forEach((h, i) => {
        let t = String(h).toUpperCase().trim();
        if (t === 'ID' || t === 'TASK ID' || t === 'CLICKUP ID') colIdClickup = i;
        if (t === 'STATUS' || t === 'STATUS CLICKUP' || t === 'SITUACAO') colStatus = i;
        if (t === 'STATUS COR' || t === 'STATUS COLOR' || t === 'STATUS CLICKUP COR' || t === 'COR STATUS') colStatusColor = i;
      });

      if (colIdClickup !== -1 && colStatus !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][colIdClickup]).trim() === String(clickupId).trim()) {
            ws.getRange(i + 1, colStatus + 1).setValue("Finalizada");
            if (colStatusColor !== -1) {
              ws.getRange(i + 1, colStatusColor + 1).setValue(mapClickupStatusColor("Finalizada"));
            }
            SpreadsheetApp.flush(); 
            break; 
          }
        }
      }
    }

    // Limpa o cache para forçar atualização
    try {
      CacheService.getScriptCache().remove("payload_dashboard_v6");
    } catch(e) {}

    return apiResponse(true, { finalized: true }, null);

  } catch(e) {
    console.error("Erro em finalizarTarefaBackend: " + e.message);
    const needsAuth = isAuthError(e);
    return apiResponse(false, null, {
      message: needsAuth ? "Autorize o app para continuar." : e.message,
      needsAuth: needsAuth
    });
  }
}

// ============================================================================
// ATUALIZAR STATUS NO CLICKUP (GENÃ‰RICO)
// ============================================================================
function atualizarStatusClickupBackend(clickupId, novoStatus, rotaId) {
  const authProbe = safeExecute(() => SpreadsheetApp.getActiveSpreadsheet().getId());
  if (!authProbe.ok) {
    return apiResponse(false, null, {
      message: authProbe.message,
      needsAuth: authProbe.needsAuth
    });
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ws = ss.getSheetByName(SHEET_MAIN);

    if (!clickupId) {
      return apiResponse(false, null, "ID do ClickUp não informado");
    }

    const sucessoApi = enviarStatusParaClickup(clickupId, novoStatus);
    if (!sucessoApi) {
      return apiResponse(false, null, "Erro na API do ClickUp");
    }

    let statusColor = mapClickupStatusColor(novoStatus);

    if (ws && clickupId) {
      const data = ws.getDataRange().getValues();
      const headers = data[0];
      let colIdClickup = -1;
      let colStatus = -1;
      let colStatusColor = -1;

      headers.forEach((h, i) => {
        let t = String(h).toUpperCase().trim();
        if (t === 'ID' || t === 'TASK ID' || t === 'CLICKUP ID') colIdClickup = i;
        if (t === 'STATUS' || t === 'STATUS CLICKUP' || t === 'SITUACAO') colStatus = i;
        if (t === 'STATUS COR' || t === 'STATUS COLOR' || t === 'STATUS CLICKUP COR' || t === 'COR STATUS') colStatusColor = i;
      });

      if (colIdClickup !== -1 && colStatus !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][colIdClickup]).trim() === String(clickupId).trim()) {
            ws.getRange(i + 1, colStatus + 1).setValue(novoStatus);
            if (colStatusColor !== -1) {
              ws.getRange(i + 1, colStatusColor + 1).setValue(statusColor);
            }
            SpreadsheetApp.flush();
            break;
          }
        }
      }
    }

    try {
      CacheService.getScriptCache().remove("payload_dashboard_v6");
    } catch(e) {}

    return apiResponse(true, { status: novoStatus, color: statusColor }, null);
  } catch(e) {
    console.error("Erro em atualizarStatusClickupBackend: " + e.message);
    const needsAuth = isAuthError(e);
    return apiResponse(false, null, {
      message: needsAuth ? "Autorize o app para continuar." : e.message,
      needsAuth: needsAuth
    });
  }
}

function mapClickupStatusColor(status) {
  const s = String(status || "").toLowerCase().trim();
  if (!s) return "";
  if (s.includes("final") || s.includes("conclu") || s.includes("fechad")) return "#10B981";
  if (s.includes("cancel")) return "#EF4444";
  if (s.includes("pernoite")) return "#F59E0B";
  return "#3B82F6";
}








