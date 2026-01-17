function backupToDrive() {
  var scriptId = ScriptApp.getScriptId(); // Pega o ID deste projeto
  // Se quiser baixar OUTRO projeto, coloque o ID dele entre aspas abaixo:
  // var scriptId = "ID_DO_OUTRO_PROJETO"; 
  
  var url = "https://script.google.com/feeds/download/export?id=" + scriptId + "&format=json";
  
  var params = {
    method: "GET",
    headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(url, params);
  
  if (response.getResponseCode() !== 200) {
    console.error("Erro ao baixar: " + response.getContentText());
    return;
  }
  
  var blob = response.getBlob().setName("Backup_Script_" + new Date().toISOString() + ".json");
  DriveApp.createFile(blob);
  console.log("Backup salvo no Google Drive!");
}

/**
 * 🔍 FUNÇÃO DEBUG - RASTREAMENTO COMPLETO DA DATA DE SAÍDA
 * Plano Alvo: 6102867018-1
 */

function DEBUG_DATA_SAIDA_COMPLETO() {
  const PLANO_TESTE = "6102867018-1";
  console.log(`🎯 INICIANDO DEBUG PARA: ${PLANO_TESTE}\n`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsMain = ss.getSheetByName("ENTREGAS");
  
  if (!wsMain) {
    console.error("❌ ABA ENTREGAS NÃO ENCONTRADA!");
    return;
  }
  
  // === PASSO 1: VERIFICAR HEADERS ===
  console.log("📋 === PASSO 1: ANALISANDO CABEÇALHOS ===");
  const headers = wsMain.getRange(1, 1, 1, wsMain.getLastColumn()).getValues()[0];
  
  console.log(`Total de Colunas: ${headers.length}\n`);
  
  // Procura colunas relacionadas a "Data" ou "Saída"
  const colunasData = [];
  headers.forEach((h, idx) => {
    let texto = String(h).toUpperCase().trim();
    if (texto.includes("DATA") || texto.includes("SAIDA") || texto.includes("SAÍDA")) {
      colunasData.push({ indice: idx, nome: h, nomeUpper: texto });
      console.log(`   ✓ Coluna ${idx}: "${h}" (Upper: "${texto}")`);
    }
  });
  
  if (colunasData.length === 0) {
    console.error("\n❌ NENHUMA COLUNA DE DATA ENCONTRADA!");
    console.log("💡 Colunas disponíveis:");
    headers.forEach((h, i) => console.log(`   ${i}: ${h}`));
    return;
  }
  
  // === PASSO 2: ENCONTRAR A LINHA DO PLANO ===
  console.log("\n📋 === PASSO 2: BUSCANDO LINHA DO PLANO ===");
  
  const colPlano = headers.findIndex(h => {
    let t = String(h).toUpperCase().trim();
    return t === "NOME" || t === "PLANO" || t === "ROTA";
  });
  
  if (colPlano === -1) {
    console.error("❌ COLUNA DE PLANO NÃO ENCONTRADA!");
    return;
  }
  
  console.log(`✓ Coluna do Plano: ${colPlano} ("${headers[colPlano]}")\n`);
  
  const dataMain = wsMain.getDataRange().getValues();
  let linhaEncontrada = -1;
  
  for (let i = 1; i < dataMain.length; i++) {
    let valorPlano = String(dataMain[i][colPlano]).trim();
    if (valorPlano.includes(PLANO_TESTE)) {
      linhaEncontrada = i;
      console.log(`✅ PLANO ENCONTRADO NA LINHA: ${i + 1}`);
      break;
    }
  }
  
  if (linhaEncontrada === -1) {
    console.error(`❌ PLANO "${PLANO_TESTE}" NÃO ENCONTRADO NA PLANILHA!`);
    return;
  }
  
  // === PASSO 3: EXTRAIR VALORES DE DATA ===
  console.log("\n📋 === PASSO 3: VALORES DAS COLUNAS DE DATA ===");
  
  const linhaData = dataMain[linhaEncontrada];
  
  colunasData.forEach(col => {
    let valorRaw = linhaData[col.indice];
    let tipo = typeof valorRaw;
    let valorStr = String(valorRaw);
    
    console.log(`\n   Coluna: "${col.nome}"`);
    console.log(`   Índice: ${col.indice}`);
    console.log(`   Valor RAW: ${valorRaw}`);
    console.log(`   Tipo: ${tipo}`);
    console.log(`   String: "${valorStr}"`);
    
    if (valorRaw instanceof Date) {
      console.log(`   ✓ É uma DATA válida!`);
      console.log(`   Formatada: ${Utilities.formatDate(valorRaw, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")}`);
    } else if (valorStr && valorStr !== "" && valorStr !== "undefined") {
      console.log(`   ⚠️ É texto, tentando converter...`);
      try {
        let tentativaData = new Date(valorRaw);
        if (!isNaN(tentativaData.getTime())) {
          console.log(`   ✓ Conversão bem-sucedida: ${Utilities.formatDate(tentativaData, Session.getScriptTimeZone(), "dd/MM/yyyy")}`);
        } else {
          console.log(`   ❌ Conversão falhou (Data inválida)`);
        }
      } catch(e) {
        console.log(`   ❌ Erro ao converter: ${e.message}`);
      }
    } else {
      console.log(`   ⚠️ VAZIO ou NULL`);
    }
  });
  
  // === PASSO 4: SIMULAR O MAPEAMENTO DO DASHBOARD ===
  console.log("\n📋 === PASSO 4: SIMULANDO mapDashboardCols ===");
  
  let colDataSaida = -1;
  headers.forEach((h, i) => {
    let t = String(h).trim().toUpperCase();
    if ((t.includes('DATA') && (t.includes('SAIDA') || t.includes('SAÍDA'))) || 
        t === 'SAIDA' || 
        t === 'DE SAÍDA') {
      colDataSaida = i;
      console.log(`   ✅ MATCH! Coluna ${i}: "${h}"`);
    }
  });
  
  if (colDataSaida === -1) {
    console.error("\n❌ A FUNÇÃO mapDashboardCols NÃO CONSEGUIU MAPEAR!");
    console.log("\n💡 Solução: A coluna precisa ter um desses nomes:");
    console.log("   - 'Data de Saída'");
    console.log("   - 'SAIDA'");
    console.log("   - 'DE SAÍDA'");
    console.log("\nOu ajustar a regex no código do backend.");
  } else {
    console.log(`\n✅ Coluna mapeada com sucesso: ${colDataSaida}`);
    
    // Testar formatação final
    let valorFinal = linhaData[colDataSaida];
    console.log(`\n📋 === FORMATAÇÃO FINAL (Como vai pro Frontend) ===`);
    console.log(`   Valor RAW: ${valorFinal}`);
    
    if (valorFinal instanceof Date) {
      let formatado = Utilities.formatDate(valorFinal, Session.getScriptTimeZone(), "dd/MM");
      console.log(`   ✅ SUCESSO! Será enviado: "${formatado}"`);
    } else {
      console.log(`   ❌ FALHA! Não é uma data válida, será enviado: "---"`);
    }
  }
  
  console.log("\n" + "=".repeat(60));
  console.log("🏁 DEBUG CONCLUÍDO!");
  console.log("=".repeat(60));
}
