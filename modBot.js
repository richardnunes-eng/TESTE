function cicloPrincipal() {
  try {
    // Executa as funções na ordem
    FORCAR_RESET_COMPLETO();
    ExecutarIntegracaoMestre();
    sincronizarGreenMileStable();
  } catch (erro) {
    console.error("Erro no ciclo principal: " + erro);
  }

  // Após finalizar, cria um acionador para executar novamente em 1 minuto
  criarTrigger1MinutoDepois();
}

function criarTrigger1MinutoDepois() {
  // Remove triggers existentes dessa função
  removerTriggersDaFuncao("cicloPrincipal");

  // Cria um novo acionador para daqui a 1 minuto
  ScriptApp.newTrigger("cicloPrincipal")
    .timeBased()
    .after(1 * 60 * 1000) // 1 minuto em milissegundos
    .create();
}

function removerTriggersDaFuncao(nomeFuncao) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === nomeFuncao) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
