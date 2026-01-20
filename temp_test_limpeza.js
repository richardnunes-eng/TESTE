
const REGEX_EMOJI = /([\u2700-\u27BF]|[\uE000-\uF8FF]|\uD83C[\uDC00-\uDFFF]|\uD83D[\uDC00-\uDFFF]|[\u2011-\u26FF]|\uD83E[\uDD10-\uDDFF]|\u200D|\uFE0F)/g;

function limparNomeColuna(nome) {
  if (!nome) return "";
  return nome.toString()
    .replace(REGEX_EMOJI, '') // Remove emojis
    .replace(/[^\w\s\-\(\)\[\]\.]/g, '') // Permite também pontos
    .replace(/\s+/g, ' ') // Normaliza espaços
    .replace(/^\s+|\s+$/g, '') // Remove espaços das bordas
    .replace(/^[\d\-\.]+$/, 'Campo_' + nome) // Se for só números, adiciona prefixo
    .substring(0, 100); // Limita tamanho do cabeçalho
}

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
    "Saída",
    "Localização",
    ""
  ];
  
  console.log("=== TESTE DE LIMPEZA DE COLUNAS (ANTES) ===");
  exemplos.forEach(exemplo => {
    const limpo = limparNomeColuna(exemplo);
    console.log(`"${exemplo}" → "${limpo}"`);
  });
}

testarLimpezaColunas();
