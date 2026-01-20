/**
 * ==============================================================================
 * ARQUIVO: Config.gs (CONFIGURAÇÕES GERAIS E EMOJIS)
 * ==============================================================================
 */



// --- CREDENCIAIS WHATSAPP (ATUALIZAR SE MUDAR) ---
const API_TOKEN = "668f2722c46b849674d7e01c";
const API_URL = "https://api.attemics.com.br/core/v2/api/chats/send-text";

// --- CREDENCIAIS GREENMILE ---
const GM_USERNAME = "richardthx";
const GM_PASSWORD = "GM@thx2025"; 
const GM_URL_BASE = "https://3coracoes.greenmile.com/StopView/Summary";

// --- CREDENCIAIS CLICKUP ---
const CLICKUP_TOKEN = "pk_87986690_9X1MC60UE18B1X9PEJFRMEFTT6GNHHFS";

// --- REGRAS ---
const LISTA_SUPERVISORES = ["35998182404", "14982271212", "41992312058"];
const LIMITE_ATRASO_MINUTOS = 40;
const INTERVALO_RELATORIO_MIN = 10; 
const INTERVALO_COBRANCA_MIN = 30;
const HORA_INICIO_EXPEDIENTE = 8; 
const HORA_FIM_EXPEDIENTE = 19;   

// --- LISTA DE EMOJIS COMPLETA ---
const EMOJI = {
  RED: "🔴", GREEN: "🟢", YELLOW: "🟡", 
  SUN: "☀️", TECHMAN: "👨‍💻", CLIP: "📋",
  PHONE: "📱", STOP: "🛑", MONEY: "💰", 
  HANDSHAKE: "🤝", TRUCK_FAST: "🚛💨",
  BULLET: "▪️", SIREN: "🚨", CAMERA: "📷", 
  SANDGLASS: "⏳", MEMO: "📝", TIME: "⏱️", 
  FLAG: "🏁", ROCKET: "🚀", ARROW: "👉", 
  BOX: "📦", OK: "✅", DRIVER: "👤", 
  WARNING: "⚠️", TROPHY: "🏆", 
  SYNC: "🔄",       
  NO_ENTRY: "🚫",   
  CALENDAR: "📅", CLOCK_FACE: "🕒", 
  PAPER: "📄", PIN: "📍", SMILE: "😄", 
  DINHEIRO: "💰", CHECK: "✅",
  
  // Perfis
  TRUCK: "🚛", VUC: "🚚", VAN: "🚐", 
  MOTO: "🛵", 
  
  // Manuais
  LUA: "🌙", PRATO: "🍽️", TCHAU: "👋"
};
