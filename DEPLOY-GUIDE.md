# Guia de Deploy e Configuração - THX LOG Dashboard

## ⚠️ IMPORTANTE: Apps Script e iframes

**O Google Apps Script Web App NÃO suporta OAuth em iframe por política de segurança.**

Isso é uma limitação da plataforma Google (frame-ancestors, X-Frame-Options) e **não pode ser removida via código**.

### Solução Implementada

✅ **Escape automático de iframe**
- Detecta se está rodando em iframe
- Tenta redirecionar `window.top` automaticamente
- Abre nova aba automaticamente se redirecionamento falhar
- Mostra tela com botão manual apenas como último recurso

---

## 📋 Checklist de Implementação

### A) ✅ Detecção e Escape de Iframe
- [x] Função `attemptIframeEscape()` implementada
- [x] Tentativa de redirecionamento automático do `window.top`
- [x] Abertura automática em nova aba (fallback)
- [x] Tela de erro com botões manuais (último recurso)

### B) ✅ Fluxo de Autenticação Sem OAuth Dialog
- [x] Função `checkAuth()` retorna estrutura padrão `{ok, data, error, ts}`
- [x] Valida sessão via `SpreadsheetApp.getActiveSpreadsheet().getId()`
- [x] Obtém email do usuário (quando permitido pela organização)
- [x] Tratamento de erro quando política bloqueia acesso ao email
- [x] Removidas todas as referências a `userCodeAppPanel` e `createOAuthDialog`

### C) ✅ Padronização de Respostas do Backend
- [x] Wrapper `apiResponse(ok, data, error)` criado
- [x] Todas as funções retornam `{ok, data, error, ts}`
- [x] Front-end valida estrutura via `validateApiResponse()`
- [x] Tratamento específico para erros de autenticação (`needsAuth`)

---

## 🔧 Arquivos Alterados

### 1. **modDashboard.js** (Backend)

**Mudanças:**
- Adicionado wrapper `apiResponse(ok, data, error)` (linhas 17-24)
- Função `checkAuth()` atualizada para retornar estrutura padrão (linhas 45-77)
- Todas as funções padronizadas:
  - `getDashboardData()` → retorna `apiResponse(true, payload, null)`
  - `exportDashboardCsv()` → retorna `apiResponse(true, {url, downloadUrl, ...}, null)`
  - `salvarOcorrencia()` → retorna `apiResponse(true, {saved: true}, null)`
  - `finalizarTarefaBackend()` → retorna `apiResponse(true, {finalized: true}, null)`
  - `atualizarStatusClickupBackend()` → retorna `apiResponse(true, {status, color}, null)`

**Documentação adicionada** (linhas 9-48):
- Explicação sobre limitação de iframe
- Instruções de deploy
- Configuração do `appsscript.json`
- Alternativa para embed (front-end separado)

### 2. **JS-Logica.html** (Front-end)

**Mudanças:**
- Função `validateApiResponse(response, context)` criada (linhas 53-92)
- Função `attemptIframeEscape()` implementada (linhas 245-277)
- Função `checkAuthAndStart()` atualizada com validação (linhas 326-351)
- Função `loadData()` atualizada com validação (linhas 502-527)
- Função `exportCsv()` atualizada com validação (linhas 897-918)
- Função `setClickupStatus()` atualizada com validação (linhas 1787-1811)
- Função `actionFinalizar()` atualizada com validação (linhas 1877-1900)

**Correções de bugs:**
- Template literals convertidos para concatenação de strings (linhas 1727, 1832)
- Marcadores de conflito Git removidos (linhas 1871-1879)

---

## 🚀 Como Fazer Deploy

### 1. **Configuração do appsscript.json**

Verifique se o arquivo `appsscript.json` está configurado corretamente:

```json
{
  "timeZone": "America/Sao_Paulo",
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "webapp": {
    "executeAs": "USER_DEPLOYING",
    "access": "ANYONE_ANONYMOUS"
  },
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/forms",
    "https://www.googleapis.com/auth/documents"
  ]
}
```

**Explicação:**
- `"executeAs": "USER_DEPLOYING"` → O script roda com as permissões de quem fez o deploy
- `"access": "ANYONE_ANONYMOUS"` → Permite acesso sem login Google (qualquer pessoa com o link)

### 2. **Push do Código**

```bash
# Na raiz do projeto (onde está o .clasp.json)
clasp push
```

### 3. **Criar Deployment**

**Opção A: Via clasp (CLI)**
```bash
clasp deploy --description "v1.0 - Producao"
```

**Opção B: Via Apps Script Editor**
1. Abra o projeto no Apps Script Editor
2. Clique em **Deploy** > **New Deployment**
3. Selecione **Web App**
4. Configure:
   - **Execute as:** `Me (seu email)`
   - **Who has access:** `Anyone` ou `Anyone with Google account`
5. Clique em **Deploy**
6. **Copie a URL do Web App**

### 4. **Obter URL do Web App**

A URL será no formato:
```
https://script.google.com/macros/s/{SCRIPT_ID}/exec
```

**Onde usar esta URL:**
- Esta é a URL que você distribui aos usuários
- O próprio script obtém automaticamente via `ScriptApp.getService().getUrl()`
- Não é necessário hardcode em lugar nenhum

---

## 🔍 Onde Configurar WEBAPP_URL (Opcional)

**Resposta curta:** Não é necessário configurar manualmente.

**Explicação:**
O script obtém automaticamente a URL do Web App via:
```javascript
// Backend (modDashboard.js)
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// Front-end (JS-Logica.html)
function fetchWebAppUrl(onSuccess, onFailure) {
  if (webAppUrlCache) {
    onSuccess(webAppUrlCache);
    return;
  }
  google.script.run
    .withSuccessHandler(url => {
      webAppUrlCache = String(url || '');
      if (webAppUrlCache) onSuccess(webAppUrlCache);
      else if (onFailure) onFailure(new Error('URL vazia'));
    })
    .withFailureHandler(err => {
      if (onFailure) onFailure(err);
    })
    .getWebAppUrl();
}
```

Se você quiser **hardcode** a URL (não recomendado), pode fazer em `JS-Logica.html`:
```javascript
// No início do arquivo, após as constantes
const WEBAPP_URL = "https://script.google.com/macros/s/SEU_SCRIPT_ID_AQUI/exec";
```

---

## ✅ Checklist Final (Obrigatório)

Antes de dar o deploy final, verifique:

- [ ] **Abrir o Web App diretamente funciona e autentica**
  - Teste: Abra a URL do Web App em uma nova aba
  - Resultado esperado: Dashboard carrega normalmente

- [ ] **Abrir dentro de iframe não trava: tenta sair automaticamente**
  - Teste: Embuta a URL em um `<iframe>` em outra página
  - Resultado esperado: Redireciona automaticamente ou abre nova aba

- [ ] **Nenhuma chamada retorna vazio; sempre retorna `{ok, data, error, ts}`**
  - Teste: Verifique o console do navegador durante operações
  - Resultado esperado: Todas as respostas têm estrutura padrão

- [ ] **Nenhuma URL `userCodeAppPanel/createOAuthDialog` permanece no projeto**
  - Teste: Busque no projeto por essas strings
  - Resultado esperado: Nenhum resultado encontrado ✅

- [ ] **Documentação atualizada no código**
  - Verificar comentário no topo de `modDashboard.js` ✅

---

## 🛠️ Troubleshooting

### Problema: "Autorização necessária" mesmo após autorizar

**Causa:** O deployment pode estar configurado como `USER_ACCESSING` em vez de `USER_DEPLOYING`.

**Solução:**
1. Abra `appsscript.json`
2. Altere `"executeAs": "USER_DEPLOYING"`
3. Faça `clasp push` novamente
4. Crie um novo deployment

### Problema: Fica preso na tela "Redirecionando..."

**Causa:** Browser está bloqueando pop-ups ou redirecionamento.

**Solução:**
1. Permita pop-ups para o domínio `script.google.com`
2. Clique manualmente no botão "Abrir em nova aba"
3. Copie o link e abra em uma nova aba do navegador

### Problema: "Resposta vazia do servidor"

**Causa:** Alguma função do backend não está retornando estrutura `apiResponse`.

**Solução:**
1. Verifique os logs no Apps Script (View > Logs)
2. Identifique qual função está retornando vazio
3. Certifique-se de que usa `return apiResponse(ok, data, error)`

---

## 🔮 Alternativa: Front-end Externo + Apps Script como API

Se você **realmente precisa** de embed em iframe (ex.: dentro de um sistema interno), considere:

### Arquitetura Recomendada:
```
┌─────────────────────────────────────┐
│  Front-end (Firebase/Vercel/CF)    │
│  - HTML/CSS/JS seu domínio          │
│  - Google Identity Services (GIS)   │
│  - Pode ser embutido em iframe      │
└─────────────┬───────────────────────┘
              │ API REST
              ↓
┌─────────────────────────────────────┐
│  Google Apps Script (Backend API)   │
│  - doGet/doPost retorna JSON        │
│  - Valida token do GIS              │
│  - Acessa Sheets/Drive/etc          │
└─────────────────────────────────────┘
```

### Passos:
1. Crie front-end separado em Firebase Hosting/Vercel/Cloudflare Pages
2. Use Google Identity Services para login
3. Envie ID token ao Apps Script
4. Apps Script valida token e retorna dados
5. Front-end pode ser embutido em iframe (seu domínio, suas regras CSP)

**Referência:** [Google Identity Services](https://developers.google.com/identity/gsi/web/guides/overview)

---

## 📞 Suporte

Se encontrar problemas:
1. Verifique os logs no Apps Script Editor (View > Logs ou Ctrl+Enter)
2. Abra o console do navegador (F12) e veja erros JavaScript
3. Revise este guia e o checklist
4. Consulte a documentação oficial: https://developers.google.com/apps-script/guides/web

---

**Última atualização:** 2026-01-18
**Versão do projeto:** 2.0 (Com escape de iframe e validação de respostas)
