# Frontend Changes Required

## 1. Modifiche alla Chiamata `checkAdminStatus`

La funzione `checkAdminStatus` sul backend è stata modificata per accettare `userId` e `tenantId` come parametri anziché estrarli dall'header `x-ms-client-principal`.

### Esempio di modifica nel frontend:

```javascript
// Prima
async function checkAdminStatus() {
  const data = await callApi('/api/checkAdminStatus');
  // resto della logica
}

// Dopo
async function checkAdminStatus(userId, tenantId) {
  // Chiamata con query parameters
  const data = await callApi(`/api/checkAdminStatus?userId=${encodeURIComponent(userId)}&tenantId=${encodeURIComponent(tenantId)}`);
  
  // Alternativa: chiamata POST con body JSON
  /*
  const response = await fetch('/api/checkAdminStatus', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ userId, tenantId })
  });
  const data = await response.json();
  */
  
  // resto della logica per mostrare/nascondere UI in base a isAdmin
  return data.isAdmin;
}
```

## 2. Modifica a `saveConfiguration`

Assicurarsi che la funzione `saveConfiguration` utilizzi il `currentTenantId` che è stato popolato durante `handleSuccessfulToken`. La chiamata a `/api/saveConfiguration` che invia il `tenantId` nel corpo rimane invariata.

### Esempio:

```javascript
async function saveConfiguration(sharepointUrls) {
  // Assicurarsi che currentTenantId sia stato popolato
  if (!currentTenantId) {
    console.error('Tentativo di salvare la configurazione senza tenantId');
    throw new Error('Tentativo di salvare la configurazione senza tenantId');
  }
  
  const response = await fetch('/api/saveConfiguration', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      tenantId: currentTenantId,
      sharepointUrls: sharepointUrls
    })
  });
  
  return await response.json();
}
```

## 3. Implementazione Logout MSAL

Aggiungere un pulsante di logout e relativa funzione:

```javascript
// Aggiungi un pulsante nel tuo HTML/JSX
// <button id="logoutButton">Logout</button>

// Funzione di logout
function logout() {
  msalInstance.logoutRedirect();
  // In alternativa, per una esperienza popup
  // msalInstance.logoutPopup();
}

// Event listener per il pulsante di logout
document.getElementById('logoutButton').addEventListener('click', logout);
```

## 4. Modifiche al Login/Autenticazione

Assicurarsi che dopo un login riuscito, siano salvati `userId` e `tenantId` per le chiamate API successive:

```javascript
// Esempio di funzione handleSuccessfulToken aggiornata
async function handleSuccessfulToken(response) {
  // Estrai informazioni dell'utente dal token
  const account = response.account;
  const idTokenClaims = account.idTokenClaims;
  
  // Salva userId e tenantId
  currentUserId = account.homeAccountId || idTokenClaims.oid || idTokenClaims.sub;
  currentTenantId = account.tenantId || idTokenClaims.tid;
  
  // Ora verifica se l'utente è un amministratore
  const adminStatus = await checkAdminStatus(currentUserId, currentTenantId);
  
  // Mostra l'UI appropriata basata sullo stato admin
  if (adminStatus) {
    showAdminUI();
  } else {
    showRegularUserUI();
  }
}
```

## 5. Note Importanti

- Le funzioni del frontend devono essere adattate in base alla struttura dell'app esistente
- Assicurarsi che sia presente una gestione errori adeguata
- Ricordarsi di passare `userId` e `tenantId` in tutti i punti dove `checkAdminStatus` viene chiamato
- Il backend continua a verificare l'header `x-ms-client-principal` come controllo di sicurezza aggiuntivo 