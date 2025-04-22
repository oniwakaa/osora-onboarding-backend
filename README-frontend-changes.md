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

La funzione `saveConfiguration` sul backend è stata modificata per non dipendere più dall'header `x-ms-client-principal`. Ora accetta due parametri aggiuntivi nel corpo della richiesta:

- `userDisplayName`: Il nome visualizzato dell'utente (per scopi di visualizzazione)
- `userIdentifier`: Un identificatore univoco dell'utente (per tracciamento e audit)

### Esempio di modifica nel frontend:

```javascript
async function saveConfiguration(sharepointUrls) {
  // Assicurarsi che currentTenantId sia stato popolato
  if (!currentTenantId) {
    console.error('Tentativo di salvare la configurazione senza tenantId');
    throw new Error('Tentativo di salvare la configurazione senza tenantId');
  }
  
  // Ottieni il nome utente e l'identificatore dall'account MSAL
  const currentAccount = msalInstance.getActiveAccount();
  const userDisplayName = currentAccount ? currentAccount.name : 'Unknown User';
  const userIdentifier = currentAccount ? (currentAccount.homeAccountId || currentAccount.username) : 'unknown';
  
  const response = await fetch('/api/saveConfiguration', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      tenantId: currentTenantId,
      sharepointUrls: sharepointUrls,
      userDisplayName: userDisplayName, // Aggiungi il nome dell'utente
      userIdentifier: userIdentifier    // Aggiungi l'identificatore dell'utente
    })
  });
  
  return await response.json();
}
```

### Informazioni salvate nel backend:

```javascript
// Nel backend, il configData ora include:
const configData = {
  tenantId,
  sharepointSites: validUrls,
  timestamp: new Date().toISOString(),
  updatedBy: userIdentifier,             // Identificatore univoco dell'utente
  updatedByDisplayName: userDisplayName  // Nome visualizzato dell'utente
};
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

- Le funzioni del backend sono state modificate per non dipendere più dall'header `x-ms-client-principal`
- Assicurarsi che sia presente una gestione errori adeguata
- Ricordarsi di passare `userId` e `tenantId` in tutti i punti dove `checkAdminStatus` viene chiamato
- Per `saveConfiguration`, assicurarsi di passare anche `userDisplayName` e `userIdentifier` dal frontend
- L'identificatore utente viene sanitizzato nel backend prima di essere salvato 