// osora-onboarding-api/checkAdminStatus/index.js

// *** MODIFICA: Usa DefaultAzureCredential ***
const { DefaultAzureCredential } = require("@azure/identity");
const { Client } = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");

const ADMIN_ROLE_TEMPLATE_IDS = [
    "62e90394-69f5-4237-9190-012177145e10", // Global Administrator
    "fe930be7-5e62-47db-91af-98c3a49a38b1", // SharePoint Administrator
    "29232cdf-9323-42fd-ade2-1d097af3e4de", // Teams Administrator
    "f2ef992c-3afb-46b9-b7cf-a126ee74c451", // Exchange Administrator
    "b0f54661-2d74-4c50-afa3-1ec803f12efe"  // User Administrator
];

// *** RIMOSSO: Non serve più l'ID Cliente esplicito ***
// const MANAGED_IDENTITY_CLIENT_ID = "1877d094-1a3c-4efc-bdfd-4c894cfa2a53";

module.exports = async function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request for checkAdminStatus.');

    let userId = null;
    let tenantId = null;
    let isAdmin = false;

    const clientPrincipalHeader = req.headers['x-ms-client-principal'];
    if (!clientPrincipalHeader) {
        context.log.warn("Missing x-ms-client-principal header.");
        context.res = { status: 401, body: { message: "Utente non autenticato." } };
        return;
    }

    try {
        const header = Buffer.from(clientPrincipalHeader, 'base64').toString('ascii');
        const clientPrincipal = JSON.parse(header);
        
        // Inspect the client principal structure
        context.log(`Client principal structure: ${JSON.stringify(clientPrincipal, null, 2)}`);
        
        userId = clientPrincipal.userId;
        context.log(`User OID extracted from header: ${userId}`);

        if (!userId) {
             throw new Error("UserID not found in client principal.");
        }
        
        // Extract tenant ID from claims
        if (clientPrincipal.claims && Array.isArray(clientPrincipal.claims)) {
            const tidClaim = clientPrincipal.claims.find(claim => claim.typ === 'tid' || claim.type === 'tid');
            if (tidClaim) {
                tenantId = tidClaim.val || tidClaim.value;
                context.log(`Tenant ID extracted from claims: ${tenantId}`);
            }
        }
        
        if (!tenantId) {
            throw new Error("Tenant ID not found in client principal claims.");
        }

        // *** MODIFICA: Ottieni la credenziale SENZA parametri ***
        // Lascia che DefaultAzureCredential scopra la MI assegnata al sistema
        // della Function App dall'ambiente.
        context.log(`Attempting to get credential using DefaultAzureCredential (auto-discovery)...`);
        const credential = new DefaultAzureCredential();

        context.log("Attempting to get token for Microsoft Graph scope 'https://graph.microsoft.com/.default'...");
        const tokenResponse = await credential.getToken("https://graph.microsoft.com/.default");
        context.log("Successfully obtained Graph token using Managed Identity.");

         const authProvider = {
            getAccessToken: async () => {
                context.log("Providing access token to Graph client.");
                // Nota: DefaultAzureCredential gestisce internamente il caching e il refresh,
                // quindi per chiamate successive potresti richiamare credential.getToken qui
                // per assicurarti un token fresco, ma per una singola chiamata va bene così.
                // Per semplicità, restituiamo il token appena ottenuto.
                // Se ci fossero problemi di scadenza, una logica più robusta richiamerebbe
                // getToken qui dentro invece di usare la variabile esterna tokenResponse.
                return tokenResponse.token;
            }
        };

        const graphClient = Client.initWithMiddleware({ authProvider: authProvider });

        // Use absolute URL with tenant ID instead of relative path
        const graphApiUrl = `https://graph.microsoft.com/v1.0/users/${userId}/transitiveMemberOf/microsoft.graph.directoryRole`;
        context.log(`Calling Graph API using absolute URL: '${graphApiUrl}' with tenant ID: ${tenantId}`);
        
        const roleMembership = await graphClient
            .api(graphApiUrl)
            .header('ConsistencyLevel', 'eventual')
            .select('id,displayName,roleTemplateId')
            .get();
            
        context.log("Graph API call successful.");

        if (roleMembership && roleMembership.value && roleMembership.value.length > 0) {
            context.log(`User belongs to ${roleMembership.value.length} directory roles. Checking against admin list...`);
            // Converti gli ID template admin in lowercase una sola volta per efficienza
            const lowerCaseAdminIds = ADMIN_ROLE_TEMPLATE_IDS.map(id => id.toLowerCase());
            isAdmin = roleMembership.value.some(role => {
                const roleTemplateIdLower = role.roleTemplateId?.toLowerCase(); // Usa optional chaining e converti
                const isRoleAdmin = roleTemplateIdLower && lowerCaseAdminIds.includes(roleTemplateIdLower);
                context.log(`- Checking role: ${role.displayName} (Template ID: ${role.roleTemplateId}) -> IsAdmin: ${isRoleAdmin}`);
                return isRoleAdmin;
            });
        } else {
            context.log("User does not belong to any directory roles.");
        }

        context.log(`Admin status determined: ${isAdmin}`);

        context.res = { status: 200, body: { isAdmin: isAdmin } };

    } catch (error) {
        // Logging dell'errore migliorato
        context.log.error("-------------------- ERROR Start --------------------");
        context.log.error(`Error in checkAdminStatus for user ${userId || 'UNKNOWN'} in tenant ${tenantId || 'UNKNOWN'}`);
        context.log.error(`Error Type: ${error.name}`);
        context.log.error(`Error Message: ${error.message}`);
        if (error.statusCode) context.log.error(`Status Code: ${error.statusCode}`);
        if (error.code) context.log.error(`Error Code: ${error.code}`);
        if (error.requestId) context.log.error(`Request ID: ${error.requestId}`);
        // Logga la causa se presente (utile per errori concatenati)
        if (error.cause) {
             try {
                 context.log.error("Cause:", JSON.stringify(error.cause, Object.getOwnPropertyNames(error.cause), 2));
             } catch (e) {
                 context.log.error("Cause (could not stringify):", error.cause);
             }
        }
        // Logga lo stack trace per il debug dettagliato
        if (error.stack) {
            context.log.error("Stack Trace:", error.stack);
        }
        context.log.error("-------------------- ERROR End ----------------------");

        // More descriptive error message based on specific failure reasons
        let errorMessage = "Errore interno del server durante la verifica dello stato amministratore.";
        if (error.message.includes("Tenant ID not found")) {
            errorMessage = "Impossibile determinare il tenant ID dell'utente. Verificare l'autenticazione.";
        } else if (error.message.includes("UserID not found")) {
            errorMessage = "Impossibile determinare l'ID utente. Verificare l'autenticazione.";
        }

        context.res = {
            status: 500,
            body: { message: errorMessage }
        };
    }
};