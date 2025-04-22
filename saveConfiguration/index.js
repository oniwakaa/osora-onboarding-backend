const { DefaultAzureCredential } = require('@azure/identity');
const { BlobServiceClient } = require('@azure/storage-blob');

module.exports = async function (context, req) {
    context.log('SharePoint Configuration Save function processed a request.');

    // REQUEST VALIDATION SECTION
    if (req.method !== 'POST') {
        context.log.error('Invalid HTTP method. Only POST is supported.');
        context.res = {
            status: 400,
            body: { error: "Invalid HTTP method. Only POST is supported." }
        };
        return;
    }

    if (!req.body) {
        context.log.error('Request body is missing.');
        context.res = {
            status: 400,
            body: { error: "Request body is missing." }
        };
        return;
    }

    const { tenantId, sharepointUrls, userDisplayName, userIdentifier } = req.body;
    // Extract userIdentifier from request body, with fallback to 'unknown'
    const sanitizedUserIdentifier = (userIdentifier || '').replace(/[^\w\s@.-]/g, '') || 'unknown';

    if (!tenantId || typeof tenantId !== 'string') {
        context.log.error('Invalid or missing tenantId in request body.');
        context.res = {
            status: 400,
            body: { error: "Invalid or missing tenantId in request body." }
        };
        return;
    }

    if (!sharepointUrls || !Array.isArray(sharepointUrls)) {
        context.log.error('Invalid or missing sharepointUrls in request body.');
        context.res = {
            status: 400,
            body: { error: "Invalid or missing sharepointUrls in request body. Must be an array." }
        };
        return;
    }

    // Validate each SharePoint URL
    const validUrls = sharepointUrls.filter(url => {
        try {
            new URL(url);
            return url.includes('.sharepoint.com/');
        } catch (e) {
            context.log.warn(`Invalid SharePoint URL detected: ${url}`);
            return false;
        }
    });

    // // USER AUTHENTICATION SECTION
    // const clientPrincipalHeader = req.headers['x-ms-client-principal'];
    // if (!clientPrincipalHeader) {
    //     context.log.error('Authentication header is missing.');
    //     context.res = {
    //         status: 401,
    //         body: { error: "Authentication required." }
    //     };
    //     return;
    // }

    // let clientPrincipal;
    // try {
    //     const decodedHeader = Buffer.from(clientPrincipalHeader, 'base64').toString('utf8');
    //     clientPrincipal = JSON.parse(decodedHeader);
    //     context.log(`User authenticated: ${clientPrincipal.userDetails}`);
    // } catch (error) {
    //     context.log.error(`Error parsing authentication header: ${error.message}`);
    //     context.res = {
    //         status: 401,
    //         body: { error: "Invalid authentication token." }
    //     };
    //     return;
    // }

    // if (!clientPrincipal || !clientPrincipal.userDetails) {
    //     context.log.error('Invalid client principal data.');
    //     context.res = {
    //         status: 401,
    //         body: { error: "Invalid authentication data." }
    //     };
    //     return;
    // }

    // AZURE STORAGE SECTION
    try {
        const credential = new DefaultAzureCredential();
        const accountName = process.env.AZURE_STORAGE_ACCOUNT_NAME;
        
        if (!accountName) {
            throw new Error("Storage account name environment variable is not configured.");
        }
        
        const blobServiceUrl = `https://${accountName}.blob.core.windows.net`;
        const blobServiceClient = new BlobServiceClient(blobServiceUrl, credential);
        
        const containerName = process.env.AZURE_STORAGE_CONTAINER_NAME;
        if (!containerName) {
            throw new Error("Storage container name environment variable is not configured.");
        }
        
        const containerClient = blobServiceClient.getContainerClient(containerName);
        
        // Ensure container exists before proceeding
        await containerClient.createIfNotExists();
        
        const blobName = `${tenantId}.json`;
        
        // Prepare configuration data
        const configData = {
            tenantId,
            sharepointSites: validUrls,
            timestamp: new Date().toISOString(),
            updatedBy: sanitizedUserIdentifier,
            updatedByDisplayName: userDisplayName || 'Unknown User'
        };
        
        // Upload to blob storage
        const blockBlobClient = containerClient.getBlockBlobClient(blobName);
        const data = JSON.stringify(configData);
        
        await blockBlobClient.upload(data, data.length, {
            blobHTTPHeaders: { blobContentType: 'application/json' }
        });
        
        context.log(`Configuration for tenant ${tenantId} saved successfully.`);
        
        // SUCCESS RESPONSE SECTION
        context.res = {
            status: 200,
            body: { message: "Configuration saved successfully" }
        };
    } catch (error) {
        context.log.error(`Error saving configuration: ${error.message}`);
        context.res = {
            status: 500,
            body: { error: `Failed to save configuration: ${error.message}` }
        };
    }
};