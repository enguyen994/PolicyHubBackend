// ============================================
// AZURE FUNCTION 1: Search Policies
// File: SearchPolicies/index.js
// ============================================

const { SearchClient, AzureKeyCredential } = require("@azure/search-documents");

module.exports = async function (context, req) {
    context.log('Processing policy search request');

    try {
        // Get search parameters from query string or body
        const searchQuery = req.query.q || req.body?.q || "*";
        const category = req.query.category;
        const department = req.query.department;
        const page = parseInt(req.query.page) || 1;
        const pageSize = parseInt(req.query.pageSize) || 10;

        // Initialize Azure Cognitive Search client
        const searchClient = new SearchClient(
            process.env.SEARCH_ENDPOINT,
            process.env.SEARCH_INDEX_NAME,
            new AzureKeyCredential(process.env.SEARCH_API_KEY)
        );

        // Build search options
        const searchOptions = {
            searchMode: "all",
            queryType: "full",
            skip: (page - 1) * pageSize,
            top: pageSize,
            includeTotalCount: true,
            highlightFields: "name,description,content",
            highlightPreTag: "<mark>",
            highlightPostTag: "</mark>",
            select: [
                "id",
                "name", 
                "category",
                "department",
                "effectiveDate",
                "version",
                "description",
                "fileUrl",
                "lastModified",
                "owner",
                "accessLevel"
            ],
            facets: ["category", "department", "effectiveDate"],
            orderBy: ["@search.score() desc", "lastModified desc"]
        };

        // Add filters based on user input
        const filters = [];
        
        if (category && category !== "All") {
            filters.push(`category eq '${category}'`);
        }
        
        if (department && department !== "All") {
            filters.push(`department eq '${department}'`);
        }

        // Add role-based access control filter
        const userRole = req.headers['x-user-role'] || 'Employee';
        const userEmail = req.headers['x-user-email'];
        
        if (userRole !== 'Admin') {
            // Filter based on access level
            filters.push(`(accessLevel eq 'All Employees' or accessLevel eq '${userRole}')`);
        }

        if (filters.length > 0) {
            searchOptions.filter = filters.join(" and ");
        }

        // Execute search
        const searchResults = await searchClient.search(searchQuery, searchOptions);

        // Process results
        const results = [];
        for await (const result of searchResults.results) {
            results.push({
                ...result.document,
                highlights: result.highlights,
                score: result.score
            });
        }

        // Get facets for filters
        const facets = {};
        if (searchResults.facets) {
            for (const [key, value] of Object.entries(searchResults.facets)) {
                facets[key] = value.map(f => ({
                    value: f.value,
                    count: f.count
                }));
            }
        }

        // Return successful response
        context.res = {
            status: 200,
            headers: {
                'Content-Type': 'application/json'
            },
            body: {
                results,
                totalCount: searchResults.count,
                facets,
                page,
                pageSize,
                query: searchQuery
            }
        };

        // Log search for analytics
        context.log(`Search completed: "${searchQuery}" - ${results.length} results`);

    } catch (error) {
        context.log.error('Search error:', error);
        context.res = {
            status: 500,
            body: { 
                error: 'Failed to execute search', 
                details: error.message 
            }
        };
    }
};

// ============================================
// AZURE FUNCTION 2: Get Policy Details
// File: GetPolicy/index.js
// ============================================

const { Client } = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");

function getAuthenticatedClient(accessToken) {
    return Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });
}

module.exports = async function (context, req) {
    context.log('Fetching policy details');

    try {
        const policyId = req.params.id;
        const accessToken = req.headers.authorization?.replace('Bearer ', '');

        if (!accessToken) {
            context.res = {
                status: 401,
                body: { error: 'No authorization token provided' }
            };
            return;
        }

        const client = getAuthenticatedClient(accessToken);

        // SharePoint site and drive information
        const siteId = process.env.SHAREPOINT_SITE_ID;
        const driveId = process.env.SHAREPOINT_DRIVE_ID;

        // Get file metadata from SharePoint
        const fileMetadata = await client
            .api(`/sites/${siteId}/drives/${driveId}/items/${policyId}`)
            .expand('listItem($expand=fields)')
            .get();

        // Get download URL
        const downloadUrl = await client
            .api(`/sites/${siteId}/drives/${driveId}/items/${policyId}`)
            .select('@microsoft.graph.downloadUrl')
            .get();

        // Extract custom SharePoint fields
        const fields = fileMetadata.listItem?.fields || {};

        // Return policy details
        context.res = {
            status: 200,
            body: {
                id: fileMetadata.id,
                name: fields.PolicyName || fileMetadata.name,
                category: fields.Category,
                department: fields.Department,
                effectiveDate: fields.EffectiveDate,
                version: fields.Version,
                description: fields.Description,
                owner: fields.Owner,
                accessLevel: fields.AccessLevel,
                status: fields.Status,
                reviewDate: fields.ReviewDate,
                fileUrl: fileMetadata.webUrl,
                downloadUrl: downloadUrl['@microsoft.graph.downloadUrl'],
                lastModified: fileMetadata.lastModifiedDateTime,
                size: fileMetadata.size,
                mimeType: fileMetadata.file?.mimeType
            }
        };

        context.log(`Policy retrieved: ${policyId}`);

    } catch (error) {
        context.log.error('Get policy error:', error);
        
        if (error.statusCode === 404) {
            context.res = {
                status: 404,
                body: { error: 'Policy not found' }
            };
        } else if (error.statusCode === 403) {
            context.res = {
                status: 403,
                body: { error: 'Access denied to this policy' }
            };
        } else {
            context.res = {
                status: 500,
                body: { 
                    error: 'Failed to fetch policy', 
                    details: error.message 
                }
            };
        }
    }
};

// ============================================
// AZURE FUNCTION 3: Index Documents
// File: IndexDocuments/index.js
// ============================================

const { SearchClient, AzureKeyCredential } = require("@azure/search-documents");
const { Client } = require("@microsoft/microsoft-graph-client");
const { ClientSecretCredential } = require("@azure/identity");
const pdf = require('pdf-parse');

// Get Microsoft Graph client with app-only authentication
function getGraphClientWithAppAuth() {
    const credential = new ClientSecretCredential(
        process.env.AZURE_AD_TENANT_ID,
        process.env.AZURE_AD_CLIENT_ID,
        process.env.AZURE_AD_CLIENT_SECRET
    );

    return Client.initWithMiddleware({
        authProvider: {
            getAccessToken: async () => {
                const token = await credential.getToken(['https://graph.microsoft.com/.default']);
                return token.token;
            }
        }
    });
}

module.exports = async function (context, req) {
    context.log('Starting document indexing from SharePoint');

    try {
        // Get Graph client with app permissions
        const graphClient = getGraphClientWithAppAuth();

        // SharePoint configuration
        const siteId = process.env.SHAREPOINT_SITE_ID;
        const driveId = process.env.SHAREPOINT_DRIVE_ID;

        // Get all documents from SharePoint library
        context.log('Fetching documents from SharePoint...');
        const items = await graphClient
            .api(`/sites/${siteId}/drives/${driveId}/root/children`)
            .expand('listItem($expand=fields)')
            .top(999)
            .get();

        context.log(`Found ${items.value.length} items in SharePoint`);

        // Initialize Azure Cognitive Search client
        const searchClient = new SearchClient(
            process.env.SEARCH_ENDPOINT,
            process.env.SEARCH_INDEX_NAME,
            new AzureKeyCredential(process.env.SEARCH_API_KEY)
        );

        const documentsToIndex = [];
        let processedCount = 0;
        let errorCount = 0;

        // Process each file
        for (const item of items.value) {
            try {
                // Skip folders
                if (item.folder) {
                    context.log(`Skipping folder: ${item.name}`);
                    continue;
                }

                // Get SharePoint list item fields (metadata)
                const fields = item.listItem?.fields || {};
                
                // Extract text from PDF if applicable
                let textContent = '';
                if (item.file?.mimeType === 'application/pdf') {
                    try {
                        context.log(`Extracting text from PDF: ${item.name}`);
                        
                        // Get file content
                        const fileBuffer = await graphClient
                            .api(`/sites/${siteId}/drives/${driveId}/items/${item.id}/content`)
                            .getStream();

                        // Convert stream to buffer
                        const chunks = [];
                        for await (const chunk of fileBuffer) {
                            chunks.push(chunk);
                        }
                        const buffer = Buffer.concat(chunks);

                        // Extract text from PDF
                        const pdfData = await pdf(buffer);
                        textContent = pdfData.text;
                        
                        context.log(`Extracted ${textContent.length} characters from ${item.name}`);
                    } catch (pdfError) {
                        context.log.warn(`Failed to extract PDF text from ${item.name}: ${pdfError.message}`);
                    }
                }

                // Create search document
                const searchDocument = {
                    id: item.id,
                    name: fields.PolicyName || item.name.replace(/\.[^/.]+$/, ''), // Remove extension
                    category: fields.Category || 'General',
                    department: fields.Department || 'General',
                    effectiveDate: fields.EffectiveDate || item.createdDateTime,
                    version: fields.Version || '1.0',
                    description: fields.Description || '',
                    fileUrl: item.webUrl,
                    lastModified: item.lastModifiedDateTime,
                    owner: fields.Owner || 'Unknown',
                    accessLevel: fields.AccessLevel || 'All Employees',
                    status: fields.Status || 'Active',
                    reviewDate: fields.ReviewDate || null,
                    content: textContent,
                    fileType: item.file?.mimeType || 'unknown',
                    fileSize: item.size
                };

                documentsToIndex.push(searchDocument);
                processedCount++;

            } catch (itemError) {
                context.log.error(`Error processing item ${item.name}:`, itemError);
                errorCount++;
            }
        }

        // Upload documents to Azure Cognitive Search in batches
        context.log(`Uploading ${documentsToIndex.length} documents to search index...`);
        
        const batchSize = 100;
        let uploadedCount = 0;
        
        for (let i = 0; i < documentsToIndex.length; i += batchSize) {
            const batch = documentsToIndex.slice(i, i + batchSize);
            
            try {
                const result = await searchClient.uploadDocuments(batch);
                
                const succeeded = result.results.filter(r => r.succeeded).length;
                const failed = result.results.filter(r => !r.succeeded).length;
                
                uploadedCount += succeeded;
                
                if (failed > 0) {
                    context.log.warn(`Batch ${i / batchSize + 1}: ${succeeded} succeeded, ${failed} failed`);
                } else {
                    context.log(`Batch ${i / batchSize + 1}: ${succeeded} documents uploaded`);
                }
            } catch (batchError) {
                context.log.error(`Error uploading batch ${i / batchSize + 1}:`, batchError);
                errorCount += batch.length;
            }
        }

        // Return summary
        context.res = {
            status: 200,
            body: {
                message: 'Indexing completed',
                summary: {
                    totalItemsInSharePoint: items.value.length,
                    documentsProcessed: processedCount,
                    documentsIndexed: uploadedCount,
                    errors: errorCount,
                    timestamp: new Date().toISOString()
                }
            }
        };

        context.log(`Indexing complete: ${uploadedCount} documents indexed, ${errorCount} errors`);

    } catch (error) {
        context.log.error('Indexing error:', error);
        context.res = {
            status: 500,
            body: { 
                error: 'Failed to index documents', 
                details: error.message 
            }
        };
    }
};

// ============================================
// AZURE FUNCTION 4: Log Audit Event
// File: LogAudit/index.js
// ============================================

const { TableClient, AzureNamedKeyCredential } = require("@azure/data-tables");

module.exports = async function (context, req) {
    context.log('Logging audit event');

    try {
        const { userId, policyId, action, metadata } = req.body;

        if (!userId || !action) {
            context.res = {
                status: 400,
                body: { error: 'userId and action are required' }
            };
            return;
        }

        // Create Azure Table Storage client
        const credential = new AzureNamedKeyCredential(
            process.env.STORAGE_ACCOUNT_NAME,
            process.env.STORAGE_ACCOUNT_KEY
        );

        const tableClient = new TableClient(
            `https://${process.env.STORAGE_ACCOUNT_NAME}.table.core.windows.net`,
            "PolicyAuditLogs",
            credential
        );

        // Ensure table exists
        await tableClient.createTable().catch(() => {
            // Table might already exist, ignore error
        });

        // Create audit entry
        const timestamp = new Date();
        const rowKey = `${timestamp.getTime()}_${Math.random().toString(36).substr(2, 9)}`;

        const auditEntry = {
            partitionKey: userId,
            rowKey: rowKey,
            userId: userId,
            policyId: policyId || 'N/A',
            action: action, // 'view', 'download', 'search'
            timestamp: timestamp.toISOString(),
            metadata: JSON.stringify(metadata || {}),
            ipAddress: req.headers['x-forwarded-for'] || req.connection?.remoteAddress || 'unknown',
            userAgent: req.headers['user-agent'] || 'unknown'
        };

        // Insert into table
        await tableClient.createEntity(auditEntry);

        context.res = {
            status: 200,
            body: { 
                message: 'Audit log created successfully',
                logId: rowKey
            }
        };

        context.log(`Audit logged: ${userId} - ${action} - ${policyId || 'N/A'}`);

    } catch (error) {
        context.log.error('Audit log error:', error);
        context.res = {
            status: 500,
            body: { 
                error: 'Failed to create audit log', 
                details: error.message 
            }
        };
    }
};

// ============================================
// FUNCTION.JSON FILES
// ============================================

// SearchPolicies/function.json
/*
{
  "bindings": [
    {
      "authLevel": "function",
      "type": "httpTrigger",
      "direction": "in",
      "name": "req",
      "methods": ["get", "post"],
      "route": "SearchPolicies"
    },
    {
      "type": "http",
      "direction": "out",
      "name": "res"
    }
  ]
}
*/

// GetPolicy/function.json
/*
{
  "bindings": [
    {
      "authLevel": "function",
      "type": "httpTrigger",
      "direction": "in",
      "name": "req",
      "methods": ["get"],
      "route": "GetPolicy/{id}"
    },
    {
      "type": "http",
      "direction": "out",
      "name": "res"
    }
  ]
}
*/

// IndexDocuments/function.json - HTTP trigger for manual indexing
/*
{
  "bindings": [
    {
      "authLevel": "function",
      "type": "httpTrigger",
      "direction": "in",
      "name": "req",
      "methods": ["post"],
      "route": "IndexDocuments"
    },
    {
      "type": "http",
      "direction": "out",
      "name": "res"
    }
  ]
}
*/

// IndexDocuments/function.json - Timer trigger for automatic nightly indexing
/*
{
  "bindings": [
    {
      "name": "myTimer",
      "type": "timerTrigger",
      "direction": "in",
      "schedule": "0 0 2 * * *"
    }
  ]
}
*/

// LogAudit/function.json
/*
{
  "bindings": [
    {
      "authLevel": "function",
      "type": "httpTrigger",
      "direction": "in",
      "name": "req",
      "methods": ["post"],
      "route": "LogAudit"
    },
    {
      "type": "http",
      "direction": "out",
      "name": "res"
    }
  ]
}
*/

// ============================================
// CONFIGURATION FILES
// ============================================

// package.json
/*
{
  "name": "policyhub-backend",
  "version": "1.0.0",
  "description": "PolicyHub Azure Functions Backend",
  "scripts": {
    "start": "func start",
    "test": "jest"
  },
  "dependencies": {
    "@azure/search-documents": "^12.0.0",
    "@azure/data-tables": "^13.2.2",
    "@azure/identity": "^3.3.0",
    "@microsoft/microsoft-graph-client": "^3.0.7",
    "isomorphic-fetch": "^3.0.0",
    "pdf-parse": "^1.1.1"
  },
  "devDependencies": {
    "azure-functions-core-tools": "^4.0.5",
    "jest": "^29.7.0"
  }
}
*/

// host.json
/*
{
  "version": "2.0",
  "logging": {
    "applicationInsights": {
      "samplingSettings": {
        "isEnabled": true,
        "maxTelemetryItemsPerSecond": 20
      }
    }
  },
  "extensionBundle": {
    "id": "Microsoft.Azure.Functions.ExtensionBundle",
    "version": "[4.*, 5.0.0)"
  },
  "functionTimeout": "00:10:00"
}
*/

// local.settings.json (for local development)
/*
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    "FUNCTIONS_WORKER_RUNTIME": "node",
    "SEARCH_ENDPOINT": "https://your-search-service.search.windows.net",
    "SEARCH_INDEX_NAME": "policies-index",
    "SEARCH_API_KEY": "your-search-admin-key",
    "SHAREPOINT_SITE_ID": "your-sharepoint-site-id",
    "SHAREPOINT_DRIVE_ID": "your-sharepoint-drive-id",
    "STORAGE_ACCOUNT_NAME": "policyhubstorage",
    "STORAGE_ACCOUNT_KEY": "your-storage-account-key",
    "AZURE_AD_TENANT_ID": "your-azure-ad-tenant-id",
    "AZURE_AD_CLIENT_ID": "your-azure-ad-client-id",
    "AZURE_AD_CLIENT_SECRET": "your-azure-ad-client-secret"
  },
  "Host": {
    "CORS": "*"
  }
}
*/