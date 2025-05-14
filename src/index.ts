import express, { Request, Response } from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";

// Environment variables configuration
const config = {
  // Authentication settings
  tenantId: process.env.TENANT_ID || "",
  clientId: process.env.CLIENT_ID || "",
  clientSecret: process.env.CLIENT_SECRET || "",
  grantType: process.env.GRANT_TYPE || "client_credentials",
  defaultScope: process.env.DEFAULT_SCOPE || "0cdb527f-a8d1-4bf8-9436-b352c68682b2/.default",
  
  // Dynamics 365 settings
  fnoId: process.env.FNO_ID || "",
  
  // Service URLs
  azureAuthUrl: process.env.AZURE_AUTH_URL || "https://login.microsoftonline.com",
  dynamicsTokenUrl: process.env.DYNAMICS_TOKEN_URL || "https://securityservice.operations365.dynamics.com/token",
  inventoryServiceUrl: process.env.INVENTORY_SERVICE_URL || "https://inventoryservice.wus-il301.gateway.prod.island.powerapps.com",
  
  // Server configuration
  port: process.env.PORT || 3001,
  serviceBaseUrl: process.env.SERVICE_BASE_URL || "",
};

// Type definitions
interface AzureADTokenParams {
  tenant_id: string;
  client_id: string;
  client_secret: string;
  grant_type: string;
}

interface DynamicsTokenParams {
  bearer_token: string;
  grant_type: string;
  fno_id: string;
}

interface InventoryQueryParams {
  access_token: string;  // This should be the Dynamics token, NOT the Azure AD token
  fno_id: string;
  product_id: string;
  organization_id: string;
}

interface AuthenticateDynamicsParams {
  tenant_id: string;
  client_id: string;
  client_secret: string;
  grant_type: string;
  fno_id: string;
}

// Initialize the MCP server
const server = new McpServer({
  name: "jokesMCP",
  description: "A server that provides jokes and Dynamics 365 inventory information",
  version: "1.0.0",
  tools: [
    {
      name: "get-chuck-joke",
      description: "Get a random Chuck Norris joke",
      parameters: {},
    },
    {
      name: "get-chuck-categories",
      description: "Get all available categories for Chuck Norris jokes",
      parameters: {},
    },
    {
      name: "get-dad-joke",
      description: "Get a random dad joke",
      parameters: {},
    },
    {
      name: "get-yo-mama-joke",
      description: "Get a random Yo Mama joke",
      parameters: {},
    },
    {
      name: "get-azure-ad-token",
      description: "Get Azure AD token from Microsoft",
      parameters: {
        type: "object",
        properties: {
          tenant_id: { type: "string", description: "Azure tenant ID" },
          client_id: { type: "string", description: "Client ID" },
          client_secret: { type: "string", description: "Client secret" },
          grant_type: { type: "string", description: "Grant type, typically 'client_credentials'" }
        },
        required: ["tenant_id", "client_id", "client_secret", "grant_type"]
      },
    },
    {
      name: "get-dynamics-token",
      description: "Get access token for Dynamics 365 operations",
      parameters: {
        type: "object",
        properties: {
          bearer_token: { type: "string", description: "Azure AD bearer token" },
          grant_type: { type: "string", description: "Grant type" },
          fno_id: { type: "string", description: "Finance and Operations ID" }
        },
        required: ["bearer_token", "grant_type", "fno_id"]
      },
    },
    {
      name: "query-inventory",
      description: "Query inventory from Dynamics 365 using the Dynamics token (NOT the Azure AD token)",
      parameters: {
        type: "object",
        properties: {
          access_token: { type: "string", description: "Dynamics 365 access token (NOT the Azure AD token)" },
          fno_id: { type: "string", description: "Finance and Operations ID" },
          product_id: { type: "string", description: "Product ID to query (example: V0001)" },
          organization_id: { type: "string", description: "Organization ID (example: USMF)" }
        },
        required: ["access_token", "fno_id", "product_id", "organization_id"]
      }
    },
    {
      name: "authenticate-dynamics",
      description: "Complete authentication flow for Dynamics 365 (gets both Azure AD token and Dynamics token)",
      parameters: {
        type: "object",
        properties: {
          tenant_id: { type: "string", description: "Azure tenant ID" },
          client_id: { type: "string", description: "Client ID" },
          client_secret: { type: "string", description: "Client secret" },
          grant_type: { type: "string", description: "Grant type, typically 'client_credentials'" },
          fno_id: { type: "string", description: "Finance and Operations ID" }
        },
        required: ["tenant_id", "client_id", "client_secret", "grant_type", "fno_id"]
      }
    },
    {
      name: "token-guide",
      description: "Get clear instructions on how to authenticate and query Dynamics 365 inventory",
      parameters: {}
    }
  ]
});

// Get Chuck Norris joke tool
const getChuckJoke = server.tool(
  "get-chuck-joke",
  "Get a random Chuck Norris joke",
  async () => {
    const response = await fetch("https://api.chucknorris.io/jokes/random");
    const data = await response.json();
    return {
      content: [
        {
          type: "text",
          text: data.value,
        },
      ],
    };
  }
);

// Get Chuck Norris joke categories tool
const getChuckCategories = server.tool(
  "get-chuck-categories",
  "Get all available categories for Chuck Norris jokes",
  async () => {
    const response = await fetch("https://api.chucknorris.io/jokes/categories");
    const data = await response.json();
    return {
      content: [
        {
          type: "text",
          text: data.join(", "),
        },
      ],
    };
  }
);

// Get Dad joke tool
const getDadJoke = server.tool(
  "get-dad-joke",
  "Get a random dad joke",
  async () => {
    const response = await fetch("https://icanhazdadjoke.com/", {
      headers: {
        Accept: "application/json",
      },
    });
    const data = await response.json();
    return {
      content: [
        {
          type: "text",
          text: data.joke,
        },
      ],
    };
  }
);

// Get Yo Mama joke tool
const getYoMamaJoke = server.tool(
  "get-yo-mama-joke",
  "Get a random Yo Mama joke",
  async () => {
    const response = await fetch(
      "https://www.yomama-jokes.com/api/v1/jokes/random"
    );
    const data = await response.json();
    return {
      content: [
        {
          type: "text",
          text: data.joke,
        },
      ],
    };
  }
);

// Get Azure AD Token tool
const getAzureADToken = server.tool(
  "get-azure-ad-token",
  "Get Azure AD token from Microsoft",
  async (params: any) => {
    const { tenant_id, client_id, client_secret, grant_type } = params as AzureADTokenParams;
    
    // Use provided parameters or fall back to environment variables
    const tenantId = tenant_id || config.tenantId;
    const clientId = client_id || config.clientId;
    const clientSecret = client_secret || config.clientSecret;
    const grantType = grant_type || config.grantType;
    const scope = config.defaultScope;
    
    if (!tenantId || !clientId || !clientSecret) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: "Missing credentials. Please provide tenant_id, client_id, and client_secret or set environment variables.",
            }, null, 2),
          },
        ],
      };
    }
    
    const formData = new URLSearchParams();
    formData.append("client_id", clientId);
    formData.append("client_secret", clientSecret);
    formData.append("grant_type", grantType);
    formData.append("scope", scope);

    try {
      const response = await fetch(
        `${config.azureAuthUrl}/${tenantId}/oauth2/v2.0/token`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/x-www-form-urlencoded",
          },
          body: formData,
        }
      );

      const data = await response.json();
      
      // Add guidance for next steps
      const responseData = {
        ...data,
        next_steps: {
          description: "To get a Dynamics token, use the get-dynamics-token tool with this access_token as the bearer_token parameter",
          tool: "get-dynamics-token",
          parameters: {
            bearer_token: data.access_token,
            grant_type: grantType,
            fno_id: config.fnoId || "[Your FnO ID]"
          }
        }
      };
      
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(responseData, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: "Failed to obtain Azure AD token",
              details: error instanceof Error ? error.message : String(error)
            }, null, 2),
          },
        ],
      };
    }
  }
);

// Get Dynamics Token tool
const getDynamicsToken = server.tool(
  "get-dynamics-token",
  "Get access token for Dynamics 365 operations",
  async (params: any) => {
    const { bearer_token, grant_type, fno_id } = params as DynamicsTokenParams;

    // Use provided parameters or fall back to environment variables
    const bearerToken = bearer_token;
    const grantType = grant_type || config.grantType;
    const fnoId = fno_id || config.fnoId;
    
    if (!bearerToken) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: "Missing bearer token. Please provide bearer_token.",
            }, null, 2),
          },
        ],
      };
    }

    if (!fnoId) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: "Missing FnO ID. Please provide fno_id or set the FNO_ID environment variable.",
            }, null, 2),
          },
        ],
      };
    }

    try {
      const response = await fetch(
        config.dynamicsTokenUrl,
        {
          method: "POST",
          headers: {
            "Api-Version": "1.0",
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            grant_type: grantType,
            client_assertion_type: "aad_app",
            client_assertion: bearerToken,
            scope: "https://inventoryservice.operations365.dynamics.com/.default",
            context: fnoId,
            context_type: "finops-env"
          }),
        }
      );

      const data = await response.json();
      
      // Add guidance for next steps
      const responseData = {
        ...data,
        next_steps: {
          description: "To query inventory, use the query-inventory tool with this access_token",
          tool: "query-inventory",
          parameters: {
            access_token: data.access_token,
            fno_id: fnoId,
            product_id: "V0001",  // Example from Postman
            organization_id: "USMF"  // Example from Postman
          }
        }
      };
      
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(responseData, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: "Failed to obtain Dynamics token",
              details: error instanceof Error ? error.message : String(error)
            }, null, 2),
          },
        ],
      };
    }
  }
);

// Query Inventory tool
const queryInventory = server.tool(
  "query-inventory",
  "Query inventory from Dynamics 365 using the Dynamics token (NOT the Azure AD token)",
  async (params: any) => {
    const { access_token, fno_id, product_id, organization_id } = params as InventoryQueryParams;

    // Use provided parameters or fall back to environment variables
    const accessToken = access_token;  // This should be the DYNAMICS token
    const fnoId = fno_id || config.fnoId;
    const productId = product_id;
    const organizationId = organization_id;
    
    if (!accessToken) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: "Missing Dynamics access token. Please provide access_token. This should be the DYNAMICS token, not the Azure AD token.",
              help: "You can get the Dynamics token by using the authenticate-dynamics tool first, then use the 'dynamics_token' value from the response."
            }, null, 2),
          },
        ],
      };
    }

    if (!fnoId) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: "Missing FnO ID. Please provide fno_id or set the FNO_ID environment variable.",
            }, null, 2),
          },
        ],
      };
    }

    if (!productId || !organizationId) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: "Missing product_id or organization_id. Both are required.",
            }, null, 2),
          },
        ],
      };
    }

    try {
      const response = await fetch(
        `${config.inventoryServiceUrl}/api/environment/${fnoId}/onhand/indexquery`,
        {
          method: "POST",
          headers: {
            "Authorization": `Bearer ${accessToken}`,
            "Api-Version": "2.0",
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            filters: {
              ProductId: [productId],
              OrganizationId: [organizationId]
            },
            groupByValues: ["batchId"],
            returnNegative: false,
            queryATP: false
          }),
        }
      );

      const data = await response.json();
      
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(data, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: "Failed to query inventory",
              details: error instanceof Error ? error.message : String(error)
            }, null, 2),
          },
        ],
      };
    }
  }
);

// Combined Authentication flow for Dynamics 365
const authenticateDynamics = server.tool(
  "authenticate-dynamics",
  "Complete authentication flow for Dynamics 365 (gets both Azure AD token and Dynamics token)",
  async (params: any) => {
    const { tenant_id, client_id, client_secret, grant_type, fno_id } = params as AuthenticateDynamicsParams;
    
    // Use provided parameters or fall back to environment variables
    const tenantId = tenant_id || config.tenantId;
    const clientId = client_id || config.clientId;
    const clientSecret = client_secret || config.clientSecret;
    const grantType = grant_type || config.grantType;
    const fnoId = fno_id || config.fnoId;
    const scope = config.defaultScope;
    
    if (!tenantId || !clientId || !clientSecret) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: "Missing credentials. Please provide tenant_id, client_id, and client_secret or set environment variables.",
            }, null, 2),
          },
        ],
      };
    }
    
    if (!fnoId) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: "Missing FnO ID. Please provide fno_id or set the FNO_ID environment variable.",
            }, null, 2),
          },
        ],
      };
    }
    
    try {
      // Step 1: Get Azure AD Token
      const formData = new URLSearchParams();
      formData.append("client_id", clientId);
      formData.append("client_secret", clientSecret);
      formData.append("grant_type", grantType);
      formData.append("scope", scope);

      const aadResponse = await fetch(
        `${config.azureAuthUrl}/${tenantId}/oauth2/v2.0/token`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/x-www-form-urlencoded",
          },
          body: formData,
        }
      );

      const aadData = await aadResponse.json();
      
      if (!aadData.access_token) {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                error: "Failed to obtain Azure AD token",
                details: aadData
              }, null, 2),
            },
          ],
        };
      }
      
      // Step 2: Get Dynamics Token
      const dynamicsResponse = await fetch(
        config.dynamicsTokenUrl,
        {
          method: "POST",
          headers: {
            "Api-Version": "1.0",
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            grant_type: grantType,
            client_assertion_type: "aad_app",
            client_assertion: aadData.access_token,
            scope: "https://inventoryservice.operations365.dynamics.com/.default",
            context: fnoId,
            context_type: "finops-env"
          }),
        }
      );

      const dynamicsData = await dynamicsResponse.json();
      
      // Return combined results with next steps
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              authentication_flow: "Two-step authentication completed successfully",
              azure_ad_token: {
                token: aadData.access_token,
                note: "This token is only used for getting the Dynamics token, not for API calls"
              },
              dynamics_token: {
                token: dynamicsData.access_token,
                note: "THIS is the token you need for inventory queries (as access_token)"
              },
              token_type: dynamicsData.token_type,
              expires_in: dynamicsData.expires_in,
              next_steps: {
                description: "To query inventory, use the query-inventory tool with the DYNAMICS token (not the Azure AD token) as the access_token",
                tool: "query-inventory",
                parameters: {
                  access_token: dynamicsData.access_token,  // THIS is the correct token for inventory queries
                  fno_id: fnoId,
                  product_id: "V0001",  // Example from Postman
                  organization_id: "USMF"  // Example from Postman
                }
              }
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: "Failed to complete Dynamics 365 authentication flow",
              details: error instanceof Error ? error.message : String(error)
            }, null, 2),
          },
        ],
      };
    }
  }
);

// Token Guide tool
const tokenGuide = server.tool(
  "token-guide",
  "Get clear instructions on how to authenticate and query Dynamics 365 inventory",
  async () => {
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify({
            title: "Dynamics 365 Authentication and Inventory Query Guide",
            overview: "This guide explains how to authenticate with Dynamics 365 and query inventory data correctly.",
            authentication_flow: {
              step1: {
                description: "First, you need to get both an Azure AD token and a Dynamics token in one step:",
                tool: "authenticate-dynamics",
                parameters: {
                  tenant_id: "your-azure-tenant-id",
                  client_id: "your-client-id",
                  client_secret: "your-client-secret",
                  grant_type: "client_credentials",
                  fno_id: "your-finance-operations-id"
                },
                note: "This returns TWO different tokens - only use the DYNAMICS token for inventory queries!"
              },
              step2: {
                description: "The authenticate-dynamics response contains TWO tokens:",
                response_example: {
                  azure_ad_token: {
                    token: "eyJ0eXAiOiJKV1QiLCJhbG...",
                    note: "DO NOT USE THIS TOKEN for inventory queries - it's only for internal auth"
                  },
                  dynamics_token: {
                    token: "eyJ0eXAiOiJKV1QiLCJhbG...",
                    note: "⭐ USE THIS TOKEN as the access_token for inventory queries"
                  }
                }
              }
            },
            inventory_query: {
              description: "After authentication, use the query-inventory tool with these parameters:",
              tool: "query-inventory",
              parameters: {
                access_token: "dynamics_token_from_authentication_step",  // NOT the Azure AD token!
                fno_id: "same-fno-id-from-authentication",
                product_id: "V0001",  // Example from Postman
                organization_id: "USMF"  // Example from Postman
              },
              http_request_details: {
                url: "https://inventoryservice.wus-il301.gateway.prod.island.powerapps.com/api/environment/your-fno-id/onhand/indexquery",
                method: "POST",
                headers: {
                  "Authorization": "Bearer your-dynamics-token",
                  "Api-Version": "2.0",
                  "Content-Type": "application/json"
                },
                body: {
                  "filters": {
                    "ProductId": ["V0001"],
                    "OrganizationId": ["USMF"]
                  },
                  "groupByValues": ["batchId"],
                  "returnNegative": false,
                  "queryATP": false
                }
              }
            },
            common_errors: {
              wrong_token: "Using the Azure AD token instead of the Dynamics token for inventory queries",
              missing_parameters: "Not providing product_id or organization_id",
              incorrect_values: "Not using correct format for product_id (e.g., 'V0001') or organization_id (e.g., 'USMF')"
            }
          }, null, 4),
        },
      ],
    };
  }
);

const app = express();

// to support multiple simultaneous connections we have a lookup object from
// sessionId to transport
const transports: { [sessionId: string]: SSEServerTransport } = {};

app.get("/sse", async (req: Request, res: Response) => {
  // Get the full URI from the request
  const host = req.get("host") || "";
  // Use configured base URL if available, otherwise build from request
  const fullUri = config.serviceBaseUrl || `https://${host}/jokes`;
  const transport = new SSEServerTransport(fullUri, res);
  transports[transport.sessionId] = transport;
  res.on("close", () => {
    delete transports[transport.sessionId];
  });
  await server.connect(transport);
});

app.post("/jokes", async (req: Request, res: Response) => {
  const sessionId = req.query.sessionId as string;
  const transport = transports[sessionId];
  if (transport) {
    await transport.handlePostMessage(req, res);
  } else {
    res.status(400).send("No transport found for sessionId");
  }
});

app.get("/", (_req, res) => {
  res.send(`The Jokes and Dynamics 365 MCP server is running! Environment: ${process.env.NODE_ENV || 'development'}`);
});

// Health check endpoint for Azure
app.get("/health", (_req, res) => {
  res.status(200).json({ status: "healthy", version: "1.0.0" });
});

const PORT = config.port;
app.listen(PORT, () => {
  console.log(`✅ Server is running at http://localhost:${PORT}`);
  console.log(`✅ Environment: ${process.env.NODE_ENV || 'development'}`);
});
