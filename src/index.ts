import express, { Request, Response } from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";

// Environment variables configuration - only inventory-related settings
const config = {
  // Authentication settings (required for inventory access)
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

// Type definitions - only inventory-related interfaces
interface InventoryQueryParams {
  access_token: string;
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

// Initialize the MCP server with only inventory-related tools
const server = new McpServer({
  name: "inventoryMCP",
  description: "A server that provides Dynamics 365 inventory information",
  version: "1.0.0",
  tools: [
    {
      name: "authenticate-dynamics",
      description: "Complete authentication flow for Dynamics 365 inventory access",
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
      name: "query-inventory",
      description: "Query inventory from Dynamics 365",
      parameters: {
        type: "object",
        properties: {
          access_token: { type: "string", description: "Dynamics 365 access token" },
          fno_id: { type: "string", description: "Finance and Operations ID" },
          product_id: { type: "string", description: "Product ID to query (example: V0001)" },
          organization_id: { type: "string", description: "Organization ID (example: USMF)" }
        },
        required: ["access_token", "fno_id", "product_id", "organization_id"]
      }
    }
  ]
});

// Combined Authentication flow for Dynamics 365 - needed for inventory access
const authenticateDynamics = server.tool(
  "authenticate-dynamics",
  "Complete authentication flow for Dynamics 365 inventory access",
  async (params: any) => {
    const { tenant_id, client_id, client_secret, grant_type, fno_id } = params as AuthenticateDynamicsParams;
    
    const tenantId = tenant_id || config.tenantId;
    const clientId = client_id || config.clientId;
    const clientSecret = client_secret || config.clientSecret;
    const grantType = grant_type || config.grantType;
    const fnoId = fno_id || config.fnoId;
    const scope = config.defaultScope;
    
    if (!tenantId || !clientId || !clientSecret || !fnoId) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: "Missing required credentials for inventory access.",
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
                error: "Failed to obtain Azure AD token for inventory access",
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
      
      // Return just what's needed for inventory queries
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              dynamics_token: dynamicsData.access_token,
              token_type: dynamicsData.token_type,
              expires_in: dynamicsData.expires_in,
              inventory_query_example: {
                tool: "query-inventory",
                parameters: {
                  access_token: dynamicsData.access_token,
                  fno_id: fnoId,
                  product_id: "V0001",
                  organization_id: "USMF"
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
              error: "Failed to complete authentication for inventory access",
              details: error instanceof Error ? error.message : String(error)
            }, null, 2),
          },
        ],
      };
    }
  }
);

// Query Inventory tool - the core functionality
const queryInventory = server.tool(
  "query-inventory",
  "Query inventory from Dynamics 365",
  async (params: any) => {
    const { access_token, fno_id, product_id, organization_id } = params as InventoryQueryParams;

    const accessToken = access_token;
    const fnoId = fno_id || config.fnoId;
    const productId = product_id;
    const organizationId = organization_id;
    
    if (!accessToken || !fnoId || !productId || !organizationId) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: "Missing required parameters for inventory query.",
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

const app = express();

// Support for SSE connections
const transports: { [sessionId: string]: SSEServerTransport } = {};

app.get("/sse", async (req: Request, res: Response) => {
  const host = req.get("host") || "";
  const fullUri = config.serviceBaseUrl || `https://${host}/inventory`;
  const transport = new SSEServerTransport(fullUri, res);
  transports[transport.sessionId] = transport;
  res.on("close", () => {
    delete transports[transport.sessionId];
  });
  await server.connect(transport);
});

app.post("/inventory", async (req: Request, res: Response) => {
  const sessionId = req.query.sessionId as string;
  const transport = transports[sessionId];
  if (transport) {
    await transport.handlePostMessage(req, res);
  } else {
    res.status(400).send("No transport found for sessionId");
  }
});

const PORT = config.port;
app.listen(PORT, () => {
  console.log(`✅ Inventory Visibility Server running at http://localhost:${PORT}`);
});
