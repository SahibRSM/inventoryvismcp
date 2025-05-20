import express, { Request, Response } from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { z, ZodRawShape } from "zod"; // Import ZodRawShape for clarity if needed, though not strictly for usage here
import { CallToolResult, isInitializeRequest } from "@modelcontextprotocol/sdk/types.js";
import { randomUUID } from "node:crypto";

// Environment variables configuration
const config = {
  tenantId: process.env.TENANT_ID || "",
  clientId: process.env.CLIENT_ID || "",
  clientSecret: process.env.CLIENT_SECRET || "",
  grantType: process.env.GRANT_TYPE || "client_credentials",
  defaultScope: process.env.DEFAULT_SCOPE || "0cdb527f-a8d1-4bf8-9436-b352c68682b2/.default",
  fnoId: process.env.FNO_ID || "",
  azureAuthUrl: process.env.AZURE_AUTH_URL || "https://login.microsoftonline.com",
  dynamicsTokenUrl: process.env.DYNAMICS_TOKEN_URL || "https://securityservice.operations365.dynamics.com/token",
  inventoryServiceUrl: process.env.INVENTORY_SERVICE_URL || "https://inventoryservice.wus-il301.gateway.prod.island.powerapps.com",
  port: process.env.PORT || 3001,
  serviceBaseUrl: process.env.SERVICE_BASE_URL || "",
};

const server = new McpServer({
  name: "inventoryMCP",
  description: "A server that provides Dynamics 365 inventory information",
  version: "1.0.0",
});

// Define Zod RAW SHAPES for tool parameters
const authenticateDynamicsParamsRawSchema = { // This is ZodRawShape
  tenant_id: z.string().optional().describe("Azure tenant ID"),
  client_id: z.string().optional().describe("Client ID"),
  client_secret: z.string().optional().describe("Client secret"),
  grant_type: z.string().optional().describe("Grant type, typically 'client_credentials'"),
  fno_id: z.string().optional().describe("Finance and Operations ID"),
};

const queryInventoryParamsRawSchema = { // This is ZodRawShape
  access_token: z.string().describe("Dynamics 365 access token"),
  fno_id: z.string().optional().describe("Finance and Operations ID"),
  product_id: z.string().describe("Product ID to query (example: V0001)"),
  organization_id: z.string().describe("Organization ID (example: USMF)"),
};

server.tool(
  "authenticate-dynamics",
  "Complete authentication flow for Dynamics 365 inventory access",
  authenticateDynamicsParamsRawSchema, // Pass the raw shape
  async (params): Promise<CallToolResult> => {
    const tenantId = params.tenant_id || config.tenantId;
    const clientId = params.client_id || config.clientId;
    const clientSecret = params.client_secret || config.clientSecret;
    const grantType = params.grant_type || config.grantType;
    const fnoId = params.fno_id || config.fnoId;
    const scope = config.defaultScope;

    if (!tenantId || !clientId || !clientSecret || !fnoId) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({ error: "Missing required credentials for inventory access." }, null, 2),
        }],
        isError: true,
      };
    }

    try {
      const formData = new URLSearchParams();
      formData.append("client_id", clientId);
      formData.append("client_secret", clientSecret);
      formData.append("grant_type", grantType);
      formData.append("scope", scope);

      const aadResponse = await fetch(
        `${config.azureAuthUrl}/${tenantId}/oauth2/v2.0/token`,
        {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: formData,
        }
      );
      const aadData = await aadResponse.json();

      if (!aadData.access_token) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              error: "Failed to obtain Azure AD token for inventory access",
              details: aadData,
            }, null, 2),
          }],
          isError: true,
        };
      }

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
            context_type: "finops-env",
          }),
        }
      );
      const dynamicsData = await dynamicsResponse.json();

      if (!dynamicsData.access_token) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              error: "Failed to obtain Dynamics token for inventory access",
              details: dynamicsData,
            }, null, 2),
          }],
          isError: true,
        };
      }
      
      return {
        content: [{
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
                organization_id: "USMF",
              },
            },
          }, null, 2),
        }],
      };
    } catch (error) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: "Failed to complete authentication for inventory access",
            details: error instanceof Error ? error.message : String(error),
          }, null, 2),
        }],
        isError: true,
      };
    }
  }
);

server.tool(
  "query-inventory",
  "Query inventory from Dynamics 365",
  queryInventoryParamsRawSchema, // Pass the raw shape
  async (params): Promise<CallToolResult> => {
    const { access_token, product_id, organization_id } = params;
    const fnoId = params.fno_id || config.fnoId;

    try {
      const response = await fetch(
        `${config.inventoryServiceUrl}/api/environment/${fnoId}/onhand/indexquery`,
        {
          method: "POST",
          headers: {
            "Authorization": `Bearer ${access_token}`,
            "Api-Version": "2.0",
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            filters: {
              ProductId: [product_id],
              OrganizationId: [organization_id],
            },
            groupByValues: ["batchId"],
            returnNegative: false,
            queryATP: false,
          }),
        }
      );
      const data = await response.json();

      if (!response.ok) {
         return {
           content: [{
             type: "text",
             text: JSON.stringify({
               error: `Inventory service request failed with status: ${response.status}`,
               details: data 
             }, null, 2),
           }],
           isError: true,
         };
      }

      return {
        content: [{
          type: "text",
          text: JSON.stringify(data, null, 2),
        }],
      };
    } catch (error) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: "Failed to query inventory",
            details: error instanceof Error ? error.message : String(error),
          }, null, 2),
        }],
        isError: true,
      };
    }
  }
);

const app = express();
app.use(express.json());

const transports: { [sessionId: string]: StreamableHTTPServerTransport } = {};

app.all("/mcp", async (req: Request, res: Response) => {
  try {
    const sessionId = req.headers["mcp-session-id"] as string | undefined;
    let transport: StreamableHTTPServerTransport;

    if (sessionId && transports[sessionId]) {
      transport = transports[sessionId];
    } else if ((req.method === "POST" && !sessionId && isInitializeRequest(req.body))) {
      transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => randomUUID(),
        onsessioninitialized: (newSessionId: string) => { // Added string type for newSessionId
          transports[newSessionId] = transport;
          console.log(`Session initialized: ${newSessionId}`);
        },
      });
      transport.onclose = () => {
        if (transport.sessionId && transports[transport.sessionId]) {
          delete transports[transport.sessionId];
          console.log(`Session closed and removed: ${transport.sessionId}`);
        }
      };
      await server.connect(transport);
    } else {
      res.status(400).json({
        jsonrpc: "2.0",
        error: { code: -32000, message: "Bad Request: Valid session ID required or proper initialization." },
        id: null,
      });
      return;
    }
    await transport.handleRequest(req, res, req.body);
  } catch (error) {
    console.error("Error handling MCP request:", error);
    if (!res.headersSent) {
      res.status(500).json({
        jsonrpc: "2.0",
        error: { code: -32603, message: "Internal server error" },
        id: null,
      });
    }
  }
});

const PORT = config.port;
app.listen(PORT, () => {
  console.log(`✅ Inventory Visibility Server (Streamable HTTP) running at http://localhost:${PORT}/mcp`);
});

process.on('SIGINT', async () => {
  console.log('Shutting down server...');
  for (const sessionId in transports) {
    await transports[sessionId].close();
  }
  process.exit(0);
});
