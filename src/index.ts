import express, { Request, Response } from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js"; // Changed from SSEServerTransport [cite: 28, 59]
import { z } from "zod"; // For schema definition [cite: 19, 24, 717]
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
  serviceBaseUrl: process.env.SERVICE_BASE_URL || "", // Used for constructing the full SSE URI if needed
};

// Initialize the MCP server
const server = new McpServer({
  name: "inventoryMCP",
  description: "A server that provides Dynamics 365 inventory information",
  version: "1.0.0",
});

// Define Zod schemas for tool parameters for type safety and validation [cite: 19, 24, 717]
const authenticateDynamicsParamsSchema = z.object({
  tenant_id: z.string().optional().describe("Azure tenant ID"),
  client_id: z.string().optional().describe("Client ID"),
  client_secret: z.string().optional().describe("Client secret"),
  grant_type: z.string().optional().describe("Grant type, typically 'client_credentials'"),
  fno_id: z.string().optional().describe("Finance and Operations ID"),
});

const queryInventoryParamsSchema = z.object({
  access_token: z.string().describe("Dynamics 365 access token"),
  fno_id: z.string().optional().describe("Finance and Operations ID"),
  product_id: z.string().describe("Product ID to query (example: V0001)"),
  organization_id: z.string().describe("Organization ID (example: USMF)"),
});

// Tool: authenticate-dynamics
server.tool(
  "authenticate-dynamics",
  "Complete authentication flow for Dynamics 365 inventory access",
  authenticateDynamicsParamsSchema, // Use Zod schema here [cite: 24, 961, 717]
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
        isError: true, // Indicate tool error [cite: 165, 44]
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
          isError: true, // Indicate tool error [cite: 165, 44]
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
          isError: true, // Indicate tool error [cite: 165, 44]
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
        isError: true, // Indicate tool error [cite: 165, 44]
      };
    }
  }
);

// Tool: query-inventory
server.tool(
  "query-inventory",
  "Query inventory from Dynamics 365",
  queryInventoryParamsSchema, // Use Zod schema here [cite: 24, 961, 717]
  async (params): Promise<CallToolResult> => {
    const { access_token, product_id, organization_id } = params;
    const fnoId = params.fno_id || config.fnoId;

    // access_token, product_id, organization_id are guaranteed by Zod schema
    // fnoId has a fallback

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

      if (!response.ok) { // Check if the fetch itself was not okay
         return {
           content: [{
             type: "text",
             text: JSON.stringify({
               error: `Inventory service request failed with status: ${response.status}`,
               details: data 
             }, null, 2),
           }],
           isError: true, // Indicate tool error [cite: 165, 44]
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
        isError: true, // Indicate tool error [cite: 165, 44]
      };
    }
  }
);

const app = express();
app.use(express.json()); // Middleware to parse JSON bodies

// Store transports by session ID for StreamableHTTPServerTransport [cite: 29]
const transports: { [sessionId: string]: StreamableHTTPServerTransport } = {};

// Using StreamableHTTPServerTransport [cite: 28, 59]
// This single endpoint handles POST for client-to-server, GET for server-to-client (SSE), and DELETE for session termination.
app.all("/mcp", async (req: Request, res: Response) => { // Changed from /inventory to /mcp for convention [cite: 29, 33, 60]
  try {
    const sessionId = req.headers["mcp-session-id"] as string | undefined;
    let transport: StreamableHTTPServerTransport;

    if (sessionId && transports[sessionId]) {
      transport = transports[sessionId]; // Reuse existing transport [cite: 29]
    } else if ((req.method === "POST" && !sessionId && isInitializeRequest(req.body))) { // isInitializeRequest is a type guard [cite: 29]
      // New initialization request
      transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => randomUUID(), // Generate a new session ID [cite: 29]
        onsessioninitialized: (newSessionId) => { // Store transport upon session initialization [cite: 29]
          transports[newSessionId] = transport;
          console.log(`Session initialized: ${newSessionId}`);
        },
      });

      transport.onclose = () => { // Clean up transport when closed [cite: 30]
        if (transport.sessionId && transports[transport.sessionId]) {
          delete transports[transport.sessionId];
          console.log(`Session closed and removed: ${transport.sessionId}`);
        }
      };
      await server.connect(transport); // Connect server to the new transport
    } else {
      // Invalid request if not initialization and no valid session ID
      res.status(400).json({
        jsonrpc: "2.0",
        error: { code: -32000, message: "Bad Request: Valid session ID required or proper initialization." },
        id: null,
      }); // [cite: 31]
      return;
    }
    // Handle the request using the transport (either new or existing)
    await transport.handleRequest(req, res, req.body); // Pass req.body for POST [cite: 32]
  } catch (error) {
    console.error("Error handling MCP request:", error);
    if (!res.headersSent) {
      res.status(500).json({
        jsonrpc: "2.0",
        error: { code: -32603, message: "Internal server error" },
        id: null,
      }); // [cite: 35]
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
    await transports[sessionId].close(); // Close all active transports
  }
  process.exit(0);
});
