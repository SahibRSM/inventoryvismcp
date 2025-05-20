import express, { Request, Response } from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"; // [cite: 2]
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js"; // [cite: 2, 3]
import { z } from "zod"; // [cite: 3, 4]
import { CallToolResult, InitializeRequestSchema } from "@modelcontextprotocol/sdk/types.js"; // Changed from isInitializeRequest
import { randomUUID } from "node:crypto"; // [cite: 4, 148]

// Environment variables configuration (remains the same)
const config = {
  tenantId: process.env.TENANT_ID || "", // [cite: 5]
  clientId: process.env.CLIENT_ID || "", // [cite: 5, 6]
  clientSecret: process.env.CLIENT_SECRET || "", // [cite: 6]
  grantType: process.env.GRANT_TYPE || "client_credentials", // [cite: 6]
  defaultScope: process.env.DEFAULT_SCOPE || "0cdb527f-a8d1-4bf8-9436-b352c68682b2/.default", // [cite: 6]
  fnoId: process.env.FNO_ID || "", // [cite: 6, 7]
  azureAuthUrl: process.env.AZURE_AUTH_URL || "https://login.microsoftonline.com", // [cite: 7]
  dynamicsTokenUrl: process.env.DYNAMICS_TOKEN_URL || "https://securityservice.operations365.dynamics.com/token", // [cite: 7]
  inventoryServiceUrl: process.env.INVENTORY_SERVICE_URL || "https://inventoryservice.wus-il301.gateway.prod.island.powerapps.com", // [cite: 7]
  port: process.env.PORT || 3001, // [cite: 7, 8]
  serviceBaseUrl: process.env.SERVICE_BASE_URL || "", // [cite: 8]
};

const server = new McpServer({ // [cite: 9]
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
  authenticateDynamicsParamsRawSchema, // Use the raw shape here
  async (params): Promise<CallToolResult> => {
    const tenantId = params.tenant_id || config.tenantId;
    const clientId = params.client_id || config.clientId;
    const clientSecret = params.client_secret || config.clientSecret;
    const grantType = params.grant_type || config.grantType;
    const fnoId = params.fno_id || config.fnoId;
    const scope = config.defaultScope;

    if (!tenantId || !clientId || !clientSecret || !fnoId) {
      return { // [cite: 13]
        content: [{
          type: "text",
          text: JSON.stringify({ error: "Missing required credentials for inventory access." }, null, 2),
        }],
        isError: true,
      };
    }

    try {
      const formData = new URLSearchParams(); // [cite: 14]
      formData.append("client_id", clientId);
      formData.append("client_secret", clientSecret);
      formData.append("grant_type", grantType);
      formData.append("scope", scope); // [cite: 14]

      const aadResponse = await fetch( // [cite: 15]
        `${config.azureAuthUrl}/${tenantId}/oauth2/v2.0/token`,
        {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: formData,
        }
      );
      const aadData = await aadResponse.json(); // [cite: 16]

      if (!aadData.access_token) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              error: "Failed to obtain Azure AD token for inventory access",
              details: aadData,
            }, null, 2), // [cite: 17]
          }],
          isError: true, // [cite: 17]
        };
      } // [cite: 18]

      const dynamicsResponse = await fetch( // [cite: 18]
        config.dynamicsTokenUrl,
        {
          method: "POST",
          headers: {
            "Api-Version": "1.0",
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ // [cite: 19]
            grant_type: grantType,
            client_assertion_type: "aad_app",
            client_assertion: aadData.access_token,
            scope: "https://inventoryservice.operations365.dynamics.com/.default",
            context: fnoId,
            context_type: "finops-env",
          }),
        } // [cite: 20]
      );
      const dynamicsData = await dynamicsResponse.json(); // [cite: 20]

      if (!dynamicsData.access_token) { // [cite: 21]
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              error: "Failed to obtain Dynamics token for inventory access",
              details: dynamicsData,
            }, null, 2),
          }], // [cite: 22]
          isError: true, // [cite: 22]
        };
      } // [cite: 23]
      
      return { // [cite: 23]
        content: [{
          type: "text",
          text: JSON.stringify({
            dynamics_token: dynamicsData.access_token,
            token_type: dynamicsData.token_type,
            expires_in: dynamicsData.expires_in,
            inventory_query_example: { // [cite: 24]
              tool: "query-inventory",
              parameters: {
                access_token: dynamicsData.access_token,
                fno_id: fnoId,
                product_id: "V0001",
                organization_id: "USMF",
              }, // [cite: 25]
            },
          }, null, 2),
        }],
      };
    } catch (error) { // [cite: 26]
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: "Failed to complete authentication for inventory access",
            details: error instanceof Error ? error.message : String(error),
          }, null, 2),
        }],
        isError: true, // [cite: 27]
      };
    } // [cite: 28]
  }
);

server.tool(
  "query-inventory",
  "Query inventory from Dynamics 365",
  queryInventoryParamsRawSchema, // Use the raw shape here
  async (params): Promise<CallToolResult> => {
    const { access_token, product_id, organization_id } = params;
    const fnoId = params.fno_id || config.fnoId;

    try {
      const response = await fetch(
        `${config.inventoryServiceUrl}/api/environment/${fnoId}/onhand/indexquery`,
        { // [cite: 29]
          method: "POST",
          headers: {
            "Authorization": `Bearer ${access_token}`,
            "Api-Version": "2.0",
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            filters: { // [cite: 30]
              ProductId: [product_id],
              OrganizationId: [organization_id],
            },
            groupByValues: ["batchId"],
            returnNegative: false,
            queryATP: false,
          }),
        } // [cite: 31]
      );
      const data = await response.json(); // [cite: 31]

      if (!response.ok) {
         return {
           content: [{
             type: "text",
             text: JSON.stringify({
               error: `Inventory service request failed with status: ${response.status}`,
               details: data // [cite: 32]
             }, null, 2),
           }],
           isError: true, // [cite: 32]
         };
      } // [cite: 33]

      return { // [cite: 33]
        content: [{
          type: "text",
          text: JSON.stringify(data, null, 2),
        }],
      };
    } catch (error) { // [cite: 34]
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: "Failed to query inventory",
            details: error instanceof Error ? error.message : String(error),
          }, null, 2),
        }],
        isError: true, // [cite: 35]
      };
    } // [cite: 36]
  }
);

const app = express();
app.use(express.json()); // [cite: 36]
const transports: { [sessionId: string]: StreamableHTTPServerTransport } = {}; // [cite: 36]

app.all("/mcp", async (req: Request, res: Response) => { // [cite: 38]
  try {
    const sessionId = req.headers["mcp-session-id"] as string | undefined;
    let transport: StreamableHTTPServerTransport;

    // Use InitializeRequestSchema directly for the check
    const isInitReq = InitializeRequestSchema.safeParse(req.body).success;

    if (sessionId && transports[sessionId]) {
      transport = transports[sessionId];
    } else if ((req.method === "POST" && !sessionId && isInitReq)) {
      transport = new StreamableHTTPServerTransport({ // [cite: 39]
        sessionIdGenerator: () => randomUUID(),
        onsessioninitialized: (newSessionId: string) => { // Explicitly type newSessionId
          transports[newSessionId] = transport;
          console.log(`Session initialized: ${newSessionId}`);
        },
      });
      transport.onclose = () => { // [cite: 39]
        if (transport.sessionId && transports[transport.sessionId]) { // [cite: 40]
          delete transports[transport.sessionId];
          console.log(`Session closed and removed: ${transport.sessionId}`); // [cite: 41]
        }
      };
      await server.connect(transport); // [cite: 41]
    } else {
      res.status(400).json({
        jsonrpc: "2.0",
        error: { code: -32000, message: "Bad Request: Valid session ID required or proper initialization." },
        id: null,
      }); // [cite: 42]
      return; // [cite: 43]
    }
    await transport.handleRequest(req, res, req.body); // [cite: 43]
  } catch (error) {
    console.error("Error handling MCP request:", error); // [cite: 44]
    if (!res.headersSent) { // [cite: 45]
      res.status(500).json({
        jsonrpc: "2.0",
        error: { code: -32603, message: "Internal server error" },
        id: null,
      }); // [cite: 45]
    } // [cite: 46]
  }
});

const PORT = config.port; // [cite: 46]
app.listen(PORT, () => { // [cite: 47]
  console.log(`✅ Inventory Visibility Server (Streamable HTTP) running at http://localhost:${PORT}/mcp`);
});

process.on('SIGINT', async () => { // [cite: 48]
  console.log('Shutting down server...');
  for (const sessionId in transports) {
    await transports[sessionId].close();
  }
  process.exit(0);
});
