import express, { Request, Response } from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { z } from "zod";
import { CallToolResult, InitializeRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import { randomUUID } from "node:crypto";

// Environment variables configuration
const config = {
  tenantId: process.env.TENANT_ID || "",
  clientId: process.env.CLIENT_ID || "",
  clientSecret: process.env.CLIENT_SECRET || "",
  grantType: process.env.GRANT_TYPE || "client_credentials",
  defaultScope: process.env.DEFAULT_SCOPE || "0cdb527f-a8d1-4bf8-9436-b352c68682b2/.default",
  fnoId: process.env.FNO_ID || "", // This will be the single source of truth for FNO ID
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

// In-memory cache for the Dynamics token
let cachedDynamicsToken: {
  token: string;
  expiresAt: number; // Timestamp (seconds since epoch) when the token expires
} | null = null;

// Helper function to perform authentication and cache the token
async function getDynamicsAccessToken(): Promise<{ token: string; error?: undefined } | { token?: undefined; error: CallToolResult }> {
  // Check cache first
  if (cachedDynamicsToken && cachedDynamicsToken.expiresAt > (Date.now() / 1000) + 60) { // Check if token is valid for at least 60 more seconds
    return { token: cachedDynamicsToken.token };
  }

  console.log("No valid cached token. Authenticating with Dynamics...");

  const { tenantId, clientId, clientSecret, grantType, fnoId, defaultScope } = config;

  if (!tenantId || !clientId || !clientSecret || !fnoId) {
    return {
      error: {
        content: [{ type: "text", text: JSON.stringify({ error: "Server is missing required credentials for Dynamics 365 access." }, null, 2) }],
        isError: true,
      }
    };
  }

  try {
    // Step 1: Get Azure AD Token
    const formData = new URLSearchParams();
    formData.append("client_id", clientId);
    formData.append("client_secret", clientSecret);
    formData.append("grant_type", grantType);
    formData.append("scope", defaultScope);

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
        error: {
          content: [{ type: "text", text: JSON.stringify({ error: "Failed to obtain Azure AD token.", details: aadData }, null, 2) }],
          isError: true,
        }
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
          context: fnoId, // fnoId from server config used here
          context_type: "finops-env",
        }),
      }
    );
    const dynamicsData = await dynamicsResponse.json();

    if (!dynamicsData.access_token || !dynamicsData.expires_in) {
      return {
        error: {
          content: [{ type: "text", text: JSON.stringify({ error: "Failed to obtain Dynamics token.", details: dynamicsData }, null, 2) }],
          isError: true,
        }
      };
    }

    // Cache the token and its expiry (expires_in is in seconds)
    cachedDynamicsToken = {
      token: dynamicsData.access_token,
      expiresAt: Math.floor(Date.now() / 1000) + dynamicsData.expires_in,
    };
    console.log("Successfully authenticated with Dynamics and cached token.");
    return { token: cachedDynamicsToken.token };

  } catch (error) {
    console.error("Authentication error:", error);
    cachedDynamicsToken = null; // Clear cache on error
    return {
      error: {
        content: [{ type: "text", text: JSON.stringify({ error: "Exception during authentication flow.", details: error instanceof Error ? error.message : String(error) }, null, 2) }],
        isError: true,
      }
    };
  }
}

// Define Zod RAW SHAPE for the updated query-inventory tool parameters
const queryInventoryParamsRawSchema = {
  // access_token is handled internally
  // fno_id is no longer a parameter; config.fnoId will be used exclusively.
  product_id: z.string().describe("Product ID to query (example: V0001)"),
  organization_id: z.string().describe("Organization ID (example: USMF)"),
};

// Tool: query-inventory (now uses server-configured fno_id exclusively)
server.tool(
  "query-inventory",
  "Queries inventory from Dynamics 365 using the server's configured F&O environment.",
  queryInventoryParamsRawSchema,
  async (params): Promise<CallToolResult> => {
    const authResult = await getDynamicsAccessToken();
    if (authResult.error) {
      return authResult.error;
    }
    const accessToken = authResult.token;

    const { product_id, organization_id } = params;
    const fnoIdToUse = config.fnoId; // Always use the server's configured FNO ID

    if (!fnoIdToUse) { // Checks if the server has fnoId configured
         return {
           content: [{
             type: "text",
             text: JSON.stringify({ error: "FNO ID is not configured on the server." }, null, 2),
           }],
           isError: true,
         };
    }

    try {
      const response = await fetch(
        `${config.inventoryServiceUrl}/api/environment/${fnoIdToUse}/onhand/indexquery`,
        {
          method: "POST",
          headers: {
            "Authorization": `Bearer ${accessToken}`,
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

    const isInitReq = InitializeRequestSchema.safeParse(req.body).success;

    if (sessionId && transports[sessionId]) {
      transport = transports[sessionId];
    } else if ((req.method === "POST" && !sessionId && isInitReq)) {
      transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => randomUUID(),
        onsessioninitialized: (newSessionId: string) => {
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
