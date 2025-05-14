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
      description: "Query inventory from Dynamics 365",
      parameters: {
        type: "object",
        properties: {
          access_token: { type: "string", description: "Dynamics 365 access token" },
          fno_id: { type: "string", description: "Finance and Operations ID" },
          product_id: { type: "string", description: "Product ID to query" },
          organization_id: { type: "string", description: "Organization ID" }
        },
        required: ["access_token", "fno_id", "product_id", "organization_id"]
      }
    }
  ],
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
  access_token: string;
  fno_id: string;
  product_id: string;
  organization_id: string;
}

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
              error: "Failed to obtain Azure AD token",
              details: error.message
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
              error: "Failed to obtain Dynamics token",
              details: error.message
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
  "Query inventory from Dynamics 365",
  async (params: any) => {
    const { access_token, fno_id, product_id, organization_id } = params as InventoryQueryParams;

    // Use provided parameters or fall back to environment variables
    const accessToken = access_token;
    const fnoId = fno_id || config.fnoId;
    const productId = product_id;
    const organizationId = organization_id;
    
    if (!accessToken) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: "Missing access token. Please provide access_token.",
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
              details: error.message
            }, null, 2),
          },
        ],
      };
    }
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
