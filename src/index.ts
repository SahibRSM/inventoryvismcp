import express, { Request, Response } from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";

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

// Get Azure AD Token tool
const getAzureADToken = server.tool(
  "get-azure-ad-token",
  "Get Azure AD token from Microsoft",
  async (params) => {
    const { tenant_id, client_id, client_secret, grant_type } = params;
    
    const formData = new URLSearchParams();
    formData.append("client_id", client_id);
    formData.append("client_secret", client_secret);
    formData.append("grant_type", grant_type);
    formData.append("scope", "0cdb527f-a8d1-4bf8-9436-b352c68682b2/.default");

    const response = await fetch(
      `https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`,
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
  }
);

// Get Dynamics Token tool
const getDynamicsToken = server.tool(
  "get-dynamics-token",
  "Get access token for Dynamics 365 operations",
  async (params) => {
    const { bearer_token, grant_type, fno_id } = params;

    const response = await fetch(
      "https://securityservice.operations365.dynamics.com/token",
      {
        method: "POST",
        headers: {
          "Api-Version": "1.0",
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          grant_type: grant_type,
          client_assertion_type: "aad_app",
          client_assertion: bearer_token,
          scope: "https://inventoryservice.operations365.dynamics.com/.default",
          context: fno_id,
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
  }
);

// Query Inventory tool
const queryInventory = server.tool(
  "query-inventory",
  "Query inventory from Dynamics 365",
  async (params) => {
    const { access_token, fno_id, product_id, organization_id } = params;

    const response = await fetch(
      `https://inventoryservice.wus-il301.gateway.prod.island.powerapps.com/api/environment/${fno_id}/onhand/indexquery`,
      {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${access_token}`,
          "Api-Version": "2.0",
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          filters: {
            ProductId: [product_id],
            OrganizationId: [organization_id]
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
  }
);

const app = express();

// to support multiple simultaneous connections we have a lookup object from
// sessionId to transport
const transports: { [sessionId: string]: SSEServerTransport } = {};

app.get("/sse", async (req: Request, res: Response) => {
  // Get the full URI from the request
  const host = req.get("host");
  const fullUri = `https://${host}/jokes`;
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
  res.send("The Jokes and Dynamics 365 MCP server is running!");
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`✅ Server is running at http://localhost:${PORT}`);
});
