import "dotenv/config";
import express from "express";
import { createServer } from "http";
import net from "net";
import { createExpressMiddleware } from "@trpc/server/adapters/express";
import { registerOAuthRoutes } from "./oauth";
import { registerStorageProxy } from "./storageProxy";
import { appRouter } from "../routers";
import { createContext } from "./context";
import { serveStatic, setupVite } from "./vite";
import { powerAutomateTranscriptHandler, teamsGraphNotificationHandler } from "../teamsIntegration";
import { registerLocalFileRoutes } from "../localFiles";

function isPortAvailable(port: number): Promise<boolean> {
  return new Promise(resolve => {
    const server = net.createServer();
    server.listen(port, () => {
      server.close(() => resolve(true));
    });
    server.on("error", () => resolve(false));
  });
}

async function findAvailablePort(startPort: number = 3000): Promise<number> {
  for (let port = startPort; port < startPort + 20; port++) {
    if (await isPortAvailable(port)) {
      return port;
    }
  }
  throw new Error(`No available port found starting from ${startPort}`);
}

async function startServer() {
  const app = express();
  const server = createServer(app);
  // Configure body parser with larger size limit for file uploads
  // A one-hour screen recording can be close to 180 MB. The browser sends a
  // data URL and the local server decodes it immediately into a file.
  app.use(express.json({ limit: "260mb" }));
  app.use(express.urlencoded({ limit: "260mb", extended: true }));
  registerStorageProxy(app);
  registerLocalFileRoutes(app);
  registerOAuthRoutes(app);
  // Microsoft Graph validates this endpoint with validationToken and later posts artifact events.
  app.get("/api/integrations/teams/notifications", teamsGraphNotificationHandler);
  app.post("/api/integrations/teams/notifications", teamsGraphNotificationHandler);
  // Power Automate can call this endpoint after retrieving the .vtt transcript from Teams.
  app.post("/api/integrations/teams/transcript", powerAutomateTranscriptHandler);
  // tRPC API
  app.use(
    "/api/trpc",
    createExpressMiddleware({
      router: appRouter,
      createContext,
    })
  );
  // development mode uses Vite, production mode uses static files
  if (process.env.NODE_ENV === "development") {
    await setupVite(app, server);
  } else {
    serveStatic(app);
  }

  const preferredPort = parseInt(process.env.PORT || "3000");
  const host = process.env.HOST || "0.0.0.0";
  const port = await findAvailablePort(preferredPort);

  if (port !== preferredPort) {
    console.log(`Port ${preferredPort} is busy, using port ${port} instead`);
  }

  server.listen(port, host, () => {
    console.log(`Server running on http://localhost:${port}/`);
    console.log(`LAN access: http://<IP-DE-ESTE-PC>:${port}/`);
  });
}

startServer().catch(console.error);
