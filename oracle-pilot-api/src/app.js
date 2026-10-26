import crypto from "node:crypto";
import cors from "cors";
import express from "express";
import rateLimit from "express-rate-limit";
import helmet from "helmet";
import { createAuthMiddleware } from "./auth.js";
import { parseListQuery } from "./repository.js";

export function createApp({ config, repository, verifyToken }) {
  const app = express();
  app.disable("x-powered-by");
  app.set("trust proxy", 1);
  app.use((req, res, next) => {
    req.requestId = String(req.headers["x-request-id"] || crypto.randomUUID()).slice(0, 128);
    res.setHeader("x-request-id", req.requestId);
    next();
  });
  app.use(helmet({ contentSecurityPolicy: false }));
  app.use(cors({
    origin(origin, callback) {
      if (!origin || config.allowedOrigins.has(String(origin).toLowerCase())) return callback(null, true);
      return callback(Object.assign(new Error("origin not allowed"), { status: 403 }));
    },
    methods: ["GET"],
    allowedHeaders: ["Authorization", "Content-Type", "X-Request-Id"],
    maxAge: 600
  }));
  app.use(rateLimit({ windowMs: 60_000, limit: 120, standardHeaders: "draft-8", legacyHeaders: false }));
  app.use(express.json({ limit: "16kb" }));

  app.get("/healthz", (_req, res) => res.json({ status: "ok" }));
  app.get("/readyz", async (_req, res, next) => {
    try { await repository.ready(); res.json({ status: "ready" }); } catch (error) { next(error); }
  });

  const authenticate = createAuthMiddleware(config, verifyToken, repository.authorizeIdentity?.bind(repository));
  app.get("/v1/me", authenticate, (req, res) => {
    res.json({ uid: req.identity.uid, email: req.identity.email, role: req.identity.role });
  });
  app.get("/v1/catalogo", authenticate, async (req, res, next) => {
    try {
      const page = parseListQuery(req.query);
      const items = await repository.catalog(page);
      res.json({ items, page: { ...page, returned: items.length } });
    } catch (error) { next(error); }
  });
  app.get("/v1/mappa", authenticate, async (req, res, next) => {
    try {
      const page = parseListQuery(req.query);
      const items = await repository.map(page);
      res.json({ items, page: { ...page, returned: items.length } });
    } catch (error) { next(error); }
  });

  app.use((_req, res) => res.status(404).json({ error: "not_found" }));
  app.use((error, req, res, _next) => {
    const status = Number(error.status) >= 400 && Number(error.status) < 500 ? Number(error.status) : 503;
    if (status >= 500) console.error(JSON.stringify({ level: "error", requestId: req.requestId, message: error.message }));
    res.status(status).json({ error: status >= 500 ? "service_unavailable" : error.message, requestId: req.requestId });
  });
  return app;
}
