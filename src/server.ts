import { readFileSync } from "node:fs";
import { createServer } from "node:https";
import cors from "cors";
import express from "express";
import type { JWTVerifyGetKey } from "jose";
import { requireAuth } from "./auth/validateToken.js";
import { config } from "./config.js";
import { errorHandler, requestLog } from "./middleware/errors.js";
import { rateLimit, UserLimits } from "./middleware/rateLimit.js";
import { api } from "./routes/api.js";

export function createApp(opts: { getKey?: JWTVerifyGetKey; limits?: UserLimits } = {}) {
  const limits = opts.limits ?? new UserLimits(config.requestsPerMinute, config.dailyTokenBudget);
  const app = express();
  app.disable("x-powered-by");
  app.use(cors({ origin: config.addinOrigins, allowedHeaders: ["Authorization", "Content-Type"] }));
  app.use(requestLog);
  app.get("/api/health", (_req, res) => void res.json({ ok: true }));
  app.use("/api", requireAuth(opts.getKey), rateLimit(limits), (req, res, next) => {
    res.on("close", () => limits.addTokens(req.user!.idHash, res.locals.tokens ?? 0));
    next();
  });
  app.use(express.json({ limit: "2mb" }));
  app.use("/api", api);
  app.use(errorHandler);
  return app;
}

if (import.meta.url === `file://${process.argv[1]}`) {
  if (!config.entraClientId && !config.devAuthBypass) throw new Error("ENTRA_CLIENT_ID is required");
  const app = createApp();
  const log = () => console.log(`Mail Assistant API on :${config.port}`);
  // Local dev: reuse the office-addin-dev-certs pair. In production TLS terminates at the host/ingress.
  if (process.env.TLS_CERT_FILE && process.env.TLS_KEY_FILE) {
    createServer(
      { cert: readFileSync(process.env.TLS_CERT_FILE), key: readFileSync(process.env.TLS_KEY_FILE) },
      app,
    ).listen(config.port, log);
  } else {
    app.listen(config.port, log);
  }
}
