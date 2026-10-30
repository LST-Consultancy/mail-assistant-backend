import { randomUUID } from "node:crypto";
import type { ErrorRequestHandler, RequestHandler } from "express";
import { ZodError } from "zod";
import { friendlyError } from "../llm/client.js";

/** Metadata-only access log: never log bodies or AI output (privacy, see README). */
export const requestLog: RequestHandler = (req, res, next) => {
  const id = randomUUID();
  const start = Date.now();
  res.locals.requestId = id;
  res.locals.tokens = 0;
  res.setHeader("X-Request-Id", id);
  res.on("close", () => {
    console.log(
      JSON.stringify({
        id,
        path: req.originalUrl.split("?")[0],
        status: res.statusCode,
        user: req.user?.idHash,
        tokens: res.locals.tokens,
        ms: Date.now() - start,
        error: res.locals.errorKind,
      }),
    );
  });
  next();
};

export const errorHandler: ErrorRequestHandler = (err, _req, res, _next) => {
  if (err instanceof ZodError) {
    res.locals.errorKind = "validation";
    return void res.status(400).json({ error: "Invalid request", issues: err.issues.map((i) => i.path.join(".")) });
  }
  if ((err as { type?: string }).type === "entity.too.large") {
    return void res.status(413).json({ error: "Request too large" });
  }
  const { status, message } = friendlyError(err);
  res.locals.errorKind = (err as Error)?.constructor?.name ?? "unknown";
  if (res.headersSent) return void res.end();
  res.status(status).json({ error: message });
};
