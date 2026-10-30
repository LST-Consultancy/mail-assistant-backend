import { createHash } from "node:crypto";
import type { RequestHandler } from "express";
import { createRemoteJWKSet, jwtVerify, type JWTVerifyGetKey } from "jose";
import { config } from "../config.js";

export interface AuthUser {
  /** sha256(tid:oid) — the only user identifier we log or key limits on. */
  idHash: string;
  tid: string;
}

declare module "express-serve-static-core" {
  interface Request {
    user?: AuthUser;
  }
}

// v2 keys cover work/school and personal (consumer tenant) accounts.
const entraJwks = createRemoteJWKSet(new URL("https://login.microsoftonline.com/common/discovery/v2.0/keys"));

export interface ValidateOptions {
  clientId: string;
  allowedTenantIds: string[];
  getKey?: JWTVerifyGetKey;
}

export async function validateToken(token: string, opts: ValidateOptions): Promise<AuthUser> {
  const { payload } = await jwtVerify(token, opts.getKey ?? entraJwks, {
    audience: [opts.clientId, `api://${opts.clientId}`],
    algorithms: ["RS256"],
  });
  const tid = String(payload.tid ?? "");
  const oid = String(payload.oid ?? payload.sub ?? "");
  // Multi-tenant authority: the issuer must match the token's own tenant.
  if (!tid || payload.iss !== `https://login.microsoftonline.com/${tid}/v2.0`) {
    throw new Error("invalid issuer");
  }
  if (opts.allowedTenantIds.length && !opts.allowedTenantIds.includes(tid)) {
    throw new Error("tenant not allowed");
  }
  const scopes = String(payload.scp ?? "").split(" ");
  if (!scopes.includes("access_as_user")) throw new Error("missing scope");
  return { tid, idHash: createHash("sha256").update(`${tid}:${oid}`).digest("hex").slice(0, 16) };
}

export function requireAuth(getKey?: JWTVerifyGetKey): RequestHandler {
  return async (req, res, next) => {
    const token = req.headers.authorization?.match(/^Bearer (.+)$/)?.[1];
    if (!token) return void res.status(401).json({ error: "Missing bearer token" });
    if (config.devAuthBypass && token === "dev") {
      req.user = { tid: "dev", idHash: "dev" };
      return next();
    }
    try {
      req.user = await validateToken(token, {
        clientId: config.entraClientId,
        allowedTenantIds: config.allowedTenantIds,
        getKey,
      });
      next();
    } catch {
      res.status(401).json({ error: "Invalid or expired token" });
    }
  };
}
