import { getApps, initializeApp } from "firebase-admin/app";
import { getAuth } from "firebase-admin/auth";
import { createRemoteJWKSet, jwtVerify } from "jose";

function bearer(req) {
  const match = /^Bearer\s+(.+)$/i.exec(String(req.headers.authorization || ""));
  return match?.[1] || "";
}

function normalizeIdentity(claims) {
  return {
    uid: String(claims.uid || claims.sub || ""),
    email: String(claims.email || "").trim().toLowerCase(),
    emailVerified: claims.email_verified === true,
    accessClaim: claims.varga_access === true,
    roleClaim: String(claims.varga_role || "").trim().toLowerCase()
  };
}

export function createTokenVerifier(config) {
  if (config.authMode === "firebase") {
    const app = getApps()[0] || initializeApp({ projectId: config.firebaseProjectId });
    return async (token) => normalizeIdentity(await getAuth(app).verifyIdToken(token, config.checkRevokedTokens));
  }
  const jwks = createRemoteJWKSet(new URL(config.oidcJwksUri));
  return async (token) => {
    const { payload } = await jwtVerify(token, jwks, {
      issuer: config.oidcIssuer,
      audience: config.oidcAudience,
      algorithms: ["RS256", "ES256"]
    });
    return normalizeIdentity(payload);
  };
}

export function createAuthMiddleware(config, verifyToken = createTokenVerifier(config), authorizeIdentity) {
  return async function authenticate(req, res, next) {
    const token = bearer(req);
    if (!token) return res.status(401).json({ error: "authentication_required" });
    let identity;
    try {
      identity = await verifyToken(token);
    } catch {
      return res.status(401).json({ error: "invalid_or_expired_token" });
    }
    if (!identity.uid || !identity.email || !identity.emailVerified) {
      return res.status(403).json({ error: "verified_identity_required" });
    }
    try {
      const authzSource = config.authzSource || "env";
      const envAdmin = config.adminEmails.has(identity.email) || identity.roleClaim === "admin";
      const envAllowed = envAdmin || config.allowedEmails.has(identity.email) || identity.accessClaim;
      let databaseAuth = { matched: false, allowed: false, role: "reader" };
      if (authzSource !== "env") {
        if (typeof authorizeIdentity !== "function") throw new Error("database authorization unavailable");
        databaseAuth = await authorizeIdentity(identity);
      }
      const databaseAdmin = databaseAuth.allowed && databaseAuth.role === "admin";
      const isAdmin = envAdmin || databaseAdmin;
      const isAllowed = authzSource === "database"
        ? databaseAuth.allowed
        : authzSource === "hybrid"
          ? (databaseAuth.matched ? databaseAuth.allowed : envAllowed)
          : envAllowed;
      if (!isAllowed) return res.status(403).json({ error: "pilot_access_denied" });
      req.identity = { ...identity, role: isAdmin ? "admin" : "reader" };
      return next();
    } catch (error) { return next(error); }
  };
}
