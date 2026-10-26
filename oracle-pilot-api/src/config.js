function csv(value) {
  return new Set(String(value || "").split(",").map((item) => item.trim().toLowerCase()).filter(Boolean));
}

function required(env, name) {
  const value = String(env[name] || "").trim();
  if (!value) throw new Error(`Missing required environment variable: ${name}`);
  return value;
}

export function loadConfig(env = process.env) {
  const authMode = String(env.AUTH_MODE || "firebase").trim().toLowerCase();
  if (!new Set(["firebase", "oidc"]).has(authMode)) throw new Error("AUTH_MODE must be firebase or oidc");
  const config = {
    host: String(env.HOST || "127.0.0.1"),
    port: Number(env.PORT || 18082),
    databaseUrl: required(env, "DATABASE_URL"),
    dbSsl: String(env.DB_SSL || "false").toLowerCase() === "true",
    authMode,
    firebaseProjectId: String(env.FIREBASE_PROJECT_ID || "hera-app-6cd2b"),
    allowedEmails: csv(env.PILOT_ALLOWED_EMAILS),
    adminEmails: csv(env.PILOT_ADMIN_EMAILS),
    allowedOrigins: csv(env.ALLOWED_ORIGINS),
    oidcIssuer: String(env.OIDC_ISSUER || ""),
    oidcAudience: String(env.OIDC_AUDIENCE || ""),
    oidcJwksUri: String(env.OIDC_JWKS_URI || "")
  };
  if (!Number.isInteger(config.port) || config.port < 1 || config.port > 65535) throw new Error("PORT is invalid");
  if (authMode === "oidc" && (!config.oidcIssuer || !config.oidcAudience || !config.oidcJwksUri)) {
    throw new Error("OIDC_ISSUER, OIDC_AUDIENCE and OIDC_JWKS_URI are required in oidc mode");
  }
  if (!config.allowedOrigins.size) throw new Error("ALLOWED_ORIGINS must contain at least one exact origin");
  return config;
}
