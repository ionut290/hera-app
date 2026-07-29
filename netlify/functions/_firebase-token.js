"use strict";

const crypto = require("crypto");

const CERT_URL = "https://www.googleapis.com/robot/v1/metadata/x509/securetoken@system.gserviceaccount.com";
let certificateCache = { expiresAt: 0, certificates: null };

function decodeBase64Url(value) {
  const normalized = String(value || "").replace(/-/g, "+").replace(/_/g, "/");
  const padding = normalized.length % 4 ? "=".repeat(4 - (normalized.length % 4)) : "";
  return Buffer.from(normalized + padding, "base64");
}

function parseJsonSegment(segment, label) {
  try {
    return JSON.parse(decodeBase64Url(segment).toString("utf8"));
  } catch (_) {
    throw new Error(`Token Firebase non valido (${label}).`);
  }
}

function cacheLifetime(headers) {
  const cacheControl = String(headers.get("cache-control") || "");
  const match = cacheControl.match(/max-age=(\d+)/i);
  const seconds = match ? Number(match[1]) : 3600;
  return Math.max(300, Math.min(seconds, 86400)) * 1000;
}

async function getCertificates() {
  if (certificateCache.certificates && Date.now() < certificateCache.expiresAt) {
    return certificateCache.certificates;
  }
  const response = await fetch(CERT_URL, { headers: { Accept: "application/json" } });
  if (!response.ok) throw new Error(`Certificati Firebase non disponibili (HTTP ${response.status}).`);
  const certificates = await response.json();
  certificateCache = {
    certificates,
    expiresAt: Date.now() + cacheLifetime(response.headers)
  };
  return certificates;
}

function bearerToken(headers = {}) {
  const authorization = headers.authorization || headers.Authorization || "";
  const match = String(authorization).match(/^Bearer\s+(.+)$/i);
  if (!match) throw new Error("Token Firebase mancante.");
  return match[1].trim();
}

async function verifyFirebaseToken(token, projectId) {
  const segments = String(token || "").split(".");
  if (segments.length !== 3) throw new Error("Token Firebase non valido.");
  const [headerSegment, payloadSegment, signatureSegment] = segments;
  const header = parseJsonSegment(headerSegment, "header");
  const payload = parseJsonSegment(payloadSegment, "payload");
  if (header.alg !== "RS256" || !header.kid) throw new Error("Firma del token Firebase non valida.");

  const certificates = await getCertificates();
  const certificate = certificates[header.kid];
  if (!certificate) {
    certificateCache.expiresAt = 0;
    const refreshed = await getCertificates();
    if (!refreshed[header.kid]) throw new Error("Certificato Firebase non riconosciuto.");
  }
  const publicCertificate = certificate || certificateCache.certificates[header.kid];
  const verified = crypto.verify(
    "RSA-SHA256",
    Buffer.from(`${headerSegment}.${payloadSegment}`),
    publicCertificate,
    decodeBase64Url(signatureSegment)
  );
  if (!verified) throw new Error("Firma del token Firebase non valida.");

  const now = Math.floor(Date.now() / 1000);
  const expectedIssuer = `https://securetoken.google.com/${projectId}`;
  if (payload.aud !== projectId || payload.iss !== expectedIssuer) throw new Error("Token Firebase destinato a un altro progetto.");
  if (!payload.sub || String(payload.sub).length > 128) throw new Error("Utente Firebase non valido.");
  if (!Number.isFinite(payload.exp) || payload.exp <= now) throw new Error("Sessione Firebase scaduta.");
  if (!Number.isFinite(payload.iat) || payload.iat > now + 300) throw new Error("Data del token Firebase non valida.");
  return payload;
}

async function authenticateEvent(event) {
  const projectId = String(process.env.FIREBASE_PROJECT_ID || "hera-app-6cd2b").trim();
  return verifyFirebaseToken(bearerToken(event.headers || {}), projectId);
}

module.exports = { authenticateEvent, verifyFirebaseToken, bearerToken };
