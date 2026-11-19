"use strict";

// Download the CURRENT active, app-specific iOS App Store profile on a private
// macOS GitHub Actions runner. Never print the profile, key, JWT, or response.
const fs = require("node:fs");
const crypto = require("node:crypto");

const EXPECTED_NAME = "VARGA CANTIERI - TestFlight";
const EXPECTED_BUNDLE = "it.vargacantieri.hera";
const API = "https://api.appstoreconnect.apple.com";

function token(keyId, issuerId, privateKey) {
  const issuedAt = Math.floor(Date.now() / 1000);
  const header = Buffer.from(JSON.stringify({ alg: "ES256", kid: keyId, typ: "JWT" })).toString("base64url");
  const payload = Buffer.from(JSON.stringify({
    iss: issuerId,
    aud: "appstoreconnect-v1",
    iat: issuedAt,
    exp: issuedAt + 300
  })).toString("base64url");
  const data = header + "." + payload;
  const signature = crypto.sign("sha256", Buffer.from(data), {
    key: privateKey,
    dsaEncoding: "ieee-p1363"
  }).toString("base64url");
  return data + "." + signature;
}

async function fetchActiveProfile(env, request = fetch) {
  for (const field of ["IOS_APPSTORE_KEY_ID", "IOS_APPSTORE_ISSUER_ID", "IOS_APPSTORE_PRIVATE_KEY_BASE64"]) {
    if (!env[field]) throw new Error("Missing existing Apple API credential: " + field);
  }
  const privateKey = Buffer.from(env.IOS_APPSTORE_PRIVATE_KEY_BASE64, "base64").toString("utf8");
  if (!privateKey.includes("-----BEGIN PRIVATE KEY-----")) {
    throw new Error("The configured Apple API key is invalid.");
  }
  const jwt = token(env.IOS_APPSTORE_KEY_ID, env.IOS_APPSTORE_ISSUER_ID, privateKey);
  const url = new URL("/v1/profiles", API);
  url.searchParams.set("filter[name]", EXPECTED_NAME);
  url.searchParams.set("filter[profileType]", "IOS_APP_STORE");
  url.searchParams.set("filter[profileState]", "ACTIVE");
  url.searchParams.set("include", "bundleId");
  url.searchParams.set("limit", "200");
  const response = await request(url.toString(), {
    headers: { Authorization: "Bearer " + jwt, Accept: "application/json" },
    signal: AbortSignal.timeout(20000)
  });
  if (!response.ok) {
    throw new Error("Apple provisioning API unavailable (HTTP " + response.status + "); check App Store Connect API key access.");
  }
  const result = await response.json();
  const bundleIds = new Map((result.included || [])
    .filter((entry) => entry.type === "bundleIds")
    .map((entry) => [entry.id, entry.attributes?.identifier]));
  const candidates = (result.data || []).filter((entry) => {
    const a = entry.attributes || {};
    const linkedId = entry.relationships?.bundleId?.data?.id;
    return entry.type === "profiles" && a.name === EXPECTED_NAME &&
      a.profileState === "ACTIVE" && a.profileType === "IOS_APP_STORE" &&
      bundleIds.get(linkedId) === EXPECTED_BUNDLE;
  });
  if (candidates.length !== 1) {
    throw new Error("Expected exactly one active App Store profile bound to " + EXPECTED_BUNDLE +
      " (found " + candidates.length + "). No signing asset was written.");
  }
  const encoded = candidates[0].attributes?.profileContent;
  if (typeof encoded !== "string" || encoded.length < 100 ||
      !/^[A-Za-z0-9+/\r\n]+=*$/.test(encoded)) {
    throw new Error("Apple API did not return valid provisioning profile data.");
  }
  const data = Buffer.from(encoded.replace(/\s/g, ""), "base64");
  if (data.length < 100) throw new Error("Provisioning profile is unexpectedly empty.");
  return data;
}

async function main() {
  const destination = process.argv[2];
  if (!destination || !process.env.RUNNER_TEMP || !destination.startsWith(process.env.RUNNER_TEMP + "/")) {
    throw new Error("A private absolute destination under RUNNER_TEMP is required.");
  }
  const bytes = await fetchActiveProfile(process.env);
  process.umask(0o077);
  fs.writeFileSync(destination, bytes, { mode: 0o600, flag: "wx" });
  console.log("Retrieved active VARGA CANTIERI App Store provisioning profile from Apple API; saved privately.");
}

module.exports = { fetchActiveProfile, EXPECTED_NAME, EXPECTED_BUNDLE };
if (require.main === module) main().catch((error) => {
  console.error("Apple profile download failed:", error.message);
  process.exitCode = 1;
});
