import assert from "node:assert/strict";
import test from "node:test";
import { createApp } from "../src/app.js";

const config = {
  allowedOrigins: new Set(["http://localhost:18081"]),
  allowedEmails: new Set(["reader@example.test"]),
  adminEmails: new Set(["admin@example.test"]),
  authMode: "firebase",
  firebaseProjectId: "test"
};
const repository = {
  async ready() {},
  async catalog(page) { return [{ collection_name: page.collection, document_id: "c1" }]; },
  async map(page) { return [{ collection_name: page.collection, document_id: "m1", latitudine: "44", longitudine: "11" }]; }
};
const verifyToken = async (token) => {
  if (token === "reader") return { uid: "u1", email: "reader@example.test", emailVerified: true, accessClaim: false, roleClaim: "" };
  if (token === "admin") return { uid: "u2", email: "admin@example.test", emailVerified: true, accessClaim: false, roleClaim: "" };
  if (token === "denied") return { uid: "u3", email: "other@example.test", emailVerified: true, accessClaim: false, roleClaim: "" };
  throw new Error("bad token");
};

async function withServer(run) {
  const server = createApp({ config, repository, verifyToken }).listen(0, "127.0.0.1");
  await new Promise((resolve) => server.once("listening", resolve));
  try { await run(`http://127.0.0.1:${server.address().port}`); }
  finally { await new Promise((resolve) => server.close(resolve)); }
}

test("health is public but data requires authentication", () => withServer(async (base) => {
  assert.equal((await fetch(`${base}/healthz`)).status, 200);
  assert.equal((await fetch(`${base}/v1/catalogo`)).status, 401);
}));

test("deny-by-default authorization and role mapping", () => withServer(async (base) => {
  assert.equal((await fetch(`${base}/v1/me`, { headers: { authorization: "Bearer denied" } })).status, 403);
  const response = await fetch(`${base}/v1/me`, { headers: { authorization: "Bearer admin" } });
  assert.equal(response.status, 200);
  assert.equal((await response.json()).role, "admin");
}));

test("catalog validates filters and caps pagination", () => withServer(async (base) => {
  const invalid = await fetch(`${base}/v1/catalogo?collection=personale`, { headers: { authorization: "Bearer reader" } });
  assert.equal(invalid.status, 400);
  const response = await fetch(`${base}/v1/catalogo?collection=impianti&limit=9999`, { headers: { authorization: "Bearer reader" } });
  assert.equal(response.status, 200);
  const body = await response.json();
  assert.equal(body.page.limit, 500);
  assert.equal(body.items[0].collection_name, "impianti");
}));

test("CORS rejects an unknown browser origin", () => withServer(async (base) => {
  const response = await fetch(`${base}/v1/me`, { headers: { origin: "https://evil.example", authorization: "Bearer reader" } });
  assert.equal(response.status, 403);
}));
