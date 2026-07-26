"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const weatherFunction = require("../netlify/functions/weather.js");

async function run() {
  const originalFetch = global.fetch;
  const calls = [];
  try {
    global.fetch = async (url) => {
      calls.push(String(url));
      return {
        ok: true,
        status: 200,
        json: async () => ({
          current: { temperature_2m: 24, wind_speed_10m: 6, weather_code: 3 },
          hourly: { time: ["2026-07-26T10:00"], temperature_2m: [24] }
        })
      };
    };

    const success = await weatherFunction.handler({
      httpMethod: "GET",
      queryStringParameters: { lat: "44.4949", lon: "11.3426", operational: "1" }
    });
    assert.equal(success.statusCode, 200);
    assert.equal(JSON.parse(success.body).provider, "Open-Meteo proxy");
    assert.match(calls[0], /latitude=44\.4949/);
    assert.match(calls[0], /longitude=11\.3426/);
    assert.match(calls[0], /minutely_15=/);

    const invalid = await weatherFunction.handler({
      httpMethod: "GET",
      queryStringParameters: { lat: "200", lon: "11.3426" }
    });
    assert.equal(invalid.statusCode, 400);
    assert.equal(calls.length, 1);

    global.fetch = async () => ({ ok: false, status: 503 });
    const originalConsoleError = console.error;
    try {
      console.error = () => {};
      const unavailable = await weatherFunction.handler({
        httpMethod: "GET",
        queryStringParameters: { lat: "44.4949", lon: "11.3426" }
      });
      assert.equal(unavailable.statusCode, 502);
      assert.equal(unavailable.headers["Cache-Control"], "no-store");
    } finally {
      console.error = originalConsoleError;
    }
  } finally {
    global.fetch = originalFetch;
  }

  const app = fs.readFileSync("app.js", "utf8");
  const index = fs.readFileSync("index.html", "utf8");
  const serviceWorker = fs.readFileSync("sw.js", "utf8");
  const netlify = fs.readFileSync("netlify.toml", "utf8");

  assert.match(app, /WEATHER_PROXY_PUBLIC_URL/);
  assert.match(app, /Open-Meteo proxy/);
  assert.match(index, /app\.js\?v=20260726-weather1/);
  assert.match(serviceWorker, /hera-app-shell-v30/);
  assert.match(serviceWorker, /app\.js\?v=20260726-weather1/);
  assert.match(netlify, /from = "\/api\/weather"/);

  console.log("Weather fallback checks passed.");
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
