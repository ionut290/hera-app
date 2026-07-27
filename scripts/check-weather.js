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
    assert.equal(JSON.parse(success.body).provider, "Open-Meteo Best Match");
    assert.match(calls[0], /latitude=44\.4949/);
    assert.match(calls[0], /longitude=11\.3426/);
    assert.match(calls[0], /minutely_15=/);
    assert.match(calls[0], /models=best_match/);
    assert.match(calls[0], /cell_selection=land/);

    const invalid = await weatherFunction.handler({
      httpMethod: "GET",
      queryStringParameters: { lat: "200", lon: "11.3426" }
    });
    assert.equal(invalid.statusCode, 400);
    assert.equal(calls.length, 1);

    global.fetch = async (url) => {
      calls.push(String(url));
      if (String(url).includes("open-meteo.com")) return { ok: false, status: 503 };
      return {
        ok: true,
        status: 200,
        json: async () => ({
          type: "Feature",
          geometry: { coordinates: [11.3426, 44.4949, 54] },
          properties: {
            meta: {
              units: {
                air_temperature: "celsius",
                relative_humidity: "%",
                precipitation_amount: "mm",
                probability_of_precipitation: "%"
              }
            },
            timeseries: [{
              time: "2026-07-26T10:00:00Z",
              data: {
                instant: {
                  details: {
                    air_temperature: 23,
                    relative_humidity: 55,
                    wind_speed: 3,
                    wind_from_direction: 180,
                    wind_speed_of_gust: 5
                  }
                },
                next_1_hours: {
                  summary: { symbol_code: "partlycloudy_day" },
                  details: { precipitation_amount: 0, probability_of_precipitation: 10 }
                }
              }
            }]
          }
        })
      };
    };
    const fallback = await weatherFunction.handler({
      httpMethod: "GET",
      queryStringParameters: { lat: "44.4949", lon: "11.3426", operational: "1" }
    });
    const fallbackBody = JSON.parse(fallback.body);
    assert.equal(fallback.statusCode, 200);
    assert.equal(fallbackBody.provider, "MET Norway fallback");
    assert.equal(fallbackBody.current.temperature_2m, 23);
    assert.equal(fallbackBody.hourly.wind_speed_10m[0], 10.8);
    assert.match(calls.at(-1), /api\.met\.no/);

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
  const weatherProxy = fs.readFileSync("netlify/functions/weather.js", "utf8");

  assert.match(app, /WEATHER_PROXY_PUBLIC_URL/);
  assert.match(app, /Open-Meteo Best Match/);
  assert.match(weatherProxy, /MET Norway fallback/);
  assert.doesNotMatch(app, /fetchOpenWeatherPrimary/);
  const appAsset = index.match(/app\.js\?v=[^"']+/)?.[0];
  assert.ok(appAsset, "index.html deve caricare app.js con una versione cache-busting");
  assert.match(serviceWorker, /hera-app-shell-v\d+/);
  assert.ok(serviceWorker.includes(`./${appAsset}`), "Il Service Worker deve precaricare la stessa versione di app.js usata da index.html");
  assert.match(netlify, /from = "\/api\/weather"/);

  console.log("Weather fallback checks passed.");
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
