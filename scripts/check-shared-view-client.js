"use strict";

const assert = require("assert");
const fs = require("fs");

const client = fs.readFileSync("shared-static-views-client.js", "utf8");

assert.match(client, /subscribePersonale = \(\) => subscribe\("personale"\)/);
assert.match(client, /subscribeMezzi = \(\) => subscribe\("mezzi"\)/);
assert.match(client, /collection\("sharedStaticViews"\)\.doc\("registri__corrente"\)/);
assert.strictEqual((client.match(/\.onSnapshot\(/g) || []).length, 1);
assert.match(client, /subscribeHoursStats = gatedHoursStats/);
assert.match(client, /bindCapture\("open-hours-btn", enableHoursSource\)/);
assert.match(client, /function stopStaticCalendarForFullHours\(\)/);
assert.match(client, /lazyStartup\.calendarUnsubscribe\(\)/);
assert.match(client, /lazyStartup\.calendarUnsubscribe = null/);
assert.match(client, /if \(lazyStartup\.hoursSourceEnabled\) \{/);
assert.match(client, /aggiornamento calendario ridotto ignorato: ore complete attive/);
assert.match(client, /stopStaticCalendarForFullHours\(\);[\s\S]*sourceSubscriptions\.hoursStats\(\)/);
assert.match(client, /ignoredStaticCalendarUpdates/);
assert.doesNotMatch(client, /CALENDAR_VIEW_FALLBACK_MS/);
assert.doesNotMatch(client, /monthly-query-fallback/);

console.log("Shared registries and full-hours priority checks passed.");
