"use strict";

const assert = require("assert");
const fs = require("fs");

const wrapper = fs.readFileSync("shared-static-views-client.js", "utf8");
const core = fs.readFileSync("shared-static-views-client-core.js", "utf8");
const explicitGuard = fs.readFileSync("hours-source-explicit-guard.js", "utf8");

assert.match(wrapper, /function stopPrematureHoursSubscriptions\(\)/);
assert.match(wrapper, /typeof unsubscribeHoursStats === "function"/);
assert.match(wrapper, /unsubscribeHoursStats\(\);\s*unsubscribeHoursStats = null/);
assert.match(wrapper, /typeof unsubscribeHoursApprovals === "function"/);
assert.match(wrapper, /unsubscribeHoursApprovals\(\);\s*unsubscribeHoursApprovals = null/);
assert.match(wrapper, /stopPrematureHoursSubscriptions\(\);[\s\S]*loadCore\(\);/);
assert.match(wrapper, /version: "1\.3\.0"/);
assert.match(wrapper, /shared-static-views-client-core\.js\?v=20260806-explicit-hours-v3/);

assert.match(core, /subscribePersonale = \(\) => subscribe\("personale"\)/);
assert.match(core, /subscribeMezzi = \(\) => subscribe\("mezzi"\)/);
assert.match(core, /collection\("sharedStaticViews"\)\.doc\("registri__corrente"\)/);
assert.strictEqual((core.match(/\.onSnapshot\(/g) || []).length, 1);
assert.match(core, /subscribeHoursStats = gatedHoursStats/);
assert.match(core, /bindCapture\("open-hours-btn", enableHoursSource\)/);
assert.match(core, /function stopStaticCalendarForFullHours\(\)/);
assert.match(core, /lazyStartup\.calendarUnsubscribe\(\)/);
assert.match(core, /lazyStartup\.calendarUnsubscribe = null/);
assert.match(core, /if \(lazyStartup\.hoursSourceEnabled\) \{/);
assert.match(core, /aggiornamento calendario ridotto ignorato: ore complete attive/);
assert.match(core, /stopStaticCalendarForFullHours\(\);[\s\S]*sourceSubscriptions\.hoursStats\(\)/);
assert.match(core, /ignoredStaticCalendarUpdates/);
assert.doesNotMatch(core, /CALENDAR_VIEW_FALLBACK_MS/);
assert.doesNotMatch(core, /monthly-query-fallback/);

assert.match(explicitGuard, /trigger\.isTrusted === true/);
assert.match(explicitGuard, /trigger\.forceSharedCalendarFallback === true/);
assert.match(explicitGuard, /avvio automatico delle ore complete bloccato/);

console.log("Shared views startup and explicit full-hours checks passed.");
