"use strict";

const assert = require("assert");
const fs = require("fs");

const client = fs.readFileSync("shared-static-views-client.js", "utf8");

assert.match(client, /subscribePersonale = \(\) => subscribe\("personale"\)/);
assert.match(client, /subscribeMezzi = \(\) => subscribe\("mezzi"\)/);
assert.match(client, /collection\("sharedStaticViews"\)\.doc\("registri__corrente"\)/);
assert.strictEqual((client.match(/\.onSnapshot\(/g) || []).length, 1);
assert.match(client, /Array\.isArray\(view\.payload\.personale\)/);
assert.match(client, /Array\.isArray\(view\.payload\.mezzi\)/);
assert.match(client, /const CALENDAR_VIEW_FALLBACK_MS = 5000/);
assert.match(client, /subscribeHoursStats = gatedHoursStats/);
assert.doesNotMatch(client, /bindCapture\("open-hours-btn", enableHoursSource\)/);
assert.match(client, /bindCapture\("open-hours-btn", \(\) => subscribeStaticCalendar\(visibleCalendarMonth\(\)\)\)/);
assert.match(client, /hoursApprovalRequests = approvalReports/);
assert.match(client, /hoursApprovalsLoaded = true/);
assert.match(client, /renderHoursApprovalRequests\(\)/);
assert.match(client, /where\("date", ">=", from\)/);
assert.match(client, /where\("date", "<=", to\)/);
assert.match(client, /monthly-query-fallback/);
assert.match(client, /changeCalendarMonth = function/);
assert.match(client, /showCalendarToday = function/);
assert.match(client, /selectCalendarDate = function/);
assert.match(client, /calendarFallbackLoads/);

console.log("Shared registries and monthly calendar client checks passed.");
