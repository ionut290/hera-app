#!/usr/bin/env node
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const source = fs.readFileSync(path.join(__dirname, "..", "accounting-view-guard.js"), "utf8");
let receivedBounds = null;
let invalidMarkerCalls = 0;

const context = {
  window: {
    addImpiantoMarkerToMapLayer() {
      invalidMarkerCalls += 1;
      return {};
    },
    renderMap() {
      context.map.fitBounds([
        ["", "11.3426"],
        [0, 11.3426],
        ["44,6339", "11,6679"],
        [44.58395, 11.60916]
      ], { maxZoom: 11 });
    }
  },
  map: {
    fitBounds(bounds) {
      receivedBounds = bounds;
      return this;
    }
  }
};

vm.createContext(context);
vm.runInContext(source, context);

const checks = [];
const validCoordinates = context.window.HeraMapAutoFit.getPlantCoordinates({
  gpsY: "44,6339",
  gpsX: "11,6679"
});
checks.push(["coordinate italiane normalizzate", JSON.stringify(validCoordinates) === JSON.stringify([44.6339, 11.6679])]);
checks.push(["coordinate vuote escluse", context.window.HeraMapAutoFit.getPlantCoordinates({ gpsY: "", gpsX: "" }) === null]);
checks.push(["coordinate zero escluse", context.window.HeraMapAutoFit.getPlantCoordinates({ gpsY: 0, gpsX: 11.3426 }) === null]);

context.window.addImpiantoMarkerToMapLayer({ gpsY: "", gpsX: "11.3426" });
checks.push(["marker senza GPS non creato", invalidMarkerCalls === 0]);

context.window.renderMap();
checks.push([
  "fitBounds riceve tutti e soli gli impianti validi",
  JSON.stringify(receivedBounds) === JSON.stringify([[44.6339, 11.6679], [44.58395, 11.60916]])
]);

let failed = false;
for (const [name, passed] of checks) {
  console.log((passed ? "OK " : "FAIL ") + name);
  failed ||= !passed;
}
if (failed) process.exit(1);
