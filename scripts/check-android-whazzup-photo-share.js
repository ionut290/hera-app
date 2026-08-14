#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const app = fs.readFileSync("app.js", "utf8");
const mainActivity = fs.readFileSync("android/app/src/main/java/it/vargacantieri/hera/MainActivity.java", "utf8");
const plugin = fs.readFileSync(
  "android/app/src/main/java/it/vargacantieri/hera/whatsapp/HeraWhazzupPhotosPlugin.java",
  "utf8"
);

assert.match(mainActivity, /import it\.vargacantieri\.hera\.whatsapp\.HeraWhazzupPhotosPlugin;/);
assert.match(mainActivity, /registerPlugin\(HeraWhazzupPhotosPlugin\.class\);/);

const requiredNativeMarkers = [
  '@CapacitorPlugin(name = "HeraWhazzupPhotos")',
  'Intent.ACTION_SEND_MULTIPLE',
  'Intent.ACTION_SEND',
  'intent.setPackage(packageName);',
  'intent.putExtra(Intent.EXTRA_TEXT, message);',
  'intent.putExtra(Intent.EXTRA_STREAM, photoUris.get(0));',
  'intent.putParcelableArrayListExtra(Intent.EXTRA_STREAM, photoUris);',
  'FileProvider.getUriForFile(',
  'Intent.FLAG_GRANT_READ_URI_PERMISSION',
  'getContext().grantUriPermission(packageName, uri, Intent.FLAG_GRANT_READ_URI_PERMISSION);',
  'private static final String WHATSAPP = "com.whatsapp";',
  'private static final String WHATSAPP_BUSINESS = "com.whatsapp.w4b";',
  'result.put("fallback", false);'
];
requiredNativeMarkers.forEach((marker) => assert.ok(plugin.includes(marker), `Marker nativo mancante: ${marker}`));

[
  "Intent.ACTION_VIEW",
  "web.whatsapp.com",
  "https://wa.me/",
  "api.whatsapp.com",
  "Intent.createChooser"
].forEach((marker) => assert.ok(!plugin.includes(marker), `Fallback o chooser vietato nel plugin foto: ${marker}`));

assert.match(app, /function getDedicatedAndroidWhazzupPhotoPlugin\(\)/);
assert.match(app, /registerPlugin\?\.\("HeraWhazzupPhotos"\)/);
assert.match(app, /async function shareWhazzupPhotosDedicatedAndroid\(plugin, orderedFiles, message\)/);
assert.match(app, /await plugin\.begin\(\)/);
assert.match(app, /await plugin\.addPhoto\(\{/);
assert.match(app, /return await plugin\.share\(\{/);
assert.match(app, /const dedicatedPlugin = getDedicatedAndroidWhazzupPhotoPlugin\(\);/);
assert.match(app, /return shareWhazzupPhotosDedicatedAndroid\(dedicatedPlugin, orderedFiles, message\);/);

const dedicatedStart = app.indexOf("async function shareWhazzupPhotosDedicatedAndroid");
const nativeStart = app.indexOf("async function shareWhazzupPhotosNativeAndroid", dedicatedStart);
const dedicatedSource = app.slice(dedicatedStart, nativeStart);
assert.ok(dedicatedSource.indexOf("await plugin.addPhoto") < dedicatedSource.indexOf("return await plugin.share"));
assert.ok(dedicatedSource.indexOf("fileName:") < dedicatedSource.indexOf("text: message"));

console.log("Android Whazzup photo-share checks passed.");
