#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const app = fs.readFileSync("app.js", "utf8");
const orderFix = fs.readFileSync("android-whazzup-photo-order.js", "utf8");
const nativeRuntime = fs.readFileSync("native-android-runtime.js", "utf8");
const capacitorPrepare = fs.readFileSync("scripts/prepare-capacitor-web.js", "utf8");
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
  'intent.putExtra(Intent.EXTRA_STREAM, photoUris.get(0));',
  'intent.putParcelableArrayListExtra(Intent.EXTRA_STREAM, photoUris);',
  'FileProvider.getUriForFile(',
  'Intent.FLAG_GRANT_READ_URI_PERMISSION',
  'getContext().grantUriPermission(packageName, uri, Intent.FLAG_GRANT_READ_URI_PERMISSION);',
  'private static final String WHATSAPP = "com.whatsapp";',
  'private static final String WHATSAPP_BUSINESS = "com.whatsapp.w4b";',
  'startActivityForResult(call, intent, "photosActivityResult");',
  '@ActivityCallback',
  'callResult.put("separateTextRequired", true);',
  'callResult.put("fallback", false);'
];
requiredNativeMarkers.forEach((marker) => assert.ok(plugin.includes(marker), `Marker nativo mancante: ${marker}`));

[
  "Intent.ACTION_VIEW",
  "web.whatsapp.com",
  "https://wa.me/",
  "api.whatsapp.com",
  "Intent.createChooser",
  "Intent.EXTRA_TEXT",
  "Intent.EXTRA_SUBJECT"
].forEach((marker) => assert.ok(!plugin.includes(marker), `Fallback o chooser vietato nel plugin foto: ${marker}`));

assert.match(app, /function getDedicatedAndroidWhazzupPhotoPlugin\(\)/);
assert.match(app, /registerPlugin\?\.\("HeraWhazzupPhotos"\)/);
assert.match(orderFix, /async function sharePhotosThroughDedicatedPlugin\(plugin, orderedFiles\)/);
assert.match(orderFix, /await plugin\.begin\(\)/);
assert.match(orderFix, /await plugin\.addPhoto\(\{/);
assert.match(orderFix, /return await plugin\.share\(\{ sessionId \}\);/);
assert.match(orderFix, /async function shareWhazzupPhotosNativeAndroidInOrder\(orderedFiles, message\)/);
assert.match(orderFix, /window\.shareWhazzupPhotosNativeAndroid = shareWhazzupPhotosNativeAndroidInOrder;/);
assert.match(orderFix, /safeOpenWhatsAppMessage\(message\)[\s\S]*await waitBeforeWhazzupPhotos\(\)[\s\S]*sharePhotosThroughDedicatedPlugin/);
assert.match(orderFix, /safeOpenWhatsAppMessage\(message\)[\s\S]*await waitBeforeWhazzupPhotos\(\)[\s\S]*await plugins\.share\.share/);
assert.doesNotMatch(orderFix, /text:\s*message/);
assert.doesNotMatch(orderFix, /title:\s*"Impianto fatto"/);

const messageIndex = orderFix.indexOf("safeOpenWhatsAppMessage(message)");
const dedicatedShareIndex = orderFix.indexOf("await sharePhotosThroughDedicatedPlugin(dedicatedPlugin, orderedFiles)");
const genericShareIndex = orderFix.indexOf("await plugins.share.share");
assert.ok(messageIndex >= 0 && messageIndex < dedicatedShareIndex, "Il messaggio deve aprirsi prima del plugin foto dedicato");
assert.ok(messageIndex >= 0 && messageIndex < genericShareIndex, "Il messaggio deve aprirsi prima del fallback foto");
assert.match(orderFix, /PHOTO_SHARE_AFTER_MESSAGE_DELAY_MS\s*=\s*8000/);

assert.match(nativeRuntime, /function loadAndroidWhazzupPhotoOrderFix\(\)/);
assert.match(nativeRuntime, /script\.src = "android-whazzup-photo-order\.js\?v=20260814a"/);
assert.match(nativeRuntime, /const start = \(\) => \{\s*loadAndroidWhazzupPhotoOrderFix\(\);/);
assert.match(capacitorPrepare, /"android-whazzup-photo-order\.js"/);

console.log("Android Whazzup message-first photo-share checks passed.");
