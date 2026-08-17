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
assert.match(orderFix, /files: fileUris,[\s\S]*safeOpenWhatsAppMessage\(message\)/);
assert.doesNotMatch(orderFix, /text:\s*message/);
assert.doesNotMatch(orderFix, /title:\s*"Impianto fatto"/);
assert.doesNotMatch(orderFix, /PHOTO_SHARE_AFTER_MESSAGE_DELAY_MS/);
assert.doesNotMatch(orderFix, /waitBeforeWhazzupPhotos/);

const dedicatedShareIndex = orderFix.indexOf("await sharePhotosThroughDedicatedPlugin(dedicatedPlugin, orderedFiles)");
const dedicatedMessageIndex = orderFix.indexOf("safeOpenWhatsAppMessage(message)", dedicatedShareIndex);
assert.ok(dedicatedShareIndex >= 0 && dedicatedShareIndex < dedicatedMessageIndex, "Le foto devono essere condivise prima del messaggio nel plugin dedicato");
const genericShareIndex = orderFix.indexOf("await plugins.share.share");
const genericMessageIndex = orderFix.indexOf("safeOpenWhatsAppMessage(message)", genericShareIndex);
assert.ok(genericShareIndex >= 0 && genericShareIndex < genericMessageIndex, "Le foto devono essere condivise prima del messaggio nel fallback nativo");

assert.match(nativeRuntime, /function loadAndroidWhazzupPhotoOrderFix\(\)/);
assert.match(nativeRuntime, /script\.src = "android-whazzup-photo-order\.js\?v=20260814a"/);
assert.match(nativeRuntime, /const start = \(\) => \{\s*loadAndroidWhazzupPhotoOrderFix\(\);/);
assert.match(capacitorPrepare, /"android-whazzup-photo-order\.js"/);

console.log("Android Whazzup photo-first share checks passed.");
