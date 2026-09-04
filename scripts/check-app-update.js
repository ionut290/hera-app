const assert = require("node:assert/strict");
const fs = require("node:fs");

const html = fs.readFileSync("index.html", "utf8");
const css = fs.readFileSync("style.css", "utf8");
const app = fs.readFileSync("app.js", "utf8");
const androidInstall = fs.readFileSync("android-play-store-install.js", "utf8");
const serviceWorker = fs.readFileSync("sw.js", "utf8");
const nativeMainActivity = fs.readFileSync("android/app/src/main/java/it/vargacantieri/hera/MainActivity.java", "utf8");
const nativeUpdatePlugin = fs.readFileSync("android/app/src/main/java/it/vargacantieri/hera/update/HeraAppUpdatePlugin.java", "utf8");
const nativeGradle = fs.readFileSync("android/app/capacitor.build.gradle", "utf8");
const capacitor = JSON.parse(fs.readFileSync("capacitor.config.json", "utf8"));

assert.match(html, /id="update-app-btn"[^>]*>[\s\S]*?Aggiorna app[\s\S]*?<\/button>/);
assert.match(css, /\.update-app-btn\s*\{[\s\S]*?width:\s*30px;[\s\S]*?height:\s*30px;/);
assert.match(css, /\.logo-head\s*\{[\s\S]*?grid-template-columns:\s*minmax\(0, 1fr\) auto minmax\(0, 1fr\)/);
assert.match(app, /updateAppBtn\?\.addEventListener\("click", openApplicationUpdate\)/);
assert.match(app, new RegExp(`play\\.google\\.com/store/apps/details\\?id=${capacitor.appId.replaceAll(".", "\\.")}`));
assert.match(app, /navigator\.serviceWorker\?\.getRegistration/);
assert.match(app, /Capacitor\?\.getPlatform\?\.\(\) === "android"/);
assert.match(html, /android-play-store-install\.js\?v=20260828a/);
assert.match(serviceWorker, /android-play-store-install\.js\?v=20260828a/);
assert.match(androidInstall, new RegExp(`play\\.google\\.com/store/apps/details\\?id=${capacitor.appId.replaceAll(".", "\\.")}`));
assert.match(androidInstall, /function isAndroidDevice\(\)/);
assert.match(androidInstall, /Vuoi aprire Google Play e installare l'app Android\?/);
assert.match(androidInstall, /OK = SÌ, INSTALLA/);
assert.match(androidInstall, /addEventListener\("click", handleAndroidInstallClick, true\)/);
assert.match(androidInstall, /window\.open\(PLAY_STORE_URL, "_blank", "noopener,noreferrer"\)/);
assert.doesNotMatch(app, /window\.location\.assign\(updateUrl/);
assert.match(html, /PWA_EMERGENCY_CACHE_RESET_VERSION\s*=\s*"20260902-verde-levato2"/);
assert.match(html, /ANDROID_EMERGENCY_CACHE_RESET_VERSION\s*=\s*"20260812-android1"/);
assert.match(html, /const cacheResetVersion\s*=\s*isNativeAndroid/);
assert.doesNotMatch(html, /if\s*\(isNativeAndroid\s*\|\|/);
assert.match(html, /name\.startsWith\("hera-app-shell-"\)[\s\S]*name\.startsWith\("varga-cantieri-shell-"\)/);
assert.match(html, /navigator\.serviceWorker\.register\("\.\/sw\.js(?:\?v=[^"]+)?", \{ updateViaCache: "none" \}\)/);
assert.doesNotMatch(html, /localStorage\.clear\(\)/);
const serviceWorkerResetMatch = serviceWorker.match(/CACHE_RESET_VERSION\s*=\s*"([^"]+)"/);
assert.ok(serviceWorkerResetMatch, "Versione reset cache Service Worker non trovata");
const serviceWorkerRegistrationMatch = html.match(/navigator\.serviceWorker\.register\("\.\/sw\.js\?v=([^"]+)"/);
assert.ok(serviceWorkerRegistrationMatch, "Versione registrazione Service Worker non trovata");
assert.equal(
  serviceWorkerRegistrationMatch[1],
  serviceWorkerResetMatch[1],
  "La registrazione del Service Worker deve usare la stessa versione del reset cache"
);
assert.match(serviceWorker, /self\.clients\.matchAll\(\{ type: "window", includeUncontrolled: true \}\)/);
assert.match(serviceWorker, /type:\s*"HERA_SW_UPDATE_READY"/);
assert.doesNotMatch(
  serviceWorker,
  /client\.navigate\s*\(/,
  "L'attivazione del Service Worker non deve ricaricare forzatamente una schermata in uso"
);
const cacheVersionMatch = serviceWorker.match(/CACHE_NAME\s*=\s*"varga-cantieri-shell-v(\d+)"/);
assert.ok(cacheVersionMatch, "Versione cache PWA non trovata");
assert.ok(Number(cacheVersionMatch[1]) >= 115, "La cache PWA deve essere almeno v115");
assert.match(nativeMainActivity, /clearWebViewCacheAfterAppUpdate\(\)/);
assert.match(nativeMainActivity, /registerPlugin\(HeraAppUpdatePlugin\.class\)/);
assert.match(nativeMainActivity, /getLongVersionCode\(\)/);
assert.match(nativeMainActivity, /webView\.clearCache\(true\)/);
assert.match(nativeMainActivity, /preferences\.getLong\(CACHE_VERSION_CODE_KEY, -1L\)/);
assert.doesNotMatch(nativeMainActivity, /deleteAllData|removeAllCookies|clearHistory|clearFormData/);
assert.match(nativeGradle, /com\.google\.android\.play:app-update:2\.1\.0/);
assert.match(nativeUpdatePlugin, /@CapacitorPlugin\(name = "HeraAppUpdate"\)/);
assert.match(nativeUpdatePlugin, /UpdateAvailability\.UPDATE_AVAILABLE/);
assert.match(nativeUpdatePlugin, /AppUpdateType\.IMMEDIATE/);
assert.match(nativeUpdatePlugin, /startUpdateFlowForResult/);
assert.match(html, new RegExp(`sw\\.js\\?v=${serviceWorkerResetMatch[1].replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`));
const updateFeature = fs.readFileSync("update-app-feature.js", "utf8");
assert.match(updateFeature, /id = "hera-update-notice"/);
assert.match(updateFeature, /Nuova versione Android disponibile/);
assert.match(updateFeature, /HERA_SW_UPDATE_READY/);
assert.match(updateFeature, /registration\.waiting/);
assert.match(updateFeature, /HeraAppUpdate/);
assert.doesNotMatch(updateFeature, /firebase\.firestore|\.collection\(/);

console.log("App update button checks passed.");
