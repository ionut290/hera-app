const assert = require("node:assert/strict");
const fs = require("node:fs");
const read = (file) => fs.readFileSync(file, "utf8");

const pngSize = (file) => {
  const data = fs.readFileSync(file);
  assert.equal(data.toString("ascii", 1, 4), "PNG", `${file} non è un PNG`);
  return [data.readUInt32BE(16), data.readUInt32BE(20)];
};

const index = read("index.html");
const styles = read("style.css");
const serviceWorker = read("sw.js");
const manifest = JSON.parse(read("manifest.webmanifest"));
const workflow = read(".github/workflows/build-android-aab.yml");
const androidManifest = read("android/app/src/main/AndroidManifest.xml");
const mainActivity = read("android/app/src/main/java/it/vargacantieri/hera/MainActivity.java");
const vargaAppTheme = read("android-resources/res/values/varga_app_theme.xml");

assert.match(index, /class="startup-loading-logo"/);
assert.match(index, /icons\/varga-cantieri-512\.png/);
const styleAsset = index.match(/style\.css\?v=[^"']+/)?.[0];
assert.ok(styleAsset, "index.html deve caricare style.css con una versione cache-busting");
assert.match(styles, /\.startup-loading\s*\{[\s\S]*position:\s*fixed;[\s\S]*inset:\s*0;/);
assert.match(styles, /\.startup-loading-logo\s*\{[\s\S]*width:\s*min\(90vw,\s*72dvh,\s*820px\);/);
assert.match(serviceWorker, /[a-z0-9-]+-shell-v\d+/i);
assert.ok(serviceWorker.includes(`./${styleAsset}`), "Il Service Worker deve precaricare la stessa versione di style.css usata da index.html");
assert.equal(manifest.background_color, "#111214");
assert.equal(manifest.theme_color, "#111214");
assert.match(workflow, /find android\/app\/src\/main\/res -type f -name "splash\.png" -delete/);
assert.match(workflow, /cp -R android-resources\/res\/\. android\/app\/src\/main\/res\//);
assert.match(workflow, /cp icons\/varga-cantieri-512\.png android\/app\/src\/main\/res\/drawable-xhdpi\/splash_logo\.png/);
assert.match(androidManifest, /android:theme="@style\/VargaAppTheme\.Launch"/);
assert.match(mainActivity, /supportRequestWindowFeature\(Window\.FEATURE_NO_TITLE\)/);
assert.match(mainActivity, /getSupportActionBar\(\)\.hide\(\)/);
assert.match(vargaAppTheme, /name="VargaAppTheme" parent="AppTheme\.NoActionBar"/);
assert.match(vargaAppTheme, /name="windowActionBar">false</);
assert.match(vargaAppTheme, /name="windowNoTitle">true</);

const splashDrawable = read("android-resources/res/drawable/splash.xml");
assert.match(splashDrawable, /#111214/);
assert.match(splashDrawable, /@drawable\/splash_logo/);
assert.match(splashDrawable, /android:gravity="center"/);
assert.deepEqual(
  pngSize("icons/varga-cantieri-512.png"),
  [512, 512],
  "Dimensioni errate per il logo sorgente dello splash Android"
);

const android12Style = read("android-resources/res/values-v31/styles.xml");
assert.match(android12Style, /windowSplashScreenBackground/);
assert.match(android12Style, /windowSplashScreenAnimatedIcon/);
assert.match(android12Style, /@mipmap\/ic_launcher_foreground/);

console.log("Startup splash checks passed for web and Android.");
