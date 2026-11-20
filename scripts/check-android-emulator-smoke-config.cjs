"use strict";

// Purely offline regression test for the REAL Android smoke shell script.
// Fake ADB and fake APK live in a disposable directory; no emulator, Firebase,
// network, signing key or device is contacted.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const cp = require("node:child_process");

const root = path.resolve(__dirname, "..");
const workflow = fs.readFileSync(path.join(root, ".github/workflows/test-android-emulator.yml"), "utf8");
const smoke = path.join(root, "scripts/android-emulator-smoke.sh");
assert.match(workflow, /emulator-options: -no-window -gpu software -no-snapshot -noaudio -no-boot-anim/);
assert.doesNotMatch(workflow, /swiftshader_indirect/, "Deprecated graphics mode must be removed, including comments");
assert.match(workflow, /script: bash scripts\/android-emulator-smoke\.sh/);
const syntax = cp.spawnSync("bash", ["-n", smoke], {encoding:"utf8"});
assert.equal(syntax.status, 0, syntax.stderr);

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "varga-android-emulator-mock-"));
try {
  const fake = path.join(dir, "bin");
  fs.mkdirSync(fake);
  const apkDir = path.join(dir, "android/app/build/outputs/apk/debug");
  fs.mkdirSync(apkDir, { recursive: true });
  fs.writeFileSync(path.join(apkDir, "app-debug.apk"), "synthetic apk");
  const mockADB = [
    "#!/usr/bin/env bash",
    "set -euo pipefail",
    "if [[ \"$*\" == *\"getprop sys.boot_completed\"* ]]; then printf '1\\r\\n'; exit 0; fi",
    "if [[ \"$*\" == *\"uiautomator dump\"* ]]; then [[ \"$MOCK_ADB_MODE\" != \"nodump\" ]] || exit 1; exit 0; fi",
    "if [[ \"$1\" == \"pull\" ]]; then",
    "  if [[ \"$MOCK_ADB_MODE\" == \"badxml\" ]]; then",
    "    printf '<hierarchy package=\"com.other.app\"/>\\n' > \"$3\"",
    "  else",
    "    printf '<hierarchy package=\"it.vargacantieri.hera\"/>\\n' > \"$3\"",
    "  fi",
    "  exit 0",
    "fi",
    "if [[ \"$*\" == *\"screencap\"* ]]; then printf 'synthetic-png'; exit 0; fi",
    "if [[ \"$*\" == *\"pidof\"* ]]; then [[ \"$MOCK_ADB_MODE\" != \"dead\" ]]; exit; fi",
    "exit 0"
  ].join("\n") + "\n";
  fs.writeFileSync(path.join(fake, "adb"), mockADB, {mode:0o755});
  fs.writeFileSync(path.join(fake, "sleep"), "#!/usr/bin/env bash\nexit 0\n", {mode:0o755});
  const run = mode => cp.spawnSync("bash", [smoke], {
    cwd:dir, encoding:"utf8", timeout:15000,
    env:{...process.env, PATH:fake + path.delimiter + process.env.PATH, MOCK_ADB_MODE:mode}
  });
  const good = run("good");
  assert.equal(good.status, 0, good.stderr + "\n" + good.stdout);
  assert.match(good.stdout, /PASS: installed, launched and inspected/);
  for (const mode of ["nodump", "badxml", "dead"]) {
    const bad = run(mode);
    assert.notEqual(bad.status, 0,
      mode + " should fail closed rather than accept an invalid Android run");
    assert.doesNotMatch(bad.stdout, /PASS: installed, launched and inspected/);
  }
  console.log("PASS: Android emulator smoke runs with supported software GPU and fails closed on missing UI, wrong app or dead process.");
} finally {
  fs.rmSync(dir, {recursive:true, force:true});
}
