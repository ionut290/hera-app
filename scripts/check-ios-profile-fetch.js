"use strict";

const assert = require("node:assert/strict");
const crypto = require("node:crypto");
const { fetchActiveProfile, EXPECTED_NAME, EXPECTED_BUNDLE } = require("./fetch-ios-appstore-profile");

const keys = crypto.generateKeyPairSync("ec", { namedCurve: "prime256v1" });
const privateKey = keys.privateKey.export({ type: "pkcs8", format: "pem" });
const env = {
  IOS_APPSTORE_KEY_ID: "TEST_KEY_1",
  IOS_APPSTORE_ISSUER_ID: "11111111-2222-3333-4444-555555555555",
  IOS_APPSTORE_PRIVATE_KEY_BASE64: Buffer.from(privateKey).toString("base64")
};
const profile = Buffer.alloc(180, 0x7a);
const responseData = (identifier = EXPECTED_BUNDLE, state = "ACTIVE") => ({
  data: [{
    type: "profiles", id: "profile-1",
    attributes: {
      name: EXPECTED_NAME,
      profileType: "IOS_APP_STORE",
      profileState: state,
      profileContent: profile.toString("base64")
    },
    relationships: { bundleId: { data: { type: "bundleIds", id: "bundle-1" } } }
  }],
  included: [{ type: "bundleIds", id: "bundle-1", attributes: { identifier } }]
});
function response(body, status = 200) {
  return { ok: status === 200, status, json: async () => body };
}
(async () => {
  let called = 0;
  const bytes = await fetchActiveProfile(env, async (uri, opts) => {
    called++;
    const url = new URL(uri);
    assert.equal(url.origin, "https://api.appstoreconnect.apple.com");
    assert.equal(url.pathname, "/v1/profiles");
    assert.equal(url.searchParams.get("filter[name]"), EXPECTED_NAME);
    assert.equal(url.searchParams.get("filter[profileType]"), "IOS_APP_STORE");
    assert.equal(url.searchParams.get("filter[profileState]"), "ACTIVE");
    assert.equal(url.searchParams.get("include"), "bundleId");
    const jwt = opts.headers.Authorization.slice("Bearer ".length);
    const pieces = jwt.split(".");
    assert.equal(pieces.length, 3);
    const signedPart = Buffer.from(pieces[0] + "." + pieces[1]);
    const verified = crypto.verify("sha256", signedPart, {
      key: keys.publicKey, dsaEncoding: "ieee-p1363"
    }, Buffer.from(pieces[2], "base64url"));
    assert.equal(verified, true);
    return response(responseData());
  });
  assert.equal(called, 1);
  assert.deepEqual(bytes, profile);
  await assert.rejects(
    fetchActiveProfile(env, async () => response(responseData("org.radioromaniidinitalia.radio"))),
    /No signing asset was written/
  );
  await assert.rejects(
    fetchActiveProfile(env, async () => response(responseData(EXPECTED_BUNDLE, "INVALID"))),
    /No signing asset was written/
  );
  await assert.rejects(
    fetchActiveProfile(env, async () => response(null, 403)),
    /HTTP 403/
  );
  await assert.rejects(
    fetchActiveProfile({ ...env, IOS_APPSTORE_PRIVATE_KEY_BASE64: "" }, async () => response(responseData())),
    /Missing existing Apple API credential/
  );
  console.log("PASS: profile auto-fetch uses scoped JWT, checks app ID and rejects invalid or unauthorized profile.");
})().catch((error) => { console.error(error); process.exitCode = 1; });
