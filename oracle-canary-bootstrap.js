(function startOracleOriginalUiCanary(global) {
  "use strict";

  const MAX_WAIT_MS = 30_000;
  const POLL_MS = 500;
  const STORAGE_KEY = "hera_oracle_canary_last_result_v1";

  function waitForOriginalCommesse() {
    const startedAt = Date.now();
    return new Promise((resolve, reject) => {
      const check = () => {
        if (typeof commesseById !== "undefined" && commesseById instanceof Map && commesseById.size > 0) {
          resolve(new Map(commesseById));
          return;
        }
        if (Date.now() - startedAt >= MAX_WAIT_MS) {
          reject(new Error("original_commesse_timeout"));
          return;
        }
        global.setTimeout(check, POLL_MS);
      };
      check();
    });
  }

  async function runComparison(user) {
    if (!user || !global.HeraOracleReadAdapter?.enabled) return;
    const [originalCommesse, oracleCatalog] = await Promise.all([
      waitForOriginalCommesse(),
      global.HeraOracleReadAdapter.loadCatalog()
    ]);
    const firebaseIds = new Set(originalCommesse.keys());
    const oracleIds = new Set(oracleCatalog.commesse.map((item) => item.id));
    const onlyFirebase = [...firebaseIds].filter((id) => !oracleIds.has(id));
    const onlyOracle = [...oracleIds].filter((id) => !firebaseIds.has(id));
    const result = Object.freeze({
      checkedAt: new Date().toISOString(),
      email: String(user.email || "").toLowerCase(),
      firebaseCommesse: firebaseIds.size,
      oracleCommesse: oracleIds.size,
      onlyFirebase,
      onlyOracle,
      matched: onlyFirebase.length === 0 && onlyOracle.length === 0
    });
    global.localStorage.setItem(STORAGE_KEY, JSON.stringify(result));
    console.info("[ORACLE CANARY READ-ONLY]", result);
  }

  global.addEventListener("load", () => {
    const auth = global.firebase?.auth?.();
    if (!auth) return;
    auth.onAuthStateChanged((user) => {
      if (!user) return;
      runComparison(user).catch((error) => {
        const failure = { checkedAt: new Date().toISOString(), error: String(error?.message || error) };
        global.localStorage.setItem(STORAGE_KEY, JSON.stringify(failure));
        console.warn("[ORACLE CANARY READ-ONLY] confronto non completato", failure);
      });
    });
  }, { once: true });
})(window);
