const { test, expect } = require("@playwright/test");

const APP_URL = process.env.E2E_APP_URL || "http://127.0.0.1:4173";

test.describe("Potature Abbattimenti - preparazione Raccolta e Ceppi", () => {
  test.use({ viewport: { width: 390, height: 844 } });

  test("apre il form fullscreen e salva le due attività senza duplicare documenti", async ({ page }) => {
    await page.goto(APP_URL, { waitUntil: "domcontentloaded" });
    await page.evaluate(() => {
      window.__potatureWrites = [];
      window.__potatureCommits = 0;
      const store = {
        collection(collectionName) {
          return {
            doc(commessaId) {
              return {
                collection(subcollectionName) {
                  return {
                    doc(documentId) {
                      return { collectionName, commessaId, subcollectionName, documentId };
                    }
                  };
                }
              };
            }
          };
        },
        batch() {
          return {
            set(ref, payload, options) { window.__potatureWrites.push({ ref, payload, options }); },
            async commit() { window.__potatureCommits += 1; }
          };
        }
      };
      window.HeraPotatureFollowup.open({
        source: {
          id: "tree-bologna-103923",
          idSap: "ALB-BOLOGNA-103923",
          denominazione: "ALBERO #103923 — Prunus cerasifera",
          comune: "Bologna",
          indirizzo: "Navile",
          gpsY: 44.5191,
          gpsX: 11.343,
          potatureAbbattimenti: true
        },
        existingItems: [],
        commessaId: "potature-abbattimenti",
        collectionName: "commesse",
        store,
        operatorUid: "user-1",
        operatorName: "Mario",
        timestampFactory: () => "SERVER_TIMESTAMP",
        onComplete: (result) => { window.__potatureResult = result; }
      });
    });

    const modal = page.locator("#potature-followup-modal");
    await expect(modal).toBeVisible();
    await expect(modal).toHaveCSS("position", "fixed");
    await expect(page.locator("#potature-followup-title")).toHaveText("Prepara la fine");
    await expect(modal).toContainText("ALBERO #103923");
    await expect(modal).toContainText("Mucchia");
    await expect(modal).toContainText("Ceppo");
    await expect(page.locator('input[name="raccolta"]')).toHaveCount(3);
    await expect(page.locator('input[name="ceppi"]')).toHaveCount(3);
    if (process.env.POTATURE_SCREENSHOT_PATH) {
      await page.screenshot({ path: process.env.POTATURE_SCREENSHOT_PATH, fullPage: true });
    }

    await page.locator('input[name="raccolta"][value="ragno"]').check();
    await page.locator('input[name="ceppi"][value="robotino"]').check();
    await page.locator(".potature-followup-save").click();
    await expect(modal).toContainText("Salvate: Raccolta e Ceppi");
    await expect(modal).toBeHidden({ timeout: 3000 });

    const saved = await page.evaluate(() => ({
      commits: window.__potatureCommits,
      writes: window.__potatureWrites,
      result: window.__potatureResult
    }));
    expect(saved.commits).toBe(1);
    expect(saved.writes).toHaveLength(2);
    expect(saved.writes.map((entry) => entry.ref.documentId)).toEqual([
      "tree-bologna-103923--raccolta",
      "tree-bologna-103923--ceppi"
    ]);
    expect(saved.writes[0].payload.potatureMetodoLabel).toBe("Con ragno");
    expect(saved.writes[1].payload.potatureMetodoLabel).toBe("Con robotino");
    expect(saved.writes.every((entry) => entry.options.merge === true)).toBe(true);
    expect(saved.result.tasks).toHaveLength(2);
  });

  test("consente di tornare senza scegliere Mucchia o Ceppo", async ({ page }) => {
    await page.goto(APP_URL, { waitUntil: "domcontentloaded" });
    await page.evaluate(() => {
      window.__potatureEmptyBatches = 0;
      window.HeraPotatureFollowup.open({
        source: { id: "tree-empty", denominazione: "ALBERO #2", potatureAbbattimenti: true },
        existingItems: [],
        commessaId: "potature-abbattimenti",
        store: {
          collection() { throw new Error("Non deve scrivere senza selezioni"); },
          batch() { window.__potatureEmptyBatches += 1; throw new Error("Non deve creare batch senza selezioni"); }
        },
        onComplete: (result) => { window.__potatureEmptyResult = result; }
      });
    });

    await page.locator(".potature-followup-save").click();
    await expect(page.locator("#potature-followup-modal")).toContainText("Nessuna attività creata");
    await expect(page.locator("#potature-followup-modal")).toBeHidden({ timeout: 3000 });
    const result = await page.evaluate(() => ({ batches: window.__potatureEmptyBatches, result: window.__potatureEmptyResult }));
    expect(result.batches).toBe(0);
    expect(result.result.tasks).toEqual([]);
  });
});

test.describe("Commesse speciali - TERMINATO separato", () => {
  test("espone TERMINATO e mostra lo stato FINITO senza cambiare FATTO", async ({ page }) => {
    await page.goto(APP_URL, { waitUntil: "domcontentloaded" });
    await page.waitForFunction(() => Boolean(window.HeraSpecialTerminato && window.getImpiantoPopupData));

    const result = await page.evaluate(() => {
      selectedCommessaId = "potature-abbattimenti";
      selectedCommessaName = "Potature Abbattimenti";
      const plant = {
        id: "tree-map-special",
        sourceIds: ["tree-map-special"],
        idSap: "ALB-SPECIAL-1",
        denominazione: "Albero speciale",
        comune: "Bologna",
        indirizzo: "Via Test",
        gpsY: 44.4949,
        gpsX: 11.3426,
        done: false
      };
      const programState = window.HeraSpecialTerminato.getDisplayState(plant);
      const finishedPlant = {
        ...plant,
        specialTerminato: true,
        specialTerminatoAt: new Date("2026-09-02T08:30:00+02:00"),
        specialTerminatoBy: "Operatore Test",
        specialTerminatoPending: false
      };
      const finishedDisplay = window.HeraSpecialTerminato.getDisplayState(finishedPlant);
      return {
        programState,
        finishedDisplay,
        finishedState: getImpiantoPopupData(finishedPlant, "Potatura"),
        coreDone: finishedPlant.done
      };
    });

    expect(result.programState.action).toBe("TERMINATO");
    expect(result.programState.state).toBe("In programma");
    expect(result.finishedDisplay.action).toBe("FINITO");
    expect(result.finishedDisplay.terminated).toBe(true);
    expect(result.finishedState.stato).toBe("Finito");
    expect(result.finishedState.operatoreSquadra).toContain("Operatore Test");
    expect(result.coreDone).toBe(false);
  });
});
