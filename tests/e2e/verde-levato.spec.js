const { test, expect } = require("@playwright/test");

const APP_URL = process.env.E2E_APP_URL || "http://127.0.0.1:4173";

test.describe("Verde Levato manuale", () => {
  test.use({ viewport: { width: 390, height: 844 } });

  test("gestisce commesse, associa cantieri, esporta Excel e configura l’amministratore", async ({ page }) => {
    await page.goto(APP_URL, { waitUntil: "domcontentloaded" });
    await page.waitForFunction(() => Boolean(window.HeraVerdeLevato));
    await expect(page.locator("#open-verde-levato-btn")).toContainText("Verde Levato");

    await page.evaluate(() => {
      window.__verdeLevatoRecords = [];
      window.__verdeLevatoWrites = [];
      window.__verdeLevatoCommesse = [];
      window.__verdeLevatoCommessaWrites = [];
      window.__verdeLevatoConfigWrites = [];
      window.__verdeLevatoExcel = null;
      auth = { currentUser: { uid: "global-admin", email: "admin@example.test", displayName: "Admin Test" } };
      canManageData = () => true;
      setAuthenticationGateState = () => {};
      document.body.classList.remove("auth-pending", "auth-required", "auth-banned");
      document.getElementById("auth-gate")?.classList.add("hidden");

      db = {
        collection(name) {
          if (name === "verdeLevatoConfig") {
            return {
              doc() {
                return {
                  async get() { return { exists: true, data: () => ({ adminEmails: [] }) }; },
                  async set(payload, options) { window.__verdeLevatoConfigWrites.push({ payload, options }); }
                };
              }
            };
          }
          if (name === "verdeLevatoRecords") {
            return {
              async get() {
                return {
                  docs: window.__verdeLevatoRecords.map((record) => ({ id: record.id, data: () => ({ ...record.payload }) }))
                };
              },
              doc(id = `record-${window.__verdeLevatoRecords.length + 1}`) {
                return {
                  async set(payload, options) {
                    window.__verdeLevatoWrites.push({ id, payload, options });
                    const index = window.__verdeLevatoRecords.findIndex((record) => record.id === id);
                    const value = { id, payload: { ...(index >= 0 ? window.__verdeLevatoRecords[index].payload : {}), ...payload } };
                    if (index >= 0) window.__verdeLevatoRecords[index] = value;
                    else window.__verdeLevatoRecords.push(value);
                  }
                };
              }
            };
          }
          if (name === "verdeLevatoCommesse") {
            return {
              async get() {
                return {
                  docs: window.__verdeLevatoCommesse.map((commessa) => ({ id: commessa.id, data: () => ({ ...commessa.payload }) }))
                };
              },
              doc(id = `commessa-${window.__verdeLevatoCommesse.length + 1}`) {
                return {
                  id,
                  async set(payload, options) {
                    window.__verdeLevatoCommessaWrites.push({ id, payload, options });
                    window.__verdeLevatoCommesse.push({ id, payload: { ...payload } });
                  }
                };
              }
            };
          }
          throw new Error(`Collection inattesa: ${name}`);
        }
      };

      window.XLSX = {
        utils: {
          json_to_sheet(rows, options) { return { rows, headers: options.header, "!ref": "A1:B2" }; },
          book_new() { return { sheets: [] }; },
          book_append_sheet(workbook, sheet, name) { workbook.sheets.push({ name, sheet }); }
        },
        writeFile(workbook, filename) { window.__verdeLevatoExcel = { workbook, filename }; }
      };

      Object.defineProperty(navigator, "geolocation", {
        configurable: true,
        value: {
          getCurrentPosition(success) {
            success({ coords: { latitude: 44.5101234, longitude: 11.3556789, accuracy: 4.2 }, timestamp: Date.parse("2026-09-02T08:00:00+02:00") });
          }
        }
      });
      const originalFetch = window.fetch.bind(window);
      window.fetch = async (url, options) => {
        if (String(url).includes("nominatim.openstreetmap.org/reverse")) {
          return {
            ok: true,
            async json() {
              return {
                display_name: "Via Levato 12, Bologna, Emilia-Romagna, Italia",
                address: {
                  city: "Bologna",
                  suburb: "Navile",
                  road: "Via Levato",
                  house_number: "12",
                  postcode: "40100",
                  county: "Bologna",
                  state: "Emilia-Romagna",
                  country: "Italia"
                }
              };
            }
          };
        }
        return originalFetch(url, options);
      };
      window.HeraVerdeLevato.open();
    });

    await expect(page.locator("#verde-levato-page")).toBeVisible();
    await expect(page.locator("[data-verde-levato-category]")).toHaveCount(3);
    await expect(page.locator("#verde-levato-new-btn")).toBeVisible();
    await page.locator("#verde-levato-new-btn").click();
    await expect(page.locator('select[name="commessaId"]')).toBeVisible();
    await page.locator("#verde-levato-show-new-commessa").click();
    await page.locator("#verde-levato-new-commessa-name").fill("Manutenzione Bologna Nord");
    await page.locator("#verde-levato-new-commessa-code").fill("VL-2026-01");
    await page.locator("#verde-levato-save-commessa").click();
    await expect(page.locator('select[name="commessaId"]')).toHaveValue("commessa-1");
    await page.locator('input[name="denominazione"]').fill("Parco Levato Nord");
    await page.locator("#verde-levato-use-location").click();
    await page.locator("#verde-levato-form").evaluate((form) => form.requestSubmit());
    await expect(page.locator("#verde-levato-record-modal")).toBeHidden();

    const cantiere = await page.evaluate(() => window.__verdeLevatoWrites[0]);
    expect(cantiere.payload.tipoRecord).toBe("cantiere");
    expect(cantiere.payload.commessaId).toBe("commessa-1");
    expect(cantiere.payload.commessaNome).toBe("Manutenzione Bologna Nord");
    expect(cantiere.payload.commessaCodice).toBe("VL-2026-01");

    await page.locator("#verde-levato-category-new-btn").click();
    await page.locator('select[name="tipoRecord"]').selectOption("albero");
    await page.locator("#verde-levato-use-location").click();

    await expect(page.locator('input[name="gpsY"]')).toHaveValue("44.5101234");
    await expect(page.locator('input[name="gpsX"]')).toHaveValue("11.3556789");
    await expect(page.locator('input[name="comune"]')).toHaveValue("Bologna");
    await expect(page.locator('input[name="via"]')).toHaveValue("Via Levato");
    await expect(page.locator('input[name="civico"]')).toHaveValue("12");
    await expect(page.locator('input[name="regione"]')).toHaveValue("Emilia-Romagna");

    await page.locator('input[name="denominazione"]').fill("Tiglio ingresso nord");
    await page.locator('input[name="numeroAlbero"]').fill("A-001");
    await page.locator('input[name="specieAlbero"]').fill("Tilia cordata");
    await page.locator("#verde-levato-form").evaluate((form) => form.requestSubmit());
    await expect(page.locator("#verde-levato-record-modal")).toBeHidden();

    const saved = await page.evaluate(() => window.__verdeLevatoWrites[1]);
    expect(saved.options).toEqual({ merge: true });
    expect(saved.payload.tipoRecord).toBe("albero");
    expect(saved.payload.denominazione).toBe("Tiglio ingresso nord");
    expect(saved.payload.specieAlbero).toBe("Tilia cordata");
    expect(saved.payload.gpsY).toBe(44.5101234);
    expect(saved.payload.gpsX).toBe(11.3556789);
    expect(saved.payload.comune).toBe("Bologna");
    expect(saved.payload.source).toBe("MANUALE_VERDE_LEVATO");

    await page.locator("#verde-levato-categories-btn").click();
    await page.locator("#verde-levato-export-btn").click();
    const excel = await page.evaluate(() => window.__verdeLevatoExcel);
    expect(excel.filename).toMatch(/^verde_levato_dati_completi_.*\.xlsx$/);
    expect(excel.workbook.sheets.map((sheet) => sheet.name)).toEqual(["Dati completi", "Commesse"]);
    expect(excel.workbook.sheets[0].sheet.rows).toHaveLength(2);
    const exportedCantiere = excel.workbook.sheets[0].sheet.rows.find((row) => row["Tipo elemento"] === "Cantieri");
    expect(exportedCantiere["Commessa"]).toBe("Manutenzione Bologna Nord");
    expect(excel.workbook.sheets[1].sheet.rows[0]["Cantieri associati"]).toBe(1);

    await page.locator("#verde-levato-admin-btn").click();
    await page.locator('#verde-levato-admin-form input[name="email"]').fill("Levato.Admin@Example.Test");
    await page.locator("#verde-levato-admin-form").evaluate((form) => form.requestSubmit());
    const config = await page.evaluate(() => window.__verdeLevatoConfigWrites[0]);
    expect(config.options).toEqual({ merge: true });
    expect(config.payload.adminEmails).toEqual(["levato.admin@example.test"]);
  });
});
