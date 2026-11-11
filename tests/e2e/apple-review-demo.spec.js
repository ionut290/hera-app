const { test, expect } = require("@playwright/test");

const base = process.env.E2E_APP_URL || "http://127.0.0.1:4173";

test.use({ viewport: { width: 390, height: 844 } });

test("Apple Review demo opens, navigates and simulates data with no external requests", async ({ page }) => {
  const externalRequests = [];
  page.on("request", (req) => {
    if (!req.url().startsWith(base + "/")) externalRequests.push(req.url());
  });
  await page.goto(base + "/apple-review-demo.html");
  await expect(page.getByText("MODALITÀ DEMO · DATI INVENTATI")).toBeVisible();
  await expect(page.getByRole("heading", { name: /Ciao, Revisore Apple/ })).toBeVisible();
  await expect(page.locator("#metric-pending")).toHaveText("3");

  await page.getByRole("button", { name: "Apri le commesse" }).click();
  await expect(page.locator("#commesse-list article")).toHaveCount(2);
  await page.getByRole("button", { name: "Visualizza impianti" }).first().click();
  await expect(page.getByRole("heading", { name: /Impianti · Verde Dimostrativo/ })).toBeVisible();
  await expect(page.locator("#plants-list article")).toHaveCount(2);
  await page.getByRole("button", { name: "Simula intervento" }).first().click();
  await expect(page.locator("#metric-pending")).toHaveText("2");
  await expect(page.locator("#plants-list .state-pill.done")).toHaveCount(1);

  await page.getByRole("button", { name: "Ore e calendario" }).click();
  await page.locator("#hours-date").fill("2026-11-11");
  await page.locator("#hours-start").fill("07:00");
  await page.locator("#hours-end").fill("15:00");
  await page.getByRole("button", { name: "Aggiungi turno demo" }).click();
  await expect(page.locator("#hours-total")).toHaveText("7,5 h");

  await page.getByRole("button", { name: "Documenti" }).click();
  await page.getByRole("button", { name: "Visualizza esempio" }).first().click();
  await expect(page.locator("#doc-preview")).toBeVisible();
  await page.getByRole("button", { name: "Panoramica" }).click();
  await page.getByRole("button", { name: "Ripristina demo" }).click();
  await expect(page.locator("#metric-pending")).toHaveText("3");

  expect(externalRequests).toEqual([]);
  await page.reload();
  await expect(page.locator("#metric-pending")).toHaveText("3");
});

test("Apple Review demo offers no real login, redirects or network resources", async ({ page }) => {
  await page.goto(base + "/apple-review-demo.html");
  await expect(page.locator('a[href="./index.html"]')).toHaveCount(2);
  expect(await page.locator("iframe").count()).toBe(0);
  expect(await page.locator("form").count()).toBe(1);
  expect(await page.locator('script[src*="firebase"],script[src*="google"],script[src^="http"]').count()).toBe(0);
  await page.getByRole("button", { name: "Impianti", exact: true }).click();
  await page.locator("#plant-query").fill("NON ESISTE");
  await expect(page.locator("#plants-empty")).toBeVisible();
  await page.locator("#plant-query").fill("DEMO-101");
  await expect(page.locator("#plants-list article")).toHaveCount(1);
});
