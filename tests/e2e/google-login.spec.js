const { test, expect } = require('@playwright/test');

const baseUrl = process.env.PLAYWRIGHT_BASE_URL || 'http://127.0.0.1:8080';
const appUrl = (path) => new URL(path, `${baseUrl.replace(/\/$/, '')}/`).toString();

async function openHarness(page, scenario, platform = 'web') {
  await page.goto(appUrl(`tests/e2e/google-login-harness.html?case=${scenario}&platform=${platform}`), {
    waitUntil: 'domcontentloaded'
  });
  await page.waitForFunction(() => Boolean(window.__googleE2E));
}

test.describe('Hera App - login Google di riserva', () => {
  test('popup riuscito usa persistenza LOCAL e salva il provider Google', async ({ page }) => {
    await openHarness(page, 'success');
    await page.locator('#auth-gate-login-btn').click();

    await expect.poll(() => page.evaluate(() => window.__googleE2E.popupCalls)).toBe(1);
    await expect.poll(() => page.evaluate(() => window.__googleE2E.persistence.includes('LOCAL'))).toBe(true);
    await expect.poll(() => page.evaluate(() => localStorage.getItem('heraLastAuthProvider'))).toBe('google');
    await expect(page.locator('#auth-email-feedback')).toContainText('Accesso Google completato');
    await expect.poll(() => page.evaluate(() => window.__googleE2E.providerParams?.prompt)).toBe('select_account');
    await expect.poll(() => page.evaluate(() => window.__googleE2E.legacyClickCalls)).toBe(0);
  });

  test('web: popup bloccato passa al redirect alternativo', async ({ page }) => {
    await openHarness(page, 'blocked', 'web');
    await page.locator('#auth-gate-login-btn').click();

    await expect.poll(() => page.evaluate(() => window.__googleE2E.popupCalls)).toBe(1);
    await expect.poll(() => page.evaluate(() => window.__googleE2E.redirectCalls)).toBe(1);
    await expect.poll(() => page.evaluate(() => window.__googleE2E.legacyClickCalls)).toBe(0);
  });

  test('Opera continua il login Google quando la persistenza non è supportata', async ({ page }) => {
    await openHarness(page, 'persistence-unsupported');
    await page.locator('#auth-gate-login-btn').click();

    await expect.poll(() => page.evaluate(() => window.__googleE2E.popupCalls)).toBe(1);
    await expect(page.locator('#auth-email-feedback')).toContainText('Accesso Google completato');
  });

  test('Android WebView: popup bloccato non usa redirect', async ({ page }) => {
    await openHarness(page, 'blocked', 'android');
    await page.locator('#auth-gate-login-btn').click();

    await expect.poll(() => page.evaluate(() => window.__googleE2E.popupCalls)).toBe(1);
    await expect.poll(() => page.evaluate(() => window.__googleE2E.redirectCalls)).toBe(0);
    await expect(page.locator('#auth-email-feedback')).toContainText('Android');
    await expect.poll(() => page.evaluate(() => window.__googleE2E.legacyClickCalls)).toBe(0);
  });

  test('offline non tenta un nuovo login Google', async ({ page }) => {
    await openHarness(page, 'offline');
    await page.locator('#auth-gate-login-btn').click();

    await expect.poll(() => page.evaluate(() => window.__googleE2E.popupCalls)).toBe(0);
    await expect.poll(() => page.evaluate(() => window.__googleE2E.redirectCalls)).toBe(0);
    await expect(page.locator('#auth-email-feedback')).toContainText('Senza Internet');
  });

  test('risultato redirect viene consumato e salva Google come provider', async ({ page }) => {
    await openHarness(page, 'redirect-result');

    await expect.poll(() => page.evaluate(() => window.__googleE2E.redirectResultCalls)).toBe(1);
    await expect.poll(() => page.evaluate(() => localStorage.getItem('heraLastAuthProvider'))).toBe('google');
    await expect(page.locator('#auth-email-feedback')).toContainText('Accesso Google completato');
  });
});
