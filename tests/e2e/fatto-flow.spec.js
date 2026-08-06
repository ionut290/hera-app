const { test, expect } = require('@playwright/test');

const BASE_URL = process.env.PLAYWRIGHT_BASE_URL || 'http://127.0.0.1:8080';
const appUrl = (path) => new URL(path, BASE_URL).toString();

test.describe('Hera App - protezioni E2E', () => {
  test('l’app reale espone gli elementi principali senza errori HTML bloccanti', async ({ page }) => {
    const pageErrors = [];
    page.on('pageerror', (error) => pageErrors.push(error.message));

    await page.goto(appUrl('/index.html'), { waitUntil: 'domcontentloaded' });

    await expect(page).toHaveTitle(/Hera App/i);
    await expect(page.locator('#home-page')).toHaveCount(1);
    await expect(page.locator('#auth-gate')).toHaveCount(1);
    await expect(page.locator('#open-panel-commesse')).toHaveCount(1);

    const blockingErrors = pageErrors.filter((message) =>
      /SyntaxError|ReferenceError:.*(?:Unexpected|is not defined)/i.test(message)
    );
    expect(blockingErrors, `Errori JavaScript bloccanti: ${blockingErrors.join(' | ')}`).toEqual([]);
  });

  test('FATTO diventa giallo, mostra la data, salva una sola operazione e apre WhatsApp', async ({ page }) => {
    await page.goto(appUrl('/tests/e2e/fatto-harness.html'), { waitUntil: 'domcontentloaded' });

    await page.waitForFunction(() =>
      window.handleImpiantoWhatsAppClick &&
      window.handleImpiantoWhatsAppClick.__heraQueueWrapped === true
    );

    const button = page.locator('#fatto-btn');
    await button.focus();
    await button.click();

    await expect(button).toHaveAttribute('data-fatto-immediate', 'true');
    await expect(button).toHaveAttribute('aria-disabled', 'true');
    await expect(button).toBeDisabled();
    await expect(button).toHaveText('FATTO');

    const background = await button.evaluate((element) =>
      getComputedStyle(element).backgroundColor
    );
    expect(background).toBe('rgb(244, 197, 66)');

    const dateNode = page.locator('[data-fatto-immediate-date="true"]');
    await expect(dateNode).toHaveCount(1);
    await expect(dateNode).toContainText(/\d{2}\/\d{2}\/\d{4}/);

    await expect.poll(async () => page.evaluate(() => window.__e2e.originalCalls)).toBe(1);
    await expect.poll(async () => page.evaluate(() => window.__e2e.whatsappCalls)).toBe(1);

    await expect.poll(async () => page.evaluate(async () => {
      const items = await window.HeraFattoSync.list();
      return items.length;
    })).toBe(0);
  });
});
