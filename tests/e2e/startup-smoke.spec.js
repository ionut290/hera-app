const { test, expect } = require('@playwright/test');

const APP_URL = process.env.E2E_APP_URL || 'http://127.0.0.1:4173';

function watchCriticalLocal404s(page) {
  const local404s = [];
  page.on('response', (response) => {
    const url = response.url();
    if (!url.startsWith(APP_URL) || response.status() !== 404) return;
    const pathname = new URL(url).pathname;
    if (/\.(?:js|html|webmanifest|json)$/i.test(pathname)) local404s.push(url);
  });
  return local404s;
}

async function countExactScript(page, pathname) {
  return page.locator('script[src]').evaluateAll((scripts, expectedPathname) => (
    scripts.filter((script) => {
      try {
        return new URL(script.src, window.location.href).pathname === expectedPathname;
      } catch (_) {
        return false;
      }
    }).length
  ), pathname);
}

test.describe('Varga Cantieri startup smoke', () => {
  test('loads the PWA shell and authentication surface', async ({ page }) => {
    const critical404s = watchCriticalLocal404s(page);
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded' });

    await expect(page.locator('#app-startup-loading')).toHaveCount(1);
    await expect(page.locator('#auth-gate')).toHaveCount(1);
    await expect(page.locator('#auth-email-form')).toHaveCount(1);
    await expect(page.locator('#auth-email-login-btn')).toHaveCount(1);
    await expect(page.locator('#auth-gate-login-btn')).toHaveCount(1);
    await expect(page.locator('#side-menu')).toHaveCount(1);
    await expect(page.locator('#open-panel-commesse')).toHaveCount(1);
    await expect(page.locator('#open-panel-squadre')).toHaveCount(1);
    await expect(page.locator('#open-hours-btn')).toHaveCount(1);
    await expect(page.locator('#open-gardening-assistant-btn')).toHaveCount(1);
    await expect(page.locator('#open-equipment-assistant-btn')).toHaveCount(1);
    await expect(page.locator('#green-assistant-overlay')).toHaveCount(1);
    await expect(page.locator('#auth-gate-message')).not.toHaveText('');

    expect(critical404s).toEqual([]);
  });

  test('survives reload without duplicating critical shell elements', async ({ page }) => {
    const critical404s = watchCriticalLocal404s(page);
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded' });
    await page.reload({ waitUntil: 'domcontentloaded' });

    await expect(page.locator('#auth-gate')).toHaveCount(1);
    await expect(page.locator('#app-startup-loading')).toHaveCount(1);
    expect(await countExactScript(page, '/app-pure-utils.js')).toBe(1);
    expect(await countExactScript(page, '/app.js')).toBe(1);
    expect(await countExactScript(page, '/green-assistant.js')).toBe(1);
    expect(critical404s).toEqual([]);
  });
});
