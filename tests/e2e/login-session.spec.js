const { test, expect } = require('@playwright/test');

const baseUrl = process.env.PLAYWRIGHT_BASE_URL || 'http://127.0.0.1:8080';
const harnessUrl = (query = '') => new URL(`tests/e2e/login-session-harness.html${query}`, `${baseUrl.replace(/\/$/, '')}/`).toString();

async function submitLogin(page) {
  await page.locator('#auth-email-form').evaluate((form) => form.requestSubmit());
}

test.describe('Hera App - login resiliente', () => {
  test('online verifica sempre la password anche con una sessione salvata', async ({ page }) => {
    await page.goto(harnessUrl('?online=1&session=valid&currentUser=valid&remembered=1'), { waitUntil: 'domcontentloaded' });
    await submitLogin(page);

    await expect.poll(() => page.evaluate(() => window.__loginTest.signInCalls)).toBe(1);
    await expect(page.locator('#auth-email-feedback')).toContainText(/Login completato/i);
    await expect.poll(() => page.evaluate(() => window.__loginTest.persistenceCalls.at(-1))).toBe('LOCAL');
  });

  test('entra con sessione salvata anche quando il dispositivo è offline', async ({ page }) => {
    await page.goto(harnessUrl('?online=0&session=valid&currentUser=valid&remembered=1'), { waitUntil: 'domcontentloaded' });
    await submitLogin(page);

    await expect.poll(() => page.evaluate(() => window.__loginTest.signInCalls)).toBe(0);
    await expect(page.locator('#auth-email-feedback')).toContainText(/Modalità offline attiva/i);
  });

  test('il primo accesso offline viene rifiutato senza tentare credenziali remote', async ({ page }) => {
    await page.goto(harnessUrl('?online=0&session=none&currentUser=none&remembered=1'), { waitUntil: 'domcontentloaded' });
    await submitLogin(page);

    await expect.poll(() => page.evaluate(() => window.__loginTest.signInCalls)).toBe(0);
    await expect(page.locator('#auth-email-feedback')).toContainText(/primo accesso.*richiede internet/i);
  });

  test('non riusa una sessione appartenente a un altro utente', async ({ page }) => {
    await page.goto(harnessUrl('?online=0&session=valid&currentUser=valid&remembered=1&email=operatore@example.com'), { waitUntil: 'domcontentloaded' });
    await page.locator('#auth-email-input').fill('altro@example.com');
    await submitLogin(page);

    await expect.poll(() => page.evaluate(() => window.__loginTest.signInCalls)).toBe(0);
    await expect(page.locator('#auth-email-feedback')).toContainText(/Rete troppo debole|primo accesso.*internet/i);
  });

  test('senza Ricordami usa persistenza di sessione e il login normale', async ({ page }) => {
    await page.goto(harnessUrl('?online=1&session=none&currentUser=none&remembered=0'), { waitUntil: 'domcontentloaded' });
    await submitLogin(page);

    await expect.poll(() => page.evaluate(() => window.__loginTest.signInCalls)).toBe(1);
    await expect.poll(() => page.evaluate(() => window.__loginTest.persistenceCalls.at(-1))).toBe('SESSION');
    await expect(page.locator('#auth-email-feedback')).toContainText(/Login completato/i);
  });

  test('la password non viene salvata in localStorage', async ({ page }) => {
    const password = 'PasswordTest123!';
    await page.goto(harnessUrl('?online=1&session=none&currentUser=none&remembered=1'), { waitUntil: 'domcontentloaded' });
    await page.locator('#auth-password-input').fill(password);
    await submitLogin(page);

    await expect.poll(() => page.evaluate(() => window.__loginTest.signInCalls)).toBe(1);
    const storageDump = await page.evaluate(() => JSON.stringify({ ...localStorage }));
    expect(storageDump).not.toContain(password);
  });
});
