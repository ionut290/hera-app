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

test.describe('Catasto arboreo', () => {
  test.use({ viewport: { width: 390, height: 844 } });

  test('offers point/code search and toggles the tree map fullscreen', async ({ page }) => {
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded' });

    await expect(page.locator('label[for="tree-number"]')).toHaveText('Numero punto o codice albero');
    await expect(page.locator('#tree-number')).toHaveAttribute('placeholder', 'Es. 64228 oppure 103VT');
    await expect(page.locator('#tree-map-fullscreen-btn')).toHaveText(/SCHERMO INTERO/);

    await page.evaluate(() => document.getElementById('tree-map-fullscreen-btn').click());
    await expect(page.locator('#tree-map-card')).toHaveClass(/tree-map-card--fullscreen/);
    await expect(page.locator('body')).toHaveClass(/tree-map-fullscreen-open/);
    await expect(page.locator('#tree-map-fullscreen-btn')).toHaveText(/CHIUDI MAPPA/);

    await page.evaluate(() => document.getElementById('tree-map-fullscreen-btn').click());
    await expect(page.locator('#tree-map-card')).not.toHaveClass(/tree-map-card--fullscreen/);
    await expect(page.locator('body')).not.toHaveClass(/tree-map-fullscreen-open/);
  });

  test('normalizes a tree code and presents every matching tree on the map', async ({ page }) => {
    const records = [
      { num_pt: '64228', cod_alb: '103VT', classe: 'Platanus acerifolia', geo_point_2d: { lat: 44.49, lon: 11.34 } },
      { num_pt: '64229', cod_alb: '103VT', classe: 'Platanus acerifolia', geo_point_2d: { lat: 44.50, lon: 11.35 } }
    ];

    await page.route('**/alberi-manutenzioni/records?**', async (route) => {
      const decodedUrl = decodeURIComponent(route.request().url());
      const payload = decodedUrl.includes("cod_alb='103VT'")
        ? { total_count: records.length, results: records }
        : { total_count: 0, results: [] };
      await route.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify(payload) });
    });
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded' });
    await page.evaluate(() => {
      const map = {
        setView() { return this; },
        on() { return this; },
        invalidateSize() { return this; },
        fitBounds() { return this; },
        getZoom() { return 13; },
        getCenter() { return { lat: 44.49, lng: 11.34 }; },
        getBounds() { return { getNorthEast: () => ({ lat: 44.50, lng: 11.35 }) }; }
      };
      window.L = {
        map: () => map,
        tileLayer: () => ({ addTo() { return this; }, remove() {} }),
        layerGroup: () => ({
          layers: [],
          addTo() { return this; },
          remove() {},
          getLayers() { return this.layers; }
        }),
        marker: () => ({
          bindPopup() { return this; },
          on() { return this; },
          addTo(layer) { layer.layers.push(this); return this; },
          remove() {}
        }),
        divIcon: (options) => options,
        latLngBounds: () => ({ pad() { return this; } })
      };
    });
    await page.evaluate(() => document.getElementById('open-tree-search-btn').click());
    await page.locator('#tree-number').fill(' 103vt ');
    await page.locator('#tree-search-form').evaluate((form) => form.requestSubmit());

    await expect(page.locator('#tree-number')).toHaveValue('103VT');
    await expect(page.locator('#tree-search-status')).toContainText('2 alberi trovati con il codice 103VT');
    await expect(page.locator('#tree-result h2')).toHaveText('Codice albero 103VT');
    await expect(page.locator('#tree-map-status')).toContainText('2 alberi con codice 103VT');
  });
});
