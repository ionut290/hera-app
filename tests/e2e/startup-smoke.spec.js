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
    await expect(page.locator('#open-urban-furniture-btn')).toHaveCount(1);
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

test.describe('Arredo urbano', () => {
  test.use({ viewport: { width: 390, height: 844 } });

  test('searches Overpass and opens map, sheet, 360 route and native Whazzup', async ({ page }) => {
    await page.route('**/api/urban-furniture?**', (route) => route.fulfill({
      status: 200,
      contentType: 'application/json',
      body: JSON.stringify({ elements: [{ type: 'node', id: 991, lat: 44.4949, lon: 11.3426, tags: { amenity: 'bench', name: 'Panchina Piazza', material: 'wood', wheelchair: 'yes', operator: 'Comune di Bologna' } }] })
    }));
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded' });
    await page.evaluate(() => {
      const bounds = { valid: false, extend() { this.valid = true; return this; }, isValid() { return this.valid; }, pad() { return this; } };
      const map = { setView() { return this; }, fitBounds() { return this; }, invalidateSize() { return this; } };
      window.L = {
        map: () => map,
        tileLayer: () => ({ addTo() { return this; }, remove() {}, bringToBack() {} }),
        layerGroup: () => ({ addTo() { return this; }, clearLayers() {} }),
        marker: () => ({ addTo() { return this; }, bindPopup() { return this; } }),
        circleMarker: () => ({ addTo() { return this; }, bindPopup() { return this; }, remove() {} }),
        divIcon: (options) => options,
        latLngBounds: () => bounds
      };
      window.__urbanWhazzupUrl = '';
      window.Capacitor = { isNativePlatform: () => true, getPlatform: () => 'android', Plugins: { HeraWhatsApp: { open({ url }) { window.__urbanWhazzupUrl = url; return Promise.resolve({ opened: true }); } } } };
    });

    await page.evaluate(() => document.getElementById('open-urban-furniture-btn').click());
    await expect(page.locator('#urban-furniture-category option')).toHaveCount(40);
    await expect(page.locator('#urban-furniture-category option[value="playground"]')).toHaveText('Aree giochi');
    await page.locator('#urban-furniture-category').selectOption('bench');
    await page.locator('#urban-furniture-form').evaluate((form) => form.requestSubmit());
    await expect(page.locator('#urban-furniture-status')).toContainText('1 elemento trovato');
    await expect(page.locator('.urban-furniture-result')).toContainText('Panchina Piazza');
    await page.locator('[data-urban-result-index="0"]').click();
    await expect(page.locator('#urban-furniture-sheet')).not.toHaveClass(/hidden/);
    await expect(page.locator('#urban-furniture-sheet-title')).toContainText('Panchina Piazza');
    await expect(page.locator('#urban-furniture-sheet-body')).toContainText('Caratteristiche e servizi');
    await expect(page.locator('#urban-furniture-sheet-body')).toContainText('Comune di Bologna');
    await expect(page.locator('#urban-furniture-sheet-body')).toContainText('Accessibile in sedia a rotelle');
    await expect(page.locator('#urban-furniture-sheet-body')).toContainText('Apri la fonte originale');

    await page.evaluate(() => {
      window.__urbanStreetView = null;
      window.HeraStreetViewCards = { openForCoordinates(coords, trigger, options) { window.__urbanStreetView = { coords, label: trigger.textContent, options }; return Promise.resolve(true); } };
    });
    await page.locator('#urban-furniture-street-view').click();
    await page.waitForFunction(() => Boolean(window.__urbanStreetView));
    const streetView = await page.evaluate(() => window.__urbanStreetView);
    expect(streetView.coords).toEqual({ lat: 44.4949, lng: 11.3426 });
    expect(streetView.options.targetLabel).toBe('Panchina');

    await page.locator('#urban-furniture-whazzup').click();
    const message = await page.evaluate(() => new URL(window.__urbanWhazzupUrl).searchParams.get('text'));
    expect(message).toContain('🪑 *ARREDO URBANO*');
    expect(message).toContain('*Elemento:* Panchina Piazza');
    expect(message).toContain('destination=44.4949,11.3426');
  });
});

test.describe('Catasto arboreo', () => {
  test.use({ viewport: { width: 390, height: 844 } });

  test('offers point/code search and toggles the tree map fullscreen', async ({ page }) => {
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded' });

    await expect(page.locator('label[for="tree-number"]')).toHaveText('Numero punto o codice albero');
    await expect(page.locator('#tree-number')).toHaveAttribute('placeholder', 'Es. 64228 oppure 103VT');
    await expect(page.locator('#tree-map-fullscreen-btn')).toHaveText(/SCHERMO INTERO/);
    await expect(page.locator('#tree-map-planting-filter')).toHaveValue('all');
    await expect(page.locator('#tree-map-planting-filter option')).toHaveText([
      'Tutti gli alberi',
      '🌱 Ultimo anno',
      '🌱 Ultimi 3 anni',
      '🌱 Ultimi 5 anni'
    ]);

    await page.evaluate(() => document.getElementById('tree-map-fullscreen-btn').click());
    await expect(page.locator('#tree-map-card')).toHaveClass(/tree-map-card--fullscreen/);
    await expect(page.locator('body')).toHaveClass(/tree-map-fullscreen-open/);
    await expect(page.locator('#tree-map-fullscreen-btn')).toHaveText(/CHIUDI MAPPA/);

    await page.evaluate(() => document.getElementById('tree-map-fullscreen-btn').click());
    await expect(page.locator('#tree-map-card')).not.toHaveClass(/tree-map-card--fullscreen/);
    await expect(page.locator('body')).not.toHaveClass(/tree-map-fullscreen-open/);
  });

  test('filters the visible map by the official planting date', async ({ page }) => {
    let filteredRequestUrl = '';
    await page.route('**/alberi-manutenzioni/records?**', async (route) => {
      const decodedUrl = decodeURIComponent(route.request().url());
      if (decodedUrl.includes('data_impnt')) filteredRequestUrl = decodedUrl;
      await route.fulfill({
        status: 200,
        contentType: 'application/json',
        body: JSON.stringify({
          total_count: 1,
          results: [{
            num_pt: '130403',
            cod_alb: '016S',
            classe: 'Acer freemanii',
            data_impnt: '2026-06-07',
            geo_point_2d: { lat: 44.451009, lon: 11.359420 }
          }]
        })
      });
    });
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded' });
    await page.evaluate(() => {
      const bounds = {
        getNorthEast: () => ({ lat: 44.452, lng: 11.361 }),
        pad() { return this; },
        contains() { return true; }
      };
      const map = {
        setView() { return this; },
        on() { return this; },
        invalidateSize() { return this; },
        getZoom() { return 17; },
        getCenter() { return { lat: 44.451, lng: 11.359 }; },
        getBounds() { return bounds; }
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
        divIcon: (options) => options
      };
    });
    await page.evaluate(() => document.getElementById('open-tree-search-btn').click());
    await page.locator('#tree-map-planting-filter').selectOption('3');
    await expect.poll(() => filteredRequestUrl).toContain('data_impnt');
    expect(filteredRequestUrl).toContain("data_impnt >= date'");
    expect(filteredRequestUrl).toContain("data_impnt <= date'");
    await expect(page.locator('#tree-map-status')).toContainText('nuovi impianti negli ultimi 3 anni visualizzati');
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

  test('shows six tree details, expands the rest and prepares a concise Whazzup message', async ({ page }) => {
    const tree = {
      num_pt: '118907',
      cod_alb: '183VT',
      classe: 'Fraxinus excelsior',
      cl_h: 'Cl2: 6mt - 12mt',
      classe_circonferenza_diametro: 'Cl5: 60 - 90 (19-28 cm)',
      quartiere: 'Santo Stefano',
      dimora: 'Prato',
      anni_impnt: -1,
      area_statistica: 'IRNERIO-2',
      data_agg: '2021-02-08',
      geo_point_2d: { lat: 44.503598, lon: 11.352738 }
    };

    await page.route('**/alberi-manutenzioni/records?**', async (route) => {
      await route.fulfill({
        status: 200,
        contentType: 'application/json',
        body: JSON.stringify({ total_count: 1, results: [tree] })
      });
    });
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded' });
    await page.evaluate(() => {
      const map = {
        setView() { return this; },
        on() { return this; },
        invalidateSize() { return this; },
        getZoom() { return 19; },
        getCenter() { return { lat: 44.503598, lng: 11.352738 }; },
        getBounds() {
          return {
            getNorthEast: () => ({ lat: 44.504, lng: 11.353 }),
            pad() { return this; },
            contains() { return true; }
          };
        }
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
          openPopup() { return this; },
          on() { return this; },
          addTo(layer) { layer.layers?.push(this); return this; },
          remove() {}
        }),
        divIcon: (options) => options
      };
    });
    await page.evaluate(() => document.getElementById('open-tree-search-btn').click());
    await page.locator('#tree-number').fill('118907');
    await page.locator('#tree-search-form').evaluate((form) => form.requestSubmit());

    await expect(page.locator('.tree-result-grid > div:visible')).toHaveCount(6);
    await expect(page.locator('.tree-details-toggle')).toHaveAttribute('aria-expanded', 'false');
    await page.locator('.tree-details-toggle').click();
    await expect(page.locator('.tree-result-grid > div:visible')).toHaveCount(Object.keys(tree).length);
    await expect(page.locator('.tree-details-toggle')).toHaveText('MOSTRA SOLO I PRIMI 6 DETTAGLI');

    await page.evaluate(() => {
      window.__treeWhazzupUrl = '';
      window.Capacitor = {
        isNativePlatform: () => true,
        getPlatform: () => 'android',
        Plugins: {
          HeraWhatsApp: {
            open({ url }) {
              window.__treeWhazzupUrl = url;
              return Promise.resolve({ opened: true });
            }
          }
        }
      };
    });
    await page.locator('.tree-whazzup-share').click();
    const message = await page.evaluate(() => new URL(window.__treeWhazzupUrl).searchParams.get('text'));
    expect(message).toContain('🌳 *SCHEDA ALBERO*');
    expect(message.match(/^• /gm)).toHaveLength(6);
    expect(message).toContain('*Quartiere:* Santo Stefano');
    expect(message).not.toContain('*Dimora:*');
    expect(message).toContain('📍 *NAVIGA VERSO L’ALBERO*');
    expect(message).toContain('destination=44.503598,11.352738');

    await page.evaluate(() => {
      window.__treeStreetViewRequest = null;
      window.HeraStreetViewCards = {
        installed: true,
        openForCoordinates(coords, trigger, options) {
          window.__treeStreetViewRequest = {
            coords,
            label: trigger.textContent,
            options
          };
          return Promise.resolve(true);
        }
      };
    });
    await expect(page.locator('.tree-street-view')).toHaveText('🌐 VISTA 360° E PERCORSO');
    await page.locator('.tree-street-view').click();
    await page.waitForFunction(() => Boolean(window.__treeStreetViewRequest));
    const streetViewRequest = await page.evaluate(() => window.__treeStreetViewRequest);
    expect(streetViewRequest.coords).toEqual({ lat: 44.503598, lng: 11.352738 });
    expect(streetViewRequest.options.targetLabel).toBe('Albero');
    expect(streetViewRequest.options.modalTitle).toContain('#118907');
  });
});
