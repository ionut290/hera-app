(() => {
  'use strict';

  const FLAG = 'priceListLiveFixInstalled';
  const MAX_MATCHES = 60;
  const MAX_VISIBLE = 10;

  function install(PV) {
    if (!PV || PV[FLAG]) return;
    if (!PV.priceListMatrixInstalled || !PV.renderPriceListEditor || !PV.importPriceFile || !PV.buildPriceItemOptions) {
      window.setTimeout(() => install(window.HeraPreventivi), 60);
      return;
    }
    PV[FLAG] = true;

    const originalRenderPriceListEditor = PV.renderPriceListEditor;
    PV.renderPriceListEditor = (existingList) => {
      originalRenderPriceListEditor(existingList);
      PV.page()?.querySelector('.pv-shell')?.classList.add('pv-shell-full-width', 'pv-price-list-editor-shell');
    };

    PV.findHeaderIndex = (headers, aliases, exclusions = []) => {
      const normalizedHeaders = headers.map(PV.normalizeText);
      const normalizedAliases = aliases.map(PV.normalizeText).filter(Boolean);
      const normalizedExclusions = exclusions.map(PV.normalizeText).filter(Boolean);
      const allowed = (header) => !normalizedExclusions.some((excluded) => header.includes(excluded));
      const exact = normalizedHeaders.findIndex((header) => allowed(header) && normalizedAliases.includes(header));
      if (exact >= 0) return exact;
      return normalizedHeaders.findIndex((header) => allowed(header)
        && normalizedAliases.some((alias) => header.startsWith(alias) || header.includes(alias)));
    };

    PV.importPriceFile = async (file) => {
      if (!file) return;
      if (!window.XLSX) {
        PV.setFeedback('Libreria Excel non disponibile. Ricarica l’app e riprova.', 'error');
        return;
      }
      try {
        const workbook = window.XLSX.read(await file.arrayBuffer(), { type: 'array' });
        const sheetName = workbook.SheetNames.find((name) => PV.normalizeText(name) === 'prezziario') || workbook.SheetNames[0];
        const rows = window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: true });
        if (!rows.length) throw new Error('Il file è vuoto.');

        const detected = rows.findIndex((row) => {
          const text = PV.normalizeText(row.join(' '));
          return text.includes('codice') && text.includes('descrizione') && text.includes('prezzo');
        });
        const headerRow = detected >= 0 ? detected : 0;
        const headers = rows[headerRow].map((cell) => String(cell || ''));
        const chapterIndex = PV.findHeaderIndex(headers, ['capitolo', 'categoria', 'gruppo']);
        const codeIndex = PV.findHeaderIndex(headers, ['codice prezzo', 'codice', 'cod.', 'code', 'articolo']);
        const descriptionIndex = PV.findHeaderIndex(headers, ['descrizione', 'lavorazione', 'voce', 'prestazione']);
        const unitIndex = PV.findHeaderIndex(headers, ['unita di misura', 'unità di misura', 'u.m.', 'um', 'unita']);
        const priceIndex = PV.findHeaderIndex(
          headers,
          ['prezzo unitario', 'importo unitario', 'prezzo netto', 'prezzo', 'importo', 'tariffa'],
          ['codice prezzo']
        );
        const discountIndex = PV.findHeaderIndex(headers, ['ribasso %', 'ribasso', 'sconto %', 'sconto']);
        if ([chapterIndex, codeIndex, descriptionIndex, unitIndex, priceIndex, discountIndex].some((index) => index < 0)) {
          throw new Error('Non riconosco tutte le colonne richieste del prezziario.');
        }

        const existingByCode = new Map(PV.collectPriceItems().map((item) => [PV.normalizeText(item.code), item]));
        const imported = rows.slice(headerRow + 1).map((row) => {
          const code = String(row[codeIndex] || '').trim();
          const previous = existingByCode.get(PV.normalizeText(code));
          return {
            id: previous?.id || PV.uid('item'),
            chapter: String(row[chapterIndex] || '').trim(),
            code,
            description: String(row[descriptionIndex] || '').trim(),
            unit: String(row[unitIndex] || '').trim(),
            unitPrice: PV.parseNumber(row[priceIndex]),
            discount: PV.parseNumber(row[discountIndex])
          };
        }).filter((item) => item.code || item.description || item.unit || item.chapter || item.unitPrice || item.discount);

        if (!imported.length) throw new Error('Non ho trovato righe compilate nel foglio PREZZIARIO.');
        const validation = PV.validatePriceItems(imported);
        if (validation) throw new Error(validation);
        const container = PV.page()?.querySelector('[data-pv-price-items]');
        if (!container) throw new Error('Editor del prezziario non disponibile.');
        container.innerHTML = '';
        imported.forEach(PV.appendPriceItem);
        const pricedItems = imported.filter((item) => item.unitPrice > 0).length;
        PV.setFeedback(
          `Caricate ${imported.length} voci, di cui ${pricedItems} con prezzo unitario. Premi “Salva prezziario” per confermare.`,
          pricedItems ? 'success' : 'warning'
        );
      } catch (error) {
        console.warn('Preventivi: importazione prezziario non riuscita.', error);
        PV.setFeedback(error?.message || 'Importazione non riuscita.', 'error');
      }
    };

    PV.availablePriceItems = (selectedIds, filter = '') => {
      const query = PV.normalizeText(filter);
      if (!query) return [];
      const terms = query.split(/\s+/).filter(Boolean);
      const result = [];
      selectedIds.forEach((listId) => {
        const list = PV.getPriceList(listId);
        (list?.items || []).forEach((item) => {
          const chapter = PV.normalizeText(item.chapter);
          const code = PV.normalizeText(item.code);
          const description = PV.normalizeText(item.description);
          const unit = PV.normalizeText(item.unit);
          const text = `${chapter} ${code} ${description} ${unit}`;
          if (!terms.every((term) => text.includes(term))) return;
          const words = description.split(/[^a-z0-9]+/).filter(Boolean);
          let score = 0;
          if (description === query || code === query) score += 1000;
          if (description.startsWith(query)) score += 800;
          if (words.some((word) => word.startsWith(query))) score += 700;
          if (code.startsWith(query)) score += 650;
          if (chapter.startsWith(query)) score += 500;
          if (description.includes(query)) score += 300;
          result.push({ list, item, score });
        });
      });
      return result.sort((a, b) => b.score - a.score
        || String(a.item.description || '').localeCompare(String(b.item.description || ''), 'it'))
        .slice(0, MAX_MATCHES);
    };

    const originalBuildPriceItemOptions = PV.buildPriceItemOptions;
    PV.buildPriceItemOptions = (selectedValue = '', filter = '') => {
      const selectedIds = PV.selectedPriceListIds();
      if (!selectedIds.length) return originalBuildPriceItemOptions(selectedValue, filter);
      const query = String(filter || '').trim();
      const matches = query ? PV.availablePriceItems(selectedIds, query) : [];
      if (selectedValue && !matches.some(({ list, item }) => `${list.id}::${item.id}` === selectedValue)) {
        const resolved = PV.resolvePriceItem(selectedValue);
        if (resolved) matches.unshift(resolved);
      }
      const placeholder = query ? 'Seleziona la voce trovata' : 'Scrivi per vedere i suggerimenti';
      return `<option value="">${placeholder}</option>${matches.map(({ list, item }) => {
        const value = `${list.id}::${item.id}`;
        const discount = Math.max(0, PV.parseNumber(item.discount));
        return `<option value="${PV.escapeHtml(value)}" ${value === selectedValue ? 'selected' : ''}>${PV.escapeHtml(`${item.code} — ${item.description} · ${PV.formatMoney(PV.priceItemNetPrice(item))}${discount ? ` · -${discount}%` : ''}`)}</option>`;
      }).join('')}`;
    };

    function resizeSuggestions(input) {
      const select = input.closest('[data-line-id]')?.querySelector('[data-pv-line-item]');
      if (!select) return;
      const count = Math.max(0, select.options.length - 1);
      if (input.value.trim() && count) {
        select.size = Math.min(MAX_VISIBLE, count + 1);
        select.classList.add('pv-live-select-open');
      } else {
        select.size = 1;
        select.classList.remove('pv-live-select-open');
      }
    }

    document.addEventListener('input', (event) => {
      if (!event.target.matches?.('[data-pv-line-filter]') || !event.target.closest(`#${PV.pageId}`)) return;
      window.setTimeout(() => resizeSuggestions(event.target), 0);
    });
    document.addEventListener('change', (event) => {
      if (!event.target.matches?.('[data-pv-line-item]')) return;
      event.target.size = 1;
      event.target.classList.remove('pv-live-select-open');
    });
    document.addEventListener('click', (event) => {
      if (event.target.closest('[data-pv-line-filter], [data-pv-line-item]')) return;
      PV.page()?.querySelectorAll('[data-pv-line-item].pv-live-select-open').forEach((select) => {
        select.size = 1;
        select.classList.remove('pv-live-select-open');
      });
    });
  }

  install(window.HeraPreventivi);
})();
