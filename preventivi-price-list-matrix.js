(() => {
  'use strict';

  const INSTALL_FLAG = 'priceListMatrixInstalled';
  const MATRIX_HEADERS = [
    'CAPITOLO',
    'CODICE PREZZO',
    'DESCRIZIONE',
    'UNITÀ DI MISURA',
    'PREZZO UNITARIO',
    'RIBASSO %'
  ];

  function install(PV) {
    if (!PV || PV[INSTALL_FLAG]) return;
    if (!PV.renderPriceListEditor || !PV.renderPriceListOverview || !PV.renderPriceListCard || !PV.populateQuoteLine) {
      window.setTimeout(() => install(window.HeraPreventivi), 60);
      return;
    }
    PV[INSTALL_FLAG] = true;

    PV.priceItemNetPrice = (item) => {
      const basePrice = Math.max(0, PV.parseNumber(item?.unitPrice));
      const discount = Math.min(100, Math.max(0, PV.parseNumber(item?.discount)));
      return PV.roundMoney(basePrice * (1 - discount / 100));
    };

    PV.newPriceItem = () => ({
      id: PV.uid('item'),
      chapter: '',
      code: '',
      description: '',
      unit: '',
      unitPrice: 0,
      discount: 0
    });

    PV.appendPriceItem = (item = PV.newPriceItem()) => {
      const container = PV.page()?.querySelector('[data-pv-price-items]');
      if (!container) return;
      const row = document.createElement('div');
      row.className = 'pv-item-row pv-item-row-matrix';
      row.dataset.priceItemId = item.id || PV.uid('item');
      row.innerHTML = `
        <label class="pv-label"><span>Capitolo</span><input data-pv-item-chapter value="${PV.escapeHtml(item.chapter || '')}" placeholder="Es. Sfalcio"></label>
        <label class="pv-label"><span class="pv-required">Codice prezzo</span><input data-pv-item-code value="${PV.escapeHtml(item.code || '')}" required></label>
        <label class="pv-label"><span class="pv-required">Descrizione</span><input data-pv-item-description value="${PV.escapeHtml(item.description || '')}" required></label>
        <label class="pv-label"><span class="pv-required">U.M.</span><input data-pv-item-unit value="${PV.escapeHtml(item.unit || '')}" placeholder="mq, ml, ora…" required></label>
        <label class="pv-label"><span class="pv-required">Prezzo unitario €</span><input type="number" min="0" step="0.01" data-pv-item-price value="${PV.escapeHtml(PV.parseNumber(item.unitPrice))}" required></label>
        <label class="pv-label"><span>Ribasso %</span><input type="number" min="0" max="100" step="0.01" data-pv-item-discount value="${PV.escapeHtml(PV.parseNumber(item.discount))}"></label>
        <button type="button" class="pv-icon-btn" data-pv-action="remove-price-item" aria-label="Rimuovi voce">✕</button>`;
      container.appendChild(row);
    };

    PV.collectPriceItems = () => [...(PV.page()?.querySelectorAll('[data-price-item-id]') || [])].map((row) => ({
      id: row.dataset.priceItemId,
      chapter: row.querySelector('[data-pv-item-chapter]')?.value.trim() || '',
      code: row.querySelector('[data-pv-item-code]')?.value.trim() || '',
      description: row.querySelector('[data-pv-item-description]')?.value.trim() || '',
      unit: row.querySelector('[data-pv-item-unit]')?.value.trim() || '',
      unitPrice: PV.parseNumber(row.querySelector('[data-pv-item-price]')?.value),
      discount: PV.parseNumber(row.querySelector('[data-pv-item-discount]')?.value)
    }));

    PV.validatePriceItems = (items) => {
      if (!items.length) return 'Inserisci almeno una voce nel prezziario.';
      const seenCodes = new Set();
      for (const item of items) {
        if (!item.code || !item.description || !item.unit) {
          return 'Compila codice prezzo, descrizione e unità di misura per ogni voce.';
        }
        if (item.unitPrice < 0) return 'Il prezzo unitario non può essere negativo.';
        if (item.discount < 0 || item.discount > 100) return 'Il ribasso deve essere compreso tra 0 e 100.';
        const code = PV.normalizeText(item.code);
        if (seenCodes.has(code)) return `Il codice ${item.code} è presente più di una volta.`;
        seenCodes.add(code);
      }
      return '';
    };

    const originalRenderPriceListCard = PV.renderPriceListCard;
    PV.renderPriceListCard = (list) => {
      const html = originalRenderPriceListCard(list);
      const marker = '<button type="button" class="pv-btn pv-btn-danger pv-btn-small" data-pv-action="delete-price-list"';
      const downloadButton = `<button type="button" class="pv-btn pv-btn-light pv-btn-small" data-pv-action="download-price-list" data-id="${PV.escapeHtml(list.id)}">Scarica Excel</button>`;
      return html.includes(marker) ? html.replace(marker, `${downloadButton}${marker}`) : html;
    };

    const originalRenderPriceListOverview = PV.renderPriceListOverview;
    PV.renderPriceListOverview = () => {
      originalRenderPriceListOverview();
      const toolbar = PV.content()?.querySelector('.pv-toolbar');
      if (!toolbar || toolbar.querySelector('[data-pv-action="download-price-matrix"]')) return;
      const actions = document.createElement('div');
      actions.className = 'pv-matrix-toolbar-actions';
      actions.innerHTML = `
        <button type="button" class="pv-btn pv-btn-light" data-pv-action="download-price-matrix">Scarica matrice vuota</button>
        <button type="button" class="pv-btn pv-btn-primary" data-pv-action="new-price-list">+ Nuovo prezziario</button>`;
      const existingNew = toolbar.querySelector('[data-pv-action="new-price-list"]');
      existingNew?.remove();
      toolbar.appendChild(actions);
    };

    const originalRenderPriceListEditor = PV.renderPriceListEditor;
    PV.renderPriceListEditor = (existingList) => {
      originalRenderPriceListEditor(existingList);
      const page = PV.page();
      if (!page) return;

      const subtitle = page.querySelector('.pv-section-head .pv-muted');
      if (subtitle) subtitle.textContent = 'Ogni voce può contenere capitolo, codice prezzo, descrizione, unità di misura, prezzo unitario e ribasso.';

      const itemHead = page.querySelector('.pv-items-head');
      if (itemHead) itemHead.innerHTML = '<span>Capitolo</span><span>Codice prezzo</span><span>Descrizione</span><span>U.M.</span><span>Prezzo unitario</span><span>Ribasso %</span><span></span>';

      const fileBox = page.querySelector('.pv-file-box');
      if (fileBox) {
        fileBox.innerHTML = `
          <div class="pv-matrix-file-actions">
            <button type="button" class="pv-btn pv-btn-light" data-pv-action="download-current-price-list">Scarica questo prezziario</button>
            <label class="pv-btn pv-btn-secondary pv-file-label">
              Ricarica e sostituisci da Excel
              <input type="file" accept=".xlsx,.xls,.csv" data-pv-price-import hidden>
            </label>
          </div>
          <span class="pv-muted">Colonne richieste: Capitolo, Codice prezzo, Descrizione, Unità di misura, Prezzo unitario e Ribasso %. Il file sostituisce tutte le righe visualizzate.</span>`;
      }
    };

    PV.findHeaderIndex = (headers, aliases) => {
      const normalized = headers.map(PV.normalizeText);
      return normalized.findIndex((header) => aliases.some((alias) => header === alias || header.includes(alias)));
    };

    PV.importPriceFile = async (file) => {
      if (!file) return;
      if (!window.XLSX) {
        PV.setFeedback('Libreria Excel non disponibile. Ricarica l’app e riprova.', 'error');
        return;
      }
      try {
        const workbook = window.XLSX.read(await file.arrayBuffer(), { type: 'array' });
        const targetSheetName = workbook.SheetNames.find((name) => PV.normalizeText(name) === 'prezziario') || workbook.SheetNames[0];
        const rows = window.XLSX.utils.sheet_to_json(workbook.Sheets[targetSheetName], { header: 1, defval: '', raw: false });
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
        const priceIndex = PV.findHeaderIndex(headers, ['prezzo unitario', 'prezzo', 'importo', 'tariffa']);
        const discountIndex = PV.findHeaderIndex(headers, ['ribasso %', 'ribasso', 'sconto %', 'sconto']);

        if ([chapterIndex, codeIndex, descriptionIndex, unitIndex, priceIndex, discountIndex].some((index) => index < 0)) {
          throw new Error('Non riconosco tutte le colonne. Servono: Capitolo, Codice prezzo, Descrizione, Unità di misura, Prezzo unitario e Ribasso %.');
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
        PV.setFeedback(`Caricate ${imported.length} voci. L’elenco precedente è stato sostituito: premi “Salva prezziario” per confermare aggiunte, modifiche ed eliminazioni.`, 'success');
      } catch (error) {
        console.warn('Preventivi: aggiornamento completo prezziario non riuscito.', error);
        PV.setFeedback(error?.message || 'Importazione non riuscita.', 'error');
      }
    };

    PV.createMatrixWorkbook = (list = null) => {
      if (!window.XLSX) throw new Error('Libreria Excel non disponibile.');
      const items = Array.isArray(list?.items) ? list.items : [];
      const rows = [MATRIX_HEADERS, ...items.map((item) => [
        item.chapter || '',
        item.code || '',
        item.description || '',
        item.unit || '',
        PV.parseNumber(item.unitPrice),
        PV.parseNumber(item.discount)
      ])];
      const sheet = window.XLSX.utils.aoa_to_sheet(rows);
      sheet['!cols'] = [
        { wch: 24 },
        { wch: 20 },
        { wch: 55 },
        { wch: 18 },
        { wch: 18 },
        { wch: 14 }
      ];
      sheet['!autofilter'] = { ref: `A1:F${Math.max(2, rows.length)}` };

      const instructions = window.XLSX.utils.aoa_to_sheet([
        ['MATRICE PREZZIARIO – ISTRUZIONI'],
        [],
        ['CAMPO', 'COME COMPILARLO'],
        ['CAPITOLO', 'Categoria o raggruppamento della lavorazione.'],
        ['CODICE PREZZO', 'Obbligatorio e univoco: serve per riconoscere e aggiornare la voce.'],
        ['DESCRIZIONE', 'Descrizione completa della lavorazione.'],
        ['UNITÀ DI MISURA', 'Esempi: mq, ml, cad, ora, giorno, corpo.'],
        ['PREZZO UNITARIO', 'Prezzo prima del ribasso.'],
        ['RIBASSO %', 'Inserire 15 per indicare un ribasso del 15%; usare 0 quando non previsto.'],
        ['AGGIORNAMENTO COMPLETO', 'Quando il file viene ricaricato nello stesso prezziario, le righe del file sostituiscono tutte le voci presenti nell’editor. Le righe eliminate dal file vengono eliminate anche dall’app dopo il salvataggio.']
      ]);
      instructions['!cols'] = [{ wch: 28 }, { wch: 95 }];

      const workbook = window.XLSX.utils.book_new();
      window.XLSX.utils.book_append_sheet(workbook, sheet, 'PREZZIARIO');
      window.XLSX.utils.book_append_sheet(workbook, instructions, 'ISTRUZIONI');
      return workbook;
    };

    PV.safeFileName = (value, fallback = 'Prezziario') => {
      const text = String(value || fallback).trim().replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, '_');
      return text || fallback;
    };

    PV.downloadPriceListMatrix = (list = null) => {
      try {
        const workbook = PV.createMatrixWorkbook(list);
        const fileName = list
          ? `${PV.safeFileName(list.name)}.xlsx`
          : 'Matrice_Prezziario_Hera_App.xlsx';
        window.XLSX.writeFile(workbook, fileName, { compression: true });
      } catch (error) {
        console.warn('Preventivi: download matrice non riuscito.', error);
        PV.setFeedback(error?.message || 'Impossibile scaricare il file Excel.', 'error');
      }
    };

    PV.currentPriceListDraft = () => {
      const form = PV.page()?.querySelector('[data-pv-price-list-form]');
      const existing = PV.state.editingPriceListId !== 'new' ? PV.getPriceList(PV.state.editingPriceListId) : null;
      return {
        id: existing?.id || '',
        name: String(form?.elements?.namedItem('name')?.value || existing?.name || 'Matrice prezziario').trim(),
        reference: String(form?.elements?.namedItem('reference')?.value || existing?.reference || '').trim(),
        description: String(form?.elements?.namedItem('description')?.value || existing?.description || '').trim(),
        items: PV.collectPriceItems()
      };
    };

    const originalAvailablePriceItems = PV.availablePriceItems;
    PV.availablePriceItems = (selectedIds, filter = '') => {
      const query = PV.normalizeText(filter);
      if (!query) return originalAvailablePriceItems(selectedIds, filter);
      const result = [];
      selectedIds.forEach((listId) => {
        const list = PV.getPriceList(listId);
        (list?.items || []).forEach((item) => {
          if (PV.normalizeText([item.chapter, item.code, item.description, item.unit].join(' ')).includes(query)) result.push({ list, item });
        });
      });
      return result.sort((a, b) => String(a.item.description || '').localeCompare(String(b.item.description || ''), 'it'));
    };

    const originalBuildPriceItemOptions = PV.buildPriceItemOptions;
    PV.buildPriceItemOptions = (selectedValue = '', filter = '') => {
      const selectedIds = PV.selectedPriceListIds();
      if (!selectedIds.length) return originalBuildPriceItemOptions(selectedValue, filter);
      return `<option value="">Seleziona la voce trovata</option>${PV.availablePriceItems(selectedIds, filter).map(({ list, item }) => {
        const value = `${list.id}::${item.id}`;
        const discount = Math.max(0, PV.parseNumber(item.discount));
        const chapter = item.chapter ? `${item.chapter} · ` : '';
        const discountText = discount > 0 ? ` · ribasso ${String(discount).replace('.', ',')}%` : '';
        return `<option value="${PV.escapeHtml(value)}" ${value === selectedValue ? 'selected' : ''}>${PV.escapeHtml(`[${list.name}] ${chapter}${item.code} — ${item.description} (${item.unit || 'u.m.'}) · ${PV.formatMoney(PV.priceItemNetPrice(item))}${discountText}`)}</option>`;
      }).join('')}`;
    };

    PV.populateQuoteLine = (row, value) => {
      const resolved = PV.resolvePriceItem(value);
      if (!resolved) {
        PV.clearQuoteLine(row, true);
        return;
      }
      row.querySelector('[data-pv-line-filter]').value = resolved.item.description || '';
      row.querySelector('[data-pv-line-code]').value = resolved.item.code || '';
      row.querySelector('[data-pv-line-unit]').value = resolved.item.unit || '';
      row.querySelector('[data-pv-line-price]').value = String(PV.priceItemNetPrice(resolved.item));
      PV.updateQuoteTotals();
    };

    document.addEventListener('click', (event) => {
      const button = event.target.closest('[data-pv-action]');
      if (!button || !event.target.closest(`#${PV.pageId}`)) return;
      const action = button.dataset.pvAction;
      if (!['download-price-matrix', 'download-price-list', 'download-current-price-list'].includes(action)) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      if (action === 'download-price-matrix') PV.downloadPriceListMatrix();
      else if (action === 'download-price-list') PV.downloadPriceListMatrix(PV.getPriceList(button.dataset.id));
      else PV.downloadPriceListMatrix(PV.currentPriceListDraft());
    }, true);
  }

  install(window.HeraPreventivi);
})();
