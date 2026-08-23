(() => {
  'use strict';
  if (window.HeraDocumentazioneGestioneClick?.installed) return;

  const text = (value) => String(value ?? '').trim();
  const upper = (value) => text(value).replace(/\s+/g, ' ').toLocaleUpperCase('it-IT');

  function isAdmin() {
    try { if (typeof canManageData === 'function' && canManageData()) return true; } catch (_) {}
    try { if (typeof window.canManageData === 'function' && window.canManageData()) return true; } catch (_) {}
    return false;
  }

  function getPlants() {
    try { if (Array.isArray(currentImpianti)) return currentImpianti; } catch (_) {}
    return Array.isArray(window.currentImpianti) ? window.currentImpianti : [];
  }

  function plantId(plant) {
    return text(plant?.id || plant?.docId || plant?.impiantoId || plant?.physicalPlantId || plant?.idSap || plant?.idSAP || plant?.['ID SAP']);
  }

  function plantName(plant) {
    return text(plant?.denominazione || plant?.nome || plant?.impianto || plant?.['Denominazione Impianto'] || 'Impianto');
  }

  function findCard(stack) {
    let node = stack?.parentElement || null;
    const list = document.getElementById('impianti-lista');
    while (node && node !== list && node !== document.body) {
      if (node.querySelector?.('.impianto-primary-actions [data-action-key="navigate"]')) return node;
      node = node.parentElement;
    }
    return null;
  }

  function findPlant(card) {
    if (!card) return null;
    const body = upper(card.textContent);
    return getPlants().find((plant) => {
      const name = upper(plantName(plant));
      const id = upper(plantId(plant));
      return (name && body.includes(name)) || (id && body.includes(id));
    }) || null;
  }

  function openDocs(plant, event) {
    event?.preventDefault?.();
    event?.stopPropagation?.();
    if (window.HeraAllPlantsDocumentsRuntime?.open) {
      void window.HeraAllPlantsDocumentsRuntime.open(plant);
    } else if (window.HeraCantiereDocuments?.open) {
      window.HeraCantiereDocuments.open(plant);
    }
  }

  function findPdfRequestButton(actions) {
    return Array.from(actions.querySelectorAll('button')).find((button) => {
      const label = upper(button.textContent);
      return label.includes('PDF RICHIESTA') || label.includes('INSERISCI PDF');
    }) || null;
  }

  function ensureButton(stack) {
    if (!isAdmin() || !(stack instanceof HTMLElement)) return;
    const actions = stack.querySelector('.item-actions-gestione');
    if (!actions) return;

    const plant = findPlant(findCard(stack));
    if (!plant) return;

    const pdfButton = findPdfRequestButton(actions);
    if (pdfButton && !pdfButton.dataset.documentazioneGestioneClick) {
      const replacement = pdfButton.cloneNode(true);
      replacement.type = 'button';
      replacement.textContent = '📁 ALLEGA DOCUMENTAZIONE';
      replacement.dataset.documentazioneGestioneClick = '1';
      replacement.removeAttribute('id');
      replacement.addEventListener('click', (event) => openDocs(plant, event));
      pdfButton.replaceWith(replacement);
    }

    actions.querySelectorAll('[data-all-doc-runtime], [data-occasional-doc-runtime], [data-cantiere-doc-admin], [data-cantiere-doc-fallback]').forEach((button) => {
      if (!button.dataset.documentazioneGestioneClick) button.remove();
    });

    if (actions.querySelector('[data-documentazione-gestione-click]')) return;

    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'btn all-doc-runtime-btn';
    button.dataset.documentazioneGestioneClick = '1';
    button.textContent = '📁 ALLEGA DOCUMENTAZIONE';
    button.addEventListener('click', (event) => openDocs(plant, event));
    actions.appendChild(button);
  }

  function onClick(event) {
    const gear = event.target?.closest?.('.gestione-toggle-btn');
    if (!gear) return;
    const list = document.getElementById('impianti-lista');
    if (!list || !list.contains(gear)) return;
    const stack = gear.closest('.impianto-management-stack');
    if (!stack) return;
    queueMicrotask(() => ensureButton(stack));
    window.setTimeout(() => ensureButton(stack), 0);
    window.setTimeout(() => ensureButton(stack), 80);
  }

  document.addEventListener('click', onClick, true);
  window.HeraDocumentazioneGestioneClick = {
    installed: true,
    refresh() {
      document.querySelectorAll('#impianti-lista .impianto-management-stack').forEach((stack) => {
        if (!stack.querySelector('.item-actions-gestione.hidden')) ensureButton(stack);
      });
    }
  };
})();
