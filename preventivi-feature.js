(() => {
  'use strict';
  const PV = window.HeraPreventivi;
  if (!PV) throw new Error('Preventivi core non caricato.');

  function handleClick(event) {
    if (event.target.closest(`#${PV.menuId}`)) { event.preventDefault(); PV.open(); return; }
    const page = PV.page();
    if (!page || !event.target.closest(`#${PV.pageId}`)) return;
    if (event.target.closest('[data-pv-close]')) { PV.close(); return; }

    const viewButton = event.target.closest('[data-pv-view]');
    if (viewButton) {
      PV.state.view = viewButton.dataset.pvView;
      PV.state.editingQuoteId = '';
      PV.state.editingPriceListId = '';
      PV.renderCurrentView();
      return;
    }

    const button = event.target.closest('[data-pv-action]');
    if (!button) return;
    const action = button.dataset.pvAction;
    const id = button.dataset.id || '';

    if (action === 'new-quote') {
      if (!PV.state.priceLists.length) { PV.state.view = 'priceLists'; PV.state.editingPriceListId = 'new'; }
      else PV.state.editingQuoteId = 'new';
      PV.renderCurrentView();
    } else if (action === 'edit-quote') {
      PV.state.editingQuoteId = id; PV.renderCurrentView();
    } else if (action === 'delete-quote') {
      PV.deleteQuote(id);
    } else if (action === 'duplicate-quote') {
      PV.duplicateQuote(id);
    } else if (action === 'print-quote') {
      PV.printQuote(PV.getQuote(id));
    } else if (action === 'print-current-quote') {
      PV.printQuote(PV.getQuote(PV.state.editingQuoteId));
    } else if (action === 'new-price-list') {
      PV.state.editingPriceListId = 'new'; PV.renderCurrentView();
    } else if (action === 'edit-price-list') {
      PV.state.editingPriceListId = id; PV.renderCurrentView();
    } else if (action === 'delete-price-list') {
      PV.deletePriceList(id);
    } else if (action === 'cancel-editor') {
      PV.state.editingQuoteId = '';
      PV.state.editingPriceListId = '';
      PV.renderCurrentView();
    } else if (action === 'go-price-lists') {
      PV.state.editingQuoteId = '';
      PV.state.view = 'priceLists';
      PV.renderCurrentView();
    } else if (action === 'add-quote-line') {
      if (!PV.selectedPriceListIds().length) {
        PV.setFeedback('Seleziona prima almeno un prezziario.', 'error');
        PV.page()?.querySelector('[data-pv-price-list-choices]')?.scrollIntoView({ behavior: 'smooth', block: 'center' });
      } else PV.appendQuoteLine();
    } else if (action === 'remove-quote-line') {
      const row = button.closest('[data-line-id]');
      const rows = PV.page()?.querySelectorAll('[data-line-id]') || [];
      if (rows.length <= 1) { PV.clearQuoteLine(row); row.querySelector('[data-pv-line-item]').value = ''; }
      else row?.remove();
      PV.updateQuoteTotals();
    } else if (action === 'add-price-item') {
      PV.appendPriceItem();
    } else if (action === 'remove-price-item') {
      const row = button.closest('[data-price-item-id]');
      const rows = PV.page()?.querySelectorAll('[data-price-item-id]') || [];
      if (rows.length <= 1) row.querySelectorAll('input').forEach((input) => { input.value = input.type === 'number' ? '0' : ''; });
      else row?.remove();
    }
  }

  function handleInput(event) {
    if (!PV.page() || !event.target.closest(`#${PV.pageId}`)) return;
    if (event.target.matches('[data-pv-quote-search]')) {
      PV.state.quoteSearch = event.target.value;
      PV.renderQuoteOverview();
      const input = PV.page()?.querySelector('[data-pv-quote-search]');
      input?.focus(); input?.setSelectionRange(PV.state.quoteSearch.length, PV.state.quoteSearch.length);
    } else if (event.target.matches('[data-pv-price-list-search]')) {
      PV.state.priceListSearch = event.target.value;
      PV.renderPriceListOverview();
      const input = PV.page()?.querySelector('[data-pv-price-list-search]');
      input?.focus(); input?.setSelectionRange(PV.state.priceListSearch.length, PV.state.priceListSearch.length);
    } else if (event.target.matches('[data-pv-line-filter]')) {
      const row = event.target.closest('[data-line-id]');
      const select = row?.querySelector('[data-pv-line-item]');
      if (select) select.innerHTML = PV.buildPriceItemOptions(select.value, event.target.value);
    } else if (event.target.matches('[data-pv-line-quantity], [data-pv-vat]')) {
      PV.updateQuoteTotals();
    }
  }

  function handleChange(event) {
    if (!PV.page() || !event.target.closest(`#${PV.pageId}`)) return;
    if (event.target.matches('input[name="priceListIds"]')) PV.refreshLineSelectors();
    else if (event.target.matches('[data-pv-line-item]')) PV.populateQuoteLine(event.target.closest('[data-line-id]'), event.target.value);
    else if (event.target.matches('[data-pv-price-import]')) {
      PV.importPriceFile(event.target.files?.[0]);
      event.target.value = '';
    }
  }

  function handleSubmit(event) {
    const quoteForm = event.target.closest('[data-pv-quote-form]');
    if (quoteForm) { event.preventDefault(); PV.saveQuote(quoteForm); return; }
    const priceListForm = event.target.closest('[data-pv-price-list-form]');
    if (priceListForm) { event.preventDefault(); PV.savePriceList(priceListForm); }
  }

  function init() {
    PV.loadLocal();
    PV.ensureMenuButton();
    PV.ensurePage();
    document.addEventListener('click', handleClick);
    document.addEventListener('input', handleInput);
    document.addEventListener('change', handleChange);
    document.addEventListener('submit', handleSubmit);
    window.addEventListener('online', PV.scheduleSync);
    PV.setSyncBadge('💾 Dati sul dispositivo', 'warning');
    PV.connectFirebase();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();
