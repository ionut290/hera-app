(() => {
  'use strict';

  const PV = window.HeraPreventivi;
  const M = window.HeraPreventiviModels;
  if (!PV) return;

  const clean = (value) => String(value ?? '').trim();
  const normalize = (value) => clean(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();

  function parseDiscount(value) {
    if (value === null || value === undefined || value === '') return 0;
    if (typeof value === 'boolean') return value ? 0 : 0;
    const raw = clean(value);
    if (!raw) return 0;
    const normalized = normalize(raw);
    if (['no', 'n', 'false', 'non previsto', 'nessuno'].includes(normalized)) return 0;
    let number = PV.parseNumber(raw);
    if (number > 0 && number < 1 && !raw.includes('%')) number *= 100;
    return Math.min(100, Math.max(0, number));
  }

  function discountIsEnabled(item) {
    if (!item || typeof item !== 'object') return false;
    const explicitFlag = item.discountEnabled ?? item.ribassoPrevisto ?? item.ribassoSiNo ?? item.ribasso_si_no ?? item.applyDiscount;
    if (explicitFlag !== undefined && explicitFlag !== null && explicitFlag !== '') {
      const flag = normalize(explicitFlag);
      if (['no', 'n', 'false', '0', 'non previsto'].includes(flag)) return false;
      if (['si', 'sì', 's', 'true', '1', 'previsto'].includes(flag)) return parseDiscount(item.discount ?? item.ribasso ?? item.percentualeRibasso) > 0;
    }
    return parseDiscount(item.discount ?? item.ribasso ?? item.percentualeRibasso) > 0;
  }

  PV.priceItemDiscount = (item) => discountIsEnabled(item)
    ? parseDiscount(item.discount ?? item.ribasso ?? item.percentualeRibasso)
    : 0;

  PV.priceItemNetPrice = (item) => {
    const contractPrice = Number(item?.contractPrice ?? item?.baseUnitPrice ?? item?.unitPrice) || 0;
    const discount = PV.priceItemDiscount(item);
    return discount > 0 ? PV.roundMoney(contractPrice * (1 - discount / 100)) : contractPrice;
  };

  function patchLine(row, item) {
    if (!row || !item) return;
    const contractPrice = Number(item.unitPrice) || 0;
    const discount = PV.priceItemDiscount(item);
    const netPrice = PV.priceItemNetPrice(item);
    row.dataset.contractPrice = String(contractPrice);
    row.dataset.discount = String(discount);
    row.dataset.discountEnabled = discount > 0 ? '1' : '0';
    const priceInput = row.querySelector('[data-pv-line-price], [data-cons-price]');
    if (priceInput) priceInput.value = String(netPrice);
  }

  const originalPopulateQuoteLine = PV.populateQuoteLine?.bind(PV);
  if (originalPopulateQuoteLine && !PV.__conditionalDiscountPopulatePatched) {
    PV.__conditionalDiscountPopulatePatched = true;
    PV.populateQuoteLine = (row, value) => {
      originalPopulateQuoteLine(row, value);
      const resolved = PV.resolvePriceItem?.(value);
      if (resolved?.item) patchLine(row, resolved.item);
      PV.updateQuoteTotals?.();
    };
  }

  const originalCollectQuoteLines = PV.collectQuoteLines?.bind(PV);
  if (originalCollectQuoteLines && !PV.__conditionalDiscountCollectPatched) {
    PV.__conditionalDiscountCollectPatched = true;
    PV.collectQuoteLines = () => originalCollectQuoteLines().map((line, index) => {
      const row = PV.page()?.querySelectorAll('[data-line-id]')?.[index];
      const selected = PV.resolvePriceItem?.(row?.querySelector('[data-pv-line-item]')?.value);
      const item = selected?.item;
      const contractPrice = Number(row?.dataset.contractPrice ?? item?.unitPrice ?? line.unitPrice) || 0;
      const discount = item ? PV.priceItemDiscount(item) : parseDiscount(row?.dataset.discount);
      const enabled = item ? discountIsEnabled(item) : row?.dataset.discountEnabled === '1' && discount > 0;
      const unitPrice = enabled ? PV.roundMoney(contractPrice * (1 - discount / 100)) : contractPrice;
      return {
        ...line,
        contractPrice,
        discount: enabled ? discount : 0,
        discountEnabled: enabled,
        unitPrice
      };
    });
  }

  if (M?.modelData && !M.__conditionalDiscountExportPatched) {
    M.__conditionalDiscountExportPatched = true;
    const originalModelData = M.modelData.bind(M);
    M.modelData = (doc, type) => {
      const normalizedLines = (doc?.lines || []).map((line) => {
        const priceList = PV.getPriceList?.(line.priceListId);
        const item = (priceList?.items || []).find((entry) => entry.id === line.priceItemId) || null;
        const contractPrice = Number(line.contractPrice ?? item?.unitPrice ?? line.unitPrice) || 0;
        const discount = item ? PV.priceItemDiscount(item) : (line.discountEnabled ? parseDiscount(line.discount) : 0);
        const enabled = item ? discountIsEnabled(item) : Boolean(line.discountEnabled && discount > 0);
        const unitPrice = enabled ? PV.roundMoney(contractPrice * (1 - discount / 100)) : contractPrice;
        return {
          ...line,
          contractPrice,
          discount: enabled ? discount : 0,
          discountEnabled: enabled,
          unitPrice
        };
      });
      return originalModelData({ ...doc, lines: normalizedLines }, type);
    };
  }

  window.HeraPreventiviConditionalDiscount = {
    parseDiscount,
    discountIsEnabled,
    version: '20260801a'
  };
})();
