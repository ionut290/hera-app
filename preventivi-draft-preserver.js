(() => {
  'use strict';

  const PV = window.HeraPreventivi;
  if (!PV || PV.__draftPreserverInstalled) return;
  PV.__draftPreserverInstalled = true;

  const FORM_SELECTOR = '[data-pv-quote-form], [data-cons-form]';
  const CONTROL_SELECTOR = 'input, textarea, select';

  function activeEditorForm() {
    const page = PV.page?.();
    if (!page || page.classList.contains('hidden')) return null;
    return page.querySelector(FORM_SELECTOR);
  }

  function controlKey(control, index) {
    return control.name
      || control.dataset.pvmField
      || control.dataset.consDesc
      || control.dataset.consCode
      || control.dataset.consUnit
      || control.dataset.consQty
      || control.dataset.consPrice
      || control.id
      || `control-${index}`;
  }

  function captureDraft(form) {
    if (!form) return null;
    const controls = [...form.querySelectorAll(CONTROL_SELECTOR)];
    const values = controls.map((control, index) => ({
      key: controlKey(control, index),
      type: control.type || control.tagName.toLowerCase(),
      value: control.value,
      checked: Boolean(control.checked),
      multipleValues: control.multiple ? [...control.selectedOptions].map((option) => option.value) : null
    }));
    const active = document.activeElement;
    const activeIndex = controls.indexOf(active);
    return {
      values,
      activeIndex,
      selectionStart: activeIndex >= 0 && typeof active.selectionStart === 'number' ? active.selectionStart : null,
      selectionEnd: activeIndex >= 0 && typeof active.selectionEnd === 'number' ? active.selectionEnd : null,
      scrollTop: PV.page?.()?.scrollTop || 0
    };
  }

  function restoreDraft(form, draft) {
    if (!form || !draft) return;
    let controls = [...form.querySelectorAll(CONTROL_SELECTOR)];
    const savedModel = draft.values.find((saved) => saved.key === 'modelId');
    const modelSelect = form.querySelector('[data-pvm-model-select]');
    if (savedModel && modelSelect && [...modelSelect.options].some((option) => option.value === savedModel.value)) {
      modelSelect.value = savedModel.value;
      const modelFields = {};
      draft.values.forEach((saved) => {
        if (saved.key && saved.key !== 'modelId' && !saved.key.startsWith('control-')) modelFields[saved.key] = saved.value;
      });
      window.HeraPreventiviModels?.renderDynamic?.(form, { modelId: savedModel.value, modelFields });
      form.dataset.pvdSignature = '';
      controls = [...form.querySelectorAll(CONTROL_SELECTOR)];
    }
    const grouped = new Map();
    controls.forEach((control, index) => {
      const key = controlKey(control, index);
      if (!grouped.has(key)) grouped.set(key, []);
      grouped.get(key).push(control);
    });

    draft.values.forEach((saved, index) => {
      const candidates = grouped.get(saved.key) || [];
      const control = candidates.shift() || (saved.key.startsWith('control-') ? controls[index] : null);
      if (!control) return;
      if (control.type === 'checkbox' || control.type === 'radio') {
        control.checked = saved.checked;
      } else if (control.multiple && Array.isArray(saved.multipleValues)) {
        [...control.options].forEach((option) => {
          option.selected = saved.multipleValues.includes(option.value);
        });
      } else {
        control.value = saved.value;
      }
    });

    const page = PV.page?.();
    if (page) page.scrollTop = draft.scrollTop;
    const focusControl = controls[draft.activeIndex];
    if (focusControl) {
      focusControl.focus({ preventScroll: true });
      if (draft.selectionStart !== null && typeof focusControl.setSelectionRange === 'function') {
        try { focusControl.setSelectionRange(draft.selectionStart, draft.selectionEnd); } catch (_) { /* tipo input non compatibile */ }
      }
    }
    form.dispatchEvent(new CustomEvent('hera:draft-restored', { bubbles: true }));
  }

  const originalRenderCurrentView = PV.renderCurrentView?.bind(PV);
  if (!originalRenderCurrentView) return;

  PV.renderCurrentView = (...args) => {
    const form = activeEditorForm();
    const draft = captureDraft(form);
    const editingBefore = {
      quote: PV.state.editingQuoteId,
      consuntivo: PV.state.editingConsuntivoId,
      view: PV.state.view
    };

    const result = originalRenderCurrentView(...args);

    const sameEditor = editingBefore.quote === PV.state.editingQuoteId
      && editingBefore.consuntivo === PV.state.editingConsuntivoId
      && editingBefore.view === PV.state.view;
    if (draft && sameEditor) {
      queueMicrotask(() => restoreDraft(activeEditorForm(), draft));
    }
    return result;
  };
})();
