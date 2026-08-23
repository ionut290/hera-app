(() => {
  'use strict';

  if (window.__heraSquadValidationPersonnelCompat) return;
  window.__heraSquadValidationPersonnelCompat = true;

  const normalize = (value) => String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLocaleLowerCase('it-IT')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const truthy = (value) => {
    if (value === true || value === 1) return true;
    return ['si', 'yes', 'true', '1', 'x', 'attivo', 'abilitato']
      .includes(normalize(value));
  };

  const toList = (value) => {
    if (Array.isArray(value)) return value.flatMap(toList);
    if (value && typeof value === 'object') {
      const preferred = value.id || value.commessaId || value.nome || value.name || value.label;
      return preferred ? [String(preferred).trim()] : [];
    }
    return String(value ?? '')
      .split(/[;,|\n]+/)
      .map((item) => item.trim())
      .filter(Boolean);
  };

  const personEnabledValues = (person = {}) => [
    person.enabledCommessaIds,
    person.authorizedCommessaIds,
    person.assignedCommessaIds,
    person.commessaIds,
    person.commesseAbilitate,
    person.commesse,
    person.allowedCommesse
  ].flatMap(toList);

  const findCommessaIdsInDom = (commessaName) => {
    const target = normalize(commessaName);
    if (!target || typeof document === 'undefined') return [];
    const ids = new Set();
    document.querySelectorAll('option').forEach((option) => {
      const label = normalize(option.textContent || option.label || '');
      if (!label || (label !== target && !label.includes(target) && !target.includes(label))) return;
      const value = String(option.value || '').trim();
      if (value) ids.add(value);
    });
    document.querySelectorAll('[data-commessa-id], [data-id-commessa]').forEach((element) => {
      const label = normalize(element.textContent || element.getAttribute('aria-label') || '');
      if (!label || (label !== target && !label.includes(target) && !target.includes(label))) return;
      const value = String(element.dataset.commessaId || element.dataset.idCommessa || '').trim();
      if (value) ids.add(value);
    });
    return [...ids];
  };

  const textFromTrainingValue = (value) => {
    if (Array.isArray(value)) return value.map(textFromTrainingValue).filter(Boolean).join('; ');
    if (value && typeof value === 'object') {
      const enabled = value.possiede ?? value.attivo ?? value.enabled ?? value.valido;
      if (enabled === false) return '';
      return [value.nome, value.name, value.titolo, value.label, value.corso, value.abilitazione]
        .filter(Boolean).join(' ');
    }
    return String(value ?? '');
  };

  const buildTrainingText = (person = {}) => [
    person.abilitazioni,
    person.corsi,
    person.formazione,
    person.training,
    person.qualifiche,
    person.certifications,
    person.mansione,
    person.qualifica,
    person.note
  ].map(textFromTrainingValue).filter(Boolean).join('; ');

  const hasTrainingFlag = (person, key) => {
    const sources = [person?.trainingRequirements, person?.requisitiSicurezza, person?.corsiSicurezza];
    return sources.some((source) => {
      if (!source || typeof source !== 'object') return false;
      return Object.entries(source).some(([name, value]) => normalize(name).includes(key) && truthy(value?.possiede ?? value));
    });
  };

  function install(attempt = 0) {
    const originalCommessaCheck = window.isPersonAbilitataForCommessa;
    const originalCourseNormalizer = window.normalizePersonCourses;

    if (typeof originalCommessaCheck !== 'function' || typeof originalCourseNormalizer !== 'function') {
      if (attempt < 100) window.setTimeout(() => install(attempt + 1), 100);
      return;
    }
    if (originalCommessaCheck.__personnelCompatibilityFix) return;

    const compatibleCommessaCheck = function compatibleCommessaCheck(person, commessaName) {
      if (!person) return false;
      if (
        truthy(person.abilitatoTutteCommesse)
        || truthy(person.tutteLeCommesse)
        || truthy(person.allCommesseEnabled)
        || truthy(person.accessoTutteCommesse)
      ) return true;

      try {
        if (originalCommessaCheck.call(this, person, commessaName)) return true;
      } catch (error) {
        console.warn('Controllo commessa originale non riuscito, uso compatibilità.', error);
      }

      const enabled = personEnabledValues(person);
      if (!enabled.length) return false;
      const enabledKeys = new Set(enabled.map(normalize));
      const target = normalize(commessaName);
      if (target && enabledKeys.has(target)) return true;

      const matchingIds = findCommessaIdsInDom(commessaName);
      if (matchingIds.some((id) => enabled.includes(id) || enabledKeys.has(normalize(id)))) return true;

      const activeIds = new Set();
      document.querySelectorAll('select option').forEach((option) => {
        const value = String(option.value || '').trim();
        if (value && value.length >= 10) activeIds.add(value);
      });
      if (activeIds.size >= 2 && [...activeIds].every((id) => enabled.includes(id))) return true;

      return false;
    };
    compatibleCommessaCheck.__personnelCompatibilityFix = true;
    compatibleCommessaCheck.__original = originalCommessaCheck;
    window.isPersonAbilitataForCommessa = compatibleCommessaCheck;

    const compatibleCourseNormalizer = function compatibleCourseNormalizer(person = {}) {
      let result = {};
      try {
        result = originalCourseNormalizer.call(this, person) || {};
      } catch (error) {
        console.warn('Normalizzazione corsi originale non riuscita, uso compatibilità.', error);
      }

      const ensure = (key) => {
        if (!result[key] || typeof result[key] !== 'object') result[key] = { possiede: false };
        return result[key];
      };
      const trainingText = normalize(buildTrainingText(person));
      const checks = {
        'primo soccorso': /primo soccorso|soccorso emergenza/,
        'antincendio': /antincendio|rischio medio/,
        'preposto': /preposto/,
        'atex': /\batex\b/
      };

      Object.entries(checks).forEach(([key, pattern]) => {
        if (ensure(key).possiede) return;
        if (pattern.test(trainingText) || hasTrainingFlag(person, key)) {
          result[key] = { ...ensure(key), possiede: true };
        }
      });
      return result;
    };
    compatibleCourseNormalizer.__personnelCompatibilityFix = true;
    compatibleCourseNormalizer.__original = originalCourseNormalizer;
    window.normalizePersonCourses = compatibleCourseNormalizer;

    console.info('Compatibilità controllo squadra installata: commesse per ID e abilitazioni importate riconosciute.');
  }

  function getDatalistValues(input) {
    const listId = String(input.getAttribute('list') || '').trim();
    const list = listId ? document.getElementById(listId) : null;
    if (!list) return [];
    return Array.from(list.querySelectorAll('option'))
      .map((option) => String(option.value || option.textContent || '').trim())
      .filter(Boolean);
  }

  function isFreeSquadEntry(input) {
    if (!(input instanceof HTMLInputElement) || !input.classList.contains('squadra-multi-entry-input')) return false;
    const container = input.closest('.squadra-personale-list, .squadra-mezzi-list');
    if (!container) return false;
    const value = String(input.value || '').trim();
    if (!value) return false;
    const normalizedValue = normalize(value);
    const known = getDatalistValues(input).some((item) => normalize(item) === normalizedValue);
    return !known;
  }

  function decorateFreeInput(input) {
    if (!(input instanceof HTMLInputElement) || input.dataset.freeSquadEntryReady === '1') return;
    if (!input.closest('.squadra-personale-list, .squadra-mezzi-list')) return;
    input.dataset.freeSquadEntryReady = '1';
    const oldPlaceholder = String(input.placeholder || '').trim();
    if (!/scrivi/i.test(oldPlaceholder)) {
      input.placeholder = oldPlaceholder ? `${oldPlaceholder} oppure scrivi liberamente` : 'Seleziona oppure scrivi liberamente';
    }
  }

  function decorateAllFreeInputs(root = document) {
    root.querySelectorAll?.('.squadra-personale-list .squadra-multi-entry-input, .squadra-mezzi-list .squadra-multi-entry-input')
      .forEach(decorateFreeInput);
  }

  // Il core legge direttamente il valore testuale di questi input quando salva la squadra.
  // Per una voce non presente nei registri impediamo soltanto al vecchio handler blur
  // di cancellarla/validarla come record esistente. Le voci già registrate continuano
  // invece a seguire il comportamento precedente (autocomplete e controlli disponibilità).
  document.addEventListener('blur', (event) => {
    const input = event.target;
    if (!isFreeSquadEntry(input)) return;
    const value = String(input.value || '').trim();
    input.value = value;
    input.dataset.freeSquadEntry = '1';
    event.stopImmediatePropagation();
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }, true);

  const observer = new MutationObserver((mutations) => {
    mutations.forEach((mutation) => mutation.addedNodes.forEach((node) => {
      if (!(node instanceof Element)) return;
      if (node.matches?.('.squadra-multi-entry-input')) decorateFreeInput(node);
      decorateAllFreeInputs(node);
    }));
  });

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => decorateAllFreeInputs(), { once: true });
  } else {
    decorateAllFreeInputs();
  }
  observer.observe(document.documentElement, { childList: true, subtree: true });

  window.HeraSquadFreePersonnelVehicles = {
    installed: true,
    version: '1.0.0',
    refresh: decorateAllFreeInputs
  };

  install();
})();
