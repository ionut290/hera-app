(() => {
  'use strict';
  const M = window.HeraPreventiviModels;
  if (!M?.documentExportName) return;

  const text = (value) => String(value ?? '').trim();
  const normalize = (value) => text(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase();
  const previous = M.isDepurazioneOrInreteModel;

  M.isDepurazioneOrInreteModel = (model, doc) => {
    const modelIdentity = normalize([model?.name, model?.originalName, doc?.modelName].filter(Boolean).join(' '));
    if (modelIdentity.includes('STANDARD')) return false;
    if (modelIdentity.includes('DEPURAZIONE') || modelIdentity.includes('INRETE')) return true;
    return previous ? previous(model, doc) : false;
  };
})();
