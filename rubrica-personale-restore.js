(() => {
  'use strict';
  if (window.__vargaPersonaleRestore) return;
  window.__vargaPersonaleRestore = true;

  const text = value => String(value ?? '').trim();
  const lower = value => text(value).toLowerCase();
  const escapeHtml = value => text(value).replace(/[&<>"']/g, char => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[char]));

  function sourceRows() {
    const personale = Array.isArray(window.personaleRecords) ? window.personaleRecords : [];
    const utenti = Array.isArray(window.platformUsers) ? window.platformUsers : [];
    const rows = [];
    const seen = new Set();

    const add = (record, source) => {
      if (!record || typeof record !== 'object') return;
      const email = lower(record.email || record.emailAccessoApp || record.linkedUserEmail);
      const telefono = text(record.telefono || record.phone || record.cellulare || record.mobile || record.telefonoPersonale);
      const nome = text(record.displayName || record.nomeCompleto || `${text(record.nome)} ${text(record.cognome)}`) || email || telefono;
      if (!nome) return;
      const key = email || telefono.replace(/\D/g, '') || lower(nome);
      if (seen.has(key)) return;
      seen.add(key);
      rows.push({
        nome,
        telefono,
        email,
        ruolo: text(record.mansione || record.ruolo || record.role),
        azienda: text(record.azienda || record.company),
        foto: text(record.photoUrl || record.fotoUrl || record.photoURL),
        source
      });
    };

    personale.forEach(record => add(record, 'personale'));
    utenti.forEach(record => add(record, 'utente'));
    return rows.sort((a, b) => a.nome.localeCompare(b.nome, 'it', { sensitivity:'base' }));
  }

  function initials(nome) {
    return text(nome).split(/\s+/).slice(0, 2).map(part => part[0] || '').join('').toUpperCase() || '👤';
  }

  function installStyle() {
    if (document.getElementById('varga-personale-restore-style')) return;
    const style = document.createElement('style');
    style.id = 'varga-personale-restore-style';
    style.textContent = `
      .vpr-page{position:fixed;inset:0;z-index:17000;background:#eef4f2;overflow:auto;padding:calc(12px + env(safe-area-inset-top)) 14px 30px}
      .vpr-shell{max-width:760px;margin:auto}.vpr-head{display:flex;align-items:center;gap:10px}.vpr-head h2{flex:1;color:#173c35}
      .vpr-page button{min-height:46px;border:1px solid #bfd0cb;border-radius:14px;background:#fff;padding:8px 12px;font-weight:850;color:#075fae}
      .vpr-search{width:100%;min-height:54px;border:1px solid #bfd0cb;border-radius:16px;padding:0 14px;font-size:17px;margin:10px 0 14px}
      .vpr-list{display:grid;gap:10px}.vpr-card{display:flex;align-items:center;gap:12px;text-align:left;color:#173c35}
      .vpr-avatar{width:52px;height:52px;border-radius:50%;object-fit:cover;background:#dcebe7;display:grid;place-items:center;flex:none;font-weight:900}
      .vpr-main{min-width:0;flex:1}.vpr-main strong,.vpr-main small{display:block;overflow:hidden;text-overflow:ellipsis}
      .vpr-empty{text-align:center;color:#687a75;padding:36px 10px}
    `;
    document.head.appendChild(style);
  }

  function openPersonale() {
    document.querySelector('.vpr-page')?.remove();
    installStyle();
    const page = document.createElement('section');
    page.className = 'vpr-page';
    page.innerHTML = `<div class="vpr-shell"><div class="vpr-head"><button data-back>← INDIETRO</button><h2>👥 Personale</h2></div><input class="vpr-search" type="search" placeholder="Cerca nome, telefono, e-mail, ruolo o azienda…"><div class="vpr-list"></div></div>`;
    const search = page.querySelector('.vpr-search');
    const list = page.querySelector('.vpr-list');

    const render = () => {
      const query = lower(search.value);
      const rows = sourceRows().filter(row => [row.nome,row.telefono,row.email,row.ruolo,row.azienda].some(value => lower(value).includes(query)));
      list.innerHTML = rows.length ? rows.map(row => {
        const avatar = row.foto ? `<img class="vpr-avatar" src="${escapeHtml(row.foto)}" alt="Foto ${escapeHtml(row.nome)}">` : `<span class="vpr-avatar">${escapeHtml(initials(row.nome))}</span>`;
        const details = [row.telefono,row.email,row.ruolo,row.azienda].filter(Boolean).join(' · ');
        const call = row.telefono ? ` onclick="location.href='tel:${escapeHtml(row.telefono.replace(/[^\d+]/g,''))}'"` : '';
        return `<button class="vpr-card"${call}>${avatar}<span class="vpr-main"><strong>${escapeHtml(row.nome)}</strong><small>${escapeHtml(details || 'Personale')}</small></span></button>`;
      }).join('') : '<div class="vpr-empty">Elenco personale non ancora caricato. Chiudi e riapri questa sezione tra pochi secondi.</div>';
    };

    search.addEventListener('input', render);
    page.querySelector('[data-back]').addEventListener('click', () => page.remove());
    document.body.appendChild(page);
    render();
    window.setTimeout(render, 600);
    window.setTimeout(render, 1600);
  }

  window.openPersonaleRipristinato = openPersonale;
  document.addEventListener('click', event => {
    const element = event.target.closest('button,a,[role="button"]');
    if (!element || element.closest('.vpr-page')) return;
    const label = text(element.textContent).toUpperCase();
    if (label === 'PERSONALE' || label === 'RUBRICA' || label.includes('RUBRICA CONTATTI')) {
      event.preventDefault();
      event.stopImmediatePropagation();
      openPersonale();
    }
  }, true);
})();