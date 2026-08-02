(() => {
  'use strict';

  if (window.__vargaRubricaVcardShareLoaded) return;
  window.__vargaRubricaVcardShareLoaded = true;

  const text = (value) => String(value ?? '').trim();

  function escapeVcard(value) {
    return text(value)
      .replace(/\\/g, '\\\\')
      .replace(/\r?\n/g, '\\n')
      .replace(/,/g, '\\,')
      .replace(/;/g, '\\;');
  }

  function safeFilename(value) {
    const cleaned = text(value)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]+/gi, '-')
      .replace(/^-+|-+$/g, '')
      .toLowerCase();
    return `${cleaned || 'contatto'}.vcf`;
  }

  function contactFromDialog(dialog) {
    const name = text(dialog.querySelector('h3')?.textContent) || 'Contatto';
    const phoneHref = dialog.querySelector('a[href^="tel:"]')?.getAttribute('href') || '';
    const emailHref = dialog.querySelector('a[href^="mailto:"]')?.getAttribute('href') || '';
    const phone = decodeURIComponent(phoneHref.replace(/^tel:/i, ''));
    const email = decodeURIComponent(emailHref.replace(/^mailto:/i, '').split('?')[0]);
    return { name, phone, email };
  }

  function createVcard(contact) {
    const lines = [
      'BEGIN:VCARD',
      'VERSION:3.0',
      `FN:${escapeVcard(contact.name)}`,
      `N:${escapeVcard(contact.name)};;;;`
    ];
    if (contact.phone) lines.push(`TEL;TYPE=CELL:${escapeVcard(contact.phone)}`);
    if (contact.email) lines.push(`EMAIL;TYPE=INTERNET:${escapeVcard(contact.email)}`);
    lines.push('PRODID:-//VARGA CANTIERI//Rubrica//IT', 'END:VCARD');
    return `${lines.join('\r\n')}\r\n`;
  }

  function downloadFile(file) {
    const url = URL.createObjectURL(file);
    const link = document.createElement('a');
    link.href = url;
    link.download = file.name;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  async function shareContact(contact, button) {
    if (!contact.phone && !contact.email) {
      alert('Il contatto non contiene telefono o e-mail.');
      return;
    }

    const file = new File([createVcard(contact)], safeFilename(contact.name), {
      type: 'text/vcard;charset=utf-8'
    });
    const originalText = button.textContent;
    button.disabled = true;
    button.textContent = 'PREPARO IL CONTATTO…';

    try {
      const payload = {
        title: `Contatto ${contact.name}`,
        text: `Contatto ${contact.name}. Apri il file per aggiungerlo alla rubrica del telefono.`,
        files: [file]
      };
      if (navigator.share && (!navigator.canShare || navigator.canShare({ files: [file] }))) {
        await navigator.share(payload);
        return;
      }

      downloadFile(file);
      alert('Il telefono non supporta la condivisione diretta del file. Il contatto .VCF è stato scaricato: allegalo in WhatsApp. Chi lo riceve potrà aggiungerlo alla rubrica.');
    } catch (error) {
      if (error?.name !== 'AbortError') {
        console.warn('Condivisione vCard non riuscita:', error);
        downloadFile(file);
        alert('Condivisione diretta non riuscita. Il contatto .VCF è stato scaricato: puoi allegarlo manualmente in WhatsApp.');
      }
    } finally {
      button.disabled = false;
      button.textContent = originalText;
    }
  }

  function enhanceDialog(dialog) {
    if (!(dialog instanceof Element) || dialog.dataset.vcardShareBound === '1') return;
    const closeButton = dialog.querySelector('[data-close]');
    const name = text(dialog.querySelector('h3')?.textContent);
    const hasContactAction = dialog.querySelector('a[href^="tel:"], a[href^="mailto:"]');
    if (!name || !hasContactAction || !closeButton) return;

    dialog.dataset.vcardShareBound = '1';
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'rubrica-v2-action';
    button.textContent = '📤 CONDIVIDI CONTATTO';
    button.setAttribute('aria-label', `Condividi ${name} come contatto vCard`);
    button.addEventListener('click', () => shareContact(contactFromDialog(dialog), button));
    closeButton.before(button);
  }

  function scan(root = document) {
    root.querySelectorAll?.('.rubrica-v2-dialog').forEach(enhanceDialog);
    if (root.matches?.('.rubrica-v2-dialog')) enhanceDialog(root);
  }

  function init() {
    scan();
    const observer = new MutationObserver((mutations) => {
      mutations.forEach((mutation) => mutation.addedNodes.forEach((node) => {
        if (node.nodeType === Node.ELEMENT_NODE) scan(node);
      }));
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();
