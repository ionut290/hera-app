"use strict";

(() => {
  const ICONS = {
    MA: { icon: "🏗️", label: "Escavatore / movimento terra" },
    A: { icon: "🚛", label: "Camion" },
    T: { icon: "🚜", label: "Trattore grande" },
    R: { icon: "🚜", label: "Trattorino" },
    DEFAULT: { icon: "🛠️", label: "Attrezzatura" }
  };

  function classify(code) {
    const normalized = String(code || "").trim().toUpperCase();
    if (/^MA\d*/.test(normalized)) return ICONS.MA;
    if (/^A\d*/.test(normalized)) return ICONS.A;
    if (/^T\d*/.test(normalized)) return ICONS.T;
    if (/^R\d*/.test(normalized)) return ICONS.R;
    return ICONS.DEFAULT;
  }

  function installStyles() {
    if (document.getElementById("squadre-mezzi-pictograms-style")) return;
    const style = document.createElement("style");
    style.id = "squadre-mezzi-pictograms-style";
    style.textContent = `
      .squadra-mezzi-compact{display:flex;align-items:center;flex-wrap:wrap;gap:5px;margin:4px 0}
      .squadra-mezzi-compact-label{font-weight:800;font-size:.88rem;line-height:1;color:#163f35;margin-right:1px}
      .squadra-mezzo-mini{display:inline-flex;align-items:center;gap:2px;min-height:22px;padding:1px 6px;border:1px solid #b9edca;border-radius:999px;background:#f3fff7;color:#15713c;font-size:.74rem;font-weight:800;line-height:1;white-space:nowrap}
      .squadra-mezzo-mini-icon{font-size:.76rem;line-height:1}
      @media(max-width:480px){.squadra-mezzi-compact{gap:4px}.squadra-mezzi-compact-label{font-size:.82rem}.squadra-mezzo-mini{font-size:.7rem;min-height:20px;padding:1px 5px}.squadra-mezzo-mini-icon{font-size:.72rem}}
    `;
    document.head.appendChild(style);
  }

  function isMezziLabel(text) {
    return /^\s*(?:🚚|🚛)?\s*Mezzi\s*\d+\s*:/i.test(String(text || ""));
  }

  function getCodeFromElement(element) {
    return String(element?.textContent || "").trim();
  }

  function enhanceRow(row) {
    if (!(row instanceof HTMLElement) || row.dataset.mezziPictogramsDone === "1") return;
    const fullText = String(row.textContent || "").trim();
    if (!isMezziLabel(fullText)) return;

    const candidates = Array.from(row.querySelectorAll("button, .chip, .badge, span"))
      .filter((node) => {
        const code = getCodeFromElement(node);
        return /^(?:MA|A|T|R)?\d+[A-Z0-9-]*$/i.test(code) && !isMezziLabel(code);
      });

    if (!candidates.length) return;

    row.classList.add("squadra-mezzi-compact");
    row.dataset.mezziPictogramsDone = "1";

    const labelNode = Array.from(row.childNodes).find((node) =>
      node.nodeType === Node.TEXT_NODE && isMezziLabel(node.textContent)
    );
    if (labelNode) {
      const label = document.createElement("span");
      label.className = "squadra-mezzi-compact-label";
      label.textContent = String(labelNode.textContent || "").replace(/^\s*(?:🚚|🚛)?\s*/, "").trim();
      labelNode.replaceWith(label);
    } else {
      const first = row.firstElementChild;
      if (first && isMezziLabel(first.textContent)) {
        first.classList.add("squadra-mezzi-compact-label");
        first.textContent = String(first.textContent || "").replace(/^\s*(?:🚚|🚛)?\s*/, "").trim();
      }
    }

    candidates.forEach((node) => {
      const code = getCodeFromElement(node);
      const meta = classify(code);
      node.classList.add("squadra-mezzo-mini");
      node.setAttribute("title", meta.label);
      node.setAttribute("aria-label", `${meta.label} ${code}`);
      node.innerHTML = `<span class="squadra-mezzo-mini-icon" aria-hidden="true">${meta.icon}</span><span>${code}</span>`;
    });
  }

  function scan(root = document) {
    const elements = root instanceof HTMLElement ? [root, ...root.querySelectorAll("p, div, li")] : Array.from(document.querySelectorAll("p, div, li"));
    elements.forEach(enhanceRow);
  }

  installStyles();
  scan();

  const observer = new MutationObserver((mutations) => {
    mutations.forEach((mutation) => {
      mutation.addedNodes.forEach((node) => {
        if (node instanceof HTMLElement) scan(node);
      });
    });
  });
  observer.observe(document.body, { childList: true, subtree: true });
})();
