(() => {
  'use strict';

  const M = window.HeraPreventiviModels;
  const PV = M?.PV;
  if (!M || !PV) return;

  const text = (value) => String(value ?? '').trim();
  const normalized = (value) => text(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase();
  const compactDate = (value) => {
    const raw = text(value);
    const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return `${iso[1]}${iso[2]}${iso[3]}`;
    const ita = raw.match(/^(\d{2})[\/.\-](\d{2})[\/.\-](\d{4})/);
    if (ita) return `${ita[3]}${ita[2]}${ita[1]}`;
    const date = raw ? new Date(raw) : null;
    if (date && !Number.isNaN(date.getTime())) {
      return `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}`;
    }
    return '';
  };
  const cleanPart = (value) => text(value)
    .replace(/[\\/:*?"<>|]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const safeFileBase = (value) => cleanPart(value).replace(/[. ]+$/g, '') || 'DOCUMENTO';
  const safeSheetName = (value) => cleanPart(value).replace(/[\[\]]/g, ' ').slice(0, 31).trim() || 'DOCUMENTO';

  M.isDepurazioneOrInreteModel = (model, doc) => {
    const haystack = normalized([
      model?.name,
      model?.fileName,
      doc?.modelName,
      doc?.commessaName,
      doc?.commessaCode
    ].filter(Boolean).join(' '));
    return haystack.includes('DEPURAZIONE') || haystack.includes('INRETE');
  };

  M.documentExecutionDate = (doc, type) => {
    if (type === 'consuntivo') {
      return doc?.executionDate || doc?.endDate || doc?.startDate || doc?.date || '';
    }
    return doc?.executionDate || doc?.date || doc?.startDate || doc?.endDate || '';
  };

  M.documentExportName = (doc, type, model = M.getModel?.(doc?.modelId)) => {
    const tipo = type === 'consuntivo' ? 'CONSUNTIVO' : 'PREVENTIVO';
    const data = compactDate(M.documentExecutionDate(doc, type)) || compactDate(new Date());
    const specific = M.isDepurazioneOrInreteModel(model, doc);
    const subject = specific
      ? (doc?.plantName || doc?.denominazioneImpianto || doc?.plantSap || 'IMPIANTO')
      : (doc?.clientName || doc?.committente || doc?.requester || 'CLIENTE');
    return safeFileBase(`${tipo} ${data} ${cleanPart(subject)}`);
  };

  M.documentSheetName = (doc, type, model = M.getModel?.(doc?.modelId)) =>
    safeSheetName(M.documentExportName(doc, type, model));

  const originalExportOriginal = M.exportOriginal?.bind(M);
  if (originalExportOriginal) {
    M.exportOriginal = async (doc, type) => {
      M.runtime.currentExportType = type;
      try {
        return await originalExportOriginal(doc, type);
      } finally {
        M.runtime.currentExportType = '';
      }
    };
  }

  const originalExportSheet = M.exportSheet?.bind(M);
  if (originalExportSheet) {
    M.exportSheet = async (model, stored, data, doc) => {
      if (!window.XLSX) throw new Error('Motore fogli non disponibile.');
      const type = M.runtime.currentExportType || (doc?.type === 'consuntivo' ? 'consuntivo' : 'preventivo');
      const workbook = XLSX.read(stored.buffer, { type: 'array', cellStyles: true });
      workbook.SheetNames.forEach((name) => {
        const sheet = workbook.Sheets[name];
        Object.keys(sheet).filter((address) => !address.startsWith('!')).forEach((address) => {
          const cell = sheet[address];
          if (typeof cell?.v === 'string') {
            const value = M.replaceTokens(cell.v, data);
            cell.v = value;
            cell.w = value;
          }
        });
      });
      if (workbook.SheetNames.length) {
        const oldName = workbook.SheetNames[0];
        const newName = M.documentSheetName(doc, type, model);
        if (newName && newName !== oldName) {
          workbook.Sheets[newName] = workbook.Sheets[oldName];
          delete workbook.Sheets[oldName];
          workbook.SheetNames[0] = newName;
        }
      }
      const output = XLSX.write(workbook, {
        type: 'array',
        bookType: model.format === 'ods' ? 'ods' : 'xlsx',
        cellStyles: true
      });
      const mime = model.format === 'ods'
        ? 'application/vnd.oasis.opendocument.spreadsheet'
        : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      const fileName = `${M.documentExportName(doc, type, model)}.${model.format}`;
      M.download(new Blob([output], { type: mime }), fileName);
    };
  }

  const wrapNamedExporter = (name, extension) => {
    const original = M[name]?.bind(M);
    if (!original) return;
    M[name] = async (...args) => {
      const [model, stored, data, doc] = args;
      const type = M.runtime.currentExportType || (doc?.type === 'consuntivo' ? 'consuntivo' : 'preventivo');
      const originalDownload = M.download;
      M.download = (blob) => originalDownload(blob, `${M.documentExportName(doc, type, model)}.${extension || model?.format || 'pdf'}`);
      try {
        return await original(...args);
      } finally {
        M.download = originalDownload;
      }
    };
  };

  wrapNamedExporter('exportText');
  wrapNamedExporter('exportDocx', 'docx');
  wrapNamedExporter('exportOdt', 'odt');
  wrapNamedExporter('exportFillablePdf', 'pdf');

  const originalExportPdf = M.exportPdf?.bind(M);
  if (originalExportPdf) {
    M.exportPdf = async (doc, type) => {
      const originalSave = window.jspdf?.jsPDF?.API?.save;
      M.runtime.currentExportType = type;
      try {
        const holder = document.createElement('div');
        holder.className = 'pvm-offscreen';
        holder.innerHTML = M.previewHtml(doc, type);
        document.body.appendChild(holder);
        try {
          if (!window.html2canvas || !window.jspdf?.jsPDF) throw new Error('Generatore PDF non disponibile.');
          const canvas = await html2canvas(holder, { scale: 1.5, backgroundColor: '#ffffff' });
          const pdf = new jspdf.jsPDF('p', 'mm', 'a4');
          const width = 190;
          const height = canvas.height * 190 / canvas.width;
          const page = 277;
          const img = canvas.toDataURL('image/jpeg', 0.92);
          let y = 10;
          let left = height;
          pdf.addImage(img, 'JPEG', 10, y, width, height);
          left -= page;
          while (left > 0) {
            y = 10 - left;
            pdf.addPage();
            pdf.addImage(img, 'JPEG', 10, y, width, height);
            left -= page;
          }
          pdf.save(`${M.documentExportName(doc, type)}.pdf`);
          PV.setFeedback('PDF generato.', 'success');
        } catch (error) {
          PV.setFeedback(error.message, 'error');
        } finally {
          holder.remove();
        }
      } finally {
        M.runtime.currentExportType = '';
        void originalSave;
      }
    };
  }
})();
