(()=>{'use strict';
const P=window.HeraPreventivi;if(!P)throw Error('Preventivi core non caricato.');
const C=v=>String(v??'').trim().replace(/\s+/g,' '),N=v=>P.normalizeText(C(v)),F=(o,k)=>{for(const x of k)if(C(o?.[x]))return o[x];return''},V=v=>Array.isArray(v)?v:v instanceof Map?[...v.values()]:v&&typeof v==='object'?Object.values(v):[];
const NK=['impianti','plants','sites','impiantiDaFare','impiantiFatti','daFare','fatti','items','children','records','data'],SK=['idSap','ID SAP','ID_SAP','id_sap','sap','sapId','codiceSap','codiceImpianto'],PN=['denominazione','Denominazione Impianto','denominazioneImpianto','nomeImpianto','descrizioneImpianto','descrizione','nome','name','impianto','siteName'];
function reg(){const cs=new Map,ps=new Map,seen=new WeakSet,addC=(r,id='')=>{const x=C(F(r,['id','uid','commessaId','projectId','codiceCommessa','codice','numeroCommessa'])||id),name=C(F(r,['nome','name','denominazione','titolo','commessa','descrizione']))||(x?`Commessa ${x}`:'');if(!x&&!name)return null;const c={raw:r,id:x||`c-${N(name).replace(/\s/g,'-')}`,name,code:C(F(r,['codiceCommessa','codice','code','numeroCommessa'])),client:C(F(r,['cliente','committente','clientName','ragioneSociale'])),contract:C(F(r,['numeroContratto','contratto','contractNumber'])),requester:C(F(r,['richiedente','referente','referenteCliente']))};cs.set(c.id,c);return c},addP=(r,cid='')=>{const sap=C(F(r,SK)),name=C(F(r,PN));if(!name&&!sap)return;const p={raw:r,id:C(F(r,['id','uid','impiantoId','plantId',...SK])||sap||name),commessaId:C(F(r,['commessaId','projectId','idCommessa','codiceCommessa','commessa','commessa_id'])||cid),name:name||`Impianto SAP ${sap}`,sap,address:C(F(r,['indirizzo','Descrizione via','descrizioneVia','via','address','ubicazione'])),city:C(F(r,['comune','Comune','city','municipality'])),type:C(F(r,['tipologia','Tipologia impianto','tipo','type']))};ps.set(p.id,p)},walk=(v,cid='',d=0)=>{if(!v||d>7)return;if(typeof v==='object'){if(seen.has(v))return;seen.add(v)}if(Array.isArray(v)||v instanceof Map)return V(v).forEach(x=>walk(x,cid,d+1));if(typeof v!=='object')return;const isP=!!F(v,[...SK,'Tipologia impianto','tipologia','impiantoId','plantId','denominazioneImpianto','nomeImpianto']);let next=cid;if(isP)addP(v,cid);else if(F(v,['codiceCommessa','numeroCommessa','numeroContratto','committente','commessaId','nome','name','denominazione','titolo']))next=addC(v)?.id||cid;NK.forEach(k=>v[k]&&walk(v[k],next,d+1))};
const by=window.commesseById;if(by&&typeof by==='object')(by instanceof Map?[...by.entries()]:Object.entries(by)).forEach(([id,r])=>{const c=addC(r,id);NK.forEach(k=>r?.[k]&&walk(r[k],c?.id||id,1))});['commesse','projects','cantieri','impianti','impiantiById','impiantiDaFare','impiantiFatti','allImpianti','plants','sites'].forEach(k=>walk(window[k]));return{commesse:[...cs.values()],plants:[...ps.values()]}}
const label=p=>`${p.name} — ID SAP ${p.sap||'—'}`,match=(p,c)=>{if(!c)return true;const s=N(p.commessaId),t=[c.id,c.code,c.name].map(N).filter(Boolean);return!s||t.some(x=>x===s||x.includes(s)||s.includes(x))},set=(f,n,v)=>{const x=f?.elements?.namedItem(n);if(x&&'value'in x)x.value=v??''};
function fillC(f,c){if(!c)return;set(f,'commessaCode',c.code);if(!C(f.elements?.clientName?.value))set(f,'clientName',c.client);if(!C(f.elements?.contractNumber?.value))set(f,'contractNumber',c.contract);if(!C(f.elements?.requester?.value))set(f,'requester',c.requester)}
function fillP(f,p){set(f,'plantId',p.id);set(f,'plantSap',p.sap);set(f,'workLocation',[p.address,p.city].filter(Boolean).join(', '));set(f,'city',p.city);set(f,'plantType',p.type)}
function commesse(s){const old=C(s.value),list=reg().commesse.sort((a,b)=>a.name.localeCompare(b.name,'it')),sig=list.map(x=>`${x.id}:${x.name}:${x.code}`).join('|');if(s.dataset.rs===sig)return;s.innerHTML='<option value="">Seleziona commessa</option>'+list.map(x=>`<option value="${P.escapeHtml(x.id)}" ${x.id===old?'selected':''}>${P.escapeHtml(x.name)}${x.code?` — ${P.escapeHtml(x.code)}`:''}</option>`).join('');s.dataset.rs=sig}
function plantSelect(s){const f=s.closest('form'),id=C(s.value),p=reg().plants.find(x=>x.id===id),w=document.createElement('div');if(!f)return;w.className='pv-plant-search';w.innerHTML=`<input type="search" data-doc-plant-search autocomplete="off" placeholder="Cerca nome impianto o ID SAP" value="${P.escapeHtml(p?label(p):'')}" ${s.required?'required':''}><input type="hidden" name="plantId" data-doc-plant-id value="${P.escapeHtml(id)}"><div class="pv-plant-results hidden" data-doc-plant-results role="listbox"></div>`;const q=w.querySelector('[data-doc-plant-search]');if(q?.required&&!id)q.setCustomValidity('Seleziona un impianto dai risultati.');s.replaceWith(w)}
function decorate(){document.querySelectorAll(`#${P.pageId} [data-doc-commessa]`).forEach(commesse);document.querySelectorAll(`#${P.pageId} select[data-doc-plant]`).forEach(plantSelect)}
function results(q){const f=q.closest('form'),box=q.parentElement.querySelector('[data-doc-plant-results]'),cid=f.querySelector('[data-doc-commessa]')?.value,c=reg().commesse.find(x=>x.id===cid),text=N(q.value),list=reg().plants.filter(p=>match(p,c)&&(!text||N(`${p.name} ${p.sap}`).includes(text))).sort((a,b)=>label(a).localeCompare(label(b),'it')).slice(0,40);box.innerHTML=list.length?list.map(p=>`<button type="button" class="pv-plant-result" data-doc-plant-choice="${P.escapeHtml(p.id)}"><strong>${P.escapeHtml(p.name)}</strong><span>ID SAP ${P.escapeHtml(p.sap||'—')}</span></button>`).join(''):'<p class="pv-muted">Nessun impianto trovato. Cerca per nome oppure ID SAP.</p>';box.classList.remove('hidden')}
document.addEventListener('focusin',e=>{const q=e.target.closest('[data-doc-plant-search]');if(q)results(q)},true);document.addEventListener('input',e=>{const q=e.target.closest('[data-doc-plant-search]');if(!q)return;q.parentElement.querySelector('[data-doc-plant-id]').value='';if(q.required)q.setCustomValidity('Seleziona un impianto dai risultati.');results(q)},true);document.addEventListener('change',e=>{const s=e.target.closest('[data-doc-commessa]');if(!s)return;const f=s.closest('form'),c=reg().commesse.find(x=>x.id===s.value);fillC(f,c);const q=f.querySelector('[data-doc-plant-search]');if(q){q.value='';if(q.required)q.setCustomValidity('Seleziona un impianto dai risultati.')}const h=f.querySelector('[data-doc-plant-id]');if(h)h.value=''},true);document.addEventListener('click',e=>{const b=e.target.closest('[data-doc-plant-choice]');if(!b)return;e.preventDefault();const f=b.closest('form'),w=b.closest('.pv-plant-search'),p=reg().plants.find(x=>x.id===b.dataset.docPlantChoice);if(!p)return;const q=w.querySelector('[data-doc-plant-search]');q.value=label(p);q.setCustomValidity('');w.querySelector('[data-doc-plant-id]').value=p.id;fillP(f,p);w.querySelector('[data-doc-plant-results]').classList.add('hidden')},true);
const enrich=(r,fd)=>{if(!r)return r;const d=reg(),cid=C(fd?.get?.('commessaId')??r.commessaId),pid=C(fd?.get?.('plantId')??r.plantId),c=d.commesse.find(x=>x.id===cid),p=d.plants.find(x=>x.id===pid);return Object.assign(r,{commessaId:cid,commessaName:c?.name||r.commessaName||'',commessaCode:C(fd?.get?.('commessaCode')??r.commessaCode??c?.code),plantId:pid,plantName:p?.name||r.plantName||'',plantSap:C(fd?.get?.('plantSap')??r.plantSap??p?.sap),workLocation:C(fd?.get?.('workLocation')??r.workLocation??[p?.address,p?.city].filter(Boolean).join(', ')),city:C(fd?.get?.('city')??r.city??p?.city),plantType:C(fd?.get?.('plantType')??r.plantType??p?.type)})};
const sr=P.saveRemote?.bind(P);if(sr)P.saveRemote=(col,r)=>{if(r&&(col===P.collections.consuntivi||col===P.collections.quotes)){enrich(r);const a=col===P.collections.consuntivi?P.state.consuntivi:P.state.quotes,x=(a||[]).find(v=>v.id===r.id);if(x)Object.assign(x,r);P.persistLocal?.()}return sr(col,r)};const sq=P.saveQuote?.bind(P);if(sq)P.saveQuote=async f=>{const fd=new FormData(f),id=P.state.editingQuoteId,num=C(fd.get('number')),out=await sq(f),r=id&&id!=='new'?P.getQuote(id):[...(P.state.quotes||[])].reverse().find(x=>C(x.number)===num);if(r){enrich(r,fd);r.syncPending=true;P.persistLocal?.();P.scheduleSync?.()}return out};const M=window.HeraPreventiviModels;if(M?.collectDraft){const cd=M.collectDraft.bind(M);M.collectDraft=t=>{const f=P.page()?.querySelector(t==='consuntivo'?'[data-cons-form]':'[data-pv-quote-form]');return enrich(cd(t),f?new FormData(f):null)}}
if(!document.querySelector('style[data-registry-search]')){const s=document.createElement('style');s.dataset.registrySearch='1';s.textContent='.pv-plant-search{position:relative;display:grid;gap:6px}.pv-plant-results{position:absolute;z-index:120;top:calc(100% + 4px);left:0;right:0;max-height:55vh;overflow:auto;background:#fff;border:1px solid #cbd5e1;border-radius:12px;box-shadow:0 18px 40px #0f172e33;padding:6px}.pv-plant-results.hidden{display:none}.pv-plant-result{display:flex;width:100%;justify-content:space-between;gap:12px;padding:11px;border:0;border-bottom:1px solid #e2e8f0;background:#fff;text-align:left}.pv-plant-result:hover{background:#eef7f1}.pv-plant-result strong{white-space:normal}.pv-plant-result span{font-size:.82rem;font-weight:700;color:#475569}';document.head.appendChild(s)}
const o=new MutationObserver(decorate),start=()=>{decorate();o.observe(document.body,{childList:true,subtree:true})};document.readyState==='loading'?document.addEventListener('DOMContentLoaded',start,{once:true}):start();window.HeraPreventiviRegistry={registry:reg,plantLabel:label};})();
(() => {
  'use strict';
  const M = window.HeraPreventiviModels;
  const PV = M?.PV;
  if (!M || !PV) throw new Error('Preventivi Models non caricato.');

  const DOCX_LIB = 'https://cdn.jsdelivr.net/npm/docxtemplater@3.51.0/build/docxtemplater.js';
  const TOKEN_PATTERN = /\{\{\s*([^{}]+?)\s*\}\}|\[\[\s*([^\[\]]+?)\s*\]\]/g;
  const xmlEscape = (value) => String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
  const originalOutputName = (model) => {
    const extension = String(model.format || M.extension(model.originalName)).toLowerCase();
    const original = String(model.originalName || model.name || 'modello').replace(/[\\/:*?"<>|]/g, '_');
    const base = original.replace(new RegExp(`\\.${extension}$`, 'i'), '');
    return `${base}-compilato.${extension}`;
  };
  const hasRealTokens = (model) => (model.fields || []).some((field) => field.token || field.pdfFieldName);
  const lookup = (data, rawKey) => {
    const flat = M.flatten(data);
    const key = M.slug(rawKey);
    return data[key] ?? flat[rawKey] ?? flat[key] ?? '';
  };
  const replaceCounted = (text, data, escape = false) => {
    let count = 0;
    const output = String(text ?? '').replace(TOKEN_PATTERN, (token, a, b) => {
      const raw = String(a || b || '').trim();
      if (!raw || /^[#/^]/.test(raw)) return token;
      count += 1;
      const value = lookup(data, raw);
      const rendered = Array.isArray(value)
        ? value.map((row) => Object.values(row || {}).join(' | ')).join('\n')
        : String(value ?? '');
      return escape ? xmlEscape(rendered) : rendered;
    });
    return { output, count };
  };
  const ensureCompilable = (model) => {
    if (!hasRealTokens(model)) {
      throw new Error(`Il modello originale “${model.originalName || model.name}” non contiene segnaposto o campi compilabili. Aggiungi campi come {{cliente}}, {{impianto}}, {{id_sap}} oppure campi modulo PDF.`);
    }
  };
  const downloadExact = (model, blob) => M.download(blob, originalOutputName(model));

  M.exportText = async (model, stored, data) => {
    ensureCompilable(model);
    const result = replaceCounted(await stored.blob.text(), data, false);
    if (!result.count) throw new Error('Nessun campo del file originale è stato compilato.');
    const mime = ['html', 'htm'].includes(model.format) ? 'text/html'
      : model.format === 'json' ? 'application/json'
        : model.format === 'rtf' ? 'application/rtf' : 'text/plain';
    downloadExact(model, new Blob([result.output], { type: `${mime};charset=utf-8` }));
  };

  M.exportDocx = async (model, stored, data) => {
    ensureCompilable(model);
    await M.ensureScript('PizZip', M.constants.ZIP_LIB, 'pizzip');
    const Docxtemplater = await M.ensureScript(['docxtemplater', 'Docxtemplater'], DOCX_LIB, 'docxtemplater');
    if (!Docxtemplater) throw new Error('Motore DOCX non disponibile.');
    const instance = new Docxtemplater(new PizZip(stored.buffer), {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: '{{', end: '}}' },
      nullGetter: () => ''
    });
    instance.render(data);
    const blob = instance.getZip().generate({
      type: 'blob',
      mimeType: stored.type || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    downloadExact(model, blob);
  };

  const exportZipTemplate = async (model, stored, data, paths, mime) => {
    ensureCompilable(model);
    await M.ensureScript('PizZip', M.constants.ZIP_LIB, 'pizzip');
    const zip = new PizZip(stored.buffer);
    let replacements = 0;
    Object.keys(zip.files).forEach((name) => {
      if (!paths.some((pattern) => pattern.test(name))) return;
      const entry = zip.file(name);
      if (!entry) return;
      const result = replaceCounted(entry.asText(), data, true);
      if (result.count) {
        replacements += result.count;
        zip.file(name, result.output);
      }
    });
    if (!replacements) throw new Error('Nel file originale non sono stati trovati segnaposto compilabili.');
    downloadExact(model, zip.generate({ type: 'blob', mimeType: stored.type || mime }));
  };

  M.exportSheet = (model, stored, data) => exportZipTemplate(
    model,
    stored,
    data,
    [/^xl\/sharedStrings\.xml$/i, /^xl\/worksheets\/.*\.xml$/i, /^content\.xml$/i],
    model.format === 'ods'
      ? 'application/vnd.oasis.opendocument.spreadsheet'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );
  M.exportOdt = (model, stored, data) => exportZipTemplate(
    model,
    stored,
    data,
    [/^content\.xml$/i, /^styles\.xml$/i],
    'application/vnd.oasis.opendocument.text'
  );

  M.exportFillablePdf = async (model, stored, data) => {
    await M.ensureScript('PDFLib', M.constants.PDF_LIB, 'pdf-lib');
    const pdf = await PDFLib.PDFDocument.load(stored.buffer, { ignoreEncryption: true });
    const form = pdf.getForm();
    const fields = form.getFields();
    if (!fields.length) throw new Error('Il PDF originale non contiene campi modulo compilabili. Il file non è stato sostituito con un PDF simile.');
    const flat = M.flatten(data);
    let filled = 0;
    fields.forEach((pdfField) => {
      const name = pdfField.getName();
      const mapped = (model.fields || []).find((field) => field.pdfFieldName === name || M.slug(field.key) === M.slug(name));
      const value = flat[mapped?.source] ?? data[mapped?.key] ?? data[M.slug(name)];
      if (value === undefined || value === null || value === '') return;
      try {
        if (typeof pdfField.setText === 'function') pdfField.setText(String(value));
        else if (typeof pdfField.check === 'function' && Boolean(value)) pdfField.check();
        else return;
        filled += 1;
      } catch (error) {
        console.warn(`Campo PDF non compilato: ${name}`, error);
      }
    });
    if (!filled) throw new Error('Nessun campo del PDF originale corrisponde ai dati del documento.');
    downloadExact(model, new Blob([await pdf.save()], { type: stored.type || 'application/pdf' }));
  };

  M.exportOriginal = async (doc, type) => {
    const model = M.getModel(doc?.modelId);
    if (!doc || !model) return PV.setFeedback('Modello non trovato. Salva prima il documento con un modello.', 'error');
    if (model.builtIn) {
      return PV.setFeedback('Il modello standard interno non è un file caricato. Carica un modello personale per scaricare esattamente il file originale compilato.', 'error');
    }
    PV.setFeedback(`Compilazione del file originale ${String(model.originalName || '').toUpperCase()}…`, '');
    try {
      const stored = await M.getModelBlob(model);
      const data = M.modelData(doc, type);
      const format = String(model.format || '').toLowerCase();
      if (['html', 'htm', 'rtf', 'txt', 'json'].includes(format)) await M.exportText(model, stored, data, doc);
      else if (format === 'docx') await M.exportDocx(model, stored, data, doc);
      else if (['xlsx', 'ods'].includes(format)) await M.exportSheet(model, stored, data, doc);
      else if (format === 'odt') await M.exportOdt(model, stored, data, doc);
      else if (format === 'pdf') await M.exportFillablePdf(model, stored, data, doc);
      else throw new Error(`Il formato .${format || '?'} non è compilabile mantenendo il file originale.`);
      PV.setFeedback(`Scaricato “${originalOutputName(model)}”: stesso formato e stessa impaginazione del modello caricato.`, 'success');
    } catch (error) {
      console.error('Compilazione modello originale non riuscita:', error);
      PV.setFeedback(error.message || 'Impossibile compilare il file originale.', 'error');
    }
  };

  const genericPdf = M.exportPdf?.bind(M);
  if (genericPdf) M.exportPdf = (doc, type) => {
    const model = M.getModel(doc?.modelId);
    if (model && !model.builtIn) return M.exportOriginal(doc, type);
    return genericPdf(doc, type);
  };

  const refreshLabels = () => {
    document.querySelectorAll('[data-pvm-export="original"]').forEach((button) => {
      const type = button.dataset.docType || 'preventivo';
      const doc = M.getDocument?.(type, button.dataset.id || '');
      const model = M.getModel(doc?.modelId);
      button.textContent = model && !model.builtIn
        ? `Scarica modello compilato (.${String(model.format).toUpperCase()})`
        : 'Scarica modello compilato';
      button.title = 'Scarica il file originale caricato, compilato senza cambiare formato';
    });
    document.querySelectorAll('[data-pvm-export="pdf"]').forEach((button) => {
      const type = button.dataset.docType || 'preventivo';
      const doc = M.getDocument?.(type, button.dataset.id || '');
      const model = M.getModel(doc?.modelId);
      if (model && !model.builtIn) button.hidden = true;
      else button.textContent = button.classList.contains('pv-btn-small') ? 'PDF' : 'Scarica PDF';
    });
  };
  const observer = new MutationObserver(refreshLabels);
  const start = () => {
    refreshLabels();
    observer.observe(document.body, { childList: true, subtree: true });
  };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();
})();
