/* Varga Piano Neve — importazione guidata del piano neve da PDF */
(function(){
  const $=id=>document.getElementById(id);
  const dialog=$('pdfImportDialog'), fileInput=$('pdfImportFile');
  if(!dialog||!fileInput)return;
  let extracted='', pages=0;
  const clean=s=>String(s||'').replace(/\s+/g,' ').trim();
  const analyze=text=>{
    const lines=text.split(/\n+/).map(clean).filter(Boolean);
    const municipality=/\b(comune|comuni|municipio|municipalità)\b/i;
    const route=/\b(percorso|itinerario|giro|zona|tratto|viabilità)\b/i;
    const roads=/\b(via|viale|piazza|strada|corso|rotonda|tangenziale|provinciale|statale)\b/i;
    const municipalities=[...new Set(lines.filter(x=>municipality.test(x)).slice(0,80))];
    const routes=[...new Set(lines.filter(x=>route.test(x)).slice(0,120))];
    const streets=[...new Set(lines.filter(x=>roads.test(x)).slice(0,200))];
    return {lines,municipalities,routes,streets};
  };
  const show=()=>{dialog.hidden=false;$('pdfImportStepUpload').hidden=false;$('pdfImportStepReview').hidden=true;fileInput.value='';};
  const preview=(data)=>{
    const q=clean($('pdfImportSearch').value).toLowerCase();
    const lines=(data.lines||[]).filter(x=>!q||x.toLowerCase().includes(q));
    $('pdfImportPages').textContent=pages;
    $('pdfImportMunicipalities').textContent=data.municipalities.length;
    $('pdfImportRoutes').textContent=data.routes.length;
    $('pdfImportPreview').textContent=lines.slice(0,220).join('\n')||'Nessuna riga corrisponde alla ricerca.';
  };
  let data=null;
  $('routeImportPdf')?.addEventListener('click',show);
  $('pdfImportClose')?.addEventListener('click',()=>dialog.hidden=true);
  $('pdfImportBack')?.addEventListener('click',show);
  $('pdfImportSearch')?.addEventListener('input',()=>data&&preview(data));
  $('pdfImportSave')?.addEventListener('click',()=>{
    if(!data)return;
    const draft={createdAt:new Date().toISOString(),pages,municipalities:data.municipalities,routes:data.routes,streets:data.streets,rawText:extracted};
    localStorage.setItem('vargaPianoNevePdfDraft',JSON.stringify(draft));
    $('pdfImportStatus').textContent='Bozza salvata. Il piano è pronto per il controllo e la pubblicazione.';
    $('pdfImportStatus').className='pdfImportStatus ok';
    window.dispatchEvent(new CustomEvent('varga-pdf-imported',{detail:draft}));
  });
  fileInput.addEventListener('change',async()=>{
    const file=fileInput.files?.[0]; if(!file)return;
    if(!window.pdfjsLib){alert('Il lettore PDF non è ancora pronto. Riprova tra qualche secondo.');return;}
    try{
      $('pdfImportStepUpload').hidden=true;$('pdfImportStepReview').hidden=false;
      $('pdfImportStatus').textContent='Analisi del PDF in corso…';
      const buf=await file.arrayBuffer(), pdf=await window.pdfjsLib.getDocument({data:buf}).promise; pages=pdf.numPages; let out=[];
      for(let p=1;p<=pages;p++){const page=await pdf.getPage(p), content=await page.getTextContent();out.push(content.items.map(x=>x.str).join(' '));}
      extracted=out.join('\n'); data=analyze(extracted); preview(data);
      $('pdfImportStatus').textContent='Controlla il testo estratto. La pubblicazione automatica sarà aggiunta dopo la verifica dei percorsi sulla mappa.';
    }catch(e){$('pdfImportStatus').textContent='Impossibile leggere il PDF: '+(e.message||e);$('pdfImportStatus').className='pdfImportStatus';}
  });
  if(window.pdfjsLib)window.pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  window.VargaPdfImport={open:show,getDraft:()=>JSON.parse(localStorage.getItem('vargaPianoNevePdfDraft')||'null')};
})();