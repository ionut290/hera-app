(()=>{
  'use strict';

  const PV=window.HeraPreventivi;
  const M=window.HeraPreventiviModels;
  if(!PV||!M)return;

  let queued=false;
  const desiredTabs=[
    ['quotes','Preventivi'],
    ['prices','Prezziari'],
    ['consuntivi','Consuntivi'],
    ['models','Modelli']
  ];

  function ensureTabs(){
    const page=PV.page?.()||document.getElementById('preventivi-page');
    const nav=page?.querySelector('.pv-nav');
    if(!nav)return;
    desiredTabs.forEach(([view,label])=>{
      if(nav.querySelector(`[data-pv-view="${view}"]`))return;
      nav.insertAdjacentHTML('beforeend',`<button type="button" class="pv-tab" data-pv-view="${view}">${PV.escapeHtml(label)}</button>`);
    });
    nav.querySelectorAll('[data-pv-view]').forEach(button=>button.classList.toggle('active',button.dataset.pvView===PV.state.view));
  }

  function currentDoc(form){
    if(form.matches('[data-cons-form]')){
      const id=PV.state.editingConsuntivoId;
      return id&&id!=='new'?(PV.state.consuntivi||[]).find(item=>item.id===id)||{}:{};
    }
    const id=PV.state.editingQuoteId;
    return id&&id!=='new'?PV.getQuote?.(id)||{}:{};
  }

  function ensureModelSelector(form){
    if(!form||form.querySelector('[data-pvm-model-select]')||typeof M.modelSection!=='function')return;
    const type=form.matches('[data-cons-form]')?'consuntivo':'preventivo';
    const firstCard=form.querySelector('.pv-form-card');
    if(!firstCard)return;
    firstCard.insertAdjacentHTML('afterend',M.modelSection(type,currentDoc(form)));
    M.renderDynamic?.(form,currentDoc(form));
    form.dataset.pvmDecorated='1';
  }

  function repair(){
    queued=false;
    const page=PV.page?.()||document.getElementById('preventivi-page');
    if(!page)return;
    ensureTabs();
    ensureModelSelector(page.querySelector('[data-pv-quote-form]'));
    ensureModelSelector(page.querySelector('[data-cons-form]'));
  }

  function queue(){
    if(queued)return;
    queued=true;
    requestAnimationFrame(repair);
  }

  const originalEnsure=PV.ensurePage?.bind(PV);
  if(originalEnsure)PV.ensurePage=()=>{const result=originalEnsure();queue();return result};
  const originalRender=PV.renderCurrentView?.bind(PV);
  if(originalRender)PV.renderCurrentView=()=>{const result=originalRender();queue();return result};

  const observer=new MutationObserver(queue);
  const start=()=>{
    observer.observe(document.body,{childList:true,subtree:true});
    queue();
    [250,750,1500,3000].forEach(delay=>setTimeout(queue,delay));
  };
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});
  else start();
})();
