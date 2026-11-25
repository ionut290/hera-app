const assert=require('node:assert/strict');
const {chromium}=require('playwright');
(async()=>{
 const {extractOriginalPreview}=await import('../generate-original-preview.mjs');
 const {buildOriginalPreviewFrame}=await import('./src/original-preview-frame.mjs');
 const browser=await chromium.launch({headless:true});
 try{
 const context=await browser.newContext({viewport:{width:390,height:844},serviceWorkers:'block'});
 const external=[];await context.route('**/*',route=>{external.push(route.request().url());return route.abort();});
 const page=await context.newPage();const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.setContent('<iframe title="original" sandbox="allow-scripts" style="width:100%;height:800px"></iframe>');
 await page.locator('iframe').evaluate((frame,srcdoc)=>frame.srcdoc=srcdoc,buildOriginalPreviewFrame(extractOriginalPreview(),'admin'));
 const f=page.frameLocator('iframe');
 await f.locator('#open-new-commessa-btn').click();
 await f.locator('#commessa-name').fill('Verde demo');
 await f.locator('#commessa-code').fill('V1');
 await f.locator('#commessa-form button[type=submit]').click();
 await f.locator('.commessa-manage-card').filter({hasText:'Verde demo'}).waitFor();
 await f.getByRole('button',{name:'APRI',exact:true}).click();
 await f.locator('#add-management-impianto-btn').click();
 await f.locator('#manual-impianto-denominazione').fill('Parco demo');
 await f.locator('#manual-impianto-comune').fill('Comune demo');
 await f.locator('#manual-impianto-codice').fill('A11');
 await f.locator('#manual-impianto-submit').click();
 await f.locator('#impianti-management-tbody tr').filter({hasText:'Parco demo'}).waitFor();
 assert.match(await f.locator('#impianti-management-tbody').innerText(),/DA FARE/);
 await f.locator('#impianti-management-search').fill('inesistente');
 await f.getByText('Nessun impianto trovato.',{exact:true}).waitFor();
 await f.locator('#clear-impianti-management-search').click();
 await f.locator('#impianti-management-tbody tr').filter({hasText:'Parco demo'}).waitFor();
 await f.locator('#close-impianti-management-btn').click();
 assert.match(await f.locator('.commessa-manage-card').innerText(),/1 impianto/);
 await page.locator('iframe').evaluate((frame,srcdoc)=>frame.srcdoc=srcdoc,buildOriginalPreviewFrame(extractOriginalPreview(),'admin'));
 await f.getByText('Nessuna commessa disponibile.',{exact:true}).waitFor();
 assert.deepEqual(errors,[]);
 assert.deepEqual(external,[]);
 console.log('ORIGINAL PREVIEW BROWSER: passed');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1});
