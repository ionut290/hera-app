export function buildVerdeFrame(source) {
  const signature = 'export function installVerdeLevatoPilot(';
  if (typeof source !== 'string' || !source.includes(signature) || /<\/script/i.test(source)) throw new Error('Unexpected pilot source');
  const code = source.replace(signature, 'function installVerdeLevatoPilot(');
  return `<!doctype html><html lang="it"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data:; base-uri 'none'; form-action 'none'"><style>*{box-sizing:border-box}body{margin:0;font-family:system-ui,sans-serif}</style></head><body><script>
${code}
const readStore=Object.freeze({collection(name){
 if(!['verdeLevatoRecords','verdeLevatoCommesse','verdeLevatoConfig'].includes(name)) throw new Error('Collection denied');
 return Object.freeze({get:async()=>({docs:[]}),doc:()=>Object.freeze({get:async()=>({exists:false})})});
}});
const pilot=installVerdeLevatoPilot({readStore,user:{uid:'preview-only',email:'preview@example.test'},onClose:()=>parent.postMessage('varga-verde-close','*')});
pilot.open().catch(()=>{document.body.textContent='Anteprima non disponibile';});
</script></body></html>`;
}
