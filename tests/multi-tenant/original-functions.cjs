'use strict';
// Parse the original JavaScript, rather than guessing where a function ends.
// Return byte-for-byte source of the unique named declaration. Never rewrite it.
const assert=require('node:assert/strict'),acorn=require('acorn');
function originalFunctions(source){
 const ast=acorn.parse(source,{ecmaVersion:'latest',sourceType:'script',allowHashBang:true});
 const declarations=new Map(),pending=[ast];
 while(pending.length){
  const n=pending.pop();
  if(!n||typeof n!=='object')continue;
  if(n.type==='FunctionDeclaration'&&n.id){
   const list=declarations.get(n.id.name)||[];list.push(n);declarations.set(n.id.name,list);
  }
  for(const v of Object.values(n)){
   if(Array.isArray(v)){for(const child of v){if(child&&typeof child.type==='string')pending.push(child);}}
   else if(v&&typeof v.type==='string')pending.push(v);
  }
 }
 return name=>{const matches=declarations.get(name)||[];assert.equal(matches.length,1,'Expected one original function: '+name);return source.slice(matches[0].start,matches[0].end);};
}
module.exports={originalFunctions};
