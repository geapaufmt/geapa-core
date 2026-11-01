'use strict';
// Adapter for the single canonical boundary. No business authorization here.
const {createDomain}=require('./domain.cjs');
function createContains(){
 const domain=createDomain(()=>{throw Error('UNEXPECTED_IO');});
 return function contains(get,scope,path){
 if(scope?.resourcePaths?.includes(path))return true;
 if(!scope?.operational)return false;
 const docs={};
 for(let i=0;i<96;i++){
  const out=JSON.parse(domain.run(x=>{const missing=[],get=p=>{if(!(p in x.docs)){missing.push(p);return null;}return x.docs[p];};return {ok:f4ActCycleBoundary_().contains(get,x.path,x.boundary,false),missing};},{docs,path,boundary:scope.operational}));
  if(!out.missing.length)return out.ok;
  for(const p of out.missing)docs[p]=get(p)||null;
 }return false;
}
}
module.exports={createContains,contains:(...args)=>createContains()(...args)};
