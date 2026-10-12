'use strict';
const a=require('node:assert/strict'),S=require('./sdk.cjs'),W=require('./wire.cjs'),{save}=require('./session.cjs');
const semantic=v=>Array.isArray(v)?v.map(semantic):v&&typeof v==='object'?'timestampValue'in v?'TIMESTAMP':Object.fromEntries(Object.entries(v).filter(([k])=>!['updateTime','createTime','proofDigest','ultimaAtualizacao','statusChamadaAtualizadoEm'].includes(k)).sort(([a],[b])=>a.localeCompare(b)).map(([k,x])=>[k,semantic(x)])):v;
async function parity(s,host,token,outcome,clock,label,policyIds){
 const before=Object.fromEntries(host.observedBefore),paths=[...host.written],actual=await s.db.getAll(...paths.map(p=>s.db.doc(S.ROOT+'/'+p)));
 const input={before,paths,actual:actual.map(d=>d.exists?W.document(d):null),request:outcome.request,clock,policyIds};save(label+'-parity-input.json',input);
 return verify(s,host,token,input,label);
}
async function verify(s,host,token,input,label){
 const local=await host.execute({token,request:input.request,memoryDocs:input.before,clock:input.clock,policyIds:input.policyIds});a.equal(local.result.state,'COMPLETED',JSON.stringify(local.result));
 const compared=[];for(let j=0;j<input.paths.length;j++){const path=input.paths[j];if(/^(memberAdmissionTickets|commandProvenance|memberAdmissionUses)\//.test(path))continue;a.deepEqual(semantic(input.actual[j]?.fields||null),semantic(local.memoryAfter[path]?.fields||null),'PARITY '+path);compared.push(path);}
 const result={pass:true,method:'immutable v43 business logic: original in-memory adapter vs real SDK transaction',comparedPaths:compared,excluded:'generated timestamps and proof digest containing timestamps; admission tickets depend on emission instant'};save(label+'-parity.json',result);return result;
}
module.exports={parity,verify,semantic};
