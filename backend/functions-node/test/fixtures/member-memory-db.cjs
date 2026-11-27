'use strict';
const assert=require('node:assert/strict'),{Timestamp}=require('firebase-admin/firestore');
const {DEMO}=require('../../src/environment.cjs');
// Snapshot and atomic commit port for offline regression. No external recipient or database.
function database(docs){
 let version=Date.now();
 const doc=path=>({path});
 const snap=ref=>({exists:!!docs[ref.path],ref,data:()=>docs[ref.path],updateTime:Timestamp.fromMillis(version)});
 const getAll=async(...refs)=>refs.map(snap);
 function collection(name){
  assert.equal(name,'memberAuthBindings');
  let person;
  return {where(field,op,value){assert.equal(field,'idPessoa');assert.equal(op,'==');person=value;return this;},
   orderBy(){return this;},limit(n){assert.equal(n,2);return this;},
   async get(){const rows=Object.keys(docs).filter(p=>p.startsWith(name+'/')&&docs[p].idPessoa===person).sort().slice(0,2).map(p=>snap(doc(p)));return {docs:rows,size:rows.length,empty:!rows.length};}};
 }
 return {projectId:DEMO,doc,getAll,collection,async runTransaction(fn){
  const writes=[];
  const result=await fn({getAll,get:async ref=>ref.get?ref.get():snap(ref),set:(r,v)=>writes.push([r.path,v]),update:(r,v)=>writes.push([r.path,v]),delete:r=>writes.push([r.path,null])});
  for(const [p,v] of writes){if(v)docs[p]=v;else delete docs[p];}version++;return result;
 }};
}
module.exports={database};
