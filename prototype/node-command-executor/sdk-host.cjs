'use strict';
const assert=require('node:assert/strict'),{Worker}=require('node:worker_threads'),p=require('node:path'),crypto=require('node:crypto');
const {FieldPath}=require('@google-cloud/firestore'),W=require('./wire.cjs'),S=require('./sdk.cjs');
class Host{
 constructor(db,verify){assert.equal(db.projectId,S.PROJECT,'B0_SDK_PROJECT_DENIED');assert.equal(db.databaseId,'(default)','B0_DATABASE_DENIED');this.db=db;this.verify=verify;this.transactions=new Map();this.metrics=[];this.written=new Set();this.touchedAll=new Set();this.observedBefore=new Map();}
 doc(name){const prefix='projects/'+S.PROJECT+'/databases/(default)/documents/'+S.ROOT+'/';assert(name.startsWith(prefix),'B0_NAMESPACE_DENIED');const relative=name.slice(prefix.length);assert(/^([A-Za-z][A-Za-z0-9_]*\/[A-Za-z0-9_-]+)(\/[A-Za-z][A-Za-z0-9_]*\/[A-Za-z0-9_-]+)*$/.test(relative),'B0_PATH_DENIED');return this.db.doc(S.ROOT+'/'+relative);}
 async call(x){
  S.fence(x.project,x.environment);const start=performance.now();let reads=0,writes=0,rpcs=0;
  try{
   if(x.kind==='verify'){const v=await this.verify(x.token);return {uid:v.uid,expiresAt:v.exp*1000};}
   const b=x.body||{},tx=b.transaction&&this.transactions.get(b.transaction);if(b.transaction)assert(tx,'UNKNOWN_TRANSACTION');
   if(x.kind==='beginTransaction'){
    const id=crypto.randomBytes(16).toString('hex');let ready,rejectReady,commit,rejectCommit;
    const begun=new Promise((a,b)=>{ready=a;rejectReady=b;}),end=new Promise((a,b)=>{commit=a;rejectCommit=b;});
    const entry={commit,rejectCommit,reads:new Map()};this.transactions.set(id,entry);
    entry.promise=this.db.runTransaction(async native=>{entry.native=native;ready();await end;},{maxAttempts:1}).catch(e=>{rejectReady(e);throw e;});entry.promise.catch(()=>{});
    await begun;return {transaction:id};
   }
   if(x.kind==='rollback'){tx.rejectCommit(Object.assign(Error('LOCAL_ROLLBACK'),{localRollback:true}));await tx.promise.catch(e=>{if(!e.localRollback)throw e;});this.transactions.delete(b.transaction);return {};}
   if(x.kind==='get'||x.kind==='batchGet'){
    const names=x.kind==='get'?[x.name]:b.documents;assert(names.length<=180);const refs=names.map(n=>this.doc(n));reads=refs.length;rpcs=1;
    const docs=await (tx?tx.native.getAll(...refs):this.db.getAll(...refs));
    for(const d of docs){const path=d.ref.path.slice((S.ROOT+'/').length);if(!this.observedBefore.has(path))this.observedBefore.set(path,d.exists?W.document(d):null);}
    if(tx)for(const d of docs)tx.reads.set(d.ref.path,d);
    if(x.kind==='get')return docs[0].exists?W.document(docs[0]):null;
    return docs.map((d,i)=>d.exists?{found:W.document(d)}:{missing:names[i]});
   }
   if(x.kind==='runQuery'){
    const q=b.structuredQuery;assert(q.limit<=20&&!q.from[0].allDescendants);const parent=x.parent?this.doc(x.parent).path:S.ROOT;
    let query=this.db.collection(parent+'/'+q.from[0].collectionId);const filters=q.where?.compositeFilter?.filters||(q.where?[q.where]:[]),ops={EQUAL:'==',LESS_THAN:'<',LESS_THAN_OR_EQUAL:'<=',GREATER_THAN:'>',GREATER_THAN_OR_EQUAL:'>='};
    for(const {fieldFilter:f}of filters){assert(ops[f.op]);query=query.where(f.field.fieldPath,ops[f.op],W.decode(f.value,this.db));}
    for(const o of q.orderBy)query=query.orderBy(o.field.fieldPath==='__name__'?FieldPath.documentId():o.field.fieldPath,o.direction==='DESCENDING'?'desc':'asc');
    if(q.startAt)query=query.startAfter(...q.startAt.values.map(v=>W.decode(v,this.db)));rpcs=1;const snap=await query.limit(q.limit).get();reads=snap.size;return snap.docs.map(d=>({document:W.document(d)}));
   }
   if(x.kind==='commit'){
    const target=tx?tx.native:this.db.batch();writes=b.writes.length;assert(writes<=400);
    for(const w of b.writes){const ref=this.doc(w.update?.name||w.delete),pre=w.currentDocument;assert(pre,'PRECONDITION_REQUIRED');
     if(tx){const before=tx.reads.get(ref.path);assert(before,'WRITE_WITHOUT_TRANSACTIONAL_READ');assert(pre.updateTime?before.exists&&W.stamp(before.updateTime)===pre.updateTime:before.exists===pre.exists,'PRECONDITION_FAILED');}
     const condition=pre.updateTime?{lastUpdateTime:W.timestamp(pre.updateTime)}:undefined;
     if(w.delete)target.delete(ref,condition);else{const data=W.decode({mapValue:{fields:w.update.fields}},this.db);if(pre.exists===false)target.create(ref,data);else if(w.updateMask||condition&&!tx)target.update(ref,data,...(condition?[condition]:[]));else target.set(ref,data);}
     this.written.add(ref.path.slice((S.ROOT+'/').length));
     this.touchedAll.add(ref.path.slice((S.ROOT+'/').length));
    }
    rpcs=writes?1:0;let result;if(tx){tx.commit();await tx.promise;this.transactions.delete(b.transaction);}else result=await target.commit();
    const commitTime=result?.[0]?.writeTime?W.stamp(result[0].writeTime):new Date().toISOString();return {commitTime,writeResults:b.writes.map(()=>({updateTime:commitTime}))};
   }
   throw Error('B0_OPERATION_DENIED');
  }finally{this.metrics.push({kind:x.kind,stage:x.stage||null,ms:performance.now()-start,reads,writes,rpcs});}
 }
 async execute(input){
  const shared=new SharedArrayBuffer(16*1024*1024),control=new Int32Array(shared,0,4),bytes=new Uint8Array(shared,16);let timeout;
  const worker=new Worker(p.join(__dirname,'worker.cjs'),{workerData:{shared,input}});
  try{return await new Promise((resolve,reject)=>{timeout=setTimeout(()=>reject(Error('B0_LOCAL_TIMEOUT')),180000);
   worker.on('error',reject);worker.on('message',async message=>{
    if(message.type==='result'){resolve(message.value);return;}if(message.type==='error'){reject(Object.assign(Error(message.code),{detail:message.detail,stack:message.stack}));return;}
    if(message.type!=='sdk')return;let response;try{response={ok:true,value:await this.call(message.value)};}catch(e){response={ok:false,code:e.code||e.message};}
    const encoded=Buffer.from(JSON.stringify(response));assert(encoded.length<bytes.length);bytes.set(encoded);Atomics.store(control,1,encoded.length);Atomics.store(control,0,1);Atomics.notify(control,0);
   });
  });}finally{clearTimeout(timeout);await worker.terminate();for(const [id,tx]of this.transactions){tx.rejectCommit(Object.assign(Error('LOCAL_ROLLBACK'),{localRollback:true}));await tx.promise.catch(()=>{});this.transactions.delete(id);}}
 }
}
module.exports={Host};
