'use strict';
const {encode,decode}=require('./firestore.cjs'),{fail}=require('./environment.cjs');
// A server-generated plan invokes the original Core pipeline in SDK phases.
// HTTP input must never be passed here as a plan.
async function executeSystem(store,run,plan,hooks={}){
 store.operationId=plan.request.operationId;
 await store.read(()=>run('(p)=>nodeStorageSystem(p)',plan));
 if(!await store.transaction(()=>run('()=>storageSystem.existing(nodeTx)'))){
  const issued=await store.transaction(()=>run('()=>codec.encode(storageSystem.issue(nodeTx))'));
  await store.transaction(()=>run('(v)=>{storageSystem.reserve(nodeTx,codec.decode(v));return true;}',issued));
  await hooks.afterSystemReserve?.(plan);
 }
 const c=decode(await store.transaction(()=>run('()=>codec.encode(storageSystem.claim(nodeTx,"CF_GEN2_DEV_STORAGE"))')));
 if(c.state==='CLAIMED'){
  const value=decode(await store.transaction(()=>run('(h)=>codec.encode(storageSystem.finalize(nodeTx,codec.decode(h)))',encode(c.handle))));
  if(value.state!=='COMPLETED'){if(require('./environment.cjs').fence().emulator)console.error('LOCAL_SYSTEM_RESULT '+JSON.stringify(value));fail('RECONCILIATION_REQUIRED');}
  await hooks.afterSystemCommit?.(plan);
 }else if(!['COMPLETED','RECOVERED'].includes(c.state))fail('RECONCILIATION_REQUIRED');
}
module.exports={executeSystem};
