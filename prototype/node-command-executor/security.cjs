'use strict';
const a=require('node:assert/strict'),crypto=require('node:crypto'),fs=require('node:fs'),{session,save}=require('./session.cjs'),{Host}=require('./sdk-host.cjs'),S=require('./sdk.cjs'),W=require('./wire.cjs');
const id=()=>crypto.randomBytes(16).toString('hex');
async function main(){
 const s=await session(),board=await s.login('BOARD'),none=await s.login('NONE'),host=new Host(s.db,s.verify),hosts=[host],checks=fs.existsSync(__dirname+'/evidence/security.json')?JSON.parse(fs.readFileSync(__dirname+'/evidence/security.json')).checks:[];
 const doc=p=>s.db.doc(S.ROOT+'/'+p),get=async p=>(await doc(p).get()).data();
 const check=async(name,fn)=>{if(checks.some(x=>x.name===name&&x.pass))return;await fn();checks.push({name,pass:true});save('security.json',{checks,passed:checks.length,failed:0});console.log(name+' PASS');};
 const edit=async(overrides={})=>{const d=await get('canonicalActivities/dev_meeting');return host.execute({mode:'edit',token:board.token,operationId:id(),activityId:d.id,expectedRevision:d.revision,title:'B0_SECURITY',...overrides});};
 try{
  await check('unbound authenticated UID denied',()=>a.rejects(()=>edit({token:none.token}),/AUTH_BINDING_INVALID|AUTHORIZATION_DENIED|M13_AUTHZ_DENIED/));
  await check('tampered Firebase token denied',()=>a.rejects(()=>edit({token:board.token.slice(0,-6)+'xxxxxx'}),/auth\/|SDK_OPERATION_FAILED/));
  await check('non DEV environment denied before SDK',()=>a.throws(()=>S.fence(S.PROJECT,'PROD')));
  await check('non DEV project denied before SDK',()=>a.throws(()=>S.fence('geapa-prod','DEV')));
  await check('stale displayed revision denied',()=>a.rejects(()=>edit({expectedRevision:1}),/STALE_REVISION/));
  const prior=JSON.parse(fs.readFileSync(__dirname+'/evidence/edit-final.json')).samples.at(-1);
  await check('idempotent replay keeps revision and receipt',async()=>{const before=await get('canonicalActivities/dev_meeting'),receipt=await get('commandReceipts/'+prior.key.operationId);const r=await host.execute({token:board.token,request:prior.request});a.equal(r.result.state,'RECOVERED');a.deepEqual(await get('canonicalActivities/dev_meeting'),before);a.deepEqual(await get('commandReceipts/'+prior.key.operationId),receipt);});
  await check('same operation with changed payload denied',()=>a.rejects(()=>host.execute({token:board.token,request:{...prior.request,payload:{...prior.request.payload,title:'B0_CHANGED'}}}),/OPERATION_CONFLICT/));
  const policyPath='memberAuthzPolicies/ACT_POLICY',original=await get(policyPath);
  try{
   await doc(policyPath).update({capabilities:['ACTIVITY_BOARD_READ'],fieldMasks:{ACTIVITY_BOARD_READ:[]}});
   await check('missing capability denied',()=>a.rejects(()=>edit(),/AUTHORIZATION_DENIED|M13_AUTHZ_DENIED/));
  }finally{await doc(policyPath).set(original);}
  const rolePath='memberRoleCatalog/CARGO-001',role=await get(rolePath),fnPath='memberFunctions/VIG-001',fn=await get(fnPath);
  try{await doc(rolePath).update({roleKey:'DIRETOR_RECURSOS_HUMANOS'});await doc(fnPath).update({scopeKey:'RECURSOS_HUMANOS'});
   await check('incompatible role cannot borrow Presidency policy',()=>a.rejects(()=>edit(),/AUTHORIZATION_DENIED|AUTH_BINDING_INVALID|DOMAIN_STATE_INVALID|M13_AUTHZ_DENIED/));
  }finally{await doc(rolePath).set(role);await doc(fnPath).set(fn);}
  const bpath='memberAuthBindings/'+board.uid,binding=await get(bpath);
  try{await doc(bpath).update({active:false});await check('revoked binding denied',()=>a.rejects(()=>edit(),/AUTHORIZATION_DENIED|AUTH_BINDING_INVALID|M13_AUTHZ_DENIED/));}finally{await doc(bpath).set(binding);}
  await check('two edits same revision: one effect',async()=>{
   const current=await get('canonicalActivities/dev_meeting'),prepared=await host.execute({mode:'edit',token:board.token,operationId:id(),activityId:current.id,expectedRevision:current.revision,title:'B0_CONCURRENT_A',prepareOnly:true});
   const h2=new Host(s.db,s.verify);hosts.push(h2);const reqB={...prepared.request,operationId:id(),payload:{...prepared.request.payload,title:'B0_CONCURRENT_B'}};
   const outcomes=await Promise.allSettled([host.execute({token:board.token,request:prepared.request}),h2.execute({token:board.token,request:reqB})]);
   save('concurrency.json',outcomes.map(x=>x.status==='fulfilled'?{state:x.value.result.state,code:x.value.result.code}:{error:x.reason.code||x.reason.message,stack:x.reason.stack}));
   const winners=outcomes.filter(x=>x.status==='fulfilled'&&x.value.result.state==='COMPLETED');a.equal(winners.length,1,JSON.stringify(outcomes));a.equal((await get('canonicalActivities/dev_meeting')).revision,current.revision+1);
   const loser=outcomes.find(x=>!winners.includes(x));a(loser.status==='rejected'&&/STALE_REVISION|REVISION_CONFLICT|FIRESTORE_ABORTED/.test(loser.reason.message)||loser.status==='fulfilled'&&loser.value.result.state==='FAILED'&&loser.value.result.code==='REVISION_CONFLICT',JSON.stringify(loser));
  });
  await check('lost local response replay has one effect',async()=>{const start=await get('canonicalActivities/dev_meeting'),r=await edit({title:'B0_TIMEOUT_REPLAY'});/* Drop result before exposing it to a caller. */const after=await get('canonicalActivities/dev_meeting'),replayed=await host.execute({token:board.token,request:r.request});a.equal(after.revision,start.revision+1);a.equal(replayed.result.state,'RECOVERED');a.deepEqual(await get('canonicalActivities/dev_meeting'),after);});
 }finally{await s.close(hosts);}
}
main().catch(e=>{console.error({code:e.code||e.message,stack:e.stack});process.exitCode=1;});
