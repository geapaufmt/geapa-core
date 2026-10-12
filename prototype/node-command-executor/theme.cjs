'use strict';
const fs=require('node:fs'),a=require('node:assert/strict'),crypto=require('node:crypto'),{session,save}=require('./session.cjs'),{Host}=require('./sdk-host.cjs'),S=require('./sdk.cjs'),W=require('./wire.cjs'),FT=require('./fixture-tools.cjs');
async function seed(db){const sample=require('./local.cjs').sample('THEME_SUBMIT'),source=Object.fromEntries(sample.documents),docs={},payloads=[],allowed=new Set(['presentationCapacityHeads/registry']);
 const rows=sample.documents.filter(([path])=>/^(presentation|canonicalPresentations|jobOutbox)/.test(path)&&path!=='presentationCapacityHeads/registry');
 for(let i=0;i<4;i++){
  const ids=[...new Set(rows.flatMap(([path])=>path.split('/').filter((s,j)=>j%2===1)))],map=Object.fromEntries(ids.map(id=>[id,'B0_T'+i+'_'+id.slice(0,32)]));
  function remap(v,key){if(typeof v==='string')return key==='kind'?v:v.split('/').map(s=>map[s]||s).join('/');if(v?.toMillis)return v;return Array.isArray(v)?v.map(x=>remap(x)):v&&typeof v==='object'?Object.fromEntries(Object.entries(v).map(([k,x])=>[map[k]||k,remap(x,k)])):v;}
  for(const [path,d]of rows){const newPath=remap(path),data=remap(W.decode({mapValue:{fields:d.fields}},db));docs[newPath]=W.encode(data).mapValue.fields;allowed.add(newPath);}
  payloads.push(remap(sample.input.payload));
 }
 const policy=W.decode({mapValue:{fields:source['memberAuthzPolicies/ACT_POLICY'].fields}},db);policy.id='B0_THEME_SELF';policy.targets.allowedResources=[...allowed];policy.targets.resourceScope={kind:'OWN_PRESENTATIONS'};policy.issuedAt=new Date(Date.now()-1000);policy.expiresAt=new Date(Date.now()+86400000);docs['memberAuthzPolicies/B0_THEME_SELF']=W.encode(policy).mapValue.fields;
 await FT.createFixtureDocs(db,docs);return payloads;
}
async function main(){const s=await session(),board=await s.login('BOARD'),host=new Host(s.db,s.verify),doc=p=>s.db.doc(S.ROOT+'/'+p),e=fs.existsSync(__dirname+'/evidence/theme.json')?JSON.parse(fs.readFileSync(__dirname+'/evidence/theme.json')):{samples:[]};
 try{if(!e.payloads){e.payloads=await seed(s.db);save('theme.json',e);}
 if(!e.axesConfigured){const batch=s.db.batch();for(let i=0;i<4;i++){const cfg=doc('presentationCycles/B0_T'+i+'_'+'c'.repeat(32)+'/configRevisions/B0_T'+i+'_config'),axis=doc('presentationPolicyReferences/B0_T'+i+'_axis'),current=await cfg.get(),definition=await axis.get();a(current.exists&&definition.exists);batch.update(cfg,{parameters:{...current.data().parameters,axisRefs:[{kind:'policy',id:'B0_T'+i+'_axis'}]}},{lastUpdateTime:current.updateTime});batch.update(axis,{approved:true,active:true},{lastUpdateTime:definition.updateTime});}await batch.commit();e.axesConfigured=true;save('theme.json',e);}
 if(!e.freshRequests){const batch=s.db.batch();for(let i=0;i<4;i++){const ref=doc('presentationUnitHeads/B0_T'+i+'_unit'),d=await ref.get();a.equal(d.data().revision,1);batch.update(ref,{thematicRequestRef:{kind:'thematic',id:'B0_T'+i+'_theme'}},{lastUpdateTime:d.updateTime});}await batch.commit();e.freshRequests=true;save('theme.json',e);}
 for(let i=e.samples.length;i<4;i++){
  host.metrics=[];host.written.clear();host.observedBefore.clear();s.db.b0RpcEvents=[];const operationId=crypto.randomBytes(16).toString('hex'),clock=Date.now(),at=performance.now();
  const outcome=await host.execute({token:board.token,operationId,commandType:'ACT.THEME_SUBMIT',payload:e.payloads[i],policyIds:['B0_THEME_SELF']});
  const wallMs=performance.now()-at,rpcs=s.db.b0RpcEvents.map(x=>x.method),io=[...host.metrics];a.equal(outcome.result.state,'COMPLETED',JSON.stringify(outcome.result));
  const rt=performance.now(),unit=await doc('presentationUnitHeads/'+e.payloads[i].unitRef.id).get(),receipt=await doc('commandReceipts/'+operationId).get();a(unit.exists&&receipt.exists);a.equal(unit.data().revision,2);
  e.samples.push({index:i,coldLike:i===0,wallMs,readbackMs:performance.now()-rt,...outcome,rpcs,io});save('theme.json',e);if(!e.parity){e.parity=await require('./parity.cjs').parity(s,host,board.token,outcome,clock,'theme',['B0_THEME_SELF']);save('theme.json',e);}console.log({sample:i,wallMs:Math.round(wallMs),rpcs:rpcs.length});
 }
 }finally{await s.close([host]);}}
main().catch(e=>{console.error({code:e.code||e.message,stack:e.stack});process.exitCode=1;});
