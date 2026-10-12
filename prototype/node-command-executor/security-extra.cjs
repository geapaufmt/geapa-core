'use strict';
const fs=require('node:fs'),a=require('node:assert/strict'),crypto=require('node:crypto'),{session,save}=require('./session.cjs'),{Host}=require('./sdk-host.cjs'),S=require('./sdk.cjs'),W=require('./wire.cjs'),FT=require('./fixture-tools.cjs');
function member(uid,suspended){
 const r=require('./local.cjs').runtime(),docs=r.run(suspended=>{const source=suspended?fixture.event(fixture.base(),'SUSPENSAO','2026-10-01','2026-12-31'):fixture.base(),d=fixture.projectionMap(source,1,Date.now());return Object.fromEntries(Object.entries(d).map(([p,v])=>[p,f4Firestore_().encode(v).mapValue.fields]));},suspended);
 const map={'PES-000001':'PES-000080','VINC-001':'B0_LINK_080',SYN_UID:uid,'EV-001':'B0_SUSPENSION_080'};
 const result=FT.replace(docs,map),subject=require('./fixtures.cjs').seed(uid)['memberAuthzSubjects/PES-000001'].fields;result['memberAuthzSubjects/PES-000080']=FT.replace(subject,map);return result;
}
async function main(){const s=await session(),self=await s.login('SELF'),host=new Host(s.db,s.verify),e={tests:[]},doc=p=>s.db.doc(S.ROOT+'/'+p);
 try{
  const active=member(self.uid,false);a.equal(active['memberCurrentStates/PES-000080'].state.stringValue,'MEMBRO_EFETIVO');await FT.createFixtureDocs(s.db,active);
  await a.rejects(()=>host.execute({mode:'edit',token:self.token,operationId:crypto.randomBytes(16).toString('hex'),activityId:'dev_meeting',expectedRevision:1,title:'COMMON_MEMBER_CANNOT_EDIT'}),/AUTHORIZATION_DENIED|AUTH_BINDING_INVALID|M13_AUTHZ_DENIED/);e.tests.push({name:'bound ordinary active member cannot edit administratively',pass:true});
  const suspended=member(self.uid,true);a.equal(suspended['memberCurrentStates/PES-000080'].state.stringValue,'SUSPENSO');
  const b=s.db.batch();for(const [path,fields]of Object.entries(suspended)){a(path.startsWith('member')||path.startsWith('portal')||path.startsWith('projection')||path.startsWith('institutional'));b.set(doc(path),W.decode({mapValue:{fields}},s.db));}const journal=JSON.parse(fs.readFileSync(__dirname+'/evidence/seed.json'));journal.paths=[...new Set([...journal.paths,...Object.keys(suspended)])];save('seed.json',journal);await b.commit();
  await a.rejects(()=>host.execute({token:self.token,operationId:crypto.randomBytes(16).toString('hex'),commandType:'ACT.DRAFT_SAVE',payload:{draftId:'B0_SUSPENDED_DRAFT',draftKind:'TITLE',textFields:{title:'Denied'}}}),/AUTHORIZATION_DENIED|AUTH_BINDING_INVALID|M13_AUTHZ_DENIED/);e.tests.push({name:'suspended canonical membership cannot issue ordinary SELF command',pass:true});
  const restore=s.db.batch();for(const [path,fields]of Object.entries(active))restore.set(doc(path),W.decode({mapValue:{fields}},s.db));for(const path of Object.keys(suspended).filter(p=>!active[p]))restore.delete(doc(path));await restore.commit();
  const r=require('./local.cjs').runtime();r.run(()=>{const F=f4Firestore_(),uid='0WejtvmbXjfVSntKDf0cSJCaQSt1';globalThis.b0Profile={schemaVersion:'core-dev-admin-v1',projectId:'geapa-convergencia-dev',environment:'DEV',uid,active:true,revision:1,issuedAt:F.Timestamp.fromMillis(Date.now()-1000),expiresAt:F.Timestamp.fromMillis(Date.now()+60000),permissions:'FUNCTIONAL_ALL',exportable:false};});r.context.PropertiesService.getScriptProperties().setProperty('GEAPA_ENV','PROD');
  a.throws(()=>r.run(()=>f4DevBoundary_().run(()=>f4DevAdmin_().profile(()=>b0Profile,b0Profile.uid,Date.now(),true))),/DESTINATION_INVALID/);e.tests.push({name:'copied DEV_ADMIN profile in PROD runtime fails closed before authority',pass:true});
  save('security-extra.json',{...e,passed:e.tests.length,failed:0});console.log(e);
 }finally{await s.close([host]);}}
if(require.main===module)main().catch(e=>{console.error({code:e.code||e.message,stack:e.stack});process.exitCode=1;});
module.exports={member};
