'use strict';
const {test}=require('node:test'),a=require('node:assert/strict');
const {setup}=require('./activity-profile-fixture.cjs'),{encode,decode}=require('../src/firestore.cjs');
const ref=(kind,id)=>({kind,id});
function scenario(){
 const now=Date.now(),activity=ref('activity','SYN_CALL'),roster=ref('attendanceRoster','SYN_ROSTER');
 const h=setup('DIRETOR_RECURSOS_HUMANOS',docs=>{
  const policy=docs['memberAuthzPolicies/CF_DEV_ACT_POLICY'];policy.capabilities=['ACTIVITY_BOARD_READ','ACTIVITY_ATTENDANCE_MANAGE'];policy.fieldMasks={ACTIVITY_BOARD_READ:[],ACTIVITY_ATTENDANCE_MANAGE:[]};
  policy.targets.allowedResources=['canonicalActivities/SYN_CALL','activityAttendanceRosters/SYN_ROSTER','activityAttendanceCollections/SYN_CALL'];
  const put=(col,id,more)=>docs[col+'/'+id]={id,revision:1,schemaVersion:'act-v2-candidate',recordMode:'NEW_V2',...more};
  put('canonicalActivities','SYN_CALL',{status:'SCHEDULED',civilDate:new Date(now).toISOString().slice(0,10),startsAt:now,cycleRef:ref('cycle','SYN_CYCLE'),typeSnapshot:{durationMinutes:60,countsAttendance:true,countsAbsence:true},historicalReferences:Array.from({length:60},(_,i)=>ref('evidence','SYN_HISTORY_'+i))});
  put('presentationCycles','SYN_CYCLE',{status:'OPEN',historicalReferences:Array.from({length:60},(_,i)=>ref('evidence','SYN_HISTORY_'+i))});
  put('presentationCapacityHeads','registry',{membersEpoch:28});
  put('activityAttendanceCollections','SYN_CALL',{activityRef:activity,state:'DRAFT',rosterCount:0,preRosterFactCount:1,opensAt:now-60000,closesAt:now+600000,frequencyBasis:{unit:'ACTIVITY',justifiedMode:'COUNT_ABSENT',excusedMode:'COUNT_PRESENT'}});
  const records={};for(let i=0;i<21;i++)records['SYN_FACT_'+i]={personRef:ref('person','SYN_PERSON_'+i),membershipRef:ref('membership','SYN_LINK_'+i),personRevision:1,participantType:'MEMBRO',applicableAtDate:true,countsAttendance:true,countsAbsence:true};
  put('activityAttendanceRosters','SYN_ROSTER',{state:'SEALED',complete:true,activityRef:activity,activityRevision:1,membersEpoch:28,records,sourceHeads:Array.from({length:99},(_,i)=>({path:'memberAggregateHeads/SYN_PERSON_'+i,revision:1}))});
  put('activityAttendance','SYN_FACT_0',{activityRef:activity,personRef:ref('person','SYN_PERSON_0'),membershipRef:ref('membership','SYN_LINK_0'),status:'UNRECORDED',pendingJustificationRef:ref('justification','SYN_PRIOR')});
  put('activityJustifications','SYN_PRIOR',{state:'PREVIOUS',basisAttendanceRevision:1,reason:'Synthetic previous justification'});
 });
 let sequence=0;
 function prepare(command='ATTENDANCE_ROSTER_APPLY',payload={activityRef:activity,rosterRef:roster}){return JSON.parse(h.domain.run(x=>f4ActivitiesFunctionalPorts_({config:nodeConfig,identity:{uid:'CF_ADMIN',expiresAt:x.now+120000},settings:{commands:[x.command],policyIds:['CF_DEV_ACT_POLICY']},clock:()=>x.now}).prepare({operationId:x.op,commandType:'ACT.'+x.command,payload:x.payload}),{command,payload,now,op:(++sequence).toString(16).padStart(32,'0')}));}
 function plan(request,apply=true){const out=decode(JSON.parse(h.domain.run(x=>{
  const docs=Object.fromEntries(x.docs.map(([p,d])=>[p,codec.decode(d)])),get=p=>docs[p],M=f4ActivitiesModel_(),r=x.request,identity={uid:'CF_ADMIN',expiresAt:x.now+120000},q=f4ActivitiesPipeline_().authorization(get,r,identity,x.now,identity.expiresAt)[0],access=f4SharedAuthz_().authorize(get,q,x.now);
  const auth={identity:()=>({b:access.binding,h:access.actor.h,s:get('memberCurrentStates/'+access.binding.idPessoa),actor:{uid:identity.uid,personId:access.binding.idPessoa,personRef:M.ref('person',access.binding.idPessoa),bindingVersion:access.binding.bindingVersion},authorityBasis:access.proof}),ordinary:(g,uid,scope,cap)=>{M.need(cap===q.capability&&scope===q.target.scopeKey,'AUTHORIZATION_DENIED');f4SharedAuthz_().recheck(g,q,access.proof,x.now);return {g:{entries:{[cap]:access.proof}}};}};
  const raw={get:p=>({data:()=>get(p)}),getAll:ps=>ps.map(p=>({data:()=>get(p)}))};
  return codec.encode(f4ActivitiesCreate_({db:{transaction:()=>{throw Error('NO_REMOTE')}},clock:()=>x.now,authority:auth,principals:{resolve:()=>({uid:identity.uid,kind:'BOARD_ACTOR'})},executionGate:()=>{}}).plan(raw,{},f4ActivitiesPipeline_().command(r)));
 },{docs:Object.entries(h.docs).map(([p,d])=>[p,encode(d)]),request,now})));
 if(apply&&out.response.businessState==='COMMIT_APPLIED')for(const w of out.writes)h.docs[w.path]=w.data;return out;}
 return {...h,now,activity,prepare,plan};
}
test('bounded roster, draft and finalization preserve previous justification then close without absence',()=>{
 const s=scenario(),r=s.prepare();a.equal(r.target.expected.length,4);a(!r.target.expected.some(x=>x.kind==='cycle'||x.kind==='person'));
 const ok=r=>{const out=s.plan(r);a.equal(out.response.businessState,'COMMIT_APPLIED',JSON.stringify(out.response));};ok(r);
 const c=()=>s.docs['activityAttendanceCollections/SYN_CALL'],j=()=>s.docs['activityJustifications/SYN_PRIOR'];a.equal(c().rosterCount,21);a.equal(c().preRosterFactCount,1);a.equal(j().state,'PREVIOUS');
 ok(s.prepare('ATTENDANCE_SAVE_DRAFT',{activityRef:s.activity,collectionRevision:c().revision,records:{SYN_FACT_0:{statusPresenca:'PRESENTE_PRESENCIAL',observacoes:''}}}));a.equal(j().state,'PREVIOUS');a.equal(j().basisAttendanceRevision,s.docs['activityAttendance/SYN_FACT_0'].revision);
 ok(s.prepare('ATTENDANCE_FINALIZE',{activityRef:s.activity,collectionRevision:c().revision,records:{}}));a.equal(j().state,'CLOSED_WITHOUT_ABSENCE');a.equal(c().state,'CLOSED');a.equal(c().preRosterFactCount,1);a.equal(s.docs['activityAttendance/SYN_FACT_0'].status,'PRESENT');
});
test('membership epoch and activity revision changes reject stale roster without writes',()=>{for(const path of ['presentationCapacityHeads/registry','canonicalActivities/SYN_CALL']){const s=scenario(),r=s.prepare();s.docs[path].revision++;if(path.includes('registry'))s.docs[path].membersEpoch++;a.throws(()=>s.plan(r),/STALE_REVISION/);}});
test('revoked authority still rejects preparation',()=>{const s=scenario();s.docs['memberAuthzPolicies/CF_DEV_ACT_POLICY'].active=false;a.throws(()=>s.prepare(),/AUTHORIZATION_DENIED/);});
test('draft replay, changed collection and expired window cannot silently overwrite facts',()=>{
 const s=scenario();s.plan(s.prepare());const payload={activityRef:s.activity,collectionRevision:s.docs['activityAttendanceCollections/SYN_CALL'].revision,records:{}};
 const r=s.prepare('ATTENDANCE_SAVE_DRAFT',payload);a.equal(s.plan(r).response.businessState,'COMMIT_APPLIED');a.throws(()=>s.plan(r),/STALE_REVISION/);
 s.docs['activityAttendanceCollections/SYN_CALL'].closesAt=s.now-1;
 const out=s.plan(s.prepare('ATTENDANCE_FINALIZE',{...payload,collectionRevision:s.docs['activityAttendanceCollections/SYN_CALL'].revision}));a.equal(out.response.error.reasonCode,'CHAMADA_FORA_DA_JANELA');a.equal(out.writes.length,0);
});
test('canonical discovery examines 99 synthetic heads and seals 21 active participants',()=>{
 const s=scenario(),docs=s.docs,{createDomain}=require('../src/domain.cjs');
 for(const p of Object.keys(docs))if(p.startsWith('memberAggregateHeads/')&&!p.endsWith('/PES-000001'))delete docs[p];
 const person=docs['memberIdentities/PES-000001'],link=docs['memberLinks/VINC-001'];
 for(let i=1;i<99;i++){
  const id='PES-'+String(990000+i).padStart(6,'0'),lid='SYN_LINK_'+i,ip='memberIdentities/'+id,lp='memberLinks/'+lid;
  docs[ip]={...person,id,idPessoa:id};docs[lp]={...link,id:lid,idPessoa:id,...(i>20?{endDate:'2026-01-02'}:{})};
  docs['memberAggregateHeads/'+id]={schemaVersion:'m11-head-v1',id,kind:'PERSON',revision:1,paths:[ip,lp]};
 }
 const rolePath='memberRoleCatalog/SYN_PRIOR_ROLE';docs[rolePath]={...docs['memberRoleCatalog/CARGO-001'],id:'SYN_PRIOR_ROLE'};
 const prior={...docs['memberFunctions/VIG-001'],id:'SYN_PRIOR_FN',idPessoa:'PES-990001',idVinculo:'SYN_LINK_1',idCargo:'SYN_PRIOR_ROLE',schemaVersion:'m1-memberFunctions-v2'};
 docs['memberEvents/SYN_CORRECTION']={schemaVersion:'m1-memberEvents-v1',id:'SYN_CORRECTION',revision:1,origin:'SYNTHETIC_M1',idPessoa:'PES-990001',idVinculo:'SYN_LINK_1',type:'RETIFICACAO',effect:null,effectiveDate:'2026-11-11',endDate:null,supersedes:null,correctionPath:'memberFunctions/SYN_PRIOR_FN',prior};docs['memberAggregateHeads/PES-990001'].paths.push('memberEvents/SYN_CORRECTION');
 const staged=new Map(),domain=createDomain(raw=>{const q=JSON.parse(raw);if(q.kind==='inspect')return JSON.stringify({environment:'DEV',projectId:'geapa-convergencia-dev',database:'(default)',host:'firestore.googleapis.com',port:443});if(q.kind==='pilotConfig')return JSON.stringify({root:''});if(['guard','getAll'].includes(q.kind))return 'null';if(q.kind==='get')return JSON.stringify(staged.get(q.path)|| (docs[q.path]?encode(docs[q.path]):null));if(q.kind==='write'){staged.set(q.path,q.data);return 'null';}if(q.kind==='query'){const cursor=q.plan.cursor?.[0]?.referenceValue?.split('/documents/')[1];return JSON.stringify(Object.keys(docs).filter(p=>p.startsWith('memberAggregateHeads/')&&(!cursor||p>cursor)).sort().slice(0,q.plan.limit).map(path=>({path,data:encode(docs[path])})));}throw Error('UNEXPECTED_'+q.kind);});
 const run=()=>JSON.parse(domain.run(x=>{const options={config:nodeConfig,identity:{uid:'CF_ADMIN',expiresAt:x.now+120000},settings:{commands:['ATTENDANCE_ROSTER_APPLY'],policyIds:['CF_DEV_ACT_POLICY']},clock:()=>x.now};options.ports=f4ActivitiesFunctionalPorts_(options);return f4ActivitiesAttendanceRoster_(options).prepare(x.activity);},{now:s.now,activity:s.activity}));
 const r=run(),sealed=decode(staged.get('activityAttendanceRosters/'+r.id));a.equal(sealed.sourceHeads.length,99);a.equal(Object.keys(sealed.records).length,21);a.equal(sealed.state,'SEALED');a.equal(docs['activityAttendanceCollections/SYN_CALL'].preRosterFactCount,1);
 delete docs[rolePath];a.throws(run,/UNRESOLVED_DEPENDENCY/);
});
module.exports={scenario};
