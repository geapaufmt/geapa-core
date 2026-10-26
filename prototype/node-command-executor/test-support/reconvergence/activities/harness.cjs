const fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),crypto=require('node:crypto'),assert=require('node:assert/strict');
const ROOT=path.resolve(__dirname,'../../..');
global.Utilities={DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},computeDigest:(a,s)=>[...crypto.createHash('sha256').update(s).digest()],newBlob:s=>({getBytes:()=>[...Buffer.from(s)]}),formatDate:(d,z,pattern)=>{const parts=new Intl.DateTimeFormat('en-CA',{timeZone:z,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hourCycle:'h23'}).formatToParts(d),v=Object.fromEntries(parts.map(p=>[p.type,p.value]));return `${v.year}-${v.month}-${v.day}`+(pattern.includes('HH')?`T${v.hour}:${v.minute}:${v.second}`:'');},sleep:()=>{}};
global.ScriptApp={getOAuthToken:()=> 'LOCAL_SYNTHETIC_ONLY'};
global.UrlFetchApp={fetch:()=>{throw Error('REMOTE_NETWORK_FORBIDDEN');}};
vm.runInThisContext(fs.readdirSync(path.join(ROOT,'base/core'),{recursive:true}).filter(f=>f.endsWith('.gs')).sort().map(file=>fs.readFileSync(path.join(ROOT,'base/core',file),'utf8')).join('\n;\n'),{filename:'unified-core.gs'});
const M=f4ActivitiesModel_(),F=f4Firestore_(),T=F.Timestamp.fromMillis,ref=M.ref;
class Store{
 constructor(docs={}){this.docs=new Map(Object.entries(docs).map(([k,v])=>[k,M.clone(v)]));this.versions=new Map([...this.docs.keys()].map(k=>[k,1]));this.traces=[];this.hook=null;}
 get(p){return M.clone(this.docs.get(p));}
 put(p,v){if(v===undefined)this.docs.delete(p);else this.docs.set(p,M.clone(v));this.versions.set(p,(this.versions.get(p)||0)+1);}
 transaction(fn){const reads=new Map(),writes=new Map();let writing=false;const tx={get:p=>{assert(!writing,'Read after write');if(!reads.has(p))reads.set(p,this.versions.get(p)||0);return {data:()=>this.get(p)};},set:(p,d)=>{writing=true;writes.set(p,M.clone(d));},delete:p=>{writing=true;writes.set(p,undefined);}};const result=fn(tx);if(this.hook){const hook=this.hook;this.hook=null;hook(this);}for(const [p,v]of reads)if((this.versions.get(p)||0)!==v)M.fail('STALE_REVISION','CAS');for(const [p,d]of writes)this.put(p,d);this.traces.push({reads:[...reads.keys()],writes:[...writes.keys()]});return result;}
 snapshot(){return JSON.stringify([...this.docs]);}
 domainSnapshot(){return JSON.stringify([...this.docs].filter(([p])=>!p.startsWith('command')));}
}
const NOW=Date.parse('2026-10-05T12:00:00Z');
function fixture(){
 const docs={},r={cycle:ref('cycle','c'.repeat(32)),policy:ref('policy','policy'),config:ref('config','config','c'.repeat(32)),session:ref('session','session'),session2:ref('session','session2'),unit:ref('unit','unit'),unit2:ref('unit','unit2'),obligation:ref('obligation','obligation'),assignment:ref('assignment','assignment'),presentation:ref('presentation','presentation'),theme:ref('thematic','theme'),axis:ref('policy','axis'),flow:ref('flow','flow'),requirement:ref('requirement','requirement'),manifest:ref('assessment','manifest'),outbox:ref('outbox','effect'),realization:ref('realization','realization'),submission:ref('submission','submission'),review:ref('review','review'),opportunity:ref('opportunity','opportunity'),artifact:ref('artifact','artifact'),capacity:ref('capacity','c'.repeat(32)),registry:ref('capacity','registry'),draft:ref('draft','draft'),dynamic:ref('dynamic','dynamic'),receipt:ref('receipt','receipt'),regularization:ref('regularization','regularization'),publication:ref('publication','presentation'),snapshot:ref('snapshot','snapshot')};
 const person={kind:'person',id:'p1'};const put=(rr,data)=>{docs[M.path(rr)]={id:rr.id,revision:1,cycleRef:r.cycle,...data};};
 put(r.registry,{kind:'DOMAIN_FENCE',membersEpoch:1,cycleRegistryEpoch:1});
 put(r.cycle,{status:'OPEN',currentConfigRef:r.config,civilStart:'2026-01-01',civilEnd:'2027-12-31',policyVersionRef:r.policy});
 put(r.policy,{approved:true,active:true,commands:require('../../fixture-command-catalog.json'),publicScope:'PUBLIC',permittedDisplay:'Approved public label',analysisHours:48});
 put(r.axis,{approved:true});
 const params={calendarMode:'MEMBER_SELF_SCHEDULE',timezone:'America/Cuiaba',shared:true,quota:false,early:true,dynamic:true,dynamicLimit:2,clockPolicyRef:r.policy,publicationPolicyRef:r.policy};
 put(r.config,{approved:true,parameters:params,parameterClasses:'OPERATIONAL',effectiveFrom:T(NOW-1000)});
 put(r.capacity,{state:'SEALED',capacityEpoch:1,dynamicActiveCount:0,dynamicHistoricalCount:0,synchronizedMembersEpoch:1,synchronizedCycleRegistryEpoch:1,validUntil:T(NOW+86400000)});
 put(r.session,{civilDate:'2026-12-20',config:{capacity:2,destination:'ORDINARY'},slots:[],configRevision:1,compositionRevision:1});
 put(r.session2,{civilDate:'2026-12-21',config:{capacity:2,destination:'ORDINARY'},slots:[],configRevision:1,compositionRevision:1});
 put(r.obligation,{personRef:person,unitRef:r.unit,incidenceState:'INCIDENT',result:'PENDING'});
 put(r.unit,{personRef:person,obligationRef:r.obligation,presentationRef:r.presentation,thematicRequestRef:r.theme});
 put(r.unit2,{personRef:{kind:'person',id:'p2'},preAgendaBasisRef:r.policy});
 put(r.presentation,{personRef:person,unitRef:r.unit,materialRequirementRef:r.requirement,title:'Fictitious title',recordMode:'NEW_V2'});
 put(r.requirement,{personRef:person,presentationRef:r.presentation,currentRevisionRef:{id:r.requirement.id,number:1},materialStateRevision:1,reviewEpoch:1,opportunitiesEpoch:1,openCorrectionCount:0,items:[{itemId:'document'}],finalEndExclusive:T(NOW+86400000),uploadPolicyRef:r.policy});
 put(r.theme,{personRef:person,unitRef:r.unit,workflowState:'ANALYSIS',principalAxisRef:r.axis,submittedRevisionRef:{id:r.theme.id,number:1},priorityBasis:{completeAt:T(NOW-1000)},manualDecisionEpoch:0,analysisClock:{state:'RUNNING',policyRef:r.policy,dueAt:T(NOW-1000),manualDecisionEpoch:0}});
 put(r.manifest,{state:'SEALED',complete:true,items:[],digest:M.hash([]),validUntil:T(NOW+86400000),membersEpoch:1,cycleRegistryEpoch:1,configRef:r.config,personRef:person,feasible:true,eligibleDate:true,preservesAcquiredRights:true,proposedChange:{unitRef:r.unit,sessionRef:r.session},presentationRef:r.presentation,requirementRevisionRef:{id:r.requirement.id,number:1},completeCoverage:true});
 put(r.flow,{kind:'JOINT_PROPOSAL',state:'OPEN',unitRefs:[r.unit,r.unit2],proposalRevision:1,deadline:T(NOW+86400000)});
 put(r.outbox,{status:'PENDING',state:'PENDING',effectKind:'DRIVE_CREATE',targetRef:r.presentation,policyRef:r.policy,policyRevision:1,leaseFence:0,attemptCount:0,nextAttemptAt:T(NOW),authorityBasis:{},details:{}});
 docs['memberAuthBindings/u1']={uid:'u1',idPessoa:'p1',active:true,bindingVersion:1,revision:1,sourceRevision:1,domainVersion:'m11'};
 docs['memberAuthBindings/u2']={...docs['memberAuthBindings/u1'],uid:'u2',idPessoa:'p2'};
 for(const id of ['p1','p2']){docs['memberAggregateHeads/'+id]={revision:1};docs['memberCurrentStates/'+id]={idPessoa:id,state:'MEMBRO_EFETIVO',sourceRevision:1,accessEpoch:1,stale:false,blocked:false,startsAt:T(NOW-86400000),validUntil:T(NOW+86400000)};docs['memberIdentities/'+id]={active:true,technicalAccount:false};docs['portalUsers/'+(id==='p1'?'u1':'u2')]={idPessoa:id,active:true,bindingVersion:1,technicalAccount:false};}
 const db=new Store(docs),handles=new WeakMap();let tick=NOW,seq=0;
 const issue=(kind='BOARD_ACTOR',uid='u1',extra={})=>{const h={};handles.set(h,{kind,uid,...extra});return h;};
 const auth=f4ActivitiesAuthority_();const authority={...auth,identity:(get,uid,now)=>{const b=get('memberAuthBindings/'+uid),h=b&&get('memberAggregateHeads/'+b.idPessoa),st=b&&get('memberCurrentStates/'+b.idPessoa),i=b&&get('memberIdentities/'+b.idPessoa);const portal=get('portalUsers/'+uid);M.need(b?.active&&h&&st&&!st.stale&&i?.active&&!i.technicalAccount&&portal?.bindingVersion===b.bindingVersion&&portal.idPessoa===b.idPessoa&&h.revision===st.sourceRevision&&M.ms(st.startsAt)<=now&&now<M.ms(st.validUntil),'AUTHORIZATION_DENIED');return {b,h,s:st,i,actor:{uid,personId:b.idPessoa,personRef:ref('person',b.idPessoa),bindingVersion:b.bindingVersion},authorityBasis:{personRevision:h.revision,bindingVersion:b.bindingVersion,state:st.state}};},ordinary:(get,uid,scope,cap,now)=>{const b=get('memberAuthBindings/'+uid),s=get('memberCurrentStates/'+b.idPessoa);M.need(b.active&&!s.blocked&&s.state==='MEMBRO_EFETIVO','AUTHORIZATION_DENIED');return {g:{entries:{[cap]:{functionId:'FAKE_AUTHORITY',scope}}}};},suspension:(get,h)=>{M.need(h.i.active&&h.s.restriction==='SUSPENSION_ONLY','AUTHORIZATION_DENIED');return {kind:'FAKE_MEMBERS_RESTRICTION',personRevision:h.h.revision};}};
 const producer=f4ActivitiesCreate_({executionGate:()=>{},db,clock:()=>tick,authority,principals:{resolve:h=>handles.get(h)}});
 const core={envelope:producer.envelope,execute:(handle,q)=>{const v=db.transaction(tx=>{const v=producer.plan(tx,handle,q);for(const w of v.writes)w.kind==='delete'?tx.delete(w.path):tx.set(w.path,w.data);return v.response;});if(v.error){const e=Error(v.error.code);e.code=v.error.code;e.reasonCode=v.error.reasonCode;throw e;}return v;}};
 db.put(M.path(r.manifest),{...db.get(M.path(r.manifest)),basisEpoch:1});
 function update(rr,change){const k=typeof rr==='string'?rr:M.path(rr),d=db.get(k);db.put(k,{...d,...change});}
 function manifest(data){const items=data.items||[];update(r.manifest,{...data,items,digest:M.hash(items)});return r.manifest;}
 function request(cmd,p,op){const expected=[],targetRefs=[];for(const [key,d]of db.docs){const found=Object.entries(M.collections).find(([k,c])=>key===c+'/'+d.id);if(found&&d.revision!==undefined&&key.split('/').length===2&&(!d.personRef||d.personRef.id==='p1'||['JOINT_PROPOSE','JOINT_CONSENT','COMPOSITION_CONFIRM','CALENDAR_SWAP_CONFIRM','COMPATIBILITY_DECIDE','CALENDAR_MOVE'].includes(cmd))){const rr=ref(found[0],d.id);expected.push({...rr,revision:d.revision});targetRefs.push(rr);}}
  expected.push({...r.config,revision:db.get(M.path(r.config)).revision});return {schemaVersion:'geapa-command-v2-candidate',operationId:op||(++seq).toString(16).padStart(32,'0'),commandType:'ACT.'+cmd,payload:p,expected,scope:{targetRefs:targetRefs.slice(0,32)}};}
 function execute(cmd,p,handle=issue(),q=request(cmd,p)){return core.execute(handle,q);}
 return {db,r,core,issue,authority,principal:h=>handles.get(h),update,manifest,request,execute,setTime:n=>tick=n,clock:()=>tick,put:(rr,d)=>db.put(M.path(rr),{id:rr.id,revision:1,cycleRef:r.cycle,...d})};
}
module.exports={ROOT,M,F,T,ref,NOW,Store,fixture,assert};
