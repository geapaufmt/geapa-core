'use strict';
const {Timestamp}=require('firebase-admin/firestore');
const {SHARED,ROOT,ACTIVITY,EXECUTOR,fence,fail}=require('./environment.cjs');
const pilot=require('./pilot.cjs');
const {queryPlan}=require('./query.cjs');
function encode(v){
 if(v===null)return {nullValue:'NULL_VALUE'};
 if(v instanceof Timestamp)return {timestampValue:new Date(v.seconds*1000).toISOString().replace('.000Z','.'+String(v.nanoseconds).padStart(9,'0')+'Z')};
 if(v instanceof Date)return encode(Timestamp.fromDate(v));
 if(Array.isArray(v))return {arrayValue:{values:v.map(encode)}};
 if(typeof v==='object')return {mapValue:{fields:Object.fromEntries(Object.entries(v).map(([k,x])=>[k,encode(x)]))}};
 if(typeof v==='string')return {stringValue:v};if(typeof v==='boolean')return {booleanValue:v};
 if(typeof v==='number'&&Number.isFinite(v))return Number.isInteger(v)?{integerValue:String(v)}:{doubleValue:v};fail('SCHEMA_INVALID');
}
function decode(v){if('nullValue'in v)return null;if('timestampValue'in v){const s=v.timestampValue;return new Timestamp(Math.floor(Date.parse(s)/1000),Number((s.match(/\.(\d+)Z$/)?.[1]||'').padEnd(9,'0')));}
 if('mapValue'in v)return Object.fromEntries(Object.entries(v.mapValue.fields||{}).map(([k,x])=>[k,decode(x)]));if('arrayValue'in v)return (v.arrayValue.values||[]).map(decode);if('integerValue'in v)return Number(v.integerValue);if('doubleValue'in v)return v.doubleValue;if('booleanValue'in v)return v.booleanValue;if('stringValue'in v)return v.stringValue;fail('SCHEMA_INVALID');}
class Store {
 constructor(db,identity){this.environment=fence();if(db.projectId!==this.environment.projectId)fail('DESTINATION_INVALID');this.db=db;this.identity=identity;this.action='recover';this.active=null;this.metrics={reads:0,writes:0,sdkCalls:0,transactions:0,transactionAttempts:0,discoveryRounds:0,firestoreMs:0};}
 hostRead(path){const s=this.active;this.ref(path);s.touched.add(path);if(!s.cache.has(path)){s.missing.add(path);return undefined;}return s.cache.get(path).encoded?decode(s.cache.get(path).encoded):undefined;}
 guard(){
  fence();if(!this.identity)fail('AUTH_BINDING_INVALID');
  const c=this.hostRead(pilot.CONTROL),o=this.hostRead(pilot.OWNER),r=this.hostRead('pilotResources/'+ACTIVITY),a=this.hostRead('canonicalActivities/'+ACTIVITY);
  pilot.validateControl(c,o,this.identity.uid,this.action);pilot.validateResource(r,a);
  if(this.ownershipEpoch!==undefined&&this.ownershipEpoch!==o.epoch)fail('EXECUTOR_FENCE_CONFLICT');
  this.ownershipEpoch=o.epoch;this.control=c;return c;
 }
 ref(path){if(typeof path!=='string'||!/^([A-Za-z][A-Za-z0-9_]*\/[A-Za-z0-9_-]+)(\/[A-Za-z][A-Za-z0-9_]*\/[A-Za-z0-9_-]+)*$/.test(path)||path.startsWith('coreDevValidation/')||path.startsWith('projects/')||path.length>1024)fail('DESTINATION_INVALID');return this.db.doc((ROOT?ROOT+'/':'')+path);}
 io(raw){const x=JSON.parse(raw),s=this.active;if(!s)fail('TRANSACTION_REQUIRED');
  const read=path=>{this.ref(path);if(s.writes.length&&!this.preparation)fail('FIRESTORE_READ_AFTER_WRITE');s.touched.add(path);if(!s.cache.has(path)){s.missing.add(path);return null;}return (this.preparation?s.writes.findLast(w=>w.path===path)?.data:undefined)||s.cache.get(path).encoded;};
  let result;
  if(x.kind==='preparationClock'){if(!this.preparation||!s.native)fail('PREPARATION_REQUIRED');result=s.preparationAt;}
  else if(x.kind==='guard'){this.guard();result=null;}
  else if(x.kind==='pilotConfig'){result={root:ROOT,shared:SHARED,executor:EXECUTOR,epoch:this.ownershipEpoch,settings:{commands:['ACTIVITY_SAVE'],policyIds:this.control.policyIds,admissionsPerHour:20,jobs:[],activityAgendaDirect:false}};}
  else if(x.kind==='query'){
   if(s.writes.length&&!this.preparation)fail('FIRESTORE_READ_AFTER_WRITE');
   const key=JSON.stringify(x.plan),query=queryPlan(this.db,x.plan,decode,this.specializedReaders===true);
   if(!s.queries.has(key)){s.pendingQueries.set(key,query);result=[];}
   else {result=s.queries.get(key);result.forEach(r=>s.touched.add(r.path));}
  }
  else if(x.kind==='get')result=read(x.path);
  else if(x.kind==='getAll'){if(!Array.isArray(x.paths)||x.paths.length>180||new Set(x.paths).size!==x.paths.length)fail('SCHEMA_INVALID');x.paths.forEach(read);result=null;}
  else if(x.kind==='inspect'){if(!s.native||s.writes.length&&!this.preparation)fail('M13_TRANSACTION_REQUIRED');result={environment:require('./runtime-profile.cjs').profile.environment,projectId:require('./runtime-profile.cjs').profile.projectId,database:'(default)',readSet:[...s.touched].map(path=>{const v=s.cache.get(path);return {path,exists:!!v?.exists,updateTime:v?.stamp||null,revision:v?.revision??null};})};}
  else if(x.kind==='preview'){result={transaction:'NODE_SDK',writes:x.effects.map(w=>{this.ref(w.path);const old=s.cache.get(w.path),name='projects/'+require('./runtime-profile.cjs').profile.projectId+'/databases/(default)/documents/'+(ROOT?ROOT+'/':'')+w.path;return {...(w.kind==='delete'?{delete:name}:{update:{name,fields:w.data.mapValue.fields}}),currentDocument:old?.exists?{updateTime:old.stamp}:{exists:false}};})};}
  else if(x.kind==='write'){if(!s.native)fail('TRANSACTION_REQUIRED');this.ref(x.path);s.writes.push(x);result=null;}
  else fail('SDK_OPERATION_DENIED');return JSON.stringify(result);
 }
 async measure(fn){const start=performance.now();this.metrics.sdkCalls++;try{return await fn();}finally{this.metrics.firestoreMs+=performance.now()-start;}}
 validateWrites(s){pilot.validateWrites(s.writes,p=>this.hostRead(p),decode,this.identity.uid,this.projectionViews);}
 async evaluate(fn,native){
  const s={native,preparationAt:Date.now(),cache:new Map(),queries:new Map(),pendingQueries:new Map(),missing:new Set(),touched:new Set(),writes:[]};this.active=s;
  try{for(let round=0;round<120;round++){
   this.metrics.discoveryRounds++;s.missing.clear();s.pendingQueries.clear();s.touched.clear();s.writes=[];let value,error;
   const evalStart=performance.now();
   try{this.guard();value=fn();if(value&&typeof value.then==='function')fail('ASYNC_DOMAIN_CALLBACK');}catch(e){error=e;}
   if(this.specializedReaders)this.metrics.domainEvaluationMs=(this.metrics.domainEvaluationMs||0)+performance.now()-evalStart;
   // No outcome (including a caught denial) can commit using an unknown read.
   // For list maintenance, resolve the cached view/fence/auth reads before
   // launching a speculative full-list query from an incomplete discovery pass.
   // Missing values are still re-evaluated; no decision uses that partial cache.
   const directedList=this.viewId==='activities'||this.projectionViews?.includes('activities');
   if(s.pendingQueries.size&&!(directedList&&s.missing.size)){
    if(s.queries.size+s.pendingQueries.size>80)fail('SCOPE_TOO_LARGE');
    const pending=[...s.pendingQueries],width=this.queryConcurrency===4?4:1;
    for(let offset=0;offset<pending.length;offset+=width){
     const group=pending.slice(offset,offset+width);
     const snapshots=await Promise.all(group.map(([,query])=>this.measure(()=>native?native.get(query):query.get())));
     for(let index=0;index<group.length;index++){
     const [key]=group[index],snap=snapshots[index];this.metrics.queryCalls=(this.metrics.queryCalls||0)+1;this.metrics.reads+=snap.size;this.metrics.emptyQueries=(this.metrics.emptyQueries||0)+(snap.empty?1:0);
     const rows=snap.docs.map(d=>{const path=d.ref.path.slice(ROOT?ROOT.length+1:0),data=d.data(),encoded=encode(data),stamp=encode(d.updateTime).timestampValue;this.ref(path);s.cache.set(path,{exists:true,encoded,stamp,revision:data.revision??data.sourceRevision??null});return {path,data:encoded,updateTime:stamp};});
     s.queries.set(key,rows);
     }
    }
   }
   if(s.missing.size){const paths=[...s.missing].filter(p=>!s.cache.has(p));if(s.cache.size+paths.length>720)fail('SCOPE_TOO_LARGE');for(let i=0;i<paths.length;i+=180){const part=paths.slice(i,i+180);this.metrics.reads+=part.length;const docs=await this.measure(()=> (native||this.db).getAll(...part.map(p=>this.ref(p))));for(let j=0;j<part.length;j++){const d=docs[j],data=d.exists?d.data():undefined;s.cache.set(part[j],{exists:d.exists,encoded:d.exists?encode(data):null,stamp:d.exists?encode(d.updateTime).timestampValue:null,revision:data?.revision??data?.sourceRevision??null});}}continue;}
   if(s.cache.size>720)fail('SCOPE_TOO_LARGE');
   if(s.pendingQueries.size)continue;
   if(error)throw error;
   if(native){
    const unread=[...new Set(s.writes.map(w=>w.path))].filter(path=>!s.cache.has(path));
    if(unread.length){for(let i=0;i<unread.length;i+=180){const paths=unread.slice(i,i+180),docs=await this.measure(()=>native.getAll(...paths.map(path=>this.ref(path))));this.metrics.reads+=paths.length;docs.forEach((d,j)=>s.cache.set(paths[j],{exists:d.exists,encoded:d.exists?encode(d.data()):null,stamp:d.exists?encode(d.updateTime).timestampValue:null,revision:d.data()?.revision??null}));}continue;}
    // Same global-vs-targeted projection fence decision as shared/core/20.
    const affected=s.writes.some(w=>/^(member(Identities|Identifiers|PrivateProfiles|Links|Events|Functions|AuthBindings|CurrentStates|AuthorityControl|AuthzSubjects|AuthzPolicies|GovernanceBoards|GovernanceTerms|GovernanceBoardTerms|RoleCatalog)|canonicalActivities|canonicalPresentations|presentationArtifacts|migrationEvidence|activityAttendance|activityJustifications)\//.test(w.path));
    if(affected){const gate='portalReaderControl/current';if(!s.cache.has(gate)){const d=await this.measure(()=>native.get(this.ref(gate)));this.metrics.reads++;s.cache.set(gate,{exists:d.exists,encoded:d.exists?encode(d.data()):null,stamp:d.exists?encode(d.updateTime).timestampValue:null,revision:null});continue;}
     const g=s.cache.get(gate).encoded&&decode(s.cache.get(gate).encoded);const targeted=g?.operationalWrites===true&&!s.writes.some(w=>/^member(Identities|Identifiers|PrivateProfiles|Links|Events|Functions|AuthBindings|CurrentStates|AuthorityControl|AuthzSubjects|AuthzPolicies|GovernanceBoards|GovernanceTerms|GovernanceBoardTerms|RoleCatalog)\//.test(w.path))&&s.writes.some(w=>/^commandProjectionInvalidations\/[a-f0-9]{32}$/.test(w.path))&&s.writes.some(w=>w.path==='portalReaderResourceFences/activities');
     if(g?.enabled&&!targeted){if(s.writes.some(w=>w.path===gate))fail('SCHEMA_INVALID');s.writes.push({kind:'write',path:gate,mode:'update',data:encode({...g,generation:g.generation+1})});}
    }
    if(s.writes.length>400||new Set(s.writes.map(w=>w.path)).size!==s.writes.length)fail('SCHEMA_INVALID');
    this.validateWrites(s);
    if(s.writes.length>400||new Set(s.writes.map(w=>w.path)).size!==s.writes.length)fail('SCHEMA_INVALID');
    if(s.missing.size)continue;
    for(const w of s.writes){const ref=this.ref(w.path);if(!s.cache.has(w.path))fail('PRECONDITION_REQUIRED');if(w.mode==='delete')native.delete(ref);else if(w.mode==='update')native.update(ref,decode(w.data));else native.set(ref,decode(w.data));}this.metrics.writes+=s.writes.length;await this.beforeCommit?.();
   }return value;
  }fail('READ_DISCOVERY_LIMIT');}finally{this.active=null;}
 }
 async read(fn){return this.evaluate(fn,null);}
 async transaction(fn){this.metrics.transactions++;const started=performance.now();try{return await this.db.runTransaction(async native=>{this.metrics.transactionAttempts++;return this.evaluate(fn,native);},{maxAttempts:5});}finally{this.metrics.transactionWallMs=(this.metrics.transactionWallMs||0)+performance.now()-started;}}
}
module.exports={Store,encode,decode};
