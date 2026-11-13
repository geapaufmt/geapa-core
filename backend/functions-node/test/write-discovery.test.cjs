'use strict';
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GEAPA_PILOT_MODE:'SHARED_DEV',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {test}=require('node:test'),a=require('node:assert/strict'),{Store,encode}=require('../src/firestore.cjs'),{Timestamp}=require('firebase-admin/firestore');
function fixture(cycle={id:'SYN_CYCLE',status:'OPEN'}){
 const docs={'presentationCycles/SYN_CYCLE':cycle},events=[];
 const db={projectId:'demo-geapa-functions-local',doc:path=>({path})};
 const native={getAll:async(...refs)=>{events.push({read:refs.map(r=>r.path)});return refs.map(ref=>({ref,exists:!!docs[ref.path],data:()=>docs[ref.path],updateTime:Timestamp.fromMillis(1000)}));},set:(r,d)=>events.push({write:r.path,data:d})};
 class S extends Store{guard(){}validateWrites(s){const c=this.hostRead('presentationCycles/SYN_CYCLE');if(s.missing.size)return;if(c?.status!=='OPEN')throw Object.assign(Error('PILOT_RESOURCE_DENIED'),{code:'PILOT_RESOURCE_DENIED'});}}
 const store=new S(db,{uid:'SYN'});store.preparation=true;
 const run=()=>store.evaluate(()=>{store.io(JSON.stringify({kind:'write',path:'activityAttendanceRosters/SYN_ROSTER',mode:'set',data:encode({complete:true})}));return 'prepared';},native);
 return {store,run,events};
}
test('roster validation discovers cycle once in native transaction and only then permits a write',async()=>{const h=fixture();a.equal(await h.run(),'prepared');a.equal(h.store.metrics.writes,1);a(h.store.metrics.discoveryRounds<6);const reads=h.events.filter(x=>x.read);a.equal(reads.filter(x=>x.read.includes('presentationCycles/SYN_CYCLE')).length,1);a.equal(h.events.at(-1).write,'activityAttendanceRosters/SYN_ROSTER');});
test('missing or forbidden cycle still denies roster and commits nothing',async()=>{for(const c of [null,{id:'SYN_CYCLE',status:'CLOSED'}]){const h=fixture(c);await a.rejects(h.run(),e=>e.code==='PILOT_RESOURCE_DENIED');a.equal(h.store.metrics.writes,0);a(!h.events.some(e=>e.write));a(h.store.metrics.discoveryRounds<6);}});
test('validation discovery keeps the existing 720-document budget',async()=>{const h=fixture();h.store.validateWrites=s=>{for(let i=0;i<721;i++)h.store.hostRead('presentationCycles/SYN_'+i);};await a.rejects(h.run(),e=>e.code==='SCOPE_TOO_LARGE');a.equal(h.store.metrics.writes,0);a(!h.events.some(e=>e.write));});
