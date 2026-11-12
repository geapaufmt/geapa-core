'use strict';
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GEAPA_PILOT_MODE:'SHARED_DEV',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {test}=require('node:test'),a=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
const {Store}=require('../src/firestore.cjs'),{Timestamp}=require('firebase-admin/firestore'),{createDomain}=require('../src/domain.cjs');
const {fixture}=require('./institutional-fixture.cjs');
function setup(){
 const h=fixture('PRESIDENTE'),docs=h.docs,ids=[];
 // Sixteen independent canonical aggregates, each with ten real synthetic
 // source documents. No real person or authorization is represented here.
 for(let n=0;n<16;n++){
  const id='SYN_PERSON_'+n,link='SYN_LINK_'+n;ids.push(id);
  docs['memberIdentities/'+id]={...docs['memberIdentities/PES-000001'],id,idPessoa:id,displayName:'Synthetic '+n};
  docs['memberLinks/'+link]={...docs['memberLinks/VINC-001'],id:link,idPessoa:id};
  const paths=['memberIdentities/'+id,'memberLinks/'+link];
  for(let i=0;i<8;i++){const p='memberLinks/SYN_OLD_'+n+'_'+i;docs[p]={...docs['memberLinks/VINC-001'],id:p.split('/')[1],idPessoa:id,startDate:(2010+i)+'-01-01',endDate:(2010+i)+'-12-31',exitDate:null};paths.push(p);}
  docs['memberAggregateHeads/'+id]={schemaVersion:'m11-head-v1',id,kind:'PERSON',revision:1,paths};
 }
 const requests=[];
 const db={projectId:'demo-geapa-functions-local',doc:p=>({path:p}),getAll:async(...refs)=>{requests.push(refs.map(x=>x.path));return refs.map(ref=>({ref,exists:!!docs[ref.path],data:()=>docs[ref.path],updateTime:Timestamp.fromMillis(h.now)}));}};
 class S extends Store{guard(){} }
 const store=new S(db,{uid:'SYN'}),domain=createDomain(x=>store.io(x));
 vm.runInContext(fs.readFileSync(path.resolve(__dirname,'../../../institutional/211_board_procedures.gs'),'utf8'),domain.context);
 const read=()=>store.read(()=>JSON.parse(domain.run('(x)=>f4BoardProcedures_().candidates(p=>nodeTx.get(p).data(),{config:nodeConfig,clock:()=>x.now},{actor:{derived:{currentDirectors:[]}}},{boardProcedures:{enabled:true,targetPersonIds:x.ids}})',{now:h.now,ids})));
 return {read,store,docs,requests,ids};
}
test('sixteen governance candidate aggregates converge without serial document discovery',async()=>{
 const h=setup();const out=await h.read();a.equal(out.length,16);a(out.every(x=>x.eligible));a(h.store.metrics.discoveryRounds<20);a.equal(h.store.metrics.writes,0);
});
test('absent canonical facts still fail closed after bounded prefetch',async()=>{
 const h=setup();delete h.docs['memberLinks/SYN_LINK_3'];await a.rejects(h.read());a(h.store.metrics.discoveryRounds<20);a.equal(h.store.metrics.writes,0);
});
test('invalid aggregate path is rejected before fetching a foreign collection',async()=>{
 const h=setup();h.docs['memberAggregateHeads/'+h.ids[2]].paths.push('privateSecrets/SENTINEL');await a.rejects(h.read(),e=>e.code==='DOMAIN_STATE_INVALID');a(!h.requests.flat().includes('privateSecrets/SENTINEL'));a.equal(h.store.metrics.writes,0);
});
test('governance HTTP diagnostic correlates response and log without payload or identity',async()=>{
 const {handler,ORIGIN}=require('../src/http.cjs');const trace=Array.from({length:120},(_,i)=>({round:i+1,newDocuments:1,newQueries:0,cachedDocuments:i,cachedQueries:0,path:'PRIVATE_SENTINEL'}));const e=Object.assign(Error('RAW_PRIVATE_SENTINEL'),{code:'READ_DISCOVERY_LIMIT',readDiscovery:{rounds:120,trace,uid:'PRIVATE_SENTINEL'}});
 const req={method:'POST',headers:{origin:ORIGIN,authorization:'Bearer PRIVATE_SENTINEL','content-type':'application/json'},body:{action:'actRead',request:{route:'/admin/gestao-institucional',params:{section:'governance',query:'PRIVATE_SENTINEL'}}}};
 const res={headers:{},set(k,v){this.headers[k]=v},status(n){this.statusCode=n;return this},json(v){this.body=v;return this}},logs=[];
 await handler({runtime:async()=>({}),execute:async()=>{throw e},actions:['actRead'],log:x=>logs.push(x)})(req,res);
 a.equal(res.statusCode,503);const log=JSON.parse(logs[0]);a.equal(log.route,'/admin/gestao-institucional');a.equal(log.correlationId,res.headers['X-GEAPA-Correlation-ID']);a.equal(log.correlationId,res.body.error.correlationId);a.equal(log.readDiscovery.trace.length,16);a(!JSON.stringify(logs).includes('PRIVATE_SENTINEL'));a(!JSON.stringify(res.body).includes('newDocuments'));
});
module.exports={setup};
