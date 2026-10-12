'use strict';
require('node:dns').setDefaultResultOrder('ipv4first');
const fs=require('node:fs'),p=require('node:path'),a=require('node:assert/strict'),crypto=require('node:crypto');
const S=require('./sdk.cjs'),W=require('./wire.cjs'),{Host}=require('./sdk-host.cjs');
const root=__dirname,save=(n,v)=>fs.writeFileSync(p.join(root,'evidence',n),JSON.stringify(v,null,2)+'\n');
const read=n=>JSON.parse(fs.readFileSync(p.join(root,'evidence',n)));
async function main(){
 const before=require('./scripts/preflight.cjs').localManifest();a.deepEqual(before,read('local-before.json').files);
 const opened=await S.open(),{db}=opened;
 const {initializeApp,deleteApp}=require('firebase-admin/app'),{getAuth}=require('firebase-admin/auth');
 const app=initializeApp({projectId:S.PROJECT,credential:opened.credential},'B0_LOCAL');
 const P=require('./private-session.cjs');
 const secret=P.secrets();a.equal(secret.project,S.PROJECT);const account=secret.accounts.BOARD;a(account&&account.uid);
 const login=(await P.authApi('signInWithPassword',{email:account.email,password:account.password,returnSecureToken:true})).value;
 a.equal(login.localId,account.uid);const token=login.idToken,identity=await getAuth(app).verifyIdToken(token,true);a.equal(identity.uid,account.uid);a.equal(identity.aud,S.PROJECT);
 const host=new Host(db,t=>getAuth(app).verifyIdToken(t,true)),journalFile=p.join(root,'evidence/seed.json');
 let journal=fs.existsSync(journalFile)?read('seed.json'):null;
 if(!journal){
  const seed=require('./fixtures.cjs').seed(account.uid),paths=Object.keys(seed),existing=await db.getAll(...paths.map(v=>db.doc(S.ROOT+'/'+v)));a(existing.every(d=>!d.exists),'B0_FIXTURE_ALREADY_EXISTS');
  journal={at:new Date().toISOString(),project:S.PROJECT,namespace:S.ROOT,uid:account.uid,paths,seedIntent:true,seeded:false};save('seed.json',journal);
  const b=db.batch();for(const [path,d]of Object.entries(seed))b.create(db.doc(S.ROOT+'/'+path),W.decode({mapValue:{fields:d.fields}},db));await b.commit();journal.seeded=true;save('seed.json',journal);
 }else{a(journal.seeded&&!journal.cleaned,'RECONCILE_EXISTING_B0_FIXTURE');a.equal(journal.uid,account.uid);}
 const evidenceFile='edit-final.json',e=fs.existsSync(p.join(root,'evidence',evidenceFile))?read(evidenceFile):{project:S.PROJECT,namespace:S.ROOT,actor:'Existing synthetic BOARD Firebase DEV account with synthetic President role in B0 only',samples:[],checks:[],operations:[]};
 try{
  const gateRef=db.doc(S.ROOT+'/portalReaderControl/current');if(!(await gateRef.get()).exists){journal.paths.push('portalReaderControl/current');save('seed.json',journal);await gateRef.create(W.decode({mapValue:{fields:require('./fixtures.cjs').seed(account.uid)['portalReaderControl/current'].fields}},db));}
  for(let i=e.samples.length;i<4;i++){
   const d=await db.doc(S.ROOT+'/canonicalActivities/dev_meeting').get(),expectedRevision=d.data().revision,operationId=crypto.randomBytes(16).toString('hex'),title='B0_NODE_POC_EDIT_'+i;
   e.operations.push(operationId);save(evidenceFile,e);host.metrics=[];host.written.clear();host.observedBefore.clear();db.b0RpcEvents=[];const clock=Date.now(),started=performance.now();
   const outcome=await host.execute({mode:'edit',token,operationId,activityId:'dev_meeting',expectedRevision,title});
   const wallMs=performance.now()-started,rpcs=db.b0RpcEvents.map(v=>v.method);a.equal(outcome.result.state,'COMPLETED',JSON.stringify(outcome.result));
   const readStart=performance.now(),back=await db.getAll(db.doc(S.ROOT+'/canonicalActivities/dev_meeting'),db.doc(S.ROOT+'/commandReceipts/'+operationId));
   a.equal(back[0].data().title,title);a.equal(back[0].data().revision,expectedRevision+1);a(back[1].exists);
   const invalidation=await db.doc(S.ROOT+'/commandProjectionInvalidations/'+operationId).get();a(invalidation.exists);a(invalidation.data().fences.includes('activity_dev_meeting'));const sample={index:i,coldLike:i===0,wallMs,readbackMs:performance.now()-readStart,...outcome,rpcs,io:host.metrics};e.samples.push(sample);save(evidenceFile,e);
   if(!e.parity){e.parity=await require('./parity.cjs').parity({db},host,token,outcome,clock,'edit');save(evidenceFile,e);}
   console.log({sample:i,state:outcome.result.state,wallMs:Math.round(wallMs),phases:outcome.phases,reads:host.metrics.reduce((s,m)=>s+m.reads,0),writes:host.metrics.reduce((s,m)=>s+m.writes,0)});
   if(i>=1&&wallMs>15000){e.earlyDecision='INVESTIGATE_EDIT_OVER_15S';save(evidenceFile,e);break;}
  }
 }finally{journal.writtenPaths=[...new Set([...(journal.writtenPaths||[]),...host.touchedAll])];save('seed.json',journal);await db.terminate();await deleteApp(app);a.deepEqual(require('./scripts/preflight.cjs').localManifest(),before);}
}
if(require.main===module)main().catch(e=>{save('last-error.json',{at:new Date().toISOString(),code:e.code||e.message,detail:e.detail,stack:e.stack?.split('\n').slice(0,9)});console.error({code:e.code||e.message,detail:e.detail,stack:e.stack?.split('\n').slice(0,5)});process.exitCode=1;});
