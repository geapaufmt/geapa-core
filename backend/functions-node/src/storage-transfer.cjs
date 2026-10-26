'use strict';
const crypto=require('node:crypto');
const {ActivityStore}=require('./activity-commands.cjs');
const {encode,decode}=require('./firestore.cjs');
const {createDomain}=require('./domain.cjs');
const {authenticate,appCheckGate}=require('./auth.cjs');
const {fence,fail,EXECUTOR,PROJECT}=require('./environment.cjs');
const GCS=require('./workspace-storage.cjs');
const CONTROL='actStorageControl/current';
class StorageStore extends ActivityStore{
 guard(){
  const c=super.guard(),S=require('./storage-control.cjs'),s=S.check(p=>this.hostRead(p),this.identity.uid,c.expiresAt);
  if(this.storagePinned&&this.storagePinned!==JSON.stringify(s))fail('ACCESS_REVOKED');this.storagePinned=JSON.stringify(s);
  const selected=S.effect(p=>this.hostRead(p),s,this.effectId,this.identity.uid,true,this.cycleBoundary||(this.cycleBoundary=require('./act-cycle-scope.cjs').createContains()));
  Object.assign(this,selected);return c;
 }
 io(raw){if(JSON.parse(raw).kind==='pilotConfig')return JSON.stringify({root:'',shared:true,executor:EXECUTOR,settings:{commands:['ARTIFACT_UPLOAD_REQUEST','PHOTO_UPLOAD_REQUEST','ATTENDANCE_PROOF_UPLOAD_REQUEST','OUTBOX_CLAIM','OUTBOX_OBSERVE','OUTBOX_RECONCILE'],policyIds:this.control.policyIds,admissionsPerHour:20,jobs:[],activityAgendaDirect:false}});return super.io(raw);}
 validateWrites(s){
  const technical=['commandRequests','commandProvenance','commandReceipts','commandLedger','commandTombstones','commandLeases','memberAudit','commandProjectionInvalidations'],added=[];
  for(const w of s.writes){
   if(w.mode==='delete')fail('PILOT_WRITE_DENIED');const [col,id]=w.path.split('/'),d=decode(w.data),v=d.snapshot||d;
   if(col==='jobOutbox'){if(id!==this.effectId)fail('PILOT_WRITE_DENIED');continue;}
   if(technical.includes(col)){if(!this.operationId||id!==this.operationId)fail('PILOT_WRITE_DENIED');continue;}
   if(col==='memberAdmissionTickets'||col==='memberAdmissionUses'){if(d.ownerUid!=='CORE_SYSTEM'||d.operationId!==this.operationId)fail('PILOT_WRITE_DENIED');continue;}
   if(col==='memberAdmissionCounters'){if(d.uid!=='CORE_SYSTEM')fail('PILOT_WRITE_DENIED');continue;}
   if(col==='portalReaderResourceFences'||w.path==='portalReaderControl/current')continue;
   if(col==='commandReconciliation'){if(v.effectId!==this.effectId)fail('PILOT_WRITE_DENIED');added.push(w.path);continue;}
   if(col==='presentationArtifacts'||col==='activityArtifacts'){
    if(v.provider!=='GOOGLE_CLOUD_STORAGE'||v.objectId!==require('./runtime-profile.cjs').profile.prefix+this.effectId||v.accessClass!=='PRIVATE'||(this.out.targetRef.kind==='presentation'?v.presentationRef?.id!==this.out.targetRef.id:v.attendanceRef?.id!==this.out.targetRef.id))fail('PILOT_WRITE_DENIED');added.push(w.path);continue;
   }
   if(col==='presentationHistory'){if(v.operationId!==this.operationId)fail('PILOT_WRITE_DENIED');added.push(w.path);continue;}
   fail('PILOT_WRITE_DENIED');
  }
  const fresh=added.filter(p=>!this.scope.resourcePaths.includes(p));
  if(fresh.length&&!this.scope.operational){if(this.scope.resourcePaths.length+fresh.length>480)fail('SCOPE_TOO_LARGE');s.writes.push({path:this.scopePath,mode:'set',data:encode({...this.scope,revision:this.scope.revision+1,resourcePaths:[...this.scope.resourcePaths,...fresh]})});}
 }
}
function contract(body){
 if(!body||Object.keys(body).sort().join()!=='action,request'||!['storageStart','storageStatus','storageChunk','storageFinalize'].includes(body.action))fail('COMMAND_PAYLOAD_INVALID');
 const r=body.request,keys=body.action==='storageChunk'?'contentBase64,effectRef,offset':'effectRef';
 if(!r||Object.keys(r).sort().join()!==keys||Object.keys(r.effectRef||{}).sort().join()!=='id,kind'||r.effectRef.kind!=='outbox'||!/^[a-f0-9]{32}$/.test(r.effectRef.id))fail('COMMAND_PAYLOAD_INVALID');
 return r;
}
async function transfer(options){
 const {body,db,auth,token,appCheck,appCheckToken,origin,hooks={}}=options,env=fence(),r=contract(body),start=performance.now();
 if(process.env.GEAPA_ACT_STORAGE_ENABLED!=='true'||!env.emulator&&Object.keys(hooks).length)fail('CONFIG_CAPABILITY_DISABLED');
 const identity=await authenticate(auth,token),store=new StorageStore(db,identity);store.effectId=r.effectRef.id;
 await store.read(()=>store.guard());await appCheckGate(appCheck,appCheckToken,store.control,origin);
 const domain=createDomain(x=>store.io(x)),run=(fn,x)=>{const v=domain.run(fn,x);return v===undefined?undefined:JSON.parse(v);};run('(x)=>nodeIdentity(x)',identity);
 const phase=(name,...args)=>store.transaction(()=>decode(run('(x)=>nodeStoragePhase(x)',{name,args:encode(args)})));
 store.beforeCommit=()=>authenticate(auth,token);
 const system=plan=>require('./trusted-system.cjs').executeSystem(store,run,plan,hooks);
 const held=await phase('acquire',r.effectRef,crypto.randomBytes(16).toString('hex'));
 if(held.status==='CONFIRMED')return {state:'CONFIRMED',effectRef:r.effectRef,artifactRef:held.artifactRef,metrics:store.metrics};
 const plan=held.plan;
 try{
  if(held.status==='PENDING')await system(await phase('commandPlan',r.effectRef,'OUTBOX_CLAIM',{effectRef:r.effectRef,expectedAttemptState:'PENDING'}));
  if(!held.reservation.sessionUri)await phase('attempt',r.effectRef);
  const provider=await (env.emulator&&hooks.provider?hooks.provider:GCS.runtimeProvider());
  let uri=held.reservation.sessionUri;
  if(body.action==='storageStart'&&!uri){
   if(held.reservation.initiationStartedAt)return {state:'UNKNOWN',reason:'SESSION_RECONCILIATION_REQUIRED',effectRef:r.effectRef};
   await phase('initiating',r.effectRef,held.lease);
   const created=await provider.initiate(plan);await phase('saveSession',r.effectRef,held.lease,created.sessionUri);uri=created.sessionUri;
  }
  if(body.action==='storageFinalize'){
   const observation=await provider.inspect(plan);
   if(observation.outcome!=='CONFIRMED')return {state:'UNKNOWN',effectRef:r.effectRef};
   await phase('digest',r.effectRef,observation);
   const z=(await phase('source',r.effectRef)).out,evidenceRef=await phase('evidence',r.effectRef,z,observation);
   const p=z.status==='UNKNOWN'?await phase('commandPlan',r.effectRef,'OUTBOX_RECONCILE',{effectRef:r.effectRef,attemptId:z.attemptId,reconciliationEvidenceRef:evidenceRef,provenOutcome:'CONFIRMED'}):await phase('commandPlan',r.effectRef,'OUTBOX_OBSERVE',{effectRef:r.effectRef,attemptId:z.attemptId,leaseFence:z.leaseFence,outcome:'CONFIRMED',evidenceRef,externalId:observation.externalId});
   await system(p);
   const done=(await phase('source',r.effectRef)).out;
   return {state:done.status,effectRef:r.effectRef,artifactRef:done.artifactRef,metrics:{...store.metrics,totalMs:performance.now()-start}};
  }
  if(!uri)return {state:'UNKNOWN',effectRef:r.effectRef};
  let progress=await provider.progress(plan,uri);
  if(body.action==='storageChunk'){
   if(!Number.isSafeInteger(r.offset)||typeof r.contentBase64!=='string'||r.contentBase64.length>Math.ceil(GCS.CHUNK/3)*4)fail('COMMAND_PAYLOAD_INVALID');
   const bytes=Buffer.from(r.contentBase64,'base64');if(bytes.toString('base64')!==r.contentBase64)fail('COMMAND_PAYLOAD_INVALID');
   if(progress.state==='UPLOADING'&&progress.offset===r.offset){await provider.chunk(plan,uri,r.offset,bytes);progress=await provider.progress(plan,uri);}
   else if(progress.offset===null||progress.offset<r.offset+bytes.length)fail('REVISION_CONFLICT');
  }
  return {...progress,effectRef:r.effectRef,chunkBytes:GCS.CHUNK,metrics:{...store.metrics,totalMs:performance.now()-start}};
 }finally{try{await phase('release',r.effectRef,held.lease);}catch{ /* The finite lease remains recoverable after a revoked session or lost reply. */ }}
}
module.exports={StorageStore,CONTROL,contract,transfer};
