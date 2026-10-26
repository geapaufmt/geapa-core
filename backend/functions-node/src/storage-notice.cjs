'use strict';
const {ActivityStore}=require('./activity-commands.cjs'),{decode,encode}=require('./firestore.cjs'),{createDomain}=require('./domain.cjs');
const {authenticate,appCheckGate}=require('./auth.cjs'),{fail,fence}=require('./environment.cjs'),S=require('./storage-control.cjs');
class NoticeStore extends ActivityStore{
 guard(){
  const c=super.guard(),s=S.check(p=>this.hostRead(p),this.identity.uid,c.expiresAt);
  if(this.storagePinned&&this.storagePinned!==JSON.stringify(s))fail('ACCESS_REVOKED');this.storagePinned=JSON.stringify(s);
  const contains=this.cycleBoundary||(this.cycleBoundary=require('./act-cycle-scope.cjs').createContains());
  const u=this.hostRead('presentationUnitHeads/'+this.unitId),pr=u?.presentationRef&&this.hostRead('canonicalPresentations/'+u.presentationRef.id),scope=(s.scopeIds||[]).map(id=>this.hostRead('actFlowExecutorScopes/'+id)).find(v=>v?.synthetic===!require('./environment.cjs').PRODUCTION&&v.projectId===s.projectId&&contains(p=>this.hostRead(p),v,'presentationUnitHeads/'+this.unitId)&&contains(p=>this.hostRead(p),v,'presentationCorrectionOpportunities/'+this.opportunityId));
  if(!scope||!pr?.materialRequirementRef||!scope.operational&&!(s.noticeUnitIds||[]).includes(this.unitId))fail('PILOT_RESOURCE_DENIED');
  this.noticeScope=scope;this.allowed=['presentationUnitHeads/'+this.unitId,'canonicalPresentations/'+pr.id,'presentationMaterialRequirements/'+pr.materialRequirementRef.id,'presentationCorrectionOpportunities/'+this.opportunityId];return c;
 }
 io(raw){if(JSON.parse(raw).kind==='pilotConfig')return JSON.stringify({root:'',shared:true,executor:require('./environment.cjs').EXECUTOR,settings:{commands:['CORRECTION_NOTICE_RECORD'],policyIds:this.control.policyIds,admissionsPerHour:20,jobs:[],activityAgendaDirect:false}});return super.io(raw);}
 validateWrites(s){
  const added=[];
  for(const w of s.writes){
   if(w.mode==='delete')fail('PILOT_WRITE_DENIED');const [col,id]=w.path.split('/'),d=decode(w.data),v=d.snapshot||d;
   if(this.allowed.some(p=>w.path===p||w.path.startsWith(p+'/revisions/')))continue;
   if(col==='commandReconciliation'){if(v.kind!=='PORTAL_EXPLICIT_ACK'||v.targetRef?.id!==this.opportunityId)fail('PILOT_WRITE_DENIED');if(!this.noticeScope.resourcePaths.includes(w.path))added.push(w.path);continue;}
   if(['commandRequests','commandProvenance','commandReceipts','commandLedger','commandTombstones','commandLeases','memberAudit','commandProjectionInvalidations'].includes(col)){if(id!==this.operationId)fail('PILOT_WRITE_DENIED');continue;}
   if(['memberAdmissionTickets','memberAdmissionUses'].includes(col)){if(d.ownerUid!=='CORE_SYSTEM'||d.operationId!==this.operationId)fail('PILOT_WRITE_DENIED');continue;}
   if(col==='memberAdmissionCounters'){if(d.uid!=='CORE_SYSTEM')fail('PILOT_WRITE_DENIED');continue;}
   if(col==='portalReaderResourceFences'||w.path==='portalReaderControl/current')continue;
   if(col==='presentationHistory'&&v.operationId===this.operationId)continue;
   fail('PILOT_WRITE_DENIED');
  }
  // The acknowledgment is an input to the subsequent correction command.
  // Register only this canonical output in the same protected synthetic scope.
  if(added.length&&!this.noticeScope.operational){const scope=this.noticeScope,paths=[...new Set([...scope.resourcePaths,...added])].sort();if(paths.length>480)fail('SCOPE_TOO_LARGE');s.writes.push({path:'actFlowExecutorScopes/'+scope.id,mode:'set',data:encode({...scope,revision:scope.revision+1,resourcePaths:paths})});}
 }
}
async function acknowledge(options){
 const env=fence();if(process.env.GEAPA_ACT_STORAGE_ENABLED!=='true')fail('CONFIG_CAPABILITY_DISABLED');if(!env.emulator&&Object.keys(options.hooks||{}).length)fail('DESTINATION_INVALID');
 const {body,db,auth,token,appCheck,appCheckToken,origin}=options,r=body?.request;
 if(body?.action!=='storageNotice'||Object.keys(body).sort().join()!=='action,request'||!/^APU_[a-f0-9]{32}$/.test(r?.idApresentacao)||!/^[a-f0-9]{32}$/.test(r?.opportunityId))fail('COMMAND_PAYLOAD_INVALID');
 const identity=await authenticate(auth,token),store=new NoticeStore(db,identity);store.unitId=r.idApresentacao.slice(4);store.opportunityId=r.opportunityId;
 await store.read(()=>store.guard());await appCheckGate(appCheck,appCheckToken,store.control,origin);store.beforeCommit=()=>authenticate(auth,token);
 const domain=createDomain(x=>store.io(x)),run=(fn,x)=>{const v=domain.run(fn,x);return v===undefined?undefined:JSON.parse(v);};run('(x)=>nodeIdentity(x)',identity);
 options.hooks?.configureDomain?.(domain);
 const phase=readOnly=>store.transaction(()=>decode(run('(x)=>nodeStorageNotice(x)',{request:r,readOnly})));
 const plan=await phase(false);await require('./trusted-system.cjs').executeSystem(store,run,plan);
 const done=await phase(true);return {ok:true,basis:done.basis,opportunityId:done.opportunityId,metrics:store.metrics};
}
module.exports={acknowledge,NoticeStore};
