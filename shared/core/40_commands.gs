// F4A: native synchronous port of F3.5 transport/commands/index.cjs. No external module loading.
function f4Commands_() {
  if (f4Commands_.value_) return f4Commands_.value_;
  return (f4Commands_.value_ = (function () {
'use strict';
const C=f4Core_(),P=f4Projections_(),A=f4Auth_(),activities=f4Activities_(),editorial=f4Editorial_(),outbox=f4Outbox_();
const D=f4Firestore_();
const fixed={environment:'SYNTHETIC_LOCAL',projectId:'demo-geapa-m04-local',database:'(default)',host:'127.0.0.1',port:8184};
function createBackend(options){
 C.destination(options.config||fixed);const {db,verifyPrincipal,clock=Date.now,config=fixed,ledgerFactory=null,receiptDays=7,contextValidator=null}=options;
 C.destination(config);D.trusted(db,config);C.need(typeof verifyPrincipal==='function','AUTH_BINDING_INVALID');
 const run=fn=>D.runTransaction(db,tx=>fn({get:path=>(tx.get(D.doc(db,path))).data(),set:(path,data)=>tx.set(D.doc(db,path),data),update:(path,data)=>tx.update(D.doc(db,path),data)}));
 const receipt=(r,state,code,now,before=null,after=null,epoch=null)=>P.receipts({sourceRevision:after,data:{ownerUid:r.uid,operationId:r.operationId,state,code,revisionBefore:before,revisionAfter:after,acceptedAt:r.receivedAt,updatedAt:D.Timestamp.fromMillis(now),authzEpochAtDecision:epoch,resultRef:state==='COMPLETED'?'result-'+r.operationId:null,reconcileAfter:state==='RECONCILIATION_REQUIRED'?D.Timestamp.fromMillis(now+3600000):null,purgeAfter:D.Timestamp.fromMillis(now+receiptDays*86400000)}});
 function writeProjection(t,path,p){t.set(path,p.dto);t.set('projectionProvenance/'+C.hash(path).slice(0,32),{projectionKey:C.hash(path),sourceRevision:p.provenance.sourceRevision})}
 function readCommand(t,id){C.need(C.op(id),'COMMAND_PAYLOAD_INVALID');const r=t.get('commandRequests/'+id);C.envelope(r,id);return r}
 function age(r,now){C.need(C.ms(r.receivedAt)<=now&&now-C.ms(r.receivedAt)<86400000,'DOMAIN_STATE_INVALID')}
 function register(id,principal){const trusted=verifyPrincipal(principal);C.need(trusted&&C.id(trusted.uid),'AUTH_BINDING_INVALID');return run(t=>{const r=readCommand(t,id);C.need(r.uid===trusted.uid,'AUTH_BINDING_INVALID');const now=clock();age(r,now);const auth=A.authorize(t.get,r.uid,C.scope(r),C.capability(r),now),slot=t.get('commandAdmission/'+r.uid+'/slots/'+r.admissionSlot),prior=t.get('commandProvenance/'+id);const dg=C.digest(r);
 C.need(slot&&slot.uid===r.uid&&slot.idPessoa===auth.b.idPessoa&&slot.operationId===id&&slot.slot===r.admissionSlot&&slot.grantId===r.grantId&&slot.bindingVersion===auth.b.bindingVersion&&slot.accessEpoch===auth.c.accessEpoch&&slot.policyVersion===1&&C.time(slot.issuedAt,slot.expiresAt,now)&&C.ms(slot.expiresAt)<=C.ms(auth.p.validUntil),'AUTH_BINDING_INVALID');if(prior){C.need(prior.digest===dg&&prior.uid===trusted.uid,'OPERATION_CONFLICT');return {code:'ACCEPTED'}}
 t.set('commandProvenance/'+id,{uid:trusted.uid,digest:dg,bindingVersion:auth.b.bindingVersion,accessEpoch:auth.c.accessEpoch,sourceRevision:auth.c.sourceRevision});writeProjection(t,'commandReceipts/'+id,receipt(r,'RECEIVED','ACCEPTED',now));return {code:'ACCEPTED'}})}
 function context(t,id){const r=readCommand(t,id),provenance=t.get('commandProvenance/'+id);C.need(provenance&&provenance.uid===r.uid,'AUTH_BINDING_INVALID');C.need(provenance.digest===C.digest(r),'OPERATION_CONFLICT');if(contextValidator)contextValidator(r,provenance,clock());return {r,provenance}}
 function claim(id,workerId){C.need(C.id(workerId),'LEASE_LOST');return run(t=>{const {r,provenance}=context(t,id),ledger=t.get('commandLedger/'+id),lease=t.get('commandLeases/'+id),now=clock();if(ledger){C.need(ledger.digest===C.digest(r),'OPERATION_CONFLICT');return {terminal:true,decision:ledger.decision};}if(lease&&C.ms(lease.leaseUntil)>now)C.fail('LEASE_LOST');const attempts=(lease?.attempts||0)+1,token=(lease?.fencingToken||0)+1;if(attempts>3){const decision={state:'RECONCILIATION_REQUIRED',code:'RECONCILIATION_REQUIRED'};t.set('commandLedger/'+id,ledgerFactory?ledgerFactory(r,provenance,decision,now,null):{digest:C.digest(r),decision});writeProjection(t,'commandReceipts/'+id,receipt(r,decision.state,decision.code,now));t.set('commandLeases/'+id,{workerId,fencingToken:token,leaseUntil:D.Timestamp.fromMillis(now),attempts,lastCode:decision.code});return {terminal:true,decision}}
 t.set('commandLeases/'+id,{workerId,fencingToken:token,leaseUntil:D.Timestamp.fromMillis(now+120000),attempts,lastCode:'ACCEPTED'});writeProjection(t,'commandReceipts/'+id,receipt(r,'PROCESSING','ACCEPTED',now));return {workerId,fencingToken:token}})}
 function finalize(id,handle){return run(t=>{const {r,provenance}=context(t,id),ledger=t.get('commandLedger/'+id),lease=t.get('commandLeases/'+id);if(ledger){C.need(ledger.digest===C.digest(r),'OPERATION_CONFLICT');return ledger.decision;}const now=clock();C.need(lease&&lease.workerId===handle.workerId&&lease.fencingToken===handle.fencingToken&&C.ms(lease.leaseUntil)>now,'LEASE_LOST');let result,auth,decision,failureRevision=null;
 try{age(r,now);auth=A.authorize(t.get,r.uid,C.scope(r),C.capability(r),now);C.need(provenance.bindingVersion===auth.b.bindingVersion&&provenance.accessEpoch===auth.c.accessEpoch&&provenance.sourceRevision===auth.c.sourceRevision,'ACCESS_REVOKED');result=(r.commandType==='portalPublicAdminPublishPage'?editorial:activities).decide(t.get,r,auth,now);const commitTime=clock();C.need(C.ms(lease.leaseUntil)>commitTime,'LEASE_LOST');C.need(C.time(auth.c.startsAt,auth.c.validUntil,commitTime)&&C.time(auth.p.startsAt,auth.p.validUntil,commitTime)&&C.time(auth.g.startsAt,auth.g.endsAtExclusive,commitTime),'FUNCTION_EXPIRED');decision={state:'COMPLETED',code:'OK',revisionBefore:result.before,revisionAfter:result.after,resultRef:'result-'+id}}
 catch(e){if(!C.codes.includes(e.code))throw e;failureRevision=C.rev(e.canonicalRevision)?e.canonicalRevision:null;decision={state:'FAILED',code:e.code};}
 const ok=decision.state==='COMPLETED',publicCode=ok?'OK':decision.code==='REVISION_CONFLICT'?'REVISION_CONFLICT':'DENIED';const rec=receipt(r,decision.state,publicCode,now,ok?result.before:failureRevision,ok?result.after:failureRevision,auth?.c.accessEpoch??null);
 // No writes were scheduled before all authorization and domain reads completed.
 if(ok){if(result.update)t.update(result.path,result.data);else t.set(result.path,result.data);writeProjection(t,result.projectionPath,result.projection);t.set('jobOutbox/'+id,outbox.intent(id))}
 t.set('commandLedger/'+id,ledgerFactory?ledgerFactory(r,provenance,decision,now,rec.dto):{digest:C.digest(r),decision});writeProjection(t,'commandReceipts/'+id,rec);t.set('commandLeases/'+id,{workerId:lease.workerId,fencingToken:lease.fencingToken,leaseUntil:D.Timestamp.fromMillis(now),attempts:lease.attempts,lastCode:decision.code});return decision})}
 function process(id,workerId){const h=claim(id,workerId);return h.terminal?h.decision:finalize(id,h)}
 // Revocation fixture primitive: canonical state + epoch and synchronous projections in one local transaction.
 function revoke(uid,kind){C.need(['SUSPEND','EGRESS'].includes(kind));return run(t=>{const b=t.get('memberAuthBindings/'+uid);C.need(b,'AUTH_BINDING_INVALID');const old=t.get('memberCurrentStates/'+b.idPessoa),account=t.get('institutionalAccounts/'+b.idPessoa);const now=clock();const c={schemaVersion:old.schemaVersion,idPessoa:old.idPessoa,sourceRevision:old.sourceRevision+1,accessEpoch:old.accessEpoch+1,policyVersion:old.policyVersion,blocked:kind==='SUSPEND',stale:false,state:kind==='EGRESS'?'EGRESSO':old.state,startsAt:old.startsAt,validUntil:old.validUntil,egressEndsAtExclusive:kind==='EGRESS'?D.Timestamp.fromMillis(now):old.egressEndsAtExclusive};const a={human:account.human,sourceRevision:c.sourceRevision},source={binding:b,state:c,account:a};const projections=[['portalUsers/'+uid,P.portalUsers(source)],['portalAccessStatus/'+uid,P.portalAccessStatus(source,now)],['memberStatusViews/'+b.idPessoa,P.memberStatusViews(source)]];t.set('memberCurrentStates/'+b.idPessoa,c);t.set('institutionalAccounts/'+b.idPessoa,a);for(const [path,p]of projections)writeProjection(t,path,p);return {sourceRevision:c.sourceRevision,accessEpoch:c.accessEpoch}})}
 function dispatch(id,workerId){try{return process(id,workerId)}catch(e){return {state:C.codes.includes(e.code)?'FAILED':'RECONCILIATION_REQUIRED',code:C.codes.includes(e.code)?e.code:'RECONCILIATION_REQUIRED'}}}
 return {register,claim,finalize,process,dispatch,revoke};
}
return {createBackend,fixed,createMembersLocal:options=>f4MemberCommandCoordinator_().createLocal(options)};







  })());
}
