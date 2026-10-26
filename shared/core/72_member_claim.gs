// G: additive lease contract. Internal local API, never a browser authority.
function f4MemberClaim_(){
 const C=f4Core_(),F=f4Firestore_(),N=C.need,T=F.Timestamp.fromMillis;
 const schema='member-command-lease-v13-g1';
 const fields=['schemaVersion','operationId','domain','workerId','fencingToken','deadline','revision','attempts','state','authorityEpoch','digest','claimedAt'];
 function validate(v){
  C.exact(v,fields);N(v.schemaVersion===schema&&/^[a-f0-9]{32}$/.test(v.operationId)&&['MEMBERS','ACT'].includes(v.domain)&&C.id(v.workerId)&&/^[a-f0-9]{64}$/.test(v.digest),'LEASE_LOST');
  for(const k of ['revision','fencingToken','authorityEpoch'])N(Number.isSafeInteger(v[k])&&v[k]>0,'LEASE_LOST');
  N(Number.isSafeInteger(v.attempts)&&v.attempts>=1&&v.attempts<=3&&['CLAIMED','RELEASED','TERMINAL','RECONCILIATION_REQUIRED'].includes(v.state),'LEASE_LOST');
  N(Number.isSafeInteger(C.ms(v.claimedAt))&&Number.isSafeInteger(C.ms(v.deadline))&&C.ms(v.deadline)>=C.ms(v.claimedAt)&&C.ms(v.deadline)<=C.ms(v.claimedAt)+120000,'LEASE_LOST');
  N(v.state!=='RECONCILIATION_REQUIRED'||v.attempts===3,'LEASE_LOST');return v;
 }
 function handle(v){return {workerId:v.workerId,fencingToken:v.fencingToken};}
 function control(v){f4MemberAuthority_().validate(v);N(v.mode==='CANONICAL'&&v.activeWriterVersion==='M13','ACCESS_REVOKED');return v;}
 function active(v,key,h,authority,n){
  validate(v);C.exact(h,['workerId','fencingToken']);control(authority);
  N(v.state==='CLAIMED'&&v.operationId===key.operationId&&v.digest===key.digest&&v.domain===(key.commandType.startsWith('ACT.')?'ACT':'MEMBERS')&&v.workerId===h.workerId&&v.fencingToken===h.fencingToken&&v.authorityEpoch===authority.epoch&&n>=C.ms(v.claimedAt)&&n<C.ms(v.deadline),'LEASE_LOST');return v;
 }
 // Caller reads/rechecks ticket, proofs, binding and recovery in this same transaction.
 function claim(previous,key,workerId,authority,n,expiresAt){
  control(authority);N(C.id(workerId)&&Number.isSafeInteger(n)&&Number.isSafeInteger(expiresAt)&&n<expiresAt,'LEASE_LOST');
  if(previous){
   validate(previous);N(previous.operationId===key.operationId&&previous.digest===key.digest&&previous.domain===(key.commandType.startsWith('ACT.')?'ACT':'MEMBERS'),'OPERATION_CONFLICT');
   N(previous.state!=='TERMINAL','RECONCILIATION_REQUIRED');
   if(previous.state==='RECONCILIATION_REQUIRED')return {lease:previous,acquired:false,changed:false};
   if(previous.state==='CLAIMED'&&n<C.ms(previous.deadline)&&previous.authorityEpoch===authority.epoch){
    N(previous.workerId===workerId,'LEASE_LOST');return {lease:previous,acquired:true,changed:false};
   }
   if(previous.attempts===3){const lease={...previous,revision:previous.revision+1,fencingToken:previous.fencingToken+1,state:'RECONCILIATION_REQUIRED',deadline:previous.claimedAt};validate(lease);return {lease,acquired:false,changed:true};}
  }
  const lease={schemaVersion:schema,operationId:key.operationId,domain:key.commandType.startsWith('ACT.')?'ACT':'MEMBERS',workerId,fencingToken:(previous?.fencingToken||0)+1,revision:(previous?.revision||0)+1,attempts:(previous?.attempts||0)+1,state:'CLAIMED',authorityEpoch:authority.epoch,digest:key.digest,claimedAt:T(n),deadline:T(Math.min(n+120000,expiresAt))};
  validate(lease);return {lease,acquired:true,changed:true};
 }
 function release(v,key,h,authority,n){active(v,key,h,authority,n);return validate({...v,revision:v.revision+1,fencingToken:v.fencingToken+1,state:v.attempts===3?'RECONCILIATION_REQUIRED':'RELEASED',deadline:T(n)});}
 function finish(v,n){validate(v);N(v.state==='CLAIMED'&&n>=C.ms(v.claimedAt)&&n<C.ms(v.deadline),'LEASE_LOST');return validate({...v,revision:v.revision+1,state:'TERMINAL',deadline:T(n)});}
 function exhaust(v,key,n){validate(v);N(v.operationId===key.operationId&&v.digest===key.digest&&v.attempts===3&&v.state==='CLAIMED'&&n>=C.ms(v.deadline),'LEASE_LOST');return validate({...v,revision:v.revision+1,fencingToken:v.fencingToken+1,state:'RECONCILIATION_REQUIRED',deadline:v.claimedAt});}
 return Object.freeze({schema,validate,handle,active,claim,release,finish,exhaust});
}
