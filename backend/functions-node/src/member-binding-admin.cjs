'use strict';
const {MemberStore,executeMember}=require('./member-executor.cjs');
const {decode}=require('./firestore.cjs'),{createDomain}=require('./domain.cjs');
const {authenticate,appCheckGate}=require('./auth.cjs'),{fail,PROJECT}=require('./environment.cjs');
const actions=['memberBindingContext','memberBindingProposal','memberBindingVerify','memberBindingConfirm','memberBindingRecover','memberBindingStatus'];
async function googleAccount(auth,email){
 if(typeof email!=='string'||email.length>254||email!==email.trim()||!/^\S+@\S+\.\S+$/.test(email))fail('COMMAND_PAYLOAD_INVALID');
 let u;try{u=await auth.getUserByEmail(email);}catch(e){fail(e.code==='auth/user-not-found'?'BINDING_ACCOUNT_NOT_FOUND':'SERVICE_UNAVAILABLE');}
 if(u.disabled)fail('BINDING_ACCOUNT_DISABLED');if(!u.emailVerified)fail('BINDING_EMAIL_UNVERIFIED');
 if(!u.providerData?.some(p=>p.providerId==='google.com'&&p.email?.toLowerCase()===u.email?.toLowerCase())||u.email?.toLowerCase()!==email.toLowerCase()||u.tenantId)fail('BINDING_GOOGLE_REQUIRED');
 if(!/^[A-Za-z0-9_-]{1,48}$/.test(u.uid))fail('AUTH_BINDING_INVALID');
 return {uid:u.uid,email:u.email.toLowerCase(),emailVerified:true,provider:'google.com'};
}
class BindingStore extends MemberStore{
 guard(){const c=super.guard(),b=this.hostRead('memberBindingAdminControl/current');if(!b?.enabled||b.projectId!==PROJECT||b.schemaVersion!=='member-binding-admin-control-v1')fail('ACCESS_REVOKED');if(this.bindingPinned&&this.bindingPinned!==JSON.stringify(b))fail('ACCESS_REVOKED');this.bindingPinned=JSON.stringify(b);return c;}
 validateWrites(s){
  const extra=s.writes.filter(w=>w.path.startsWith('memberBindingVerifications/'));
  for(const w of extra){const d=decode(w.data);if(this.phase!=='verify'||w.mode!=='set'||w.path!=='memberBindingVerifications/'+this.operationId||this.hostRead(w.path)||d.ownerUid!==this.identity.uid||d.personId!==this.personId)fail('PILOT_WRITE_DENIED');}
  super.validateWrites({...s,writes:s.writes.filter(w=>!extra.includes(w))});
 }
}
async function bindingAdmin(options){
 const {db,auth,appCheck,token,appCheckToken,origin,body}=options;
 if(!body||Object.keys(body).sort().join()!=='action,request'||!actions.includes(body.action))fail('COMMAND_TYPE_INVALID');
 const identity=await authenticate(auth,token),store=new BindingStore(db,identity),domain=createDomain(x=>store.io(x)),q=body.request;
 const run=(fn,x)=>JSON.parse(domain.run(fn,x));run('(v)=>{nodeIdentity(v);return null}',identity);
 const call=(method,x)=>run('(v)=>f4MemberBindingAdmin_({config:nodeConfig,identity:nodeActor,clock:Date.now})[v.method](...v.args)',{method,args:x});
 await store.read(()=>store.guard());await appCheckGate(appCheck,appCheckToken,store.control,origin);
 store.beforeCommit=async()=>{await authenticate(auth,token);await appCheckGate(appCheck,appCheckToken,store.control,origin);};
 const keys=body.action==='memberBindingContext'?['personId']:body.action==='memberBindingVerify'?['operationId','personId','email','evidence']:body.action==='memberBindingConfirm'?['operationId','personId','confirmed']:['operationId','personId'];
 if(Object.keys(q).sort().join()!==keys.sort().join()||!/^[A-Za-z0-9_-]{1,48}$/.test(q.personId)||q.operationId!==undefined&&!/^[a-f0-9]{32}$/.test(q.operationId))fail('COMMAND_PAYLOAD_INVALID');
 store.personId=q.personId;store.operationId=q.operationId;store.preparation=true;
 // Canonical authorization BEFORE lookup: directory permission never enumerates Auth.
 await store.read(()=>run('(p)=>{f4MemberBindingAdmin_({config:nodeConfig,identity:nodeActor}).access(nodeRead,p);return null}',q.personId));
 let result;
 if(body.action==='memberBindingContext'){
  const c=await store.transaction(()=>call('context',[q.personId]));let email=null;if(c.uid){try{const linked=await googleAccount(auth,(await auth.getUser(c.uid)).email);if(linked.uid!==c.uid)fail('BINDING_ACCOUNT_CHANGED');email=linked.email;}catch(e){if(!['BINDING_ACCOUNT_NOT_FOUND','BINDING_EMAIL_UNVERIFIED','BINDING_GOOGLE_REQUIRED','BINDING_ACCOUNT_DISABLED'].includes(e.code)&&e.code!=='auth/user-not-found')throw e;}}
  const after=await store.transaction(()=>call('context',[q.personId]));if(after.uid!==c.uid)fail('BINDING_CONFLICT');result={canManage:c.canManage,state:c.state,email};
 }else if(body.action==='memberBindingVerify'){
  const account=await googleAccount(auth,q.email);store.phase='verify';
  result=await store.transaction(()=>run('(x)=>f4MemberBindingAdmin_({config:nodeConfig,identity:nodeActor}).verify(nodeTx,x.input,x.account)',{input:{operationId:q.operationId,personId:q.personId,evidence:q.evidence},account}));
 }else{
  const v=await store.read(()=>{const v=store.hostRead('memberBindingVerifications/'+q.operationId);if(v&&(v.ownerUid!==identity.uid||v.personId!==q.personId))fail('ACCESS_REVOKED');return v;});
  if(!v){if(body.action==='memberBindingStatus')return {result:{state:'UNVERIFIED',operationId:q.operationId},metrics:store.metrics};fail('ACCESS_REVOKED');}
  if(body.action==='memberBindingProposal')return {result:{operationId:q.operationId,personId:v.personId,email:v.account.email,expiresAt:v.expiresAt.toMillis()},metrics:store.metrics};
  const terminal=await store.read(()=>['COMPLETED','FAILED'].includes(store.hostRead('commandReceipts/'+q.operationId)?.state));
  const recheck=async()=>{if(terminal)return;const current=await googleAccount(auth,v.account.email);if(JSON.stringify(current)!==JSON.stringify(v.account))fail('BINDING_ACCOUNT_CHANGED');await appCheckGate(appCheck,appCheckToken,store.control,origin);};
  const saved=await store.read(()=>store.hostRead('memberPreparedIntents/'+q.operationId));
  if(body.action==='memberBindingStatus'&&!saved)return {result:{state:v.expiresAt.toMillis()<=Date.now()?'EXPIRED':'VERIFIED',operationId:q.operationId},metrics:store.metrics};
  const exists=await store.read(()=>!!store.hostRead('commandRequests/'+q.operationId));
  if(!exists&&body.action==='memberBindingStatus')return {result:{state:v.expiresAt.toMillis()<=Date.now()?'EXPIRED':'NOT_ADMITTED',operationId:q.operationId},metrics:store.metrics};
  if(body.action==='memberBindingRecover'&&!exists)fail('COMMAND_TYPE_INVALID');
  let request=saved?.input;if(body.action==='memberBindingConfirm'){if(q.confirmed!==true)fail('COMMAND_PAYLOAD_INVALID');if(!exists)await recheck();store.phase='prepare';request=await store.transaction(()=>run('(x)=>f4MemberBindingAdmin_({config:nodeConfig,identity:nodeActor}).prepare(nodeTx,x)',q));}
  const key=await store.read(()=>run('(r)=>({operationId:r.operationId,commandType:r.commandType,digest:f4MemberCommandCatalog_().digest(nodeConfig.projectId,{...r,ownerUid:nodeActor.uid,expectedRevisions:nodeRead("devProtectedInputs/"+r.operationId).expectedRevisions})})',request));
  // Only this server module can enter protected admission. No UID/envelope/key escapes.
  let answer;try{answer=await executeMember({...options,body:{action:body.action==='memberBindingStatus'?'memberStatus':body.action==='memberBindingRecover'?'memberRecover':'memberExecute',request:['memberBindingStatus','memberBindingRecover'].includes(body.action)?key:request},bindingAdmin:{uid:v.account.uid,recheck:body.action==='memberBindingStatus'?async()=>{}:recheck}});}catch(e){if(body.action!=='memberBindingStatus'||e.code!=='RECONCILIATION_REQUIRED')throw e;return {result:{state:'PENDING',operationId:q.operationId},metrics:store.metrics};}
  result={state:answer.dto?.receipt?.state==='COMPLETED'?'COMMIT_APPLIED':answer.dto?.receipt?.state==='FAILED'?'FAILED':answer.dto?.businessState||answer.result?.state||'PENDING',operationId:q.operationId,receiptState:answer.dto?.receipt?.state||null};
 }
 await authenticate(auth,token);return {result,metrics:store.metrics};
}
module.exports={bindingAdmin,BindingStore,googleAccount,actions};
