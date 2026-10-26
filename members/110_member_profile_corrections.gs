// Correction requests reuse the profile command family. A SELF request only
// records a proposal; the existing administrative capability applies changes.
function f4MemberProfileCorrections_(){
 const C=f4Core_(),N=C.need,K=f4MemberCommandCatalog_(),P=f4MemberProfileContract_();
 const fields=['displayName','cpf','rga','birthDate','countryCode','courseId'];
 const statuses=['PENDENTE','EM_ANALISE','COMPLEMENTO_SOLICITADO','INDEFERIDA','CANCELADA','APLICADA'];
 const terminal=['INDEFERIDA','CANCELADA','APLICADA'];
 const own=['SUBMIT','COMPLEMENT','CANCEL_SELF'],admin=['REVIEW','REQUEST_COMPLEMENT','REJECT','APPLY'];
 const is=r=>['members.self.update','members.profile.correct'].includes(r?.commandType)&&r.payload?.operation==='CORRECTION_REQUEST';
 function text(v,max,required=true){N(typeof v==='string'&&v.length<=max&&!/[\x00-\x08\x0b\x0c\x0e-\x1f<>]/.test(v)&&(!required||v.trim().length>0),'COMMAND_PAYLOAD_INVALID');return v.trim();}
 function value(field,v,today){N(fields.includes(field),'COMMAND_PAYLOAD_INVALID');if(field==='cpf')return f4MemberEnrollmentContract_().cpf(v);if(field==='rga'){N(typeof v==='string'&&/^[0-9]{5,20}$/.test(v),'COMMAND_PAYLOAD_INVALID');return v;}return P.patch({[field]:v},false,today)[field];}
 function validate(r){
  N(is(r),'COMMAND_TYPE_INVALID');C.exact(r.target,['personId']);N(C.id(r.target.personId),'COMMAND_PAYLOAD_INVALID');
  const p=r.payload;C.exact(p,['operation','action','requestId','field','value','reason']);N(C.op(p.requestId)&&fields.includes(p.field),'COMMAND_PAYLOAD_INVALID');text(p.reason,1000,!['REVIEW','CANCEL_SELF'].includes(p.action));
  if(r.commandType==='members.self.update')N(own.includes(p.action)&&r.decisionRef===null&&r.evidenceRefs.length===0,'COMMAND_PAYLOAD_INVALID');
  else N(admin.includes(p.action)&&r.evidenceRefs.length>0,'COMMAND_PAYLOAD_INVALID');
  if(p.action==='SUBMIT'){N(p.requestId===r.operationId,'COMMAND_PAYLOAD_INVALID');value(p.field,p.value,'9999-12-31');}else N(p.value===null,'COMMAND_PAYLOAD_INVALID');return true;
 }
 function record(v){
  C.exact(v,['schemaVersion','id','personId','revision','field','previousValue','proposedValue','baseValueHash','status','createdAt','updatedAt','history']);
  N(v.schemaVersion==='member-profile-correction-v1'&&C.op(v.id)&&C.id(v.personId)&&C.rev(v.revision)&&v.revision>0&&fields.includes(v.field)&&statuses.includes(v.status)&&/^[a-f0-9]{64}$/.test(v.baseValueHash),'SCHEMA_INVALID');value(v.field,v.proposedValue,'9999-12-31');
  N((v.previousValue===null||typeof v.previousValue==='string'&&v.previousValue.length<=2000)&&v.baseValueHash===C.hash(v.previousValue),'SCHEMA_INVALID');
  N(Number.isSafeInteger(v.createdAt)&&Number.isSafeInteger(v.updatedAt)&&v.createdAt<=v.updatedAt&&Array.isArray(v.history)&&v.history.length===v.revision&&v.history.length<=100,'SCHEMA_INVALID');
  for(const h of v.history){C.exact(h,['operationId','actorUid','action','at','reason']);N(C.op(h.operationId)&&C.id(h.actorUid)&&[...own,...admin].includes(h.action)&&Number.isSafeInteger(h.at),'SCHEMA_INVALID');text(h.reason,1000,false);}return v;
 }
 function index(v,id){if(!v)return {schemaVersion:'member-profile-corrections-index-v1',id,revision:0,requestIds:[],activeByField:{}};C.exact(v,['schemaVersion','id','revision','requestIds','activeByField']);N(v.schemaVersion==='member-profile-corrections-index-v1'&&v.id===id&&C.rev(v.revision)&&Array.isArray(v.requestIds)&&v.requestIds.length<=200&&v.requestIds.every(C.op)&&new Set(v.requestIds).size===v.requestIds.length&&v.activeByField&&Object.keys(v.activeByField).every(k=>fields.includes(k)&&v.requestIds.includes(v.activeByField[k])),'SCHEMA_INVALID');return v;}
 function source(get,id,field){
  if(field==='displayName')return {path:'memberIdentities/'+id,value:get('memberIdentities/'+id)?.displayName||null};
  if(!['cpf','rga'].includes(field))return {path:'memberPrivateProfiles/'+id,value:get('memberPrivateProfiles/'+id)?.[field]??null};
  const h=get('memberAggregateHeads/'+id);N(h&&Array.isArray(h.paths),'PROFILE_MISSING');const rows=h.paths.filter(p=>p.startsWith('memberIdentifiers/')).map(p=>({path:p,data:get(p)})).filter(v=>v.data?.active&&v.data.kind===field.toUpperCase());
  N(rows.length<=1&&rows.every(v=>v.data.idPessoa===id),'RECONCILIATION_REQUIRED');return rows.length?{path:rows[0].path,value:rows[0].data.value}:{path:null,value:null};
 }
 function paths(get,r){const p=r.payload,id=r.target.personId,out=['memberProfileCorrections/'+p.requestId,'memberProfileCorrectionIndexes/'+id];if(p.action==='APPLY'){const s=source(get,id,p.field);out.push(s.path||'memberIdentifiers/I_'+r.operationId);}return out;}
 function changes(get,r,now,profileChanges){
  validate(r);const p=r.payload,id=r.target.personId,path='memberProfileCorrections/'+p.requestId,ip='memberProfileCorrectionIndexes/'+id,old=get(path),ix=index(get(ip),id),s=source(get,id,p.field),today=Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd');
  let v;if(p.action==='SUBMIT'){
   N(!old&&!ix.activeByField[p.field]&&ix.requestIds.length<200,'OPERATION_CONFLICT');const proposed=value(p.field,p.value,today);N(!K.equal(proposed,s.value),'COMMAND_PAYLOAD_INVALID');
   v={schemaVersion:'member-profile-correction-v1',id:p.requestId,personId:id,revision:0,field:p.field,previousValue:s.value,proposedValue:proposed,baseValueHash:C.hash(s.value),status:'PENDENTE',createdAt:now,updatedAt:now,history:[]};
   ix.requestIds=ix.requestIds.concat(p.requestId);ix.activeByField={...ix.activeByField,[p.field]:p.requestId};
  }else{v=record(old);N(v.personId===id&&v.field===p.field&&ix.activeByField[p.field]===v.id,'AUTH_BINDING_INVALID');N(!terminal.includes(v.status),'INVALID_TRANSITION');v={...v,history:v.history.slice()};
   if(p.action==='COMPLEMENT')N(v.status==='COMPLEMENTO_SOLICITADO','INVALID_TRANSITION');
   if(p.action==='APPLY')N(['PENDENTE','EM_ANALISE'].includes(v.status)&&v.baseValueHash===C.hash(s.value),'REVISION_CONFLICT');
   v.status=({COMPLEMENT:'PENDENTE',CANCEL_SELF:'CANCELADA',REVIEW:'EM_ANALISE',REQUEST_COMPLEMENT:'COMPLEMENTO_SOLICITADO',REJECT:'INDEFERIDA',APPLY:'APLICADA'})[p.action];
  }
  v={...v,revision:v.revision+1,updatedAt:now,history:v.history.concat({operationId:r.operationId,actorUid:r.ownerUid,action:p.action,at:now,reason:text(p.reason,1000,false)})};record(v);
  if(terminal.includes(v.status)){ix.activeByField={...ix.activeByField};delete ix.activeByField[p.field];}
  const out=[{path,expectedRevision:old?.revision||0,data:v},{path:ip,expectedRevision:ix.revision,data:{...ix,revision:ix.revision+1}}];
  if(p.action==='APPLY'){
   if(['cpf','rga'].includes(p.field)){
    const sp=s.path||'memberIdentifiers/I_'+r.operationId,prior=get(sp),revision=prior?.revision||0;
    out.push({path:sp,expectedRevision:revision,data:prior?{...prior,revision:revision+1,value:v.proposedValue}:{schemaVersion:'m1-memberIdentifiers-v1',id:sp.split('/')[1],idPessoa:id,revision:1,origin:'M13_COMMAND',kind:p.field.toUpperCase(),value:v.proposedValue,active:true}});
   }else out.push(...profileChanges(get,r,{[p.field]:value(p.field,v.proposedValue,today)},now));
  }
  for(const c of out)N(f4MemberFamilyInputs_().revision(r,c.path)===c.expectedRevision,'REVISION_CONFLICT');return out;
 }
 function change(c,id){const isIndex=c.path==='memberProfileCorrectionIndexes/'+id,v=isIndex?index(c.data,id):record(c.data);N((isIndex||c.path==='memberProfileCorrections/'+v.id&&v.personId===id)&&C.rev(c.expectedRevision)&&v.revision===c.expectedRevision+1,'M1_REVISION_INVALID');}
 function selfTarget(get,t,id){if(t.fieldMask.length!==0)return null;const requestPath=t.resources.find(p=>/^memberProfileCorrections\/[a-f0-9]{32}$/.test(p));N(t.personId===id&&t.resources.length===2&&requestPath&&t.resources.includes('memberProfileCorrectionIndexes/'+id),'M13_AUTHZ_DENIED');const v=get(requestPath),ix=index(get('memberProfileCorrectionIndexes/'+id),id);if(v){record(v);N(v.personId===id&&ix.requestIds.includes(v.id),'M13_AUTHZ_DENIED');}return {correctionRequestId:requestPath.split('/')[1],correctionRevision:v?.revision||0,correctionIndexRevision:ix.revision};}
 return {fields,statuses,terminal,is,validate,record,index,source,paths,changes,change,value,selfTarget};
}
