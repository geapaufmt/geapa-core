// F4A: native synchronous port of F3.5 activities/index.cjs. No external module loading.
function f4Activities_() {
  if (f4Activities_.value_) return f4Activities_.value_;
  return (f4Activities_.value_ = (function () {
'use strict';
const C=f4Core_(),P=f4Projections_();
const {Timestamp}=f4Firestore_();
function decide(get,r,auth,now){const v=r.payload,a=get('canonicalActivities/'+v.activityId);C.need(a&&a.activityId===v.activityId&&a.active===true,'DOMAIN_STATE_INVALID');const sending=r.commandType==='justificativaEnviar';const jid=sending?C.hash({uid:r.uid,activityId:v.activityId}).slice(0,32):v.justificationId,path='canonicalJustifications/'+jid,old=get(path);const before=old?.revision??0;C.expectedRevision(r.expectedRevision,before);let data;
if(sending){C.need(!old&&r.expectedRevision===0&&Array.isArray(a.eligiblePeople)&&a.eligiblePeople.includes(auth.b.idPessoa)&&C.time(a.opensAt,a.deadline,now),'DOMAIN_STATE_INVALID');data={justificationId:jid,activityId:v.activityId,ownerUid:r.uid,idPessoa:auth.b.idPessoa,reason:v.reason,status:'ENVIADA',submittedAt:Timestamp.fromMillis(now),updatedAt:Timestamp.fromMillis(now),analysisSummary:'',revision:1}}
else{C.need(old&&old.activityId===v.activityId&&['ENVIADA','PENDENTE','EM_ANALISE'].includes(old.status),'DOMAIN_STATE_INVALID');data={justificationId:old.justificationId,activityId:old.activityId,ownerUid:old.ownerUid,idPessoa:old.idPessoa,reason:old.reason,status:v.decision==='DEFERIR'?'DEFERIDA':'INDEFERIDA',submittedAt:old.submittedAt,updatedAt:Timestamp.fromMillis(now),analysisSummary:v.decision==='DEFERIR'?'Deferida.':'Indeferida.',revision:before+1}}
const ownerBinding=get('memberAuthBindings/'+data.ownerUid);C.need(ownerBinding&&ownerBinding.active===true&&ownerBinding.uid===data.ownerUid&&ownerBinding.idPessoa===data.idPessoa,'AUTH_BINDING_INVALID');const projection=P.memberJustificationViews({sourceRevision:data.revision,data:{schemaVersion:'safe-own-justification-v1',ownerUid:data.ownerUid,justificationId:data.justificationId,activityId:data.activityId,status:data.status,submittedAt:data.submittedAt,updatedAt:data.updatedAt,reason:data.reason,analysisSummary:data.analysisSummary}});return {path,data,update:!sending,before,after:data.revision,projectionPath:'memberJustificationViews/'+jid,projection}}
return {decide};




  })());
}
