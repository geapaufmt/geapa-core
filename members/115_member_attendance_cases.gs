// Attendance follow-up is an administrative Members case. It never changes
// membership by calculation; dismissal remains members.discipline.dismiss.
function f4MemberAttendanceCases_(){
 const C=f4Core_(),N=C.need,F=f4Firestore_();
 const actions=['INFORMAL_WARNING','FORMAL_WARNING','COMMUNICATION','FOLLOW_UP','MANIFESTATION','CLOSE'];
 const caseId=(personId,cycleId)=>'RH_'+C.hash({personId,cycleId}).slice(0,32);
 function rules(get,cycleId){
  const cycle=get('presentationCycles/'+cycleId);N(cycle,'RESOURCE_NOT_FOUND');
  const cfg=cycle.currentConfigRef?.kind==='config'&&get(f4ActivitiesModel_().path(cycle.currentConfigRef));
  const entry=cfg?.approved&&f4CycleParameters_().entry(f4CycleEffectiveConfiguration_(cfg,get).parameters,'attendanceDisciplinePolicyRef');
  if(entry?.status!=='DEFINED'||entry.value?.kind!=='policy')return {established:false,cycle};
  const p=get(f4ActivitiesModel_().path(entry.value)),v=p?.attendanceDiscipline;
  N(p?.approved===true&&p.active===true&&p.usage==='ATTENDANCE_DISCIPLINE'&&v?.schemaVersion==='attendance-discipline-rules-v1','RECONCILIATION_REQUIRED');
  N(typeof v.normativeReference==='string'&&v.normativeReference.trim()&&[v.informalBasisPoints,v.formalBasisPoints,v.limitBasisPoints].every(n=>Number.isInteger(n)&&n>0&&n<=10000)&&v.informalBasisPoints<v.formalBasisPoints&&v.formalBasisPoints<v.limitBasisPoints&&Number.isInteger(v.manifestationDays)&&v.manifestationDays>=1&&v.manifestationDays<=366,'RECONCILIATION_REQUIRED');
  return {established:true,cycle,policyId:p.id,policyRevision:p.revision,configId:cfg.id,rules:v};
 }
 function record(d){N(d?.schemaVersion==='member-attendance-case-v1'&&d.id===caseId(d.personId,d.cycleId)&&C.rev(d.revision)&&d.revision>0&&Array.isArray(d.history)&&d.history.length<=100&&['OPEN','CLOSED','DISMISSED'].includes(d.state),'RECONCILIATION_REQUIRED');return d;}
 function validate(r){
  N(r.commandType==='members.attendance.record','COMMAND_TYPE_INVALID');C.exact(r.target,['personId','linkId','cycleId']);
  N([r.target.personId,r.target.linkId,r.target.cycleId,r.decisionRef].every(C.id)&&r.evidenceRefs.length>0,'COMMAND_PAYLOAD_INVALID');
  C.exact(r.payload,['action','occurrenceDate','reason','documentReference','confirmed']);
  N(actions.includes(r.payload.action)&&r.payload.confirmed===true&&typeof r.payload.reason==='string'&&r.payload.reason.trim().length>=10&&r.payload.reason.length<=2000&&typeof r.payload.documentReference==='string'&&r.payload.documentReference.length<=800&&!/[<>\x00-\x08]/.test(r.payload.reason+r.payload.documentReference),'COMMAND_PAYLOAD_INVALID');C.civilStart(r.payload.occurrenceDate);return true;
 }
 function advance(old,r,policy,now){
  validate(r);N((policy.established||old?.disciplineSnapshot)&&policy.cycle.status==='OPEN','DOMAIN_STATE_INVALID','CYCLE_DISCIPLINE_POLICY_REQUIRED');
  const p=r.payload,day=Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd'),id=caseId(r.target.personId,r.target.cycleId);
  N(policy.cycle.civilStart<=p.occurrenceDate&&p.occurrenceDate<=policy.cycle.civilEnd&&p.occurrenceDate<=day,'COMMAND_PAYLOAD_INVALID');
  if(old){record(old);N(old.id===id&&old.linkId===r.target.linkId&&old.state==='OPEN','DOMAIN_STATE_INVALID');N(old.history.length<100,'SCOPE_TOO_LARGE');N(p.occurrenceDate>=old.history.at(-1).occurrenceDate,'DOMAIN_STATE_INVALID','CASE_CHRONOLOGY');}
  const history=old?.history||[],has=action=>history.some(h=>h.action===action);
  if(['FORMAL_WARNING','COMMUNICATION'].includes(p.action))N(p.documentReference.trim(),'HUMAN_DECISION_REQUIRED','DOCUMENT_REQUIRED');
  if(p.action==='FORMAL_WARNING')N(has('INFORMAL_WARNING'),'DOMAIN_STATE_INVALID','INFORMAL_WARNING_REQUIRED');
  if(p.action==='COMMUNICATION')N(has('FORMAL_WARNING')&&!has('COMMUNICATION'),'DOMAIN_STATE_INVALID','FORMAL_WARNING_REQUIRED');
  if(p.action==='MANIFESTATION')N(has('COMMUNICATION'),'DOMAIN_STATE_INVALID','COMMUNICATION_REQUIRED');
  N(!['INFORMAL_WARNING','FORMAL_WARNING'].includes(p.action)||!has(p.action),'DOMAIN_STATE_INVALID','WARNING_ALREADY_RECORDED');
  if(!old)N(['INFORMAL_WARNING','FOLLOW_UP'].includes(p.action),'DOMAIN_STATE_INVALID');
  const snapshot=old?.disciplineSnapshot||{policyId:policy.policyId,policyRevision:policy.policyRevision,configId:policy.configId,rules:policy.rules};
  const deadline=p.action==='COMMUNICATION'?new Date(C.civilStart(p.occurrenceDate)+(snapshot.rules.manifestationDays+1)*86400000).toISOString():old?.manifestationEndsAtExclusive||null;
  return {schemaVersion:'member-attendance-case-v1',id,revision:(old?.revision||0)+1,personId:r.target.personId,linkId:r.target.linkId,cycleId:r.target.cycleId,state:p.action==='CLOSE'?'CLOSED':'OPEN',disciplineSnapshot:snapshot,manifestationEndsAtExclusive:deadline,history:[...history,{operationId:r.operationId,action:p.action,occurrenceDate:p.occurrenceDate,recordedAt:F.Timestamp.fromMillis(now),actorUid:r.ownerUid,reason:p.reason.trim(),documentReference:p.documentReference.trim(),decisionRef:r.decisionRef,evidenceRefs:r.evidenceRefs,policyId:snapshot.policyId,policyRevision:snapshot.policyRevision,configId:snapshot.configId,normativeReference:snapshot.rules.normativeReference}]};
 }
 function resolve(get,r,now){
  const U=f4MemberFamilyInputs_();validate(r);const {request:q,proof}=U.authorization(get,r),id=r.target.personId,path='memberAttendanceCases/'+caseId(id,r.target.cycleId);
  N(q.kind==='EXPLICIT'&&q.capability==='MEMBERS_DISCIPLINE_EXECUTE'&&q.target.personId===id&&q.target.linkId===r.target.linkId&&q.target.effect===r.payload.action&&q.target.effectiveDate===r.payload.occurrenceDate,'AUTH_BINDING_INVALID');
  f4MemberActContent_().empty(get,r,now);f4MemberAuthz_().recheck(get,q,proof,now);
  const head=get('memberAggregateHeads/'+id),link=get('memberLinks/'+r.target.linkId),cyclePolicy=rules(get,r.target.cycleId),old=get(path);
  N(head?.revision===U.revision(r,'memberAggregateHeads/'+id)&&link?.idPessoa===id&&(old?.revision||0)===U.revision(r,path),'REVISION_CONFLICT');
  const range=f4MemberModel_().interval(link.startDate,link.endDate),at=C.civilStart(r.payload.occurrenceDate);N(range.a<=at&&at<range.z,'COMMAND_PAYLOAD_INVALID');
  N(f4MemberCommandCatalog_().equal([...q.target.resources].sort(),[path,'memberLinks/'+r.target.linkId].sort()),'AUTH_BINDING_INVALID');
  return {id,expected:head.revision,changes:[{path,expectedRevision:old?.revision||0,data:advance(old,r,cyclePolicy,now)}],now,boardExpected:{},mode:'M13',authorization:{request:q,proof},deadline:C.ms(proof.expiresAt)};
 }
 function change(c,id,old){record(c.data);N(c.path==='memberAttendanceCases/'+c.data.id&&c.data.personId===id&&c.data.revision===c.expectedRevision+1,'REVISION_CONFLICT');if(old){record(old);N(old.state==='OPEN'&&old.personId===id&&old.cycleId===c.data.cycleId&&old.linkId===c.data.linkId&&c.data.history.length===old.history.length+1&&C.hash(F.encode(c.data.history.slice(0,-1)))===C.hash(F.encode(old.history))&&C.hash(old.disciplineSnapshot)===C.hash(c.data.disciplineSnapshot),'REVISION_CONFLICT');}}
 function dismissal(get,r,now){
  const U=f4MemberFamilyInputs_(),context=get('memberPreparedIntents/'+r.operationId)?.disciplineContext;if(!context)return [];
  const path='memberAttendanceCases/'+caseId(r.target.personId,context.cycleId),d=record(get(path));
  N(d.id===context.caseId&&d.revision===context.caseRevision&&d.revision===U.revision(r,path)&&d.personId===r.target.personId&&d.linkId===r.target.linkId&&d.state==='OPEN'&&context.confirmed===true&&context.reason.trim().length>=10&&context.documentReference.trim(),'HUMAN_DECISION_REQUIRED');
  N(d.history.some(h=>h.action==='FORMAL_WARNING')&&d.history.some(h=>h.action==='COMMUNICATION')&&d.manifestationEndsAtExclusive&&Date.parse(d.manifestationEndsAtExclusive)<=now&&Date.parse(d.manifestationEndsAtExclusive)<=C.civilStart(r.payload.effectiveDate),'DOMAIN_STATE_INVALID','MANIFESTATION_PERIOD_NOT_COMPLETE');
  return [{path,expectedRevision:d.revision,data:{...d,revision:d.revision+1,state:'DISMISSED',dismissalEventId:'EVT_'+r.operationId,history:[...d.history,{operationId:r.operationId,action:'DISMISS',occurrenceDate:r.payload.effectiveDate,recordedAt:F.Timestamp.fromMillis(now),actorUid:r.ownerUid,reason:context.reason,documentReference:context.documentReference,decisionRef:r.decisionRef,evidenceRefs:r.evidenceRefs}]}}];
 }
 function allowedActions(d,canDismiss,now){if(d&&d.state!=='OPEN')return [];const has=a=>d?.history.some(h=>h.action===a),out=['FOLLOW_UP'];if(!has('INFORMAL_WARNING'))out.push('INFORMAL_WARNING');if(has('INFORMAL_WARNING')&&!has('FORMAL_WARNING'))out.push('FORMAL_WARNING');if(has('FORMAL_WARNING')&&!has('COMMUNICATION'))out.push('COMMUNICATION');if(has('COMMUNICATION'))out.push('MANIFESTATION');if(d)out.push('CLOSE');if(canDismiss&&d?.manifestationEndsAtExclusive&&Date.parse(d.manifestationEndsAtExclusive)<=now)out.push('DISMISS');return out;}
 return {actions,caseId,rules,record,validate,advance,resolve,change,dismissal,allowedActions};
}
