// Canonical authoring. Availability, identity and authorization remain in the
// existing ACT command pipeline. No import endpoint or historical promotion.
function f4ActivityModelDefinition_(d,load){
 const C=f4Core_(),M=f4ActivitiesModel_(),N=M.need;
 C.exact(d,['name','durationMinutes','visibility','countsAttendance','countsAbsence','allowsJustification','portalModel','purpose','organizer','participantTypes','allowedFormats','academicRequirements',...('attendancePolicyRef'in d?['attendancePolicyRef']:[])]);
 for(const k of ['name','purpose','organizer','academicRequirements'])N(typeof d[k]==='string'&&d[k].length<=(k==='name'?120:800)&&!/[<>\x00-\x08]/.test(d[k]),'VALIDATION_ERROR',k);
 N(d.name.trim()&&d.purpose.trim()&&d.organizer.trim(),'INCOMPLETE_REQUEST','MODEL_PURPOSE_ORGANIZER_REQUIRED');
 N(Number.isInteger(d.durationMinutes)&&d.durationMinutes>0&&d.durationMinutes<=1440,'VALIDATION_ERROR','MODEL_DURATION');
 N(['MEMBERS','BOARD','PUBLIC'].includes(d.visibility),'VALIDATION_ERROR','MODEL_VISIBILITY');
 for(const k of ['countsAttendance','countsAbsence','allowsJustification'])N(typeof d[k]==='boolean','VALIDATION_ERROR');
 N(!d.countsAbsence||d.countsAttendance,'VALIDATION_ERROR','ABSENCE_REQUIRES_ATTENDANCE');
 N(!d.allowsJustification||d.countsAbsence,'VALIDATION_ERROR','JUSTIFICATION_REQUIRES_ABSENCE');
 const list=(v,allowed)=>N(Array.isArray(v)&&v.length>0&&v.length<=allowed.length&&new Set(v).size===v.length&&v.every(s=>allowed.includes(s)),'VALIDATION_ERROR');
 list(d.participantTypes,['MEMBRO','INGRESSANTE','EGRESSO','PROFESSOR','CONVIDADO','EXTERNO']);list(d.allowedFormats,['PRESENCIAL','REMOTO','HIBRIDO']);
 const m=f4ActivityPortalModel_(d.portalModel);
 N(!m.requiresPrincipal||m.allowedPrincipalTypes.length>0,'INCOMPLETE_REQUEST','PRINCIPAL_TYPES_REQUIRED');
 N(!m.requiresMaterial||m.materialExtensions.length>0&&m.maxMaterialMb>0,'INCOMPLETE_REQUEST','MATERIAL_MODEL_REQUIRED');
 N(!m.memberPresentation||m.subtype==='APRESENTACAO_MEMBRO','VALIDATION_ERROR','MEMBER_PRESENTATION_SUBTYPE');
 N(!d.countsAttendance||d.attendancePolicyRef?.kind==='attendancePolicy','CONFIG_CAPABILITY_DISABLED','ATTENDANCE_POLICY_REQUIRED');
 let policyRevision=0;
 if(d.attendancePolicyRef){N(d.attendancePolicyRef.kind==='attendancePolicy','VALIDATION_ERROR');const policy=load(d.attendancePolicyRef,true);f4ActivityFrequencyPolicy_(policy);policyRevision=policy.revision;}
 return {...d,portalModel:m,policyRevision};
}
function f4CycleFrequencyPlanning_(){
 const M=f4ActivitiesModel_(),C=f4Core_(),N=M.need;
 function validate(plan,cycle,policy){
  C.exact(plan,['periodId','nature','civilStart','civilEnd','attendanceActivitiesCount','normativeMeetingsCount','classification','equivalenceEvidence','policyRef','approvalDate','approvingBody','approvalEvidence','decisionId']);
  N(M.id(plan.periodId)&&['ANNUAL','SEMESTER'].includes(plan.nature),'VALIDATION_ERROR');
  M.civil(plan.civilStart);M.civil(plan.civilEnd);
  N(M.gap(plan.civilStart,plan.civilEnd)>=0&&M.gap(cycle.civilStart,plan.civilStart)>=0&&M.gap(plan.civilEnd,cycle.civilEnd)>=0,'VALIDATION_ERROR','PLAN_PERIOD_OUTSIDE_CYCLE');
  for(const k of ['attendanceActivitiesCount','normativeMeetingsCount'])N(Number.isInteger(plan[k])&&plan[k]>=0&&plan[k]<=10000,'VALIDATION_ERROR','PLAN_COUNT');
  N(['PLANNED_MEETINGS','UNCONFIRMED'].includes(plan.classification),'VALIDATION_ERROR','PLAN_CLASSIFICATION');
  for(const k of ['equivalenceEvidence','approvalDate','approvingBody','approvalEvidence','decisionId'])N(typeof plan[k]==='string'&&plan[k].length<=800&&!/[<>\x00-\x08]/.test(plan[k]),'VALIDATION_ERROR');
  N(policy?.approved===true&&policy.active===true&&policy.usage==='FREQUENCY_PLANNING','CONFIG_CAPABILITY_DISABLED','FREQUENCY_PLANNING_POLICY_REQUIRED');
  const r=policy.frequencyPlanning;
  N(r&&r.classification==='PLANNED_MEETINGS'&&r.rounding==='FLOOR'&&Number.isInteger(r.maxAbsenceBasisPoints)&&r.maxAbsenceBasisPoints>=0&&r.maxAbsenceBasisPoints<=10000&&Number.isInteger(r.minimumAttendanceBasisPoints)&&r.minimumAttendanceBasisPoints>=0&&r.minimumAttendanceBasisPoints<=10000&&typeof r.absenceNorm==='string'&&r.absenceNorm.trim()&&typeof r.attendanceNorm==='string'&&r.attendanceNorm.trim(),'CONFIG_CAPABILITY_DISABLED','FREQUENCY_NORMATIVE_BASIS_REQUIRED');
  return {...plan,rules:{...r},policyRevision:policy.revision,absenceLimit:plan.classification==='PLANNED_MEETINGS'?Math.floor(plan.normativeMeetingsCount*r.maxAbsenceBasisPoints/10000):null};
 }
 function opening(c){if(c.frequencyPlanningRequired)N(c.frequencyPlanning?.periods?.length>0&&c.frequencyPlanning.periods.every(p=>p.state==='APPROVED'),'HUMAN_DECISION_REQUIRED','FREQUENCY_PLANNING_APPROVAL_REQUIRED');}
 return {validate,opening};
}
function f4ActivitiesModelFrequencyCommands_(x){
 const {M,N,p,load,revision,event,now,actor}=x,C=f4Core_();
 function decision(){N(actor.kind==='BOARD_ACTOR','AUTHORIZATION_DENIED');N(typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'HUMAN_DECISION_REQUIRED');return {reason:p.reason.trim(),actor,recordedAt:M.stamp(now)};}
 function history(ref,action){event('history',{subjectRef:ref,action,ownerPersonRef:'',visibility:'BOARD',summary:p.reason,occurredAt:M.stamp(now)});}
 function fence(){const r=M.ref('attendanceRegistry','current'),old=load(r,false,true);revision(r,{epoch:(old?.epoch||0)+1});}
 function model(){const d=load(p.typeRef,true);N(d.governance?.schemaVersion==='original-activity-model-v1','INVALID_STATE_TRANSITION','LEGACY_MODEL_PROTECTED');return d;}
 function cycle(){const c=load(p.cycleRef,true);N(c.status==='DRAFT','INVALID_STATE_TRANSITION','ORIGINAL_PLANNING_LOCKED');return c;}
 return {
  ACTIVITY_MODEL_DRAFT_SAVE(){
   const record=decision(),old=load(p.typeRef,true,true);N(!old||old.governance?.schemaVersion==='original-activity-model-v1'&&old.governance.state==='DRAFT','INVALID_STATE_TRANSITION','APPROVED_OR_LEGACY_MODEL_PROTECTED');
   N(p.typeRef.kind==='activityType','VALIDATION_ERROR');const definition=f4ActivityModelDefinition_(p.definition,load),{policyRevision,...fields}=definition;
   revision(p.typeRef,{...old,...fields,status:'INACTIVE',governance:{schemaVersion:'original-activity-model-v1',state:'DRAFT',definitionHash:M.hash(fields),attendancePolicyRevision:policyRevision,draftDecision:record}});
   history(p.typeRef,'ACTIVITY_ORIGINAL_MODEL_DRAFTED');fence();return {typeRef:p.typeRef,state:'DRAFT'};
  },
  ACTIVITY_MODEL_APPROVE(){
   const approval=decision(),d=model();N(d.governance.state==='DRAFT','INVALID_STATE_TRANSITION');N(p.definitionHash===d.governance.definitionHash,'STALE_REVISION','MODEL_DEFINITION_CHANGED');
   const fields={};for(const k of ['name','durationMinutes','visibility','countsAttendance','countsAbsence','allowsJustification','portalModel','purpose','organizer','participantTypes','allowedFormats','academicRequirements','attendancePolicyRef'])if(k in d)fields[k]=d[k];
   const validated=f4ActivityModelDefinition_(fields,load);N(validated.policyRevision===d.governance.attendancePolicyRevision,'STALE_REVISION','MODEL_POLICY_CHANGED');
   N(typeof p.evidence==='string'&&p.evidence.trim()&&p.evidence.length<=800,'HUMAN_DECISION_REQUIRED','MODEL_APPROVAL_EVIDENCE_REQUIRED');
   revision(p.typeRef,{...d,status:'ACTIVE',decision:{...approval,evidence:p.evidence},governance:{...d.governance,state:'APPROVED',approval:{...approval,evidence:p.evidence}}});history(p.typeRef,'ACTIVITY_ORIGINAL_MODEL_APPROVED');fence();return {typeRef:p.typeRef,state:'APPROVED'};
  },
  ACTIVITY_MODEL_AVAILABILITY_SET(){const record=decision(),d=model();N(d.governance.state==='APPROVED'&&['ACTIVE','INACTIVE'].includes(p.status),'INVALID_STATE_TRANSITION');revision(p.typeRef,{...d,status:p.status,availabilityDecision:record});history(p.typeRef,'ACTIVITY_MODEL_AVAILABILITY_CHANGED');fence();return {typeRef:p.typeRef,status:p.status};},
  CYCLE_FREQUENCY_PLAN_SAVE(){
   const record=decision(),c=cycle(),policy=load(p.plan.policyRef,true),plan=f4CycleFrequencyPlanning_().validate(p.plan,c,policy),periods=c.frequencyPlanning?.periods||[],old=periods.find(v=>v.periodId===plan.periodId);
   N(!old||old.state==='DRAFT','INVALID_STATE_TRANSITION','ORIGINAL_PLANNING_LOCKED');
   N(periods.filter(v=>v.periodId!==plan.periodId).every(v=>v.civilEnd<plan.civilStart||v.civilStart>plan.civilEnd),'CONFLICT','FREQUENCY_PERIOD_OVERLAP');N(old||periods.length<8,'SCOPE_TOO_LARGE');
   const next={...plan,state:'DRAFT',planRevision:(old?.planRevision||0)+1,draftDecision:record};next.definitionHash=M.hash(plan);
   revision(p.cycleRef,{...c,frequencyPlanningRequired:true,frequencyPlanning:{schemaVersion:'cycle-frequency-planning-v1',periods:[...periods.filter(v=>v.periodId!==plan.periodId),next]}});history(p.cycleRef,'CYCLE_FREQUENCY_PLAN_DRAFTED');fence();return {cycleRef:p.cycleRef,periodId:plan.periodId,absenceLimit:plan.absenceLimit,state:'DRAFT'};
  },
  CYCLE_FREQUENCY_PLAN_APPROVE(){
   const record=decision(),c=cycle(),periods=c.frequencyPlanning?.periods||[],d=periods.find(v=>v.periodId===p.periodId);
   N(d?.state==='DRAFT','INVALID_STATE_TRANSITION');N(d.definitionHash===p.definitionHash,'STALE_REVISION','FREQUENCY_PLAN_CHANGED');
   const policy=load(d.policyRef,true);N(policy.approved&&policy.active&&policy.revision===d.policyRevision,'STALE_REVISION','FREQUENCY_POLICY_CHANGED');
   N(d.classification==='PLANNED_MEETINGS'&&d.normativeMeetingsCount>0&&d.equivalenceEvidence.trim()&&d.approvalEvidence.trim()&&d.approvingBody.trim()&&d.decisionId.trim(),'HUMAN_DECISION_REQUIRED','ORIGINAL_PLAN_EVIDENCE_REQUIRED');
   M.civil(d.approvalDate);N(d.approvalDate<=Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd'),'VALIDATION_ERROR','FUTURE_APPROVAL_DATE');
   const next={...d,state:'APPROVED',approval:record,recognizedAt:M.stamp(now)};
   revision(p.cycleRef,{...c,frequencyPlanning:{...c.frequencyPlanning,periods:periods.map(v=>v.periodId===d.periodId?next:v)}});history(p.cycleRef,'CYCLE_FREQUENCY_PLAN_APPROVED');fence();return {cycleRef:p.cycleRef,periodId:d.periodId,absenceLimit:d.absenceLimit,state:'APPROVED'};
  }
 };
}
function f4ActivityModelFrequencyCompatRoutes_(){return {
 '/admin/atividades/modelos/original':'ACTIVITY_MODEL_DRAFT_SAVE',
 '/admin/atividades/modelos/aprovar':'ACTIVITY_MODEL_APPROVE',
 '/admin/atividades/modelos/disponibilidade':'ACTIVITY_MODEL_AVAILABILITY_SET',
 '/admin/ciclos-semestres/frequencia/salvar':'CYCLE_FREQUENCY_PLAN_SAVE',
 '/admin/ciclos-semestres/frequencia/aprovar':'CYCLE_FREQUENCY_PLAN_APPROVE'
};}
function f4ActivityModelFrequencyCompat_(options){
 const M=f4ActivitiesModel_(),C=f4Core_(),N=M.need,db=f4Firestore_().connect(options.config),ports=options.ports;
 function prepare(input){
  C.exact(input,['operationId','route','payload']);const cmd=f4ActivityModelFrequencyCompatRoutes_()[input.route],p=input.payload;N(cmd,'VALIDATION_ERROR');
  const model=cmd.startsWith('ACTIVITY_MODEL_'),ref=M.ref(model?'activityType':'cycle',p.id||(cmd==='ACTIVITY_MODEL_DRAFT_SAVE'?'cfg_'+input.operationId:''));
  const extra={ACTIVITY_MODEL_DRAFT_SAVE:['definition'],ACTIVITY_MODEL_APPROVE:['definitionHash','evidence'],ACTIVITY_MODEL_AVAILABILITY_SET:['status'],CYCLE_FREQUENCY_PLAN_SAVE:['plan'],CYCLE_FREQUENCY_PLAN_APPROVE:['periodId','definitionHash']}[cmd];
  C.exact(p,['id','revision','reason',...extra]);N(Number.isInteger(p.revision)&&p.revision>=0,'VALIDATION_ERROR');
  const auth=ports.readAuthority(['activityAttendanceRegistry/current'],false),old=db.get(M.path(ref));N((old?.revision||0)===p.revision,'STALE_REVISION');
  const payload={[model?'typeRef':'cycleRef']:ref,reason:p.reason};for(const k of extra)payload[k]=p[k];
  ports.recheckAuthority(auth);const prepared=ports.prepare({operationId:input.operationId,commandType:'ACT.'+cmd,payload});
  N(prepared.target.expected.some(e=>e.kind===ref.kind&&e.id===ref.id&&(old?e.revision===p.revision:e.exists===false)),'STALE_REVISION');return prepared;
 }
 return {prepare};
}
