// A promotion reconciles an existing future record. It never creates planning facts.
function f4ActivityPromotion_(){
 const M=f4ActivitiesModel_(),C=f4Core_(),N=M.need;
 const same=(a,b)=>C.hash(f4Firestore_().encode(a))===C.hash(f4Firestore_().encode(b));
 function instant(day,time){const m=typeof time==='string'&&time.trim().match(/^(\d{1,2})(?:h|:)(\d{2})$/);N(m&&+m[1]<24&&+m[2]<60,'UNRESOLVED_DEPENDENCY','SOURCE_TIME_REQUIRED');const v=Date.parse(day+'T'+m[1].padStart(2,'0')+':'+m[2]+':00-04:00');N(Number.isFinite(v),'VALIDATION_ERROR');return v;}
 function inspect(a,source,type,cycle,now){
  N(a?.recordMode==='HISTORICAL_PARTIAL'&&a.provenanceHead?.id&&a.evidenceScope,'INVALID_STATE_TRANSITION','PROMOTION_SOURCE_REQUIRED');
  N(source?.literal?.ID_ATIVIDADE===a.id&&source.literal.CICLO===cycle.id,'UNRESOLVED_DEPENDENCY','PROMOTION_SOURCE_MISMATCH');
  N(cycle.timezone==='America/Cuiaba','AUTHORIZATION_DENIED','PROMOTION_CYCLE_NOT_APPROVED');
  N(a.administrativeReview?.state==='PENDING'&&a.administrativeReview.cycleRef.id===cycle.id,'INVALID_STATE_TRANSITION','ACTIVITY_REVIEW_LOCKED');
  N(!a.realizationRef&&!a.realizationRecognizedByLegacyLabel&&a.legacyStatus==='PLANEJADA'&&source.literal.STATUS_OPERACIONAL==='PLANEJADA'&&!['CANCELLED','CANCELADA','REALIZADA','REALIZED','ARCHIVED','ARQUIVADA'].includes(a.status),'INVALID_STATE_TRANSITION','PROMOTION_NOT_FUTURE_PLANNED');
  N(M.civil(a.civilDate)&&a.civilDate>M.day(now),'INVALID_STATE_TRANSITION','PROMOTION_DATE_NOT_FUTURE');
  const literalDay=a.civilDate.split('-').reverse().join('/');N(source.literal.DATA_ATIVIDADE===literalDay||source.literal.DATA_ATIVIDADE===a.civilDate,'UNRESOLVED_DEPENDENCY','SOURCE_DATE_MISMATCH');
  N(type?.status==='ACTIVE'&&type.portalModel&&type.decision,'CONFIG_CAPABILITY_DISABLED','APPROVED_MODEL_REQUIRED');
  const model=type.originRef?.lineage?.sourceSnapshot;N(model&&['TIPO_ATIVIDADE','SUBTIPO_ATIVIDADE','CLASSIFICACAO_REUNIAO'].every(k=>source.literal[k]&&source.literal[k]===model[k]),'UNRESOLVED_DEPENDENCY','SOURCE_MODEL_MISMATCH');
  for(const [key,value]of [['CONTA_PRESENCA',type.countsAttendance],['CONTA_FALTA',type.countsAbsence??type.countsAttendance]])N(['SIM','NAO'].includes(source.literal[key])&&(source.literal[key]==='SIM')===value,'HUMAN_DECISION_REQUIRED','SOURCE_CRITERIA_MISMATCH');
  if(source.literal.CARGA_HORARIA!==''&&source.literal.CARGA_HORARIA!==undefined)N(Number(String(source.literal.CARGA_HORARIA).replace(',','.'))*60===type.durationMinutes,'HUMAN_DECISION_REQUIRED','SOURCE_DURATION_MISMATCH');
  const startsAt=instant(a.civilDate,source.literal.HORARIO_INICIO),endsAt=instant(a.civilDate,source.literal.HORARIO_FIM);N(endsAt>startsAt&&endsAt-startsAt<=86400000,'VALIDATION_ERROR');
  N(typeof a.title==='string'&&a.title.trim()&&typeof a.location==='string'&&a.location.trim(),'INCOMPLETE_REQUEST');
  N(a.personRef?.resolution==='RESOLVED'&&a.personRef.ref?.kind==='person'&&source.literal.ID_PESSOA_PRINCIPAL===a.personRef.ref.id,'UNRESOLVED_DEPENDENCY','PRINCIPAL_REFERENCE_REQUIRED');
  return {startsAt,endsAt,personRef:a.personRef.ref};
 }
 function handle(x){const {p,load,get,M,N,revision,event,now,operationId,actor}=x;
  N(actor.kind==='BOARD_ACTOR'&&typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'HUMAN_DECISION_REQUIRED');
  const old=load(p.activityRef,true),cycle=load(p.cycleRef,true),type=load(p.typeRef,true),source=get('migrationEvidence/'+old.provenanceHead?.id),v=inspect(old,source,type,cycle,now);
  N(p.sourceHash===C.hash(f4Firestore_().encode(source)),'STALE_REVISION','PROMOTION_SOURCE_CHANGED');
  N(type.portalModel.memberPresentation===true,'CONFIG_CAPABILITY_DISABLED','PROMOTION_MEMBER_MODEL_ONLY');
  N(!p.sessionRef&&!old.presentationSessionRef,'UNRESOLVED_DEPENDENCY','EXISTING_SESSION_REQUIRES_SEPARATE_RECONCILIATION');
  const refs=old.historicalPresentationRefs||[];N(refs.length>0&&refs.length<=2,'UNRESOLVED_DEPENDENCY','PRESENTATION_COVERAGE_REQUIRED');
  const seen=new Set();for(const ref of refs){N(ref.kind==='presentation'&&!seen.has(ref.id),'UNRESOLVED_DEPENDENCY','PRESENTATION_IDS_CHANGED');seen.add(ref.id);const pr=get(M.path(ref)),e=pr?.provenanceHead&&get(M.path(pr.provenanceHead));N(pr?.recordMode==='HISTORICAL_PARTIAL'&&pr.activityRef?.resolution==='RESOLVED'&&pr.activityRef.ref?.id===old.id&&pr.personRef?.resolution==='RESOLVED'&&pr.personRef.ref?.id===v.personRef.id&&!pr.realizationRef&&e?.literal?.ID_APRESENTACAO===pr.id&&e.literal.ID_ATIVIDADE===old.id,'UNRESOLVED_DEPENDENCY','IMPORTED_PRESENTATION_LINK_MISMATCH');}
  // Promotion establishes the existing activity's operational facts only. A
  // session, incidence, allocation, theme approval or realization is not inferred
  // from a planned date. Those commands keep their own institutional gates.
  f4ActivitiesMemberActivityPerson_(get,v.personRef,v.startsAt);
  const context=f4ActivityCycleContext_(get,p.cycleRef,p.academicSemesterRef,old.civilDate),policy=load(type.attendancePolicyRef,true),ref=M.ref('attendanceCollection',old.id),collection=load(ref,true,true);
  N(!collection,'INVALID_STATE_TRANSITION','EXISTING_ATTENDANCE_REQUIRES_RECONCILIATION');
  const modelContent=f4ActivityContent_(type.portalModel,{principal:{kind:'CORE_PERSON',participantType:'MEMBRO',participationRole:type.portalModel.principalRole,institutionLabel:'',personRef:v.personRef}},get,v.startsAt,true);
  const next={...old,...context,recordMode:'NEW_V2',status:'SCHEDULED',startsAt:M.stamp(v.startsAt),endsAt:M.stamp(v.endsAt),typeRef:p.typeRef,typeRevision:type.revision,typeSnapshot:f4ActivityOperationalSnapshot_(type,v.startsAt,v.endsAt),personRef:v.personRef,policyRef:type.attendancePolicyRef,modelContent,details:{...old.details,formato:old.format},promotion:{schemaVersion:'activity-operational-promotion-v1',fromMode:'HISTORICAL_PARTIAL',sourceRef:old.provenanceHead,sourceHash:C.hash(f4Firestore_().encode(source)),previousRevision:old.revision,typeRevision:type.revision,cycleRevision:cycle.revision,presentationRefs:refs,operationId,promotedAt:M.stamp(now),reason:p.reason.trim(),actorPersonRef:actor.personRef}};
  revision(p.activityRef,next);
  if(type.countsAttendance){const w=f4ActivityJustificationWindow_(policy,v.startsAt,v.endsAt);f4AttendanceCommands_({...x,event:(kind,data)=>{N(kind==='history'&&data.action==='ATTENDANCE_CONFIGURED','INVALID_STATE_TRANSITION');return null;},p:{activityRef:p.activityRef,attendancePolicyRef:type.attendancePolicyRef,opensAt:v.startsAt-policy.opensBeforeMinutes*60000,closesAt:v.endsAt+policy.closesAfterMinutes*60000,justificationOpensAt:w.opensAt,justificationEndsAt:w.endsAt,acceptLateJustification:policy.acceptLateJustification,decision:{reasonCode:'EXPLICIT_HISTORICAL_PROMOTION',reason:p.reason}}}).ATTENDANCE_CONFIGURE();}
  event('history',{action:'ACTIVITY_OPERATIONAL_PROMOTED',subjectRef:p.activityRef,ownerPersonRef:'',visibility:'BOARD',summary:p.reason.trim(),previousActivityRevision:old.revision,sourceRef:old.provenanceHead,occurredAt:M.stamp(now)});
  const registry=M.ref('attendanceRegistry','current'),r=load(registry,false,true);x.save(registry,{epoch:(r?.epoch||0)+1});
  return {activityRef:p.activityRef,revision:old.revision+1,reviewState:'PENDING',recordMode:'NEW_V2',presentationRefs:refs,planningCreated:false};
 }
 return {inspect,handle};
}
function f4ActivityOperationalSnapshot_(type,start,end){return {name:type.name,countsAttendance:type.countsAttendance,...(type.countsAbsence!==undefined?{countsAbsence:type.countsAbsence}:{}),durationMinutes:type.durationMinutes??Math.round((end-start)/60000),visibility:type.visibility,allowsJustification:type.allowsJustification,...(type.portalModel?{portalModel:type.portalModel}:{})};}
// Pin the complete approved intent, independent of harmless precondition ordering.
function f4ActivityPromotionIntentHash_(r){const M=f4ActivitiesModel_(),key=r=>M.path({kind:r.kind,id:r.id,...(r.parent?{parent:r.parent}:{})}),sort=v=>[...v].sort((a,b)=>key(a).localeCompare(key(b)));return f4Core_().hash(f4Firestore_().encode({schemaVersion:r.schemaVersion,operationId:r.operationId,commandType:r.commandType,payload:r.payload,target:{...r.target,scope:{...r.target.scope,targetRefs:sort(r.target.scope.targetRefs)},expected:sort(r.target.expected)},evidenceRefs:r.evidenceRefs,decisionRef:r.decisionRef}));}
