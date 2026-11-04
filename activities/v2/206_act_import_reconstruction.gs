// Pure closed-manifest planner. Authority is supplied by the canonical handler below.
// The caller must lock/compare the complete read set and commit archives, revisions,
// receipt, history and fences atomically. This function never authorizes an actor.
function f4ImportedActivityReconstructionPlan_(input) {
 const M=f4ActivitiesModel_(),N=M.need,{intent,get,now,actorRef}=input;
 const normalized=v=>v&&typeof v.toMillis==='function'?{_seconds:v.seconds,_nanoseconds:v.nanoseconds}:Array.isArray(v)?v.map(normalized):v&&typeof v==='object'?Object.fromEntries(Object.entries(v).map(([k,x])=>[k,normalized(x)])):v;
 const hash=v=>f4Core_().hash(JSON.parse(JSON.stringify(normalized(v)))),eq=(a,b)=>hash(a)===hash(b);
 N(intent?.schemaVersion==='act-import-reconstruction-intent-v1'&&M.id(intent.operationId)&&actorRef?.kind==='person'&&M.id(actorRef.id),'VALIDATION_ERROR');
 const syntheticCycle=intent.cycleId!=='GEAPA_2026'&&!f4ExecutionEnvironment_().production&&/^SYN_IMPORT_[A-Za-z0-9_]+$/.test(intent.cycleId);
 N((intent.cycleId==='GEAPA_2026'||syntheticCycle)&&intent.historicalApplicability?.status==='PROVEN'&&intent.historicalApplicability.sourceRefs?.length>0,'UNRESOLVED_DEPENDENCY','HISTORICAL_CRITERIA_NOT_PROVEN');
 N(Array.isArray(intent.readSet)&&intent.readSet.length>0&&intent.readSet.length<=250&&new Set(intent.readSet.map(x=>x.path)).size===intent.readSet.length,'VALIDATION_ERROR');
 const loaded={},reads=new Set();for(const r of intent.readSet){N(typeof r.hash==='string'&&/^[a-f0-9]{64}$/.test(r.hash),'VALIDATION_ERROR');const d=get(r.path);N(d&&hash(d)===r.hash,'STALE_REVISION',r.path);loaded[r.path]=d;}
 function load(kind,id){const path=M.path(M.ref(kind,id));N(Object.hasOwn(loaded,path),'UNRESOLVED_DEPENDENCY','UNSEALED_DEPENDENCY');reads.add(path);return loaded[path];}
 const a=load('activity',intent.activityId),source=load('evidence',a.provenanceHead.id);
 if(intent.variant)return f4ImportedNonAttendancePlan_({intent,get,load,now,actorRef,a,source,hash});
 const c=load('attendanceCollection',a.id),l={...source.literal};
 const cycle=load('cycle',intent.cycleId),type=load('activityType',intent.typeId),policy=load('attendancePolicy',intent.policyId);
 let adjustment=null;
 if(intent.documentaryAdjustmentRef){
  N(intent.documentaryAdjustmentRef.kind==='policy','VALIDATION_ERROR');adjustment=load('policy',intent.documentaryAdjustmentRef.id);
  N(adjustment.usage==='IMPORTED_ACTIVITY_DOCUMENTARY_ADJUSTMENT'&&adjustment.active===true&&adjustment.approved===true&&adjustment.activityId===a.id&&adjustment.sourceHash===hash(source)&&adjustment.historicalApprovalClaimed===false&&adjustment.reason&&adjustment.decisionEvidenceRef?.kind==='evidence','UNRESOLVED_DEPENDENCY','DOCUMENTARY_ADJUSTMENT_NOT_PROVEN');
  N(hash(load('evidence',adjustment.decisionEvidenceRef.id))===adjustment.decisionEvidenceHash,'STALE_REVISION','DOCUMENTARY_DECISION_CHANGED');
  if(adjustment.times){const t=adjustment.times;N(Object.keys(t).sort().join(',')==='durationMinutes,end,start'&&Number.isInteger(t.durationMinutes)&&t.durationMinutes>0,'VALIDATION_ERROR');Object.assign(l,{HORARIO_INICIO:t.start,HORARIO_FIM:t.end,CARGA_HORARIA:String(t.durationMinutes/60)});}
 }
 if(syntheticCycle)N(cycle.originRef?.kind==='SYNTHETIC_RECONCILIATION_FIXTURE','AUTHORIZATION_DENIED');
 N(a.recordMode==='HISTORICAL_PARTIAL'&&c.recordMode==='HISTORICAL_PARTIAL'&&c.id===a.id&&c.activityRef?.id===a.id,'INVALID_STATE_TRANSITION');
 N(l?.ID_ATIVIDADE===a.id&&l.CICLO===cycle.id&&l.STATUS_OPERACIONAL==='REALIZADA'&&a.legacyStatus==='REALIZADA'&&!a.status&&!a.realizationRef,'UNRESOLVED_DEPENDENCY','REALIZATION_SOURCE');
 M.civil(a.civilDate);N(a.civilDate<='2026-10-22'&&cycle.status==='OPEN'&&cycle.civilStart<=a.civilDate&&cycle.civilEnd>=a.civilDate,'INVALID_STATE_TRANSITION','CYCLE_OR_DATE');
 N(c.state==='CLOSED'&&!c.lastFinalizedAt&&!c.closedAt&&!c.completionRecognition,'INVALID_STATE_TRANSITION','ACTUAL_FINALIZATION_PRESERVED');
 N(type.recordMode==='NEW_V2'&&type.decision&&type.attendancePolicyRef?.id===policy.id&&policy.approved&&type.portalModel,'UNRESOLVED_DEPENDENCY','APPROVED_TEMPLATE_REQUIRED');
 const template=type.originRef?.lineage?.sourceSnapshot;
 if(intent.modelEquivalenceRef){
  N(intent.modelEquivalenceRef.kind==='policy','VALIDATION_ERROR');
  const e=load('policy',intent.modelEquivalenceRef.id);
  N(e.usage==='IMPORTED_MODEL_EQUIVALENCE'&&e.active===true&&e.approved===true&&e.activityId===a.id&&e.cycleId===cycle.id&&e.typeRef?.id===type.id&&e.typeRevision===type.revision&&e.sourceHash===hash(source)&&e.historicalApprovalClaimed===false&&e.evidence&&e.reason,'UNRESOLVED_DEPENDENCY','MODEL_EQUIVALENCE_NOT_PROVEN');
  N(eq(e.sourceClassification,Object.fromEntries(['TIPO_ATIVIDADE','SUBTIPO_ATIVIDADE','CLASSIFICACAO_REUNIAO'].map(k=>[k,l[k]])))&&eq(e.criteria,f4ActivityConfigNormativeFields_('activityType',type)),'UNRESOLVED_DEPENDENCY','MODEL_EQUIVALENCE_CRITERIA_CHANGED');
  if(type.governance?.schemaVersion==='original-activity-model-v1')N(e.definitionHash===type.governance.definitionHash&&type.status==='ACTIVE'&&type.governance.state==='APPROVED'&&(f4OriginalModelUsage_(get,type,cycle.id)||e.scope==='IMPORTED_ACTIVITY_ONLY'&&adjustment?.modelEquivalenceRef?.id===e.id),'UNRESOLVED_DEPENDENCY','ORIGINAL_MODEL_NOT_APPROVED_FOR_CYCLE');
 }else N(template&&['TIPO_ATIVIDADE','SUBTIPO_ATIVIDADE','CLASSIFICACAO_REUNIAO'].every(k=>l[k]&&l[k]===template[k]),'UNRESOLVED_DEPENDENCY','MODEL_CLASSIFICATION_MISMATCH');
 N(eq(intent.typeCriteria,f4ActivityConfigNormativeFields_('activityType',type))&&intent.policyRevision===policy.revision&&intent.typeRevision===type.revision,'STALE_REVISION','CONFIGURATION_CHANGED');
 const overrides=adjustment?.snapshotOverrides||{},pm=overrides.portalModel||{};
 N(Object.keys(overrides).every(k=>['countsAbsence','allowsJustification','portalModel'].includes(k))&&Object.keys(pm).every(k=>['requiresPrincipal','requiresAxis','requiresMaterial','generatesCertificate'].includes(k))&&Object.values(pm).every(v=>typeof v==='boolean')&&['countsAbsence','allowsJustification'].every(k=>!(k in overrides)||typeof overrides[k]==='boolean'),'VALIDATION_ERROR','DOCUMENTARY_OVERRIDE_FIELDS');
 const effective={...type,...overrides,portalModel:{...type.portalModel,...pm}};
 N(l.CONTA_PRESENCA==='SIM'&&effective.countsAttendance===true&&['SIM','NAO'].includes(l.CONTA_FALTA)&&(l.CONTA_FALTA==='SIM')===effective.countsAbsence,'UNRESOLVED_DEPENDENCY','ATTENDANCE_CRITERIA_CONFLICT');
 N(!effective.allowsJustification||effective.countsAbsence,'VALIDATION_ERROR');
 N(!l.PERMITE_JUSTIFICATIVA||['SIM','NAO'].includes(l.PERMITE_JUSTIFICATIVA)&&(l.PERMITE_JUSTIFICATIVA==='SIM')===effective.allowsJustification,'UNRESOLVED_DEPENDENCY','JUSTIFICATION_CRITERIA_CONFLICT');
 const clock=s=>typeof s==='string'&&/^\d{2}[:h]\d{2}$/.test(s)&&Number(s.slice(0,2))<24&&Number(s.slice(3))<60;
 N(clock(l.HORARIO_INICIO)&&clock(l.HORARIO_FIM),'UNRESOLVED_DEPENDENCY','DOCUMENTED_TIMES_REQUIRED');
 const start=Date.parse(a.civilDate+'T'+l.HORARIO_INICIO.replace('h',':')+':00-04:00'),end=Date.parse(a.civilDate+'T'+l.HORARIO_FIM.replace('h',':')+':00-04:00');
 const duration=Number(String(l.CARGA_HORARIA).replace(',','.'))*60;
 N(end>start&&Number.isSafeInteger(duration)&&duration>0&&duration<=1440,'UNRESOLVED_DEPENDENCY','TIME_CONVENTION_OR_DURATION');
 if(adjustment?.times)N(end-start===duration*60000,'VALIDATION_ERROR','ADMINISTRATIVE_DURATION_MISMATCH');
 N(intent.closureEvidenceRef?.kind==='evidence','UNRESOLVED_DEPENDENCY','CLOSURE_AUDIT_REQUIRED');
 const closureSource=load('evidence',intent.closureEvidenceRef.id);
 N(closureSource.schemaVersion==='act-closure-evidence-v1'&&closureSource.activityId===a.id&&closureSource.auditComplete===true&&Array.isArray(closureSource.sourceRefs)&&closureSource.sourceRefs.length>0&&Array.isArray(closureSource.events),'UNRESOLVED_DEPENDENCY','CLOSURE_AUDIT_REQUIRED');
 const events=closureSource.events.slice().sort((x,y)=>Date.parse(x.at)-Date.parse(y.at));
 for(const e of events)N(['CHAMADA_FINALIZADA','CHAMADA_REABERTA','CHAMADA_SALVA','CHAMADA_RASCUNHO_SALVO'].includes(e.type)&&e.active===true&&e.success===true&&e.activityId===a.id&&e.sourceRef&&typeof e.sourceHash==='string'&&/^[a-f0-9]{64}$/.test(e.sourceHash)&&Number.isFinite(Date.parse(e.at))&&Date.parse(e.at)<=now,'UNRESOLVED_DEPENDENCY','INVALID_CLOSURE_EVIDENCE');
 let latest=events.at(-1),reopeningRecognition=null,closure,closureMetadata;
 if(latest?.type==='CHAMADA_REABERTA'&&intent.reopeningRecognitionRef){
  const r=load('policy',intent.reopeningRecognitionRef.id),facts=intent.factIds.map(id=>load('attendance',id)),material=facts.map(f=>({id:f.id,personRef:f.personRef,status:f.status,countsAttendance:f.countsAttendance,countsAbsence:f.countsAbsence,applicableAtDate:f.applicableAtDate}));
  N(r.usage==='IMPORTED_REOPENING_RECOGNITION'&&r.approved===true&&r.active===true&&r.activityId===a.id&&r.closureEvidenceHash===hash(closureSource)&&r.afterMaterialHash===hash(material)&&r.beforeMaterialHash===r.afterMaterialHash&&r.noMaterialChange===true&&r.reason&&r.comparisonEvidenceRef?.kind==='evidence','UNRESOLVED_DEPENDENCY','REOPENING_MATERIAL_CHANGE');
  const comparison=load('evidence',r.comparisonEvidenceRef.id);
  N(hash(comparison)===r.comparisonEvidenceHash&&hash(comparison.materialBefore)===r.beforeMaterialHash&&hash(comparison.materialAfter)===r.afterMaterialHash,'STALE_REVISION','REOPENING_COMPARISON_CHANGED');
  const final=events.filter(e=>e.type==='CHAMADA_FINALIZADA').at(-1);N(final&&events.filter(e=>Date.parse(e.at)>Date.parse(final.at)).every(e=>e.type==='CHAMADA_REABERTA'),'UNRESOLVED_DEPENDENCY','REOPENING_MATERIAL_EVENT');
  reopeningRecognition={approvalRef:intent.reopeningRecognitionRef,reopenedAt:latest.at,recognizedAt:M.stamp(now),originalEventPreserved:true,materialHash:r.afterMaterialHash};latest=final;
 }
 if(latest){
  N(latest.type==='CHAMADA_FINALIZADA'&&!events.some(e=>e!==latest&&Date.parse(e.at)===Date.parse(latest.at)&&e.type!==latest.type),'UNRESOLVED_DEPENDENCY','CLOSURE_STATE_REQUIRES_REVIEW');
  N(['LEGACY_PORTAL_FINALIZATION','HISTORICAL_RECONCILIATION','HISTORICAL_ADMINISTRATIVE_RECOVERY'].includes(latest.basis),'UNRESOLVED_DEPENDENCY','INVALID_CLOSURE_EVIDENCE');
  closure=Date.parse(latest.at);N(closure>=Date.parse(a.civilDate+'T00:00:00-04:00'),'UNRESOLVED_DEPENDENCY','CLOSURE_PRECEDES_ACTIVITY');
  closureMetadata={kind:latest.basis,sourceRef:latest.sourceRef,sourceHash:latest.sourceHash,documentedAt:latest.at,notCapturedAtOriginalClosure:latest.basis!=='LEGACY_PORTAL_FINALIZATION'};
 }else{
  N(closureSource.equivalentRecordAbsent===true,'UNRESOLVED_DEPENDENCY','CLOSURE_ABSENCE_NOT_PROVEN');
  closure=Date.parse(a.civilDate+'T21:00:00-04:00');N(end<=closure&&closure<=now,'UNRESOLVED_DEPENDENCY','TIME_CONVENTION_OR_DURATION');
  N(M.day(closure)===a.civilDate&&Utilities.formatDate(new Date(closure),'America/Cuiaba','HH:mm')==='21:00','VALIDATION_ERROR');
  closureMetadata={kind:'ADMINISTRATIVE_CONVENTION',conventionalTime:'21:00',notCapturedAtOriginalClosure:true};
 }
 let personRef=null,modelContent={};
 if(adjustment?.principals){
  N(Array.isArray(adjustment.principals)&&adjustment.principals.length>0&&adjustment.principals.length<=8&&new Set(adjustment.principals.map(p=>p.personRef?.id)).size===adjustment.principals.length,'VALIDATION_ERROR','DOCUMENTED_PRINCIPALS');
  const principals=adjustment.principals.map(p=>{N(p.personRef?.kind==='person'&&p.participationRole==='PALESTRANTE','VALIDATION_ERROR');const identity=load('person',p.personRef.id);N(identity.idPessoa===p.personRef.id&&!identity.technicalAccount,'UNRESOLVED_DEPENDENCY','PRINCIPAL_IDENTITY');return {kind:'CORE_PERSON',personRef:p.personRef,participantType:p.participantType||'EXTERNO',participationRole:'PALESTRANTE',institutionLabel:''};});
  modelContent={principals};if(principals.length===1){personRef=principals[0].personRef;modelContent.principal=principals[0];}
 }else if(effective.portalModel.requiresPrincipal){
  N(type.portalModel.memberPresentation&&a.personRef?.resolution==='RESOLVED'&&a.personRef.ref?.kind==='person','UNRESOLVED_DEPENDENCY','DOCUMENTED_PRINCIPAL_REQUIRED');personRef=a.personRef.ref;
  const identity=load('person',personRef.id);N(identity.idPessoa===personRef.id&&!identity.technicalAccount,'UNRESOLVED_DEPENDENCY','PRINCIPAL_IDENTITY');
  N(a.historicalPresentationRefs?.length>0,'UNRESOLVED_DEPENDENCY','PRESENTATION_LINEAGE');
  for(const r of a.historicalPresentationRefs){const pr=load('presentation',r.id);N(pr.activityRef?.resolution==='RESOLVED'&&pr.activityRef.ref.id===a.id&&pr.personRef?.resolution==='RESOLVED'&&pr.personRef.ref.id===personRef.id,'UNRESOLVED_DEPENDENCY','PRESENTATION_LINEAGE');}
  modelContent={principal:{kind:'CORE_PERSON',personRef,participantType:'MEMBRO',participationRole:type.portalModel.principalRole,institutionLabel:''}};
 }
 const frequencyBasis=f4ActivityFrequencyPolicy_(policy),semester=load('academicSemester',intent.academicSemesterId);
 N(semester.cycleId===cycle.id&&semester.civilStart<=a.civilDate&&semester.civilEnd>=a.civilDate&&(cycle.academicSemesterRefs||[]).some(r=>r.id===semester.id),'UNRESOLVED_DEPENDENCY','SEMESTER_ASSOCIATION');
 N(Array.isArray(intent.factIds)&&intent.factIds.length>0&&intent.factIds.length<=50&&new Set(intent.factIds).size===intent.factIds.length&&c.recordedFactCount===intent.factIds.length,'UNRESOLVED_DEPENDENCY','COMPLETE_FACT_SET_REQUIRED');
 const facts=intent.factIds.map(id=>load('attendance',id)),people=new Set(),factsAfter=[],pendingWrites=[];
 const recognition={schemaVersion:'act-import-recognition-v1',kind:'DOCUMENTARY_RECOGNITION',operationId:intent.operationId,recognizedAt:M.stamp(now),recognizedBy:actorRef,manifestHash:hash(intent),historicalApplicability:intent.historicalApplicability};
 const evolve=(d,patch)=>({...d,...patch,revision:d.revision+1,recordMode:'NEW_V2',operationId:intent.operationId,originRef:d.originRef||d.provenanceHead,updatedAt:M.stamp(now),importRecognition:{...recognition,previousRevision:d.revision,previousMode:d.recordMode}});
 const summary={totalRegistros:facts.length,totalPresentes:0,totalFaltas:0,totalNaoSeAplica:0};
 for(const f of facts){
  N(f.recordMode==='HISTORICAL_PARTIAL'&&f.activityRef?.id===a.id&&f.personRef?.kind==='person'&&!people.has(f.personRef.id)&&f.historicalAttendanceFact===true&&!f.officialRecordedAt&&!f.officialCountBasis&&!f.frequencyBasis,'UNRESOLVED_DEPENDENCY','UNMODIFIED_IMPORTED_FACT_REQUIRED');people.add(f.personRef.id);
  const evidence=load('evidence',f.provenanceHead.id),literal=evidence.literal;
  N(literal?.ID_ATIVIDADE===a.id&&literal.ID_PESSOA===f.personRef.id&&literal.ID_REGISTRO_PRESENCA===f.legacyId&&literal.CICLO===cycle.id,'UNRESOLVED_DEPENDENCY','FACT_LINEAGE');
  const originalCode=({P:'PRESENT',F:'ABSENT',J:'JUSTIFIED',A:'EXCUSED','N/A':'NOT_APPLICABLE'})[literal.CODIGO_PRESENCA];
  const documentedJustification=literal.CODIGO_PRESENCA==='F'&&literal.STATUS_PRESENCA==='JUSTIFICADA'&&literal.STATUS_JUSTIFICATIVA==='DEFERIDA'&&literal.DECISAO_JUSTIFICATIVA==='FALTA_JUSTIFICADA'&&literal.VALOR_DEPOIS_JUSTIFICATIVA==='FALTA_JUSTIFICADA';
  N(originalCode===f.status||documentedJustification&&f.status==='JUSTIFIED','UNRESOLVED_DEPENDENCY','FACT_RESULT_CONFLICT');
  N(typeof f.countsAttendance==='boolean'&&typeof f.countsAbsence==='boolean'&&typeof f.applicableAtDate==='boolean','UNRESOLVED_DEPENDENCY','COUNT_BASIS_REQUIRED');
  N(effective.countsAbsence||f.countsAbsence===false,'UNRESOLVED_DEPENDENCY','ACTIVITY_ABSENCE_EXCEPTION_CONFLICT');
  N((literal.CONTA_PRESENCA==='SIM')===f.countsAttendance&&(literal.CONTA_FALTA==='SIM')===f.countsAbsence&&['SIM','NAO'].includes(literal.CONTA_PRESENCA)&&['SIM','NAO'].includes(literal.CONTA_FALTA),'UNRESOLVED_DEPENDENCY','COUNT_SOURCE_CONFLICT');
  if(f.status==='NOT_APPLICABLE')N(f.applicableAtDate===false,'UNRESOLVED_DEPENDENCY','NON_APPLICABILITY_CONFLICT');
  else {N(f.applicableAtDate===true&&f.membershipRef?.kind==='membership','UNRESOLVED_DEPENDENCY','HISTORICAL_MEMBERSHIP_REQUIRED');const link=load('membership',f.membershipRef.id);N(link.idPessoa===f.personRef.id&&['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE'].includes(link.type)&&link.startDate<=a.civilDate&&(!link.endDate||link.endDate>=a.civilDate),'UNRESOLVED_DEPENDENCY','HISTORICAL_MEMBERSHIP_MISMATCH');}
  const related=(intent.justificationIds||[]).map(id=>load('justification',id)).filter(j=>j.attendanceRef?.id===f.id),pending=related.filter(j=>j.state==='PENDING');
  N(pending.length<=1&&!related.some(j=>['PREVIOUS','CORRECTION_REQUESTED'].includes(j.state)),'UNRESOLVED_DEPENDENCY','JUSTIFICATION_REQUIRES_INDIVIDUAL_REVIEW');
  const patch={cycleId:cycle.id,academicSemesterId:semester.id,frequencyBasis,durationMinutes:duration,officialCountBasis:{countsAttendance:f.countsAttendance,countsAbsence:f.countsAbsence,applicableAtDate:f.applicableAtDate},officialRecordedAt:M.stamp(now),officialRecordedAtBasis:'DOCUMENTARY_RECOGNITION_NOT_ORIGINAL_ENTRY'};
  if(pending.length){const j=pending[0],je=load('evidence',j.provenanceHead.id);N(j.recordMode==='HISTORICAL_PARTIAL'&&j.activityRef?.id===a.id&&j.personRef?.id===f.personRef.id&&f.status==='ABSENT'&&effective.allowsJustification&&(!f.pendingJustificationRef||f.pendingJustificationRef.id===j.id)&&!j.basisAttendanceRevision&&je.literal?.ID_JUSTIFICATIVA===j.id,'UNRESOLVED_DEPENDENCY','PENDING_REBASE_NOT_PROVEN');patch.pendingJustificationRef=M.ref('justification',j.id);pendingWrites.push(evolve(j,{basisAttendanceRevision:f.revision+1,basisRecognition:{previousBasisAttendanceRevision:null,recognizedAt:M.stamp(now),operationId:intent.operationId,notOriginalSubmissionRevision:true,ordinaryDeadline:M.stamp(end+48*3600000),submissionWasLate:M.ms(j.submittedAt)>end+48*3600000,lateSubmissionAccepted:false}}));}
  else N(!f.pendingJustificationRef,'UNRESOLVED_DEPENDENCY','UNDECLARED_PENDING_REFERENCE');
  factsAfter.push(evolve(f,patch));if(f.status==='PRESENT')summary.totalPresentes++;else if(f.status==='NOT_APPLICABLE')summary.totalNaoSeAplica++;else summary.totalFaltas++;
 }
 N((intent.justificationIds||[]).every(id=>{const j=load('justification',id);return j.activityRef?.id===a.id;}),'UNRESOLVED_DEPENDENCY','JUSTIFICATION_ACTIVITY_MISMATCH');
 const historicalSummaries=events.filter(e=>e.historicalSummary).map(e=>({sourceRef:e.sourceRef,sourceHash:e.sourceHash,documentedAt:e.at,summary:e.historicalSummary}));
 const mismatched=historicalSummaries.filter(e=>Number(e.summary.totalRegistros??e.summary.totalParticipantes)!==facts.length||['totalPresentes','totalFaltas','totalNaoSeAplica'].some(k=>k in e.summary&&Number(e.summary[k])!==summary[k]));
 if(mismatched.length){
  N(intent.summaryReconciliationRef?.kind==='policy','UNRESOLVED_DEPENDENCY','SUMMARY_RECONCILIATION_REQUIRED');
  const s=load('policy',intent.summaryReconciliationRef.id);
  N(s.usage==='IMPORTED_SUMMARY_RECOGNITION'&&s.active===true&&s.approved===true&&s.activityId===a.id&&s.factResultsHash===hash(facts.map(f=>({id:f.id,status:f.status})))&&s.historicalSummariesHash===hash(historicalSummaries)&&s.resolution==='DERIVE_FROM_PRESERVED_INDIVIDUAL_FACTS'&&s.reason&&s.originalSummaryPreserved===true,'UNRESOLVED_DEPENDENCY','SUMMARY_RECONCILIATION_NOT_PROVEN');
 }
 const realizationId='rec_'+intent.operationId,realizationRef=M.ref('activityRealization',realizationId);
 const activity=evolve(a,{cycleId:cycle.id,cycleRef:M.ref('cycle',cycle.id),academicSemesterId:semester.id,personRef,modelContent,typeRef:M.ref('activityType',type.id),typeRevision:type.revision,typeSnapshot:{name:type.name,countsAttendance:type.countsAttendance,countsAbsence:effective.countsAbsence,allowsJustification:effective.allowsJustification,durationMinutes:duration,visibility:type.visibility,portalModel:effective.portalModel},policyRef:M.ref('attendancePolicy',policy.id),startsAt:M.stamp(start),endsAt:M.stamp(end),status:'OCCURRED',realizationRef,...(adjustment?{documentaryAdjustment:{approvalRef:intent.documentaryAdjustmentRef,recognizedAt:M.stamp(now),sourceRef:a.provenanceHead,originalModelCriteria:intent.typeCriteria}}:{})});
 const collection=evolve(c,{cycleId:cycle.id,cycleRef:M.ref('cycle',cycle.id),academicSemesterId:semester.id,historicalReadOnly:false,state:'CLOSED',policyRef:M.ref('attendancePolicy',policy.id),policyRevision:policy.revision,frequencyBasis,roster:Object.fromEntries(facts.map(f=>[f.id,M.ref('attendance',f.id)])),rosterCount:facts.length,rosterActivityRevision:activity.revision,unrecordedCount:0,officialSummary:summary,lastFinalizedAt:M.stamp(closure),completionRecognition:{...recognition,...closureMetadata,closureEvidenceRef:intent.closureEvidenceRef,timezone:'America/Cuiaba',reconciledAt:M.stamp(now)},justificationOpensAt:null,justificationEndsAt:M.stamp(end+48*3600000),ordinaryDeadlineBasis:'DOCUMENTED_END_AND_NC01_2025_ART8',acceptLateJustification:false});
 N(policy.justificationAfterHours===48,'UNRESOLVED_DEPENDENCY','ORDINARY_RULE_CHANGED');
 collection.summaryRecognition={kind:'DETERMINISTIC_FROM_PRESERVED_FACTS',recognizedAt:M.stamp(now),historicalSummaries,totalDiffers:mismatched.length>0,approvalRef:intent.summaryReconciliationRef||null};
 if(reopeningRecognition)collection.completionRecognition.reopeningRecognition=reopeningRecognition;
 const realization={schemaVersion:a.schemaVersion,id:realizationId,revision:1,recordMode:'NEW_V2',originRef:a.provenanceHead,operationId:intent.operationId,activityRef:M.ref('activity',a.id),occurred:true,recognized:true,occurrenceTemporalObject:{precision:'DAY',civilDate:a.civilDate},basis:{reasonCode:'IMPORTED_FACT_RECOGNIZED',sourceRef:a.provenanceHead,recognition},predecessorRef:{kind:'NONE'},recognizedAt:M.stamp(now)};
 const writes=[{ref:M.ref('activity',a.id),data:activity},{ref:M.ref('attendanceCollection',c.id),data:collection},...factsAfter.map(f=>({ref:M.ref('attendance',f.id),data:f})),...pendingWrites.map(j=>({ref:M.ref('justification',j.id),data:j})),{ref:realizationRef,data:realization}];
 writes.forEach(w=>M.validate(w.ref,w.data));
 return {schemaVersion:'act-import-reconstruction-plan-v1',operationId:intent.operationId,manifestHash:hash(intent),readSet:intent.readSet,writes:writes.map(w=>({path:M.path(w.ref),data:w.data})),factsPreservedHash:hash(facts.map(f=>({id:f.id,personRef:f.personRef,status:f.status,modality:f.modality,countsAttendance:f.countsAttendance,countsAbsence:f.countsAbsence,applicableAtDate:f.applicableAtDate,provenanceHead:f.provenanceHead}))),preservedJustificationIds:(intent.justificationIds||[]).filter(id=>!pendingWrites.some(j=>j.id===id)),historyEvent:'IMPORTED_ACTIVITY_AND_ATTENDANCE_RECOGNIZED',requiredFences:['activityAttendanceRegistry/current','presentationCapacityHeads/registry'],requiresEmptyOwnershipAndLeases:true};
}

// Two finite documentary corrections use the same sealed command, archives and
// fences. They do not promote a historical record or synthesize attendance.
function f4ImportedNonAttendancePlan_({intent,get,load,now,actorRef,a,source,hash}){
 const M=f4ActivitiesModel_(),N=M.need,variant=intent.variant,cycle=load('cycle',intent.cycleId);
 N(['HISTORICAL_RESCHEDULE','HISTORICAL_NO_CALL'].includes(variant)&&a.recordMode==='HISTORICAL_PARTIAL'&&source.literal?.ID_ATIVIDADE===a.id&&source.literal.CICLO===cycle.id&&cycle.status==='OPEN','INVALID_STATE_TRANSITION');
 if(cycle.id!=='GEAPA_2026')N(!f4ExecutionEnvironment_().production&&cycle.originRef?.kind==='SYNTHETIC_RECONCILIATION_FIXTURE','AUTHORIZATION_DENIED');
 N(intent.factIds?.length===0&&intent.justificationIds?.length===0&&!get('activityAttendanceCollections/'+a.id),'UNRESOLVED_DEPENDENCY','NO_CALL_REQUIRED');
 N(intent.documentaryAdjustmentRef?.kind==='policy','VALIDATION_ERROR');const decision=load('policy',intent.documentaryAdjustmentRef.id);
 N(decision.usage===variant&&decision.activityId===a.id&&decision.sourceHash===hash(source)&&decision.approved===true&&decision.active===true&&decision.reason&&decision.decisionEvidenceRef?.kind==='evidence'&&hash(load('evidence',decision.decisionEvidenceRef.id))===decision.decisionEvidenceHash,'UNRESOLVED_DEPENDENCY','DOCUMENTARY_DECISION_REQUIRED');
 const recognition={schemaVersion:'act-import-recognition-v1',kind:variant,operationId:intent.operationId,recognizedAt:M.stamp(now),recognizedBy:actorRef,manifestHash:hash(intent),approvalRef:intent.documentaryAdjustmentRef,sourceRef:a.provenanceHead};
 const evolve=(d,patch)=>({...d,...patch,revision:d.revision+1,operationId:intent.operationId,updatedAt:M.stamp(now)});let next,writes=[];
 if(variant==='HISTORICAL_NO_CALL'){
  N(a.legacyStatus==='REALIZADA'&&source.literal.STATUS_OPERACIONAL==='REALIZADA'&&!a.status&&!a.realizationRef,'INVALID_STATE_TRANSITION');
  next=evolve(a,{attendanceRecognition:{...recognition,state:'NO_CALL',createsFrequencyFacts:false}});
 }else{
  N(a.legacyStatus==='PLANEJADA'&&source.literal.STATUS_OPERACIONAL==='PLANEJADA'&&!a.status&&!a.realizationRef&&!a.presentationSessionRef,'INVALID_STATE_TRANSITION');
  M.civil(decision.civilDate);N(decision.civilDate>=M.day(now)&&cycle.civilStart<=decision.civilDate&&decision.civilDate<=cycle.civilEnd,'INVALID_STATE_TRANSITION','RESCHEDULE_DATE');
  const sem=load('academicSemester',decision.academicSemesterId);N(sem.cycleId===cycle.id&&sem.civilStart<=decision.civilDate&&decision.civilDate<=sem.civilEnd&&(cycle.academicSemesterRefs||[]).some(r=>r.id===sem.id),'UNRESOLVED_DEPENDENCY','SEMESTER_ASSOCIATION');
  N(/^\d{2}:\d{2}$/.test(decision.start)&&/^\d{2}:\d{2}$/.test(decision.end),'VALIDATION_ERROR');
  const start=Date.parse(decision.civilDate+'T'+decision.start+':00-04:00'),end=Date.parse(decision.civilDate+'T'+decision.end+':00-04:00');N(Number.isFinite(start)&&end>start&&end-start<=86400000,'VALIDATION_ERROR');
  N(Array.isArray(intent.calendarIds)&&intent.calendarIds.length>0&&intent.calendarIds.length<=64&&new Set(intent.calendarIds).size===intent.calendarIds.length,'UNRESOLVED_DEPENDENCY','SEALED_CALENDAR_REQUIRED');
  N(Number.isSafeInteger(intent.calendarFenceRevision)&&intent.calendarFenceRevision>=0&&(get('portalReaderResourceFences/activities')?.revision||0)===intent.calendarFenceRevision,'STALE_REVISION','CALENDAR_CHANGED');
  N(Array.isArray(intent.calendarSnapshots)&&intent.calendarSnapshots.length===intent.calendarIds.length,'UNRESOLVED_DEPENDENCY','SEALED_CALENDAR_REQUIRED');
  for(const [index,id] of intent.calendarIds.entries()){const other=get(M.path(M.ref('activity',id))),sealed=intent.calendarSnapshots[index];N(other&&sealed.id===id&&hash(f4ImportedCalendarFields_(other))===sealed.hash,'STALE_REVISION','CALENDAR_CHANGED');if(other.id===a.id||other.status==='CANCELLED'||other.legacyStatus==='CANCELADA'||other.civilDate!==decision.civilDate)continue;N(other.startsAt&&other.endsAt,'UNRESOLVED_DEPENDENCY','CALENDAR_TIME_UNKNOWN');N(M.ms(other.endsAt)<=start||M.ms(other.startsAt)>=end,'CONFLICT','CALENDAR_OVERLAP');}
  const history=[...(a.scheduleRecognition?.priorDates||[]),a.civilDate];
  next=evolve(a,{civilDate:decision.civilDate,startsAt:M.stamp(start),endsAt:M.stamp(end),academicSemesterId:sem.id,scheduleRecognition:{...recognition,priorDates:[...new Set(history)],previousStartsAt:a.startsAt||null,previousEndsAt:a.endsAt||null}});
  N(a.historicalPresentationRefs?.length>0,'UNRESOLVED_DEPENDENCY','PRESENTATION_LINEAGE');
  for(const ref of a.historicalPresentationRefs){const pr=load('presentation',ref.id);N(pr.recordMode==='HISTORICAL_PARTIAL'&&pr.activityRef?.resolution==='RESOLVED'&&pr.activityRef.ref.id===a.id,'UNRESOLVED_DEPENDENCY','PRESENTATION_LINEAGE');writes.push({path:M.path(ref),data:evolve(pr,{scheduledCivilDate:decision.civilDate,scheduleRecognition:recognition})});}
 }
 writes.unshift({path:M.path(M.ref('activity',a.id)),data:next});
 return {schemaVersion:'act-import-reconstruction-plan-v1',operationId:intent.operationId,manifestHash:hash(intent),readSet:intent.readSet,writes,factsPreservedHash:hash([]),preservedJustificationIds:[],historyEvent:variant,requiredFences:['activityAttendanceRegistry/current','presentationCapacityHeads/registry'],requiresEmptyOwnershipAndLeases:true};
}
function f4ImportedCalendarFields_(a){const M=f4ActivitiesModel_();return {id:a.id,civilDate:a.civilDate||null,status:a.status||null,legacyStatus:a.legacyStatus||null,startsAt:a.startsAt?M.ms(a.startsAt):null,endsAt:a.endsAt?M.ms(a.endsAt):null};}
