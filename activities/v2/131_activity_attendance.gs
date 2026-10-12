// Attendance is a factual ACT domain. Membership identity and discipline stay in Membros.
function f4AttendanceScope_(command) {
 return /^(ACTIVITY_|ATTENDANCE_|FREQUENCY_)/.test(command);
}

// Operational configuration can reuse approved criteria, never create normative
// authority. These are the institutional fields protected at command execution.
function f4ActivityConfigNormativeFields_(kind,d){
 if(kind==='activityType')return {countsAttendance:d.countsAttendance,countsAbsence:d.countsAbsence??d.countsAttendance,allowsJustification:d.allowsJustification,generatesCertificate:d.portalModel?.generatesCertificate===true,memberPresentation:d.portalModel?.memberPresentation===true};
 if(kind==='frequencyPeriod')return {unit:d.unit,justifiedMode:d.justifiedMode,excusedMode:d.excusedMode||'EXCLUDE'};
 if(kind==='attendancePolicy')return {previousJustificationDays:d.previousJustificationDays,previousJustificationCivilDay:d.previousJustificationCivilDay===true,justificationAfterDays:d.justificationAfterDays,justificationAfterHours:d.justificationAfterHours??null,acceptLateJustification:d.acceptLateJustification};
 f4ActivitiesModel_().need(false,'VALIDATION_ERROR');
}
function f4ActivityJustificationWindow_(policy,start,end){
 const M=f4ActivitiesModel_(),civil=f4Core_().civilStart;
 return {opensAt:(policy.previousJustificationCivilDay?civil(M.day(start)):start)-policy.previousJustificationDays*86400000,
  endsAt:policy.justificationAfterHours===undefined?civil(M.day(end))+(policy.justificationAfterDays+1)*86400000-1:end+policy.justificationAfterHours*3600000};
}

function f4AttendanceCommands_(x) {
 const {M,N,p,load,save,revision,event,now,actor}=x;
 const factRef=(activity,person)=>M.ref('attendance','a_'+M.hash({activity:activity.id,person:person.id}).slice(0,32));
 const collectionRef=activity=>M.ref('attendanceCollection',activity.id);
 const text=(value,max=2000)=>{N(typeof value==='string'&&value.trim()&&value.length<=max,'INCOMPLETE_REQUEST');return value.trim()};
 const instant=value=>{N(Number.isSafeInteger(value)&&value>0,'VALIDATION_ERROR');return value};
 function operationalCriteria(ref,old,proposed){
  let basis=old;
  if(!basis){N(p.operationalTemplateRef?.kind===ref.kind&&M.path(p.operationalTemplateRef)!==M.path(ref),'CONFIG_CAPABILITY_DISABLED','APPROVED_OPERATIONAL_TEMPLATE_REQUIRED');basis=load(p.operationalTemplateRef,true);}
  N(basis.recordMode!=='HISTORICAL_PARTIAL'&&(ref.kind==='activityType'?!!basis.decision:basis.approved===true),'CONFIG_CAPABILITY_DISABLED','APPROVED_OPERATIONAL_TEMPLATE_REQUIRED');
  N(M.hash(f4ActivityConfigNormativeFields_(ref.kind,basis))===M.hash(f4ActivityConfigNormativeFields_(ref.kind,proposed)),'AUTHORIZATION_DENIED','NORMATIVE_CHANGE_NOT_OPERATIONAL');
 }
 function decision(value){N(actor.kind==='BOARD_ACTOR','AUTHORIZATION_DENIED');N(value&&typeof value.reasonCode==='string','HUMAN_DECISION_REQUIRED');return {reasonCode:text(value.reasonCode,80),reason:text(value.reason),decidedAt:M.stamp(now),actorPersonRef:actor.personRef};}
 function fence(){const ref=M.ref('attendanceRegistry','current'),old=load(ref,false,true);save(ref,{epoch:(old?.epoch||0)+1});}
 function history(action,subjectRef,ownerPersonRef,summary){event('history',{action,subjectRef,ownerPersonRef:ownerPersonRef||'',visibility:ownerPersonRef?'OWNER':'BOARD',summary,occurredAt:M.stamp(now)});}
 function collection(activityRef,expected=true){return load(collectionRef(activityRef),expected);}
 // Drafts and roster metadata do not replace the official attendance decision.
 // Keep a pending justification attached to that unchanged fact when its document advances.
 function preserveJustificationBasis(fact){
  if(!fact?.pendingJustificationRef)return;
  const j=load(fact.pendingJustificationRef);
  if(['PREVIOUS','PENDING','CORRECTION_REQUESTED'].includes(j.state)){
   N(j.basisAttendanceRevision===fact.revision,'STALE_REVISION','JUSTIFICATION_BASIS_CHANGED');
   revision(fact.pendingJustificationRef,{...j,basisAttendanceRevision:fact.revision+1});
  }
 }
 function prefetchJustifications(refs){
  const pending=refs.map(r=>load(r,false,true)?.pendingJustificationRef).filter(Boolean);
  x.prefetch(pending.map(M.path));
  x.prefetch(pending.map(r=>M.path(r)+'/revisions/'+(load(r).revision+1)));
 }
 function sheet(finalize){
  const activity=load(p.activityRef,true),c=collection(p.activityRef);
  N(c.revision===p.collectionRevision,'STALE_REVISION','ATTENDANCE_CHANGED');
  N(actor.kind==='BOARD_ACTOR'&&activity.status!=='CANCELLED','AUTHORIZATION_DENIED');
  N(['DRAFT','OPEN','REOPENED'].includes(c.state),'INVALID_STATE_TRANSITION','CHAMADA_FINALIZADA');
  N(now>=M.ms(c.opensAt)&&now<=M.ms(c.closesAt),'DEADLINE_VIOLATION','CHAMADA_FORA_DA_JANELA');
  N(c.roster&&Object.keys(c.roster).length===c.rosterCount&&c.rosterCount>0&&c.rosterCount<=50,'UNRESOLVED_DEPENDENCY','ATTENDANCE_ROSTER');
  const membersFence=load(M.ref('capacity','registry'));N(c.rosterMembersEpoch===membersFence.membersEpoch,'STALE_REVISION','ATTENDANCE_MEMBERS_CHANGED');
  N(c.rosterActivityRevision===activity.revision,'STALE_REVISION','ATTENDANCE_ACTIVITY_CHANGED');
  N(p.records&&Object.getPrototypeOf(p.records)===Object.prototype&&Object.keys(p.records).every(id=>Object.hasOwn(c.roster,id)),'VALIDATION_ERROR','UNKNOWN_PARTICIPANT');
  const summary={totalRegistros:c.rosterCount,totalPresentes:0,totalFaltas:0,totalNaoSeAplica:0};
  x.prefetch(Object.values(c.roster).map(M.path));
  prefetchJustifications(Object.values(c.roster));
  x.prefetch(Object.values(c.roster).flatMap(r=>{const fact=load(r);return [M.path(r)+'/revisions/'+(fact.revision+1),...(finalize?[M.path(M.ref('history',M.hash({operationId:x.operationId,person:fact.personRef.id,action:'ATTENDANCE_FINALIZED'}).slice(0,32)))]:[])];}));
  for(const [id,r] of Object.entries(c.roster)){
   const fact=load(r),given=p.records[id];let entry=given||{statusPresenca:'',observacoes:''};
   if(finalize&&Object.keys(p.records).length===0){
    const saved=c.draftSavedAt&&(!c.reopenedAt||M.ms(c.draftSavedAt)>=M.ms(c.reopenedAt))?fact.draft:c.state==='REOPENED'?fact:null;
    if(saved)entry={statusPresenca:saved.status==='PRESENT'?(saved.modality==='REMOTA'?'PRESENTE_REMOTO':'PRESENTE_PRESENCIAL'):'',observacoes:saved.observations||''};
   }
   N(fact.activityRef.id===p.activityRef.id,'AUTHORIZATION_DENIED');
   N(Object.keys(entry).every(k=>['statusPresenca','observacoes'].includes(k))&&['','PRESENTE_PRESENCIAL','PRESENTE_REMOTO'].includes(entry.statusPresenca),'VALIDATION_ERROR','STATUS_PRESENCA_INVALIDO');
   N(typeof entry.observacoes==='string'&&entry.observacoes.length<=300,'VALIDATION_ERROR');
   const applicable=fact.applicableAtDate!==false;
   N(applicable||!entry.statusPresenca,'AUTHORIZATION_DENIED','PARTICIPANT_NOT_APPLICABLE');
   let status=!applicable?'NOT_APPLICABLE':entry.statusPresenca?'PRESENT':finalize?(fact.countsAbsence===false||activity.typeSnapshot.countsAbsence===false?'NOT_APPLICABLE':'ABSENT'):'UNRECORDED';
   if(status==='ABSENT'&&['JUSTIFIED','EXCUSED'].includes(fact.status))status=fact.status;
   const modality=entry.statusPresenca==='PRESENTE_REMOTO'?'REMOTA':entry.statusPresenca==='PRESENTE_PRESENCIAL'?'PRESENCIAL':'';
   const draft={status,modality,observations:entry.observacoes,recordedAt:M.stamp(now),recordedBy:actor.personRef};
   if(finalize&&fact.pendingJustificationRef){const j=load(fact.pendingJustificationRef);if(['PREVIOUS','PENDING','CORRECTION_REQUESTED'].includes(j.state)){const absent=['ABSENT','JUSTIFIED','EXCUSED'].includes(status);revision(fact.pendingJustificationRef,{...j,state:absent?(j.state==='PREVIOUS'?'PENDING':j.state):'CLOSED_WITHOUT_ABSENCE',basisAttendanceRevision:fact.revision+1,...(!absent?{closedAt:M.stamp(now)}:{})});if(!absent)fact.pendingJustificationRef='';}}
   if(!finalize)preserveJustificationBasis(fact);
   revision(r,finalize?{...fact,...draft,draft:{},officialCountBasis:{countsAttendance:fact.countsAttendance!==false,countsAbsence:fact.countsAbsence!==false,applicableAtDate:applicable},officialRecordedAt:M.stamp(now)}:{...fact,draft});
   if(status==='PRESENT')summary.totalPresentes++;else if(['ABSENT','JUSTIFIED','EXCUSED'].includes(status))summary.totalFaltas++;else if(status==='NOT_APPLICABLE')summary.totalNaoSeAplica++;
   if(finalize){const hr=M.ref('history',M.hash({operationId:x.operationId,person:fact.personRef.id,action:'ATTENDANCE_FINALIZED'}).slice(0,32));save(hr,{action:'ATTENDANCE_FINALIZED',subjectRef:r,ownerPersonRef:fact.personRef.id,visibility:'OWNER',summary:status,occurredAt:M.stamp(now)});}
  }
  revision(collectionRef(p.activityRef),{...c,state:finalize?'CLOSED':c.state==='REOPENED'?'REOPENED':'OPEN',draftSavedAt:M.stamp(now),draftSavedBy:actor.personRef,...(finalize?{unrecordedCount:0,lastFinalizedAt:M.stamp(now),lastFinalizedBy:actor.personRef,officialSummary:summary}:{})});
  if(finalize)fence();history(finalize?'ATTENDANCE_FINALIZED':'ATTENDANCE_DRAFT_SAVED',p.activityRef,null,finalize?'Chamada finalizada':'Rascunho salvo');
  return {...summary,collectionRevision:c.revision+1,statusChamada:finalize?'FINALIZADA':'SALVA',chamadaFinalizada:finalize,rascunhoSalvo:!finalize};
 }
 function eligibleMembership(personRef,membershipRef,at){
  N(personRef.kind==='person'&&membershipRef.kind==='membership','VALIDATION_ERROR');
  const person=load(personRef),link=load(membershipRef),head=x.get('memberAggregateHeads/'+personRef.id);
  N(person&&link.idPessoa===personRef.id,'AUTHORIZATION_DENIED','MEMBERSHIP_PERSON_MISMATCH');
  N(head?.paths?.includes(M.path(membershipRef))&&head.paths.includes(M.path(personRef)),'UNRESOLVED_DEPENDENCY','MEMBERSHIP_NOT_IN_AGGREGATE');
  f4MemberModel_(true).validate('memberLinks',link);
  // Membros defines exitDate as the previous active-link exit for an EGRESSO link.
  // Its canonical endDate, inclusive, is the only end of this referenced link.
  const start=link.startDate,finish=link.endDate;
  N(typeof start==='string','UNRESOLVED_DEPENDENCY','MEMBERSHIP_START_REQUIRED');
  const civil=M.day(at);N(M.gap(start,civil)>=0&&(!finish||M.gap(civil,finish)>=0),'AUTHORIZATION_DENIED','MEMBERSHIP_OUTSIDE_ACTIVITY');
  return link;
 }
 function justificationContext(){
   let fact=load(p.attendanceRef,true,true);const newFact=!fact;
   if(!fact){
    N(p.activityRef&&actor.kind==='PORTAL_AUTHENTICATED_MEMBER','UNRESOLVED_DEPENDENCY');const a=load(p.activityRef,true),c=collection(p.activityRef);
    N(a.publicationStatus==='PUBLICADO'&&['PUBLIC','MEMBERS'].includes(a.publishedVisibility)&&a.typeSnapshot.allowsJustification&&now<M.ms(a.startsAt),'AUTHORIZATION_DENIED');
    N(p.attendanceRef.id===factRef(p.activityRef,actor.personRef).id,'AUTHORIZATION_DENIED');
    const head=x.get('memberAggregateHeads/'+actor.personId),paths=head.paths.filter(path=>/^member(Identities|Links|Events)\//.test(path));x.prefetch(paths);
    const docs=Object.fromEntries(paths.map(path=>[path,x.get(path)])),member=f4MemberModel_(true),at=M.ms(a.startsAt),state=member.derive(docs,actor.personId,head.revision,at).state;
    N(['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE'].includes(state.state),'AUTHORIZATION_DENIED');
    const links=member.rows(docs,'memberLinks').filter(l=>{const r=member.interval(l.startDate,l.endDate);return r.a<=at&&at<r.z});N(links.length===1,'UNRESOLVED_DEPENDENCY');
    fact={revision:0,activityRef:p.activityRef,personRef:actor.personRef,ownerPersonRef:actor.personId,membershipRef:M.ref('membership',links[0].id),periodId:c.periodRef.id,status:'UNRECORDED',durationMinutes:a.typeSnapshot.durationMinutes,frequencyBasis:c.frequencyBasis,applicableAtDate:true,participantType:'MEMBRO',countsAttendance:a.typeSnapshot.countsAttendance,countsAbsence:a.typeSnapshot.countsAbsence!==false};
   }
   const activity=load(fact.activityRef),c=collection(fact.activityRef);N(fact.personRef.id===actor.personId&&activity.typeSnapshot.allowsJustification&&fact.applicableAtDate!==false,'AUTHORIZATION_DENIED');
   if(p.activityRef)N(p.activityRef.id===fact.activityRef.id,'AUTHORIZATION_DENIED');
   N(!['CANCELLED','NOT_OCCURRED','ARCHIVED'].includes(activity.status),'INVALID_STATE_TRANSITION');
   const previous=now<M.ms(activity.startsAt),late=now>M.ms(c.justificationEndsAt),old=fact.pendingJustificationRef&&load(fact.pendingJustificationRef);
   N(!old||old.state==='CORRECTION_REQUESTED'||old.state==='PREVIOUS'&&previous,'CONFLICT','JUSTIFICATION_ALREADY_PENDING');
   if(previous)N(c.justificationOpensAt&&now>=M.ms(c.justificationOpensAt),'DEADLINE_VIOLATION','PREVIOUS_JUSTIFICATION_WINDOW');
   else N(['ABSENT','UNRECORDED'].includes(fact.status),'INVALID_STATE_TRANSITION','NO_ABSENCE');
   N(!late||c.acceptLateJustification&&p.lateAcknowledged===true,'DEADLINE_VIOLATION','LATE_JUSTIFICATION_ACKNOWLEDGEMENT');
   if(p.proofLink)N(typeof p.proofLink==='string'&&/^https:\/\/drive\.google\.com\//.test(p.proofLink)&&p.proofLink.length<=2000,'VALIDATION_ERROR','PROOF_LINK');
   return {fact,newFact,activity,c,previous,late,old};
 }
 function setAttendance(ref,status,reason,correction=false,justificationDecision=false){
  const fact=load(ref,true),c=collection(fact.activityRef),activity=load(fact.activityRef,true);
   N(['PRESENT','ABSENT','JUSTIFIED','EXCUSED','NOT_APPLICABLE'].includes(status),'VALIDATION_ERROR');
   N(fact.applicableAtDate!==false||status==='NOT_APPLICABLE','AUTHORIZATION_DENIED','PARTICIPANT_NOT_APPLICABLE');
   N(status!=='ABSENT'||fact.countsAbsence!==false&&activity.typeSnapshot.countsAbsence!==false,'VALIDATION_ERROR','ACTIVITY_DOES_NOT_COUNT_ABSENCE');
  N(activity.status!=='CANCELLED','INVALID_STATE_TRANSITION');
  N(correction?c.state==='CLOSED':c.state==='OPEN','INVALID_STATE_TRANSITION','COLLECTION_STATE');
  N(correction||justificationDecision||now>=M.ms(c.opensAt)&&now<=M.ms(c.closesAt),'DEADLINE_VIOLATION','COLLECTION_WINDOW');
  N(!['JUSTIFIED','EXCUSED'].includes(status)||activity.typeSnapshot.allowsJustification,'CONFIG_CAPABILITY_DISABLED');
  text(reason);const next=revision(ref,{...fact,status,reason,recordedBy:actor.personRef,recordedAt:M.stamp(now)});
  if(fact.status==='UNRECORDED')revision(collectionRef(fact.activityRef),{...c,unrecordedCount:c.unrecordedCount-1});
  fence();history(correction?'ATTENDANCE_CORRECTED':'ATTENDANCE_RECORDED',ref,fact.personRef.id,status);return {attendanceRef:ref,status:next.status};
 }
 return {
  ACTIVITY_POLICY_SAVE(){
   const old=load(p.attendancePolicyRef,true,true);operationalCriteria(p.attendancePolicyRef,old,p);const period=load(p.periodRef,true);N(period.approved,'CONFIG_CAPABILITY_DISABLED');
   if(p.justificationAfterHours!==undefined)N(Number.isInteger(p.justificationAfterHours)&&p.justificationAfterHours>=0&&p.justificationAfterHours<=8760,'VALIDATION_ERROR');
   if(p.previousJustificationCivilDay!==undefined)N(typeof p.previousJustificationCivilDay==='boolean','VALIDATION_ERROR');
   for(const key of ['opensBeforeMinutes','closesAfterMinutes'])N(Number.isInteger(p[key])&&p[key]>=0&&p[key]<=10080,'VALIDATION_ERROR');
   for(const key of ['previousJustificationDays','justificationAfterDays'])N(Number.isInteger(p[key])&&p[key]>=0&&p[key]<=365,'VALIDATION_ERROR');
   if(p.proofUploadPolicyRef){N(p.proofUploadPolicyRef.kind==='attachmentPolicy','VALIDATION_ERROR');const proof=load(p.proofUploadPolicyRef,true);N(proof.approved&&proof.active&&proof.usage==='ATTENDANCE_PROOF'&&proof.workspace?.enabled,'CONFIG_CAPABILITY_DISABLED');}
   N(typeof p.acceptLateJustification==='boolean','VALIDATION_ERROR');revision(p.attendancePolicyRef,{...(p.justificationAfterHours!==undefined?{justificationAfterHours:p.justificationAfterHours}:{}),...(p.previousJustificationCivilDay!==undefined?{previousJustificationCivilDay:p.previousJustificationCivilDay}:{}),...(p.proofUploadPolicyRef?{proofUploadPolicyRef:p.proofUploadPolicyRef}:{}),periodRef:p.periodRef,opensBeforeMinutes:p.opensBeforeMinutes,closesAfterMinutes:p.closesAfterMinutes,previousJustificationDays:p.previousJustificationDays,justificationAfterDays:p.justificationAfterDays,acceptLateJustification:p.acceptLateJustification,approved:true,decision:decision(p.decision)});return {attendancePolicyRef:p.attendancePolicyRef};
  },
  ACTIVITY_TYPE_SAVE(){
   N(p.typeRef.kind==='activityType','VALIDATION_ERROR');const old=load(p.typeRef,true,true);
   N(typeof p.countsAttendance==='boolean'&&typeof p.allowsJustification==='boolean','VALIDATION_ERROR');
   N(p.durationMinutes===null||Number.isSafeInteger(p.durationMinutes)&&p.durationMinutes>0&&p.durationMinutes<=1440,'VALIDATION_ERROR');
   N(['MEMBERS','BOARD','PUBLIC'].includes(p.visibility)&&['ACTIVE','INACTIVE'].includes(p.status),'VALIDATION_ERROR');
   if(p.countsAbsence!==undefined)N(typeof p.countsAbsence==='boolean','VALIDATION_ERROR');
   const portalModel=p.portalModel?f4ActivityPortalModel_(p.portalModel):old?.portalModel,attendancePolicyRef=p.attendancePolicyRef||old?.attendancePolicyRef;
   operationalCriteria(p.typeRef,old,{...p,portalModel});
   if(attendancePolicyRef)N(load(attendancePolicyRef,true).approved,'CONFIG_CAPABILITY_DISABLED');
   revision(p.typeRef,{name:text(p.name,120),countsAttendance:p.countsAttendance,...(p.countsAbsence!==undefined?{countsAbsence:p.countsAbsence}:{}),...(attendancePolicyRef?{attendancePolicyRef}:{}),...(portalModel?{portalModel}:{}),durationMinutes:p.durationMinutes,visibility:p.visibility,allowsJustification:p.allowsJustification,status:p.status,decision:decision(p.decision)});
   history(old?'ACTIVITY_TYPE_UPDATED':'ACTIVITY_TYPE_CREATED',p.typeRef,null,p.name);return {typeRef:p.typeRef};
  },
  ACTIVITY_SAVE(){
   N(p.activityRef.kind==='activity'&&p.typeRef.kind==='activityType','VALIDATION_ERROR');const old=load(p.activityRef,true,true),type=load(p.typeRef,true);
   N(type.status==='ACTIVE','CONFIG_CAPABILITY_DISABLED');const start=instant(p.startsAt),end=instant(p.endsAt);
   N(!old||old.typeRef.id!==p.typeRef.id||Number.isSafeInteger(old.typeRevision)&&old.typeRevision>0&&old.typeSnapshot,'UNRESOLVED_DEPENDENCY','ACTIVITY_MODEL_SNAPSHOT_REQUIRED');
   N(!old?.presentationSessionRef||p.status===old.status,'INVALID_STATE_TRANSITION','USE_ACTIVITY_CANCELLATION');
   N(end>start&&end-start<=86400000,'VALIDATION_ERROR');N(['SCHEDULED','CANCELLED'].includes(p.status),'VALIDATION_ERROR');
   const c=load(collectionRef(p.activityRef),false,true),structural=old&&(old.typeRef.id!==p.typeRef.id||M.ms(old.startsAt)!==start||M.ms(old.endsAt)!==end);
   N(!old?.realizationRef&&(!structural||!c||c.state==='DRAFT'&&c.rosterCount===0&&!c.preRosterFactCount),'INVALID_STATE_TRANSITION','ACTIVITY_ALREADY_COLLECTING');
   if(old?.presentationSessionRef&&(M.ms(old.startsAt)!==start||M.ms(old.endsAt)!==end))f4ActivitiesApplyActivityCalendar_(x,old);
   if(structural&&c){
    N(old.typeRef.id===p.typeRef.id&&c.policyRef,'UNRESOLVED_DEPENDENCY','ATTENDANCE_RESCHEDULE_POLICY_REQUIRED');
    const head=load(c.policyRef),policy=head.revision===c.policyRevision?head:x.get(M.path(c.policyRef)+'/revisions/'+c.policyRevision)?.snapshot,period=load(c.periodRef);
    N(policy?.approved&&policy.revision===c.policyRevision&&period.approved&&M.gap(period.civilStart,M.day(start))>=0&&M.gap(M.day(start),period.civilEnd)>=0,'CONFIG_CAPABILITY_DISABLED','ATTENDANCE_RESCHEDULE_POLICY_REQUIRED');
    revision(collectionRef(p.activityRef),{...c,opensAt:M.stamp(start-policy.opensBeforeMinutes*60000),closesAt:M.stamp(end+policy.closesAfterMinutes*60000),justificationOpensAt:M.stamp(f4ActivityJustificationWindow_(policy,start,end).opensAt),justificationEndsAt:M.stamp(f4ActivityJustificationWindow_(policy,start,end).endsAt)});
   }
   const details=p.details?f4ActivityPortalDetails_(p.details):old?.details||{},appliedModel=old&&old.typeRef.id===p.typeRef.id?old.typeSnapshot.portalModel:type.portalModel,modelContent=p.modelContent===undefined&&old?.modelContent?old.modelContent:appliedModel?f4ActivityContent_(appliedModel,p.modelContent||old?.modelContent,x.get,start,x.memberPresentationCreation===true):null;
   const guestParticipants=p.guestParticipants===undefined?old?.guestParticipants||[]:p.guestParticipants;
   N(Array.isArray(guestParticipants)&&guestParticipants.length<=25,'SCOPE_TOO_LARGE','GUEST_ROSTER_LIMIT');
   N(!c||['DRAFT','OPEN','REOPENED'].includes(c.state)||M.hash(guestParticipants)===M.hash(old?.guestParticipants||[]),'INVALID_STATE_TRANSITION','FINALIZED_GUEST_ROSTER');
   const guestIds=new Set();
   for(const guest of guestParticipants){f4Core_().exact(guest,['personRef','participantType','institutionLabel']);N(guest.personRef?.kind==='person'&&!guestIds.has(guest.personRef.id)&&['PROFESSOR','CONVIDADO','EXTERNO','EGRESSO'].includes(guest.participantType),'VALIDATION_ERROR','GUEST_REFERENCE');guestIds.add(guest.personRef.id);N(typeof guest.institutionLabel==='string'&&guest.institutionLabel.length<=180&&!/[<>\x00-\x08]/.test(guest.institutionLabel),'VALIDATION_ERROR');const identity=load(guest.personRef),head=x.get('memberAggregateHeads/'+guest.personRef.id);N(identity?.idPessoa===guest.personRef.id&&!identity.technicalAccount&&head?.paths?.includes(M.path(guest.personRef)),'UNRESOLVED_DEPENDENCY','GUEST_CORE_IDENTITY_REQUIRED');}
   revision(p.activityRef,{...old,guestParticipants,...(modelContent?{modelContent}:{}),title:text(p.title,240),location:text(p.location,240),startsAt:M.stamp(start),endsAt:M.stamp(end),civilDate:M.day(start),typeRef:p.typeRef,typeRevision:old&&old.typeRef.id===p.typeRef.id?old.typeRevision:type.revision,typeSnapshot:old&&old.typeRef.id===p.typeRef.id?old.typeSnapshot:{name:type.name,countsAttendance:type.countsAttendance,...(type.countsAbsence!==undefined?{countsAbsence:type.countsAbsence}:{}),durationMinutes:type.durationMinutes??Math.round((end-start)/60000),visibility:type.visibility,allowsJustification:type.allowsJustification,...(type.portalModel?{portalModel:type.portalModel}:{})},details,format:details.formato||old?.format||'',publicationStatus:old?.publicationStatus||'RASCUNHO',publishedVisibility:old?.publishedVisibility||'BOARD',status:p.status,decision:decision(p.decision)});
   if(!old&&type.portalModel&&type.countsAttendance){
    N(type.attendancePolicyRef,'CONFIG_CAPABILITY_DISABLED','ATTENDANCE_POLICY_REQUIRED');const policy=load(type.attendancePolicyRef),period=load(policy.periodRef);N(policy.approved&&period.approved&&M.gap(period.civilStart,M.day(start))>=0&&M.gap(M.day(start),period.civilEnd)>=0,'CONFIG_CAPABILITY_DISABLED','FREQUENCY_PERIOD_REQUIRED');
    const opensAt=start-policy.opensBeforeMinutes*60000,closesAt=end+policy.closesAfterMinutes*60000,justificationEndsAt=f4ActivityJustificationWindow_(policy,start,end).endsAt;
    N(closesAt>opensAt&&justificationEndsAt>=end,'VALIDATION_ERROR');
    revision(collectionRef(p.activityRef),{...(policy.proofUploadPolicyRef?{proofUploadPolicyRef:policy.proofUploadPolicyRef}:{}),activityRef:p.activityRef,periodRef:policy.periodRef,state:'DRAFT',opensAt:M.stamp(opensAt),closesAt:M.stamp(closesAt),justificationOpensAt:M.stamp(f4ActivityJustificationWindow_(policy,start,end).opensAt),justificationEndsAt:M.stamp(justificationEndsAt),acceptLateJustification:policy.acceptLateJustification,rosterCount:0,unrecordedCount:0,frequencyBasis:{periodId:policy.periodRef.id,policyRevision:period.revision,unit:period.unit,justifiedMode:period.justifiedMode,...(period.excusedMode?{excusedMode:period.excusedMode}:{})},policyRef:type.attendancePolicyRef,policyRevision:policy.revision});save(policy.periodRef,{...period,usedByCount:period.usedByCount+1});
   }
   fence();history(p.status==='CANCELLED'?'ACTIVITY_CANCELLED':'ACTIVITY_SAVED',p.activityRef,null,p.title);return {activityRef:p.activityRef,atividade:{idAtividade:p.activityRef.id,tituloPublico:p.title,dataAtividade:M.day(start),statusOperacional:p.status==='CANCELLED'?'CANCELADA':'PLANEJADA',statusPublicacaoPortal:old?.publicationStatus==='PUBLICADO'?'PUBLICADA':old?.publicationStatus||'RASCUNHO',visibilidadePortal:{PUBLIC:'PUBLICA',MEMBERS:'MEMBROS',BOARD:'DIRETORIA',HIDDEN:'OCULTA'}[old?.publishedVisibility||'BOARD']}};
  },
  ACTIVITY_PUBLICATION_SET(){
   const activity=load(p.activityRef,true);N(['PUBLISH','HIDE','CANCEL','REOPEN'].includes(p.action),'VALIDATION_ERROR');
   N(!activity.realizationRef||p.action==='HIDE','INVALID_STATE_TRANSITION','ACTIVITY_ALREADY_REALIZED');
   if(p.action==='PUBLISH')N(activity.status!=='CANCELLED'&&activity.title&&activity.location&&activity.format&&activity.typeSnapshot.portalModel,'INCOMPLETE_REQUEST','ACTIVITY_NOT_READY');
   if(p.action==='REOPEN')N(activity.status==='CANCELLED','INVALID_STATE_TRANSITION');
   if(p.visibility!==undefined)N(['BOARD','MEMBERS','PUBLIC'].includes(p.visibility)&&p.action==='PUBLISH','VALIDATION_ERROR');
   f4ActivitiesApplyActivityCancellation_(x,activity);
   revision(p.activityRef,{...activity,publishedVisibility:p.action==='PUBLISH'?(p.visibility||activity.publishedVisibility||'BOARD'):p.action==='REOPEN'?'BOARD':'HIDDEN',status:p.action==='CANCEL'?'CANCELLED':p.action==='REOPEN'?'SCHEDULED':activity.status,publicationStatus:p.action==='PUBLISH'?'PUBLICADO':p.action==='REOPEN'?'RASCUNHO':'OCULTA'});fence();history('ACTIVITY_'+p.action,p.activityRef,null,activity.title);return {activityRef:p.activityRef};
  },
  ACTIVITY_REALIZE(){
   const activity=load(p.activityRef,true);N(typeof p.occurred==='boolean','VALIDATION_ERROR');N(activity.status!=='CANCELLED','INVALID_STATE_TRANSITION');
   N(activity.realizationRef?M.path(activity.realizationRef)===M.path(p.predecessorRef):p.predecessorRef.kind==='NONE','STALE_REVISION');
   const temporal=p.occurrenceTemporalObject;N(temporal&&['DAY','INSTANT'].includes(temporal.precision),'VALIDATION_ERROR');M.civil(temporal.civilDate);N(M.gap(temporal.civilDate,M.day(now))>=0,'VALIDATION_ERROR','FUTURE_OCCURRENCE');
   if(temporal.precision==='INSTANT')N(instant(temporal.instant)<=now&&M.day(temporal.instant)===temporal.civilDate,'VALIDATION_ERROR');
   const basis=decision(p.basis),ref=event('activityRealization',{activityRef:p.activityRef,occurred:p.occurred,recognized:true,occurrenceTemporalObject:temporal,basis,predecessorRef:p.predecessorRef,recognizedAt:M.stamp(now)});
   revision(p.activityRef,{...activity,status:p.occurred?'OCCURRED':'NOT_OCCURRED',realizationRef:ref});fence();history('ACTIVITY_RECOGNIZED',p.activityRef,null,p.occurred?'Ocorreu':'Não ocorreu');return {realizationRef:ref};
  },
  FREQUENCY_PERIOD_SAVE(){
   N(p.periodRef.kind==='frequencyPeriod','VALIDATION_ERROR');const old=load(p.periodRef,true,true);N(!old?.usedByCount,'INVALID_STATE_TRANSITION','PERIOD_ALREADY_USED');
   operationalCriteria(p.periodRef,old,p);
   N(M.gap(p.civilStart,p.civilEnd)>=0&&['ACTIVITY','MINUTE'].includes(p.unit)&&['EXCLUDE','COUNT_PRESENT','COUNT_ABSENT'].includes(p.justifiedMode),'VALIDATION_ERROR');
   if(p.excusedMode!==undefined)N(['EXCLUDE','COUNT_PRESENT','COUNT_ABSENT'].includes(p.excusedMode),'VALIDATION_ERROR');
   revision(p.periodRef,{label:text(p.label,120),civilStart:p.civilStart,civilEnd:p.civilEnd,timezone:'America/Cuiaba',unit:p.unit,justifiedMode:p.justifiedMode,...(p.excusedMode?{excusedMode:p.excusedMode}:{}),approved:true,usedByCount:0,decision:decision(p.decision)});history('FREQUENCY_PERIOD_CONFIGURED',p.periodRef,null,p.label);return {periodRef:p.periodRef};
  },
  ATTENDANCE_CONFIGURE(){
   const activity=load(p.activityRef,true),period=load(p.periodRef,true),ref=collectionRef(p.activityRef),old=load(ref,true,true);
   N(activity.typeSnapshot.countsAttendance&&period.approved,'CONFIG_CAPABILITY_DISABLED');N(!old||old.state==='DRAFT'&&old.rosterCount===0&&!old.preRosterFactCount,'INVALID_STATE_TRANSITION');
   N(M.gap(period.civilStart,activity.civilDate)>=0&&M.gap(activity.civilDate,period.civilEnd)>=0,'VALIDATION_ERROR','PERIOD_OUTSIDE_ACTIVITY');
   N(instant(p.closesAt)>instant(p.opensAt)&&instant(p.justificationEndsAt)>=p.closesAt,'VALIDATION_ERROR');
   N(!old||old.periodRef.id===p.periodRef.id,'INVALID_STATE_TRANSITION','PERIOD_CHANGE_REQUIRES_NEW_COLLECTION');
   if(p.justificationOpensAt!==undefined)N(instant(p.justificationOpensAt)<=M.ms(activity.startsAt),'VALIDATION_ERROR');if(p.acceptLateJustification!==undefined)N(typeof p.acceptLateJustification==='boolean','VALIDATION_ERROR');
   revision(ref,{activityRef:p.activityRef,periodRef:p.periodRef,...(p.justificationOpensAt!==undefined?{justificationOpensAt:M.stamp(p.justificationOpensAt)}:{}),acceptLateJustification:p.acceptLateJustification===true,state:'DRAFT',opensAt:M.stamp(p.opensAt),closesAt:M.stamp(p.closesAt),justificationEndsAt:M.stamp(p.justificationEndsAt),rosterCount:0,unrecordedCount:0,frequencyBasis:{periodId:p.periodRef.id,policyRevision:period.revision,unit:period.unit,justifiedMode:period.justifiedMode,...(period.excusedMode?{excusedMode:period.excusedMode}:{})},decision:decision(p.decision)});
   if(!old)save(p.periodRef,{...period,usedByCount:period.usedByCount+1});history('ATTENDANCE_CONFIGURED',p.activityRef,null,period.label);return {collectionRef:ref};
  },
  ATTENDANCE_ENROLL(){
   const activity=load(p.activityRef,true),c=collection(p.activityRef);N(c.state==='DRAFT'&&activity.status!=='CANCELLED','INVALID_STATE_TRANSITION');eligibleMembership(p.personRef,p.membershipRef,M.ms(activity.startsAt));
   const ref=factRef(p.activityRef,p.personRef);N(!load(ref,true,true),'CONFLICT','ATTENDANCE_ALREADY_ENROLLED');
   revision(ref,{activityRef:p.activityRef,personRef:p.personRef,ownerPersonRef:p.personRef.id,membershipRef:p.membershipRef,periodId:c.periodRef.id,status:'UNRECORDED',durationMinutes:activity.typeSnapshot.durationMinutes,frequencyBasis:c.frequencyBasis,enrollmentBasis:decision(p.decision)});
   const membersFence=load(M.ref('capacity','registry'));N(Number.isSafeInteger(membersFence.membersEpoch)&&(!c.rosterCount||c.rosterMembersEpoch===membersFence.membersEpoch),'STALE_REVISION','ATTENDANCE_MEMBERS_CHANGED');
   N(c.rosterCount<50,'SCOPE_TOO_LARGE','ATTENDANCE_ROSTER_LIMIT');revision(collectionRef(p.activityRef),{...c,rosterActivityRevision:activity.revision,rosterMembersEpoch:membersFence.membersEpoch,roster:{...(c.roster||{}),[ref.id]:ref},rosterCount:c.rosterCount+1,unrecordedCount:c.unrecordedCount+1});fence();history('ATTENDANCE_ENROLLED',ref,p.personRef.id,activity.title);return {attendanceRef:ref};
  },
  ATTENDANCE_ROSTER_APPLY(){
   const activity=load(p.activityRef,true),c=collection(p.activityRef),snapshot=load(p.rosterRef,true),members=load(M.ref('capacity','registry'));
   N(actor.kind==='BOARD_ACTOR'&&['DRAFT','OPEN','REOPENED'].includes(c.state)&&!['CANCELLED','NOT_OCCURRED','ARCHIVED'].includes(activity.status),'INVALID_STATE_TRANSITION');
   N(snapshot.state==='SEALED'&&snapshot.complete&&snapshot.activityRef.id===activity.id&&snapshot.activityRevision===activity.revision&&snapshot.membersEpoch===members.membersEpoch,'STALE_REVISION','ROSTER_SOURCE_CHANGED');
   const allIds=[...new Set([...Object.keys(snapshot.records),...Object.keys(c.roster||{})])];N(allIds.length>0&&allIds.length<=50,'SCOPE_TOO_LARGE','ATTENDANCE_ROSTER_LIMIT');
   const refs=allIds.map(id=>M.ref('attendance',id));x.prefetch(refs.map(M.path));x.prefetch(refs.map(r=>M.path(r)+'/revisions/'+((load(r,false,true)?.revision||0)+1)));
   prefetchJustifications(refs);
   const roster={};let unrecorded=0;
   for(const r of refs){const old=load(r,false,true),source=snapshot.records[r.id];N(!old||old.activityRef.id===activity.id,'AUTHORIZATION_DENIED');
    const applicable=source||{...old,applicableAtDate:false,countsAttendance:false,countsAbsence:false,notApplicableReason:'VINCULO_NAO_ATIVO_NA_DATA'};
    N(applicable.personRef&&(applicable.participantType!=='MEMBRO'||applicable.membershipRef),'UNRESOLVED_DEPENDENCY');const value={...old,...applicable,activityRef:p.activityRef,ownerPersonRef:applicable.personRef.id,periodId:c.periodRef.id,frequencyBasis:c.frequencyBasis,durationMinutes:activity.typeSnapshot.durationMinutes,status:old?.status||'UNRECORDED',rosterSourceRef:p.rosterRef};
    const changed=!old||Object.entries(value).some(([key,v])=>key!=='rosterSourceRef'&&f4Core_().hash(f4Firestore_().encode(v))!==f4Core_().hash(f4Firestore_().encode(old[key]===undefined?'__ABSENT__':old[key])));
    if(changed){preserveJustificationBasis(old);revision(r,value);}roster[r.id]=r;if(value.status==='UNRECORDED')unrecorded++;
   }
   revision(collectionRef(p.activityRef),{...c,roster,rosterCount:refs.length,unrecordedCount:unrecorded,rosterActivityRevision:activity.revision,rosterMembersEpoch:members.membersEpoch,rosterSourceRef:p.rosterRef});fence();history('ATTENDANCE_ROSTER_UPDATED',p.activityRef,null,'Lista de chamada atualizada');return {rosterCount:refs.length};
  },
  ATTENDANCE_SAVE_DRAFT(){return sheet(false);},
  ATTENDANCE_FINALIZE(){return sheet(true);},
  ATTENDANCE_REOPEN(){const c=collection(p.activityRef);N(actor.kind==='BOARD_ACTOR'&&c.revision===p.collectionRevision,'STALE_REVISION');N(c.state==='CLOSED','INVALID_STATE_TRANSITION');revision(collectionRef(p.activityRef),{...c,state:'REOPENED',reopenedAt:M.stamp(now),reopenedBy:actor.personRef});history('ATTENDANCE_REOPENED',p.activityRef,null,'Chamada reaberta');return {collectionRevision:c.revision+1,statusChamada:'REABERTA',chamadaFinalizada:false};},
  ATTENDANCE_OPEN(){const c=collection(p.activityRef),activity=load(p.activityRef,true);N(c.state==='DRAFT'&&c.rosterCount>0&&activity.status!=='CANCELLED','INVALID_STATE_TRANSITION');N(now<=M.ms(c.closesAt),'DEADLINE_VIOLATION');revision(collectionRef(p.activityRef),{...c,state:'OPEN'});return {state:'OPEN'};},
  ATTENDANCE_CLOSE(){const c=collection(p.activityRef);N(c.state==='OPEN'&&c.unrecordedCount===0,'INCOMPLETE_REQUEST','UNRECORDED_ATTENDANCE');revision(collectionRef(p.activityRef),{...c,state:'CLOSED',closedAt:M.stamp(now),closingDecision:decision(p.decision)});fence();history('ATTENDANCE_CLOSED',p.activityRef,null,'Coleta encerrada');return {state:'CLOSED'};},
  ATTENDANCE_RECORD(){return setAttendance(p.attendanceRef,p.status,p.reason);},
  ATTENDANCE_CORRECT(){return setAttendance(p.attendanceRef,p.status,p.reason,true);},
  ATTENDANCE_PROOF_UPLOAD_REQUEST(){return f4AttendanceProof_(x).begin(justificationContext());},
  ATTENDANCE_JUSTIFICATION_REQUEST(){
   const {fact,newFact,activity,c,previous,late,old}=justificationContext();
   if(p.clearProof!==undefined)N(typeof p.clearProof==='boolean'&&(!p.clearProof||!p.proofUploadRef&&!p.proofLink),'VALIDATION_ERROR');
   const value={attendanceRef:p.attendanceRef,activityRef:fact.activityRef,personRef:fact.personRef,ownerPersonRef:fact.personRef.id,state:previous?'PREVIOUS':'PENDING',basisAttendanceRevision:fact.revision+1,reason:text(p.reason),reasonCode:p.reasonCode?text(p.reasonCode,120):'',proofLink:p.proofLink||'',submittedAt:M.stamp(now),late,lateAcknowledged:late&&p.lateAcknowledged===true};
   const ref=old?fact.pendingJustificationRef:event('justification',value);if(old){const next={...old,...value};if(p.clearProof||p.proofLink){delete next.proofArtifactRef;delete next.proofUploadRef;}revision(ref,next);}if(p.proofUploadRef)revision(ref,{...load(ref),...f4AttendanceProof_(x).attach(ref,fact)});revision(p.attendanceRef,{...fact,pendingJustificationRef:ref,lastJustificationRef:ref});if(newFact)revision(collectionRef(fact.activityRef),{...c,preRosterFactCount:(c.preRosterFactCount||0)+1});fence();history('JUSTIFICATION_REQUESTED',ref,fact.personRef.id,'Justificativa enviada');return {justificationRef:ref,state:value.state};
  },
  ATTENDANCE_JUSTIFICATION_DECIDE(){
   const j=load(p.justificationRef,true),fact=load(j.attendanceRef,true),c=collection(fact.activityRef),activity=load(fact.activityRef,true);
   N(j.state==='PENDING'&&['APPROVE','EXCUSE','REJECT','REQUEST_CORRECTION'].includes(p.decision)&&fact.pendingJustificationRef?.id===p.justificationRef.id,'INVALID_STATE_TRANSITION');
   N(fact.revision===j.basisAttendanceRevision&&['ABSENT','UNRECORDED'].includes(fact.status),'STALE_REVISION','ATTENDANCE_CHANGED_AFTER_REQUEST');
   N(activity.typeSnapshot.allowsJustification&&!['CANCELLED','NOT_OCCURRED','ARCHIVED'].includes(activity.status),'INVALID_STATE_TRANSITION');text(p.reason);if(p.internalReason!==undefined)N(typeof p.internalReason==='string'&&p.internalReason.length<=2000,'VALIDATION_ERROR');
   const state={APPROVE:'APPROVED',EXCUSE:'EXCUSED',REJECT:'REJECTED',REQUEST_CORRECTION:'CORRECTION_REQUESTED'}[p.decision];
   N(p.decision!=='EXCUSE'||['EXCLUDE','COUNT_PRESENT','COUNT_ABSENT'].includes(fact.frequencyBasis.excusedMode),'CONFIG_CAPABILITY_DISABLED','EXCUSED_FREQUENCY_POLICY_REQUIRED');
   revision(p.justificationRef,{...j,state,decisionReason:p.reason,...(p.internalReason!==undefined?{internalReason:p.internalReason}:{}),decidedAt:M.stamp(now),decidedBy:actor.personRef,exceptional:j.late===true&&['APPROVE','EXCUSE'].includes(p.decision),basisAttendanceRevision:fact.revision+1});
   const keep=p.decision==='REQUEST_CORRECTION',next={...fact,pendingJustificationRef:keep?p.justificationRef:'',lastJustificationRef:p.justificationRef,...(p.decision==='APPROVE'||p.decision==='EXCUSE'?{status:p.decision==='APPROVE'?'JUSTIFIED':'EXCUSED',justificationDecisionRef:p.justificationRef,recordedBy:actor.personRef,recordedAt:M.stamp(now)}:{})};
   revision(j.attendanceRef,next);fence();history('JUSTIFICATION_'+state,p.justificationRef,fact.personRef.id,p.reason);return {state};
  }
 };
}

// One explicit calculation; no discipline, rounding thresholds, or inferred absences.
function f4AttendanceFrequency_(rows,complete) {
 const N=f4ActivitiesModel_().need;N(Array.isArray(rows)&&typeof complete==='boolean','VALIDATION_ERROR');
 let numerator=0,denominator=0,pending=0,basisComplete=true;const counts={PRESENT:0,ABSENT:0,JUSTIFIED:0,EXCUSED:0,NOT_APPLICABLE:0};
 const partial=rows.filter(r=>r.recordMode==='HISTORICAL_PARTIAL'&&r.historicalAttendanceFact===true&&r.frequencyBasis===null);
 for(const row of partial){N(typeof row.countable==='boolean'&&Object.hasOwn(counts,row.status),'VALIDATION_ERROR');if(!row.countable){pending++;continue;}if(row.applicableAtDate===false&&row.status!=='NOT_APPLICABLE'||row.status==='PRESENT'&&row.countsAttendance===false||['ABSENT','JUSTIFIED','EXCUSED'].includes(row.status)&&row.countsAbsence===false)continue;counts[row.status]++;if(row.status!=='NOT_APPLICABLE')basisComplete=false;}
 rows=rows.filter(r=>!partial.includes(r));
 for(const row of rows){N(typeof row.countable==='boolean'&&['PRESENT','ABSENT','JUSTIFIED','EXCUSED','UNRECORDED','NOT_APPLICABLE'].includes(row.status)&&['MINUTE','ACTIVITY'].includes(row.frequencyBasis?.unit)&&['EXCLUDE','COUNT_PRESENT','COUNT_ABSENT'].includes(row.frequencyBasis?.justifiedMode),'VALIDATION_ERROR');if(!row.countable||row.status==='UNRECORDED'){pending++;continue}const countBasis=row.officialCountBasis||row;if(countBasis.applicableAtDate===false&&row.status!=='NOT_APPLICABLE'||row.status==='PRESENT'&&countBasis.countsAttendance===false||['ABSENT','JUSTIFIED','EXCUSED'].includes(row.status)&&countBasis.countsAbsence===false)continue;counts[row.status]++;if(row.status==='NOT_APPLICABLE')continue;const weight=row.frequencyBasis.unit==='MINUTE'?row.durationMinutes:1;N(Number.isSafeInteger(weight)&&weight>0&&weight<=1440,'VALIDATION_ERROR');
  const mode=row.status==='EXCUSED'?row.frequencyBasis.excusedMode:row.frequencyBasis.justifiedMode;N(['EXCLUDE','COUNT_PRESENT','COUNT_ABSENT'].includes(mode),'CONFIG_CAPABILITY_DISABLED');if(['JUSTIFIED','EXCUSED'].includes(row.status)&&mode==='EXCLUDE')continue;
  denominator+=weight;if(row.status==='PRESENT'||['JUSTIFIED','EXCUSED'].includes(row.status)&&mode==='COUNT_PRESENT')numerator+=weight;
 }
 return {numerator:basisComplete?numerator:null,denominator:basisComplete?denominator:null,pending,counts,complete,percentage:complete&&basisComplete&&denominator>0?100*numerator/denominator:null,...(!basisComplete?{basisComplete:false}:{})};
}


