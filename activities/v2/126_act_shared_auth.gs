// ACT contracts extend the shared source-based authorizer; no ACT identity store or verifier.
function f4ActivitiesAuthzDomain_(){
 const B=f4MemberAuthzContracts_(),C=f4Core_(),M=f4ActivitiesModel_(),N=B.need;
 const caps=[...new Set(Object.values(f4ActivitiesContracts_()).flatMap(c=>c.capabilities).concat('PRESENTATION_SELF_READ','PRESENTATION_MATERIAL_READ_SELF','PRESENTATION_MATERIAL_NOTICE_SELF','PRESENTATION_BOARD_READ','ACTIVITY_BOARD_READ','ACTIVITY_ATTENDANCE_READ_SELF','ACTIVITY_AGENDA_READ_SELF'))];
 const scope=s=>s==='SELF'||s==='ACTIVITIES'||s==='PRESENTATIONS_PLAN'||/^PRESC_[A-Za-z0-9_-]{1,128}$/.test(s);
 function targets(t){B.exact(t,['ownPerson','allowedTargetScopes','allowedBoards','allowedRoles','allowedPersons','allowedResources',...('resourceScope'in t?['resourceScope']:[])]);N(typeof t.ownPerson==='boolean');B.list(t.allowedTargetScopes,scope,true);for(const k of ['allowedBoards','allowedRoles','allowedPersons'])B.list(t[k],B.id);B.list(t.allowedResources,p=>M.validPath(p));if(t.resourceScope){const s=t.resourceScope;B.exact(s,['kind',...('cycleAccess'in s?['cycleAccess']:[])]);if('cycleAccess'in s)N(s.kind==='PRESENTATIONS'&&s.cycleAccess==='ALL');N(['ACTIVITIES','PRESENTATIONS','OWN_ATTENDANCE','OWN_PRESENTATIONS'].includes(s.kind));N(s.kind==='PRESENTATIONS'?!t.ownPerson&&t.allowedTargetScopes.every(v=>v==='PRESENTATIONS_PLAN'||/^PRESC_[A-Za-z0-9_-]{1,128}$/.test(v)):s.kind!=='ACTIVITIES'?t.ownPerson&&!t.allowedTargetScopes.some(v=>v!=='SELF'):!t.ownPerson&&!t.allowedTargetScopes.some(v=>v!=='ACTIVITIES'));}return t}
 // Institutional operational distribution approved for ACT. A matching role is
 // necessary, never sufficient: the shared authorizer still checks the policy,
 // binding, membership, current function, term, scope and all source fences.
 const operationalRoles={
  ACTIVITY_MANAGE:['PRESIDENTE','VICE_PRESIDENTE','DIRETOR_ADMINISTRATIVO'],
  ACTIVITY_REALIZATION_MANAGE:['PRESIDENTE','VICE_PRESIDENTE','DIRETOR_ADMINISTRATIVO'],
  ACTIVITY_ATTENDANCE_MANAGE:['PRESIDENTE','VICE_PRESIDENTE','DIRETOR_RECURSOS_HUMANOS'],
  ACTIVITY_CONFIG_MANAGE:['PRESIDENTE','VICE_PRESIDENTE','DIRETOR_ADMINISTRATIVO'],
  PRESENTATION_PLAN_MANAGE:['PRESIDENTE','VICE_PRESIDENTE','DIRETOR_ADMINISTRATIVO'],
  PRESENTATION_ACTIVITY_CREATE:['PRESIDENTE','VICE_PRESIDENTE','DIRETOR_ADMINISTRATIVO'],
  PRESENTATION_COMPOSITION_MANAGE:['PRESIDENTE','VICE_PRESIDENTE','DIRETOR_ADMINISTRATIVO'],
  PRESENTATION_THEME_REVIEW:['PRESIDENTE','VICE_PRESIDENTE','DIRETOR_ADMINISTRATIVO'],
  PRESENTATION_COMPATIBILITY_REVIEW:['PRESIDENTE','VICE_PRESIDENTE','DIRETOR_ADMINISTRATIVO'],
  PRESENTATION_DYNAMIC_DECIDE:['PRESIDENTE','VICE_PRESIDENTE','DIRETOR_ADMINISTRATIVO'],
  PRESENTATION_REALIZATION_RECOGNIZE:['PRESIDENTE','VICE_PRESIDENTE','DIRETOR_ADMINISTRATIVO'],
  PRESENTATION_MATERIAL_REVIEW:['PRESIDENTE','VICE_PRESIDENTE','DIRETOR_ADMINISTRATIVO'],
  PRESENTATION_PHOTO_MANAGE:['PRESIDENTE','VICE_PRESIDENTE','DIRETOR_COMUNICACAO'],
  PRESENTATION_PUBLICATION_MANAGE:['PRESIDENTE','VICE_PRESIDENTE','DIRETOR_COMUNICACAO'],
  PRESENTATION_OBLIGATION_EVALUATE:['PRESIDENTE','VICE_PRESIDENTE'],
  PRESENTATION_MATERIAL_EXCEPTION:['PRESIDENTE','VICE_PRESIDENTE'],
  PRESENTATION_RECONSIDERATION_DECIDE:['PRESIDENTE','VICE_PRESIDENTE']
 };
 function roleCapability(c,s){const role=f4MemberContracts_().roles[s.roleKey];return caps.includes(c)&&!c.endsWith('_SELF')&&!!role&&role[0]===s.group&&role[1]===s.actorScope&&(!operationalRoles[c]||operationalRoles[c].includes(s.roleKey));}
 function source(p,explicit){
  B.exact(p,explicit?['schemaVersion','id','revision','active','origin','approvalRef','issuerRef','uid','idPessoa','bindingVersion','subjectEpoch','issuedAt','expiresAt','capabilities','actorScopes','targets','fieldMasks','environment','runId']:['schemaVersion','id','revision','active','origin','approvalRef','issuedAt','expiresAt','proofKind','capabilities','actorSelector','targets','fieldMasks']);
  N(p.schemaVersion===(explicit?'act-explicit-grant-v13':'act-authz-policy-v13')&&B.id(p.id)&&B.positive(p.revision)&&typeof p.active==='boolean'&&p.origin==='ACT_CANONICAL_APPROVAL'&&B.id(p.approvalRef));N(B.timestamp(p.issuedAt)<B.timestamp(p.expiresAt));B.list(p.capabilities,c=>caps.includes(c),true);targets(p.targets);B.exact(p.fieldMasks,p.capabilities);for(const cap of p.capabilities)N(Array.isArray(p.fieldMasks[cap])&&p.fieldMasks[cap].length===0);
  if(explicit){N(B.id(p.issuerRef)&&B.id(p.uid)&&B.id(p.idPessoa)&&B.positive(p.bindingVersion)&&B.positive(p.subjectEpoch)&&p.environment==='DEV'&&(p.runId===null||B.id(p.runId)));B.list(p.actorScopes,s=>B.scopes.includes(s),true);N(p.capabilities.every(c=>!c.endsWith('_SELF')));}
  else{N(['SELF','ROLE'].includes(p.proofKind));const s=p.actorSelector;if(p.proofKind==='SELF'){B.exact(s,['kind']);N(s.kind==='SELF'&&p.targets.ownPerson&&p.capabilities.every(c=>c.endsWith('_SELF')||c==='PRESENTATION_SELF_READ'));}else{B.exact(s,['kind','roleKey','group','actorScope']);const role=f4MemberContracts_().roles[s.roleKey];N(s.kind==='ROLE'&&role&&role[0]===s.group&&role[1]===s.actorScope&&p.capabilities.every(c=>!c.endsWith('_SELF')));}}
  if(!explicit&&p.proofKind==='ROLE')N(p.capabilities.every(c=>roleCapability(c,p.actorSelector)));
  return p;
 }
 function request(r){B.exact(r,['schemaVersion','uid','kind','capability','sourceId','functionId','actorScope','expected','tokenExpiresAt','admissionExpiresAt','target','evidenceRefs']);N(r.schemaVersion==='act-authz-request-v13'&&B.id(r.uid)&&['SELF','ROLE','EXPLICIT','DEV_ADMIN'].includes(r.kind)&&caps.includes(r.capability)&&B.id(r.sourceId));N(r.kind==='ROLE'?B.id(r.functionId):r.functionId===null);N(r.kind==='DEV_ADMIN'?[null,'SELF'].includes(r.actorScope):r.kind==='SELF'?r.actorScope===null:B.scopes.includes(r.actorScope));B.exact(r.expected,['authorityEpoch','policyEpoch','readAuthzEpoch','subjectEpoch','bindingVersion','personRevision','authorizationRevision']);Object.values(r.expected).forEach(v=>N(B.positive(v)));B.timestamp(r.tokenExpiresAt);B.timestamp(r.admissionExpiresAt);B.exact(r.target,['personId','scopeKey','boardId','roleKey','resources','fieldMask','command']);N(B.id(r.target.personId)&&scope(r.target.scopeKey)&&r.target.boardId===null&&r.target.roleKey===null&&r.target.fieldMask.length===0);B.list(r.target.resources,p=>M.validPath(p),!['ACTIVITY_ATTENDANCE_READ_SELF','ACTIVITY_AGENDA_READ_SELF','PRESENTATION_SELF_READ','PRESENTATION_MATERIAL_READ_SELF'].includes(r.capability));N(r.evidenceRefs.length===0);if(['PRESENTATION_SELF_READ','PRESENTATION_MATERIAL_READ_SELF','PRESENTATION_MATERIAL_NOTICE_SELF','PRESENTATION_BOARD_READ','ACTIVITY_BOARD_READ','ACTIVITY_ATTENDANCE_READ_SELF','ACTIVITY_AGENDA_READ_SELF'].includes(r.capability))N(r.target.command===null);else f4ActivitiesPipeline_().validateCommand(r.target.command);return r}
 function match(t,limits,person){targets(limits);N(limits.ownPerson?t.personId===person:limits.allowedPersons.includes(t.personId));N(limits.allowedTargetScopes.includes(t.scopeKey));B.subset(t.resources,limits.allowedResources)}
 function matchSource(get,t,limits,person){
  if(!limits.resourceScope)return match(t,limits,person);
  targets(limits);N(limits.ownPerson?t.personId===person:limits.allowedPersons.includes(t.personId));N(limits.allowedTargetScopes.includes(t.scopeKey)||limits.resourceScope.kind==='PRESENTATIONS'&&limits.resourceScope.cycleAccess==='ALL'&&/^PRESC_[A-Za-z0-9_-]{1,128}$/.test(t.scopeKey));
  const boardCollections=['activityArtifacts','activityAttachmentPolicies','activityPolicies','activityAttendanceRosters','activityTypes','canonicalActivities','activityRealizations','activityAttendanceCollections','activityAttendance','activityJustifications','activityFrequencyPeriods','activityAttendanceRegistry','activityHistory'];
  for(const path of t.resources){N(M.validPath(path)&&path.split('/').length===2);const [collection]=path.split('/');
   if(limits.resourceScope.kind==='PRESENTATIONS'){
    if(['canonicalActivities','activityTypes'].includes(collection)&&t.command?.commandType==='ACT.MEMBER_PRESENTATION_ACTIVITY_CREATE'){const p=t.command.payload;N(t.scopeKey==='PRESC_'+p.cycleRef.id);if(collection==='canonicalActivities'){const a=get(path);N(path===M.path(p.activityRef)&&(!a||a.presentationSessionRef?.id===p.sessionRef.id&&get(M.path(p.sessionRef))?.cycleRef?.id===p.cycleRef.id));}else N(path===M.path(p.typeRef)&&get(path)?.portalModel?.memberPresentation);continue;}
    if(collection==='canonicalActivities'&&t.command?.commandType==='ACT.SESSION_CONFIGURE'){const p=t.command.payload,a=get(path);N(p.activityRef&&M.path(p.activityRef)===path&&t.scopeKey==='PRESC_'+p.cycleRef.id&&a?.civilDate===p.civilDate&&a.status==='SCHEDULED');continue;}
    N(['presentationCycles','presentationPolicyReferences','presentationSessions','presentationUnitHeads','presentationObligations','presentationClaims','calendarAssignments','canonicalPresentations','presentationThematicRequests','presentationQuotaOccupancies','presentationQuotaBuckets','presentationCompatibilityAssessments','presentationDynamicDecisions','presentationDynamicOccupancies','presentationRealizations','presentationMaterialRequirements','presentationMaterialSubmissions','presentationMaterialReviews','presentationMaterialReceipts','presentationMaterialRegularizations','presentationCorrectionOpportunities','presentationArtifacts','presentationPublicationSnapshots','publicPresentations','presentationBoardViews','presentationCapacityHeads','presentationCapacityAssessments','presentationFlowCases','jobOutbox'].includes(collection));const d=get(path);if(t.scopeKey==='PRESENTATIONS_PLAN'){N(path==='presentationCapacityHeads/registry'||collection==='presentationPolicyReferences'&&d?.approved);continue;}const cycleId=t.scopeKey.slice(6),declared=t.command?.payload?.cycleRef;N(d?(collection==='presentationCycles'?d.id:d.cycleRef?.id||d.cycleId)===cycleId:declared?.id===cycleId&&t.command.scope.targetRefs?.some(r=>M.path(r)===path));continue;
   }
   if(limits.resourceScope.kind==='OWN_PRESENTATIONS'){
    if(collection==='presentationFlowCases'&&['ACT.JOINT_CONSENT',null].includes(t.command?.commandType||null)){const f=get(path);N(['JOINT_PROPOSAL','CALENDAR_CHANGE'].includes(f?.kind)&&f.unitRefs?.some(r=>get(M.path(r))?.personRef?.id===person));continue;}
    if(collection==='presentationSessions'&&t.command?.commandType==='ACT.CALENDAR_RESERVE'){const p=t.command.payload,u=get(M.path(p.unitRef)),s=get(path);N(path===M.path(p.sessionRef)&&u?.personRef?.id===person&&s?.cycleRef?.id===u.cycleRef.id);continue;}

    if(collection==='jobOutbox'){
     const out=get(path);N(['ACT.MATERIAL_SUBMIT','ACT.MATERIAL_CORRECTION_SUBMIT','ACT.PHOTO_SUBMIT'].includes(t.command?.commandType)&&out?.effectKind==='DRIVE_CREATE'&&out.targetRef?.kind==='presentation'&&get(M.path(out.targetRef))?.personRef?.id===person);continue;
    }
    N(['presentationUnitHeads','calendarAssignments','canonicalPresentations','presentationThematicRequests','presentationFlowCases','presentationMaterialRequirements','presentationMaterialSubmissions','presentationCorrectionOpportunities','presentationArtifacts','presentationUserDrafts','presentationDynamicOccupancies'].includes(collection));
    const d=get(path);let owner=d?.personRef?.id||d?.ownerPersonRef;if(!owner&&d?.unitRef)owner=get(M.path(d.unitRef))?.personRef?.id;if(!owner&&d?.presentationRef)owner=get(M.path(d.presentationRef))?.personRef?.id;N(owner===person);
   }else if(limits.resourceScope.kind==='ACTIVITIES'){
    if(collection==='presentationSessions'&&['ACT.ACTIVITY_SAVE','ACT.ACTIVITY_PUBLICATION_SET'].includes(t.command?.commandType)){
     const p=t.command.payload,a=p.activityRef&&get(M.path(p.activityRef)),s=get(path);
     N(p.sessionRef&&M.path(p.sessionRef)===path&&a?.presentationSessionRef?.id===s?.id&&s?.activityRef?.id===a?.id);continue;
    }
    N(boardCollections.includes(collection));
   }
   else{const q=t.command,p=q?.payload;
    if(['ACT.ATTENDANCE_JUSTIFICATION_REQUEST','ACT.ATTENDANCE_PROOF_UPLOAD_REQUEST'].includes(q?.commandType)&&p.activityRef){
     const a=get(M.path(p.activityRef)),fact=get(M.path(p.attendanceRef)),ownId='a_'+M.hash({activity:p.activityRef.id,person}).slice(0,32);
     N(p.attendanceRef.id===ownId&&(fact?fact.personRef?.id===person&&fact.activityRef?.id===p.activityRef.id:a?.publicationStatus==='PUBLICADO'&&['PUBLIC','MEMBERS'].includes(a.publishedVisibility)));
     if(path===M.path(p.activityRef)||path===M.path(p.attendanceRef)&&!get(path))continue;
    }
    if(collection==='jobOutbox'&&q?.commandType==='ACT.ATTENDANCE_JUSTIFICATION_REQUEST'){const out=get(path);N(p.proofUploadRef&&M.path(p.proofUploadRef)===path&&out?.ownerPersonRef===person&&out.targetRef?.id===p.attendanceRef.id&&out.details?.purpose==='ATTENDANCE_PROOF');continue;}
    N(['activityAttendance','activityJustifications','activityArtifacts'].includes(collection));const d=get(path);N(d?.personRef?.id===person&&d.ownerPersonRef===person);}
  }
 }
 function suspended(get,r,state,b,now){
  if(r.kind!=='SELF'||state.state!=='SUSPENSO'||!state.blocked)return false;
  if(['PRESENTATION_MATERIAL_READ_SELF','PRESENTATION_MATERIAL_NOTICE_SELF'].includes(r.capability)){return r.target.resources.every(path=>{const d=get(path),pr=path.startsWith('presentationUnitHeads/')&&d?.presentationRef?get(M.path(d.presentationRef)):path.startsWith('canonicalPresentations/')?d:null,rq=pr?.materialRequirementRef&&get(M.path(pr.materialRequirementRef)),rz=pr?.realizationRef&&get(M.path(pr.realizationRef));return pr?.personRef?.id===b.idPessoa&&!!rq&&rz?.recognized===true&&!rq.currentRegularizationRef;});}
  const q=r.target.command;if(!q)return false;if(!['ACT.MATERIAL_SUBMIT','ACT.MATERIAL_CORRECTION_SUBMIT','ACT.ARTIFACT_UPLOAD_REQUEST'].includes(q.commandType)||!q.payload.presentationRef)return false;
  const pr=get(M.path(q.payload.presentationRef)),rq=pr?.materialRequirementRef&&get(M.path(pr.materialRequirementRef)),rz=pr?.realizationRef&&get(M.path(pr.realizationRef));return pr?.personRef?.id===b.idPessoa&&!!rq&&rz?.recognized===true&&!rq.currentRegularizationRef;
 }
 function self(get,r,state,b,p,now){N(r.target.personId===b.idPessoa&&(r.capability.endsWith('_SELF')||r.capability==='PRESENTATION_SELF_READ'));const attendance=['ACTIVITY_AGENDA_READ_SELF','ACTIVITY_ATTENDANCE_READ_SELF','ACTIVITY_JUSTIFICATION_SELF'].includes(r.capability);N((state.state==='MEMBRO_EFETIVO'||state.state==='MEMBRO_INGRESSANTE'&&attendance)&&!state.blocked||suspended(get,r,state,b,now));}
 return {contracts:()=>({...B,caps,targets,request,policy:p=>source(p,false),grant:p=>source(p,true),match,roleCapability}),match:matchSource,self,suspended};
}


