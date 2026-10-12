// Hosted ACT adapters share the existing verifier, admission and command coordinator.
// Configuration and policy identifiers are protected Core inputs, never browser claims.
function f4ActivitiesFunctionalPorts_(options){
 const M=f4ActivitiesModel_(),N=M.need,F=f4Firestore_(),C=f4Core_(),db=F.connect(options.config),clock=options.clock||Date.now;
 let snapshot=null,readProofs=null;
 const get=p=>{if(!snapshot)return db.get(p);if(!snapshot.has(p))snapshot.set(p,db.get(p));return snapshot.get(p);};
 const identity=options.identity,settings=options.settings;
 // Reuse source reads only while assembling one authorization/preparation.
 // The enclosing reader and the command pipeline still recheck fresh sources.
 function sourceSnapshot(fn,force=false,knownPaths=[]){if(snapshot&&!force)return fn();const prior=snapshot;snapshot=new Map();try{
  const paths=[...new Set(['coreDevAdmins/'+identity.uid,'memberAuthorityControl/MEMBERS','memberAuthBindings/'+identity.uid,'portalUsers/'+identity.uid,...settings.policyIds.map(id=>'memberAuthzPolicies/'+id),...knownPaths])];
  function batch(missing){for(let i=0;i<missing.length;i+=180){const part=missing.slice(i,i+180),values=db.getMany(part);part.forEach((p,j)=>snapshot.set(p,values[j]));}}
  batch(paths);
  const binding=get('memberAuthBindings/'+identity.uid);
  if(binding?.idPessoa){batch(['memberAggregateHeads/'+binding.idPessoa,'memberCurrentStates/'+binding.idPessoa,'memberAuthzSubjects/'+binding.idPessoa].filter(p=>!snapshot.has(p)));const head=get('memberAggregateHeads/'+binding.idPessoa);if(Array.isArray(head?.paths)){batch(head.paths.filter(p=>!snapshot.has(p)));const functions=head.paths.filter(p=>p.startsWith('memberFunctions/')).map(get).filter(Boolean);batch([...new Set(functions.flatMap(f=>['memberBoardHeads/'+f.idDiretoria,'memberRoleCatalog/'+f.idCargo,'memberGovernanceTerms/'+f.idTermo]))].filter(p=>!snapshot.has(p)));batch([...new Set(functions.flatMap(f=>get('memberBoardHeads/'+f.idDiretoria)?.paths||[]))].filter(p=>!snapshot.has(p)));}}
  return fn();
 }finally{snapshot=prior;}}
 // One source snapshot for a list, followed by an independent fresh snapshot
 // for all final proofs. Nothing survives the request or authorizes a writer.
 function readWithAuthoritySnapshot(fn){
  // Nested read adapters share source reads; every collected proof is checked
  // against one independent fresh snapshot immediately before the outer return.
  if(readProofs)return fn(proofs=>readProofs.push(...proofs),get);
  return sourceSnapshot(()=>{readProofs=[];try{const result=fn(proofs=>readProofs.push(...proofs),get),proofs=readProofs.slice(),paths=[...snapshot.keys()];N(proofs.length>0,'AUTHORIZATION_DENIED');options.authorityPaths?.push(...paths);sourceSnapshot(()=>{for(const proof of proofs)f4SharedAuthz_().recheck(get,proof.q,proof.proof,clock());},true,paths);return result;}finally{readProofs=null;}},true);
 }
 function refs(value,out=new Map(),fromDocument=false){
  if(value&&typeof value==='object'){
   // Population coverage is a fenced Core proof, not the actor's target set.
   if(['act-incidence-coverage-v1','act-incidence-candidate-v1','act-quota-queue-v1'].includes(value.schemaVersion)||value.coverageKind==='act-incidence-coverage-v1')return out;
   if(value.kind&&value.id){try{const path=M.path(value);if(M.validPath(path))out.set(path,value);return out}catch{}}
   // Historical links must not make every new operation load an ever-growing
   // chain. An explicitly supplied predecessor remains an input precondition.
   const historical=['predecessorRef','predecessorCoverageRef','initialCoverageRef','predecessorRecognitionRef','predecessorSubmissionRef','lastCalendarChangeRef','changeSummary','proposedBy'];
   for(const [key,v] of Object.entries(value))if(!(fromDocument&&historical.includes(key))&&!(key==='sourceManifestRef'&&value.originRef?.kind==='CORE_COMMAND')&&!['actor','actorAuthorityBasis','authorityBasis','decision','originRef','basisRevisionVector','roster','records','sourceHeads'].includes(key))refs(v,out,fromDocument);
  }return out;
 }
 function selectors(){
  if(f4DevAdmin_().profile(get,identity.uid,clock()))return [null,'SELF'].map(actorScope=>({kind:'DEV_ADMIN',sourceId:identity.uid,functionId:null,actorScope}));
  const a=f4MemberAuthz_().identify(get,identity.uid,clock(),s=>s.state==='SUSPENSO'),result=[];
  for(const sourceId of settings.policyIds){const source=get('memberAuthzPolicies/'+sourceId);if(!source?.active)continue;
   if(source.proofKind==='SELF')result.push({kind:'SELF',sourceId,functionId:null,actorScope:null});
   if(source.proofKind==='ROLE')for(const current of a.derived.currentFunctions){const f=a.actor.d['memberFunctions/'+current.idVigencia],role=f&&a.actor.d['memberRoleCatalog/'+f.idCargo];if(role?.roleKey===source.actorSelector.roleKey&&f.scopeKey===source.actorSelector.actorScope)result.push({kind:'ROLE',sourceId,functionId:f.id,actorScope:f.scopeKey});}
  }return result;
 }
 function prepare(input,quotaPrepared=false){f4DevReadTiming_().command(input.commandType);return f4DevReadTiming_().stage('prepare',()=>sourceSnapshot(()=>prepareFromSources(input,quotaPrepared),quotaPrepared));}
 function prepareFromSources(input,quotaPrepared=false){
  C.exact(input,['operationId','commandType','payload']);N(C.op(input.operationId)&&input.commandType.startsWith('ACT.'),'VALIDATION_ERROR');
  const cmd=input.commandType.slice(4),contract=f4ActivitiesContracts_()[cmd];N(contract&&settings.commands.includes(cmd),'CONFIG_CAPABILITY_DISABLED');
  const selected=refs(input.payload);
  if(cmd==='DRAFT_SAVE'){const r=M.ref('draft',input.payload.draftId||M.hash({operationId:input.operationId,kind:'draft'}).slice(0,32));selected.set(M.path(r),r);}
  const expectedRefs=new Map(selected);
  if(f4AttendanceScope_(cmd)){
   for(const a of input.payload.activityRefs||[]){const r=M.ref('attendanceCollection',a.id);expectedRefs.set(M.path(r),r);}const activity=input.payload.activityRef;if(activity)expectedRefs.set(M.path(M.ref('attendanceCollection',activity.id)),M.ref('attendanceCollection',activity.id));
   if(cmd==='ATTENDANCE_ENROLL'){const r=M.ref('attendance','a_'+M.hash({activity:activity.id,person:input.payload.personRef.id}).slice(0,32));expectedRefs.set(M.path(r),r);}
  }
  if(cmd==='CYCLE_CREATE'){const r=M.ref('cycle',M.hash({operationId:input.operationId,kind:'cycle'}).slice(0,32));expectedRefs.set(M.path(r),r);}
  if(cmd==='INCIDENCE_EVALUATE'){const r=M.ref('obligation',M.hash({person:input.payload.personRef.id,cycle:input.payload.cycleRef.id,kind:'ORDINARY',version:1}).slice(0,32));expectedRefs.set(M.path(r),r);}
  if(cmd==='PUBLICATION_APPROVE')for(const pr of input.payload.presentationRefs){const r=M.ref('publication',pr.id);expectedRefs.set(M.path(r),r);}
  if(cmd==='CORRECTION_OPPORTUNITY_OPEN'){const r=M.ref('opportunity',M.hash({operationId:input.operationId,kind:'opportunity'}).slice(0,32));expectedRefs.set(M.path(r),r);}
  const documents=new Map();
  while(documents.size<expectedRefs.size){
   N(expectedRefs.size<=48,'SCOPE_TOO_LARGE','PREPARATION_DEPENDENCIES');
   const pending=[...expectedRefs].filter(([path])=>!documents.has(path)),values=db.getMany(pending.map(([path])=>path));
   pending.forEach(([path,ref],index)=>{const d=values[index];documents.set(path,d);if(snapshot)snapshot.set(path,d);if(d)refs(d,expectedRefs,true);if(d&&ref.kind==='cycle'){const r=M.ref('capacity',ref.id);expectedRefs.set(M.path(r),r);}if(d?.activityRef){const r=M.ref('attendanceCollection',d.activityRef.id);expectedRefs.set(M.path(r),r);}});
  }
  if(!f4AttendanceScope_(cmd)||input.payload.calendarAssessmentRef||input.payload.sessionRef){const r=M.ref('capacity','registry');expectedRefs.set(M.path(r),r);documents.set(M.path(r),get(M.path(r)));}
  function addExpected(ref){const path=M.path(ref);if(!expectedRefs.has(path)){expectedRefs.set(path,ref);documents.set(path,get(path));}}
  if(cmd==='QUOTA_CONFIGURE')for(const item of input.payload.axisLimits||[])addExpected(M.ref('bucket',M.hash({cycle:input.payload.cycleRef.id,axis:item.axisRef.id}).slice(0,32)));
  for(const [path,d] of documents){if(!d)continue;
   if(path===M.path(input.payload.unitRef||M.ref('unit','NONE'))&&cmd==='THEME_SUBMIT')addExpected(d.thematicRequestRef||M.ref('thematic',M.hash({unit:d.id,cycle:d.cycleRef.id}).slice(0,32)));
   if(cmd==='DYNAMIC_DECIDE'&&path.startsWith('presentationUnitHeads/')&&d.cycleRef)addExpected(d.dynamicOccupancyRef||M.ref('dynamic',M.hash({unit:d.id,cycle:d.cycleRef.id}).slice(0,32)));
   if(d.cycleRef){const cycle=documents.get(M.path(d.cycleRef)),cfg=cycle?.currentConfigRef&&documents.get(M.path(cycle.currentConfigRef));if(cfg?.parameters.quota)for(const axis of [d.principalAxisRef,d.candidate?.principalAxisRef,path.startsWith('presentationQuotaOccupancies/')?d.axisRef:null,cmd==='THEME_SUBMIT'&&path.startsWith('presentationUnitHeads/')?input.payload.principalAxisRef:null].filter(Boolean))addExpected(M.ref('bucket',M.hash({cycle:d.cycleRef.id,axis:axis.id}).slice(0,32)));}
  }
  const expected=[...expectedRefs].map(([path,ref])=>({...ref,...(documents.get(path)?{revision:documents.get(path).revision}:{exists:false})}));
  const direct=[...selected.values()],primary=direct.filter(r=>!['policy','artifact','config','assessment','receipt','review','realization','reconciliation'].includes(r.kind)),targetRefs=primary.length?primary:direct;N(targetRefs.length>0&&targetRefs.length<=32,'SCOPE_TOO_LARGE');
  let last;
  for(const authorization of selectors()){
   const self=authorization.kind==='SELF'||authorization.kind==='DEV_ADMIN'&&authorization.actorScope==='SELF';
   if(cmd==='DYNAMIC_CANCEL'&&input.payload.cause!=='WITHDRAWAL'&&self)continue;
   if(!contract.actors.includes(self?'PORTAL_AUTHENTICATED_MEMBER':'BOARD_ACTOR'))continue;
   // A linked session/flow can mention a partner. Do not add unrelated partner
   // documents to a member's optimistic preconditions; handlers still enforce
   // their own required reads and ownership inside the transaction.
   const binding=get('memberAuthBindings/'+identity.uid),ownExpected=!self?expected:expected.filter(ref=>{const path=M.path({kind:ref.kind,id:ref.id,...(ref.parent?{parent:ref.parent}:{})}),d=documents.get(path),owner=d?.personRef?.id||d?.ownerPersonRef||(d?.unitRef&&documents.get(M.path(d.unitRef))?.personRef?.id)||(d?.presentationRef&&documents.get(M.path(d.presentationRef))?.personRef?.id);if(!owner||owner===binding.idPessoa||selected.has(path))return true;return ['JOINT_PROPOSE','JOINT_CONSENT'].includes(cmd)&&ref.kind==='unit';});
   const request={schemaVersion:'geapa-act-command-v13',...input,target:{scope:{targetRefs},expected:ownExpected,authorization},evidenceRefs:[],decisionRef:null};
   try{const q=f4ActivitiesPipeline_().authorization(get,request,identity,clock(),identity.expiresAt)[0];f4SharedAuthz_().authorize(get,q,clock());
    if(!quotaPrepared&&cmd!=='QUOTA_CONFIGURE'){let changed=false;for(const ref of expectedRefs.values())if(ref.kind==='bucket')changed=f4ActivitiesQuotaQueue_({db,config:options.config,clock}).refresh(ref,reader=>authorizePrepared(reader,request))||changed;if(changed)return prepare(input,true);}
    return request;}catch(e){last=e;if(!['AUTH_BINDING_INVALID','ACCESS_REVOKED','AUTHORIZATION_DENIED','M13_AUTHZ_DENIED'].includes(e.code))throw e;}
  }
  N(false,'AUTHORIZATION_DENIED',last?.reasonCode||'NO_CURRENT_AUTHORITY');
 }
 function readAuthority(resources,self,readCapability,scopeKey){return sourceSnapshot(()=>readAuthorityFromSources(resources,self,readCapability,scopeKey));}
 function readAuthorityFromSources(resources,self,readCapability,scopeKey){
  if(f4DevAdmin_().profile(get,identity.uid,clock())){
   const d=f4DevAdmin_(),a=d.identify(get,identity.uid,clock()),capability=readCapability||(self?'ACTIVITY_ATTENDANCE_READ_SELF':'ACTIVITY_BOARD_READ');
   const q={schemaVersion:'act-authz-request-v13',uid:identity.uid,kind:'DEV_ADMIN',sourceId:identity.uid,functionId:null,actorScope:self?'SELF':null,capability,expected:d.expected(a),tokenExpiresAt:M.stamp(identity.expiresAt),admissionExpiresAt:M.stamp(identity.expiresAt),target:{personId:a.binding.idPessoa,scopeKey:self?'SELF':scopeKey||'ACTIVITIES',boardId:null,roleKey:null,resources,fieldMask:[],command:null},evidenceRefs:[]};
   return {personId:a.binding.idPessoa,q,proof:d.authorize(get,q,clock()).proof};
  }
  const capability=readCapability||(self?'ACTIVITY_ATTENDANCE_READ_SELF':'ACTIVITY_BOARD_READ'),identified=f4MemberAuthz_().identify(get,identity.uid,clock(),state=>['PRESENTATION_MATERIAL_READ_SELF','PRESENTATION_MATERIAL_NOTICE_SELF'].includes(capability)&&state.state==='SUSPENSO'&&state.blocked),b=identified.binding,control=identified.control,subject=identified.subject;N(['ACTIVITY_ATTENDANCE_READ_SELF','ACTIVITY_AGENDA_READ_SELF','ACTIVITY_BOARD_READ','PRESENTATION_SELF_READ','PRESENTATION_MATERIAL_READ_SELF','PRESENTATION_MATERIAL_NOTICE_SELF','PRESENTATION_BOARD_READ'].includes(capability),'VALIDATION_ERROR');
  for(const selector of selectors()){if((selector.kind==='SELF')!==self)continue;const source=get('memberAuthzPolicies/'+selector.sourceId);if(!source.capabilities.includes(capability))continue;
   const q={schemaVersion:'act-authz-request-v13',uid:identity.uid,...selector,capability,expected:{authorityEpoch:control.epoch,policyEpoch:control.policyEpoch,readAuthzEpoch:control.readAuthzEpoch,subjectEpoch:subject.authzEpoch,bindingVersion:b.bindingVersion,personRevision:identified.actor.h.revision,authorizationRevision:source.revision},tokenExpiresAt:M.stamp(identity.expiresAt),admissionExpiresAt:M.stamp(identity.expiresAt),target:{personId:b.idPessoa,scopeKey:self?'SELF':scopeKey||'ACTIVITIES',boardId:null,roleKey:null,resources,fieldMask:[],command:null},evidenceRefs:[]};
   try{const proof=f4SharedAuthz_().authorize(get,q,clock()).proof;return {personId:b.idPessoa,q,proof};}catch(e){if(!['AUTH_BINDING_INVALID','ACCESS_REVOKED','AUTHORIZATION_DENIED','M13_AUTHZ_DENIED'].includes(e.code))throw e;}
  }N(false,'AUTHORIZATION_DENIED');
 }
 function resume(key){
  f4SharedCommandCatalog_().selector(key);N(key.commandType.startsWith('ACT.')&&settings.commands.includes(key.commandType.slice(4)),'CONFIG_CAPABILITY_DISABLED');
  return db.transaction(tx=>{const get=p=>tx.get(p).data();f4ActivitiesPipeline_().statusAccess(get,identity,key,clock());const provenance=get('commandProvenance/'+key.operationId),ticket=provenance&&get('memberAdmissionTickets/'+provenance.ticketId);
   N(ticket?.ownerUid===identity.uid&&ticket.digest===key.digest&&ticket.requestEnvelope?.commandType===key.commandType,'AUTHORIZATION_DENIED');
   const r=ticket.requestEnvelope;return Object.fromEntries(['schemaVersion','operationId','commandType','target','payload','evidenceRefs','decisionRef'].map(k=>[k,r[k]]));
  });
 }
 function authorizeUpload(reader,out){
  if(out?.details?.purpose==='PRESENTATION_PHOTO')return f4ActivitiesPhotoAuthorizeUpload_(reader,out,{identity,settings,clock});
  if(out?.details?.purpose==='ATTENDANCE_PROOF')return f4AttendanceProofAuthorizeUpload_(reader,out,{identity,settings,clock});
  N(settings.commands.includes('ARTIFACT_UPLOAD_REQUEST')&&out.effectKind==='DRIVE_CREATE'&&out.actor?.uid===identity.uid&&C.op(out.createdByOperation),'AUTHORIZATION_DENIED');
  const provenance=reader('commandProvenance/'+out.createdByOperation),ticket=provenance&&reader('memberAdmissionTickets/'+provenance.ticketId),request=ticket?.requestEnvelope;
  N(provenance?.ownerUid===identity.uid&&request?.commandType==='ACT.ARTIFACT_UPLOAD_REQUEST'&&M.path(request.payload.presentationRef)===M.path(out.targetRef),'AUTHORIZATION_DENIED');
  const key={operationId:out.createdByOperation,commandType:request.commandType,digest:provenance.digest};f4ActivitiesPipeline_().statusAccess(reader,identity,key,clock());
  const pr=reader(M.path(out.targetRef)),rq=pr?.materialRequirementRef&&reader(M.path(pr.materialRequirementRef)),expected=out.details.requirementRevisionRef;
  N(pr?.personRef?.id===out.actor.personId&&rq&&rq.currentRevisionRef.id===expected.id&&rq.currentRevisionRef.number===expected.number&&(!rq.currentRegularizationRef||out.status==='CONFIRMED'),'AUTHORIZATION_DENIED','UPLOAD_CONTEXT_CHANGED');
  if(request.payload.opportunityRef){const opportunity=reader(M.path(request.payload.opportunityRef));N(opportunity?.presentationRef?.id===pr.id&&opportunity.status==='OPEN'&&clock()<=M.ms(opportunity.dueAt),'DEADLINE_VIOLATION');}
 }
 function publicFact(d){const allowed=['recordMode','historicalAttendanceFact','countsAbsence','applicableAtDate','id','revision','title','name','label','location','civilDate','startsAt','endsAt','typeRef','typeSnapshot','status','personRef','membershipRef','activityRef','periodId','durationMinutes','frequencyBasis','state','opensAt','closesAt','rosterCount','unrecordedCount','reason','submittedAt','decisionReason','decidedAt','recordedAt','recordedBy','pendingJustificationRef','countsAttendance','allowsJustification','visibility','civilStart','civilEnd','unit','justifiedMode','summary','action','subjectRef','occurredAt','realizationRef','periodRef','justificationEndsAt'];return Object.fromEntries(Object.entries(d).filter(([k])=>allowed.includes(k)));}
 function read(input){
  C.exact(input,['view','limit',...['cursor','periodId','activityId'].filter(k=>Object.hasOwn(input,k))]);N(Number.isInteger(input.limit)&&input.limit>=1&&input.limit<=20,'VALIDATION_ERROR');
  const views={types:'activityTypes',activities:'canonicalActivities',periods:'activityFrequencyPeriods',collections:'activityAttendanceCollections',attendance:'activityAttendance',justifications:'activityJustifications',history:'activityHistory',mine:'activityAttendance',myHistory:'activityHistory',frequency:'activityAttendance'},collection=views[input.view];N(collection,'VALIDATION_ERROR');
  const self=['mine','myHistory','frequency'].includes(input.view),binding=get('memberAuthBindings/'+identity.uid);N(binding?.active,'AUTHORIZATION_DENIED');
  // An empty SELF listing is authorized by its protected policy and current Members identity.
  const auth=readAuthority(self?[]:['activityAttendanceRegistry/current'],self),filters=[];
  if(self)filters.push({field:'ownerPersonRef',op:'EQUAL',value:auth.personId});
  if(input.periodId){N(M.id(input.periodId),'VALIDATION_ERROR');filters.push({field:'periodId',op:'EQUAL',value:input.periodId});}
  if(input.activityId){N(M.id(input.activityId),'VALIDATION_ERROR');filters.push({field:'activityRef',op:'EQUAL',value:M.ref('activity',input.activityId)});}
  const epoch=get('activityAttendanceRegistry/current')?.epoch||0;
  if(input.cursor)N(input.cursor.epoch===epoch&&M.id(input.cursor.id),'STALE_REVISION','LIST_CHANGED');
  const plan={collection,filters,limit:input.limit,order:[{field:'__name__',direction:'ASCENDING'}]};
  if(input.cursor)plan.cursor=[new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+collection+'/'+input.cursor.id)];
  const rows=db.query(plan),items=rows.map(({data:d})=>{
   if(self)N(d.ownerPersonRef===auth.personId,'AUTHORIZATION_DENIED');const out=publicFact(d);
   if(collection==='activityAttendance'){const activity=get(M.path(d.activityRef)),c=get('activityAttendanceCollections/'+d.activityRef.id),realization=activity?.realizationRef&&get(M.path(activity.realizationRef));out.activityTitle=activity?.title||'';out.countable=(d.recordMode==='HISTORICAL_PARTIAL'&&d.historicalAttendanceFact===true||!!c?.lastFinalizedAt&&!!d.officialRecordedAt)&&activity?.legacyStatus!=='CANCELADA'&&!['CANCELLED','NOT_OCCURRED','ARCHIVED'].includes(activity?.status);}
   return out;
  });
  f4SharedAuthz_().recheck(get,auth.q,auth.proof,clock());N((get('activityAttendanceRegistry/current')?.epoch||0)===epoch,'STALE_REVISION','LIST_CHANGED');
  const complete=rows.length<input.limit;return {items,epoch,complete,nextCursor:complete?null:{id:items[items.length-1].id,epoch},...(input.view==='frequency'?{frequency:f4AttendanceFrequency_(items,complete&&!input.cursor)}:{})};
 }
 function capabilityHints(scopeKey='ACTIVITIES'){
  if(f4DevAdmin_().profile(get,identity.uid,clock()))return new Set(f4ActivitiesAuthzDomain_().contracts().caps);
  const hints=new Set(),K=f4ActivitiesAuthzDomain_().contracts(),binding=get('memberAuthBindings/'+identity.uid);
  for(const selector of selectors())if(selector.kind==='ROLE'){
   const source=get('memberAuthzPolicies/'+selector.sourceId);
   try{K.policy(source);K.live(source,clock());}catch(e){if(['AUTH_BINDING_INVALID','ACCESS_REVOKED','M13_AUTHZ_DENIED','AUTHORIZATION_DENIED'].includes(e.code))continue;throw e;}
   if(source.targets.ownPerson||!source.targets.allowedPersons.includes(binding.idPessoa)||!(source.targets.allowedTargetScopes.includes(scopeKey)||source.targets.resourceScope?.kind==='PRESENTATIONS'&&source.targets.resourceScope.cycleAccess==='ALL'&&/^PRESC_/.test(scopeKey)))continue;
   for(const cap of source.capabilities)if(K.roleCapability(cap,source.actorSelector))hints.add(cap);
  }
  return hints;
 }
 function commandAvailable(command,scopeKey='ACTIVITIES'){
  const contract=f4ActivitiesContracts_()[command];return settings.commands.includes(command)&&!!contract&&contract.capabilities.some(cap=>capabilityHints(scopeKey).has(cap));
 }
 function context(){return readWithAuthoritySnapshot(recheck=>{
  let auth,board=false,materialOnly=false;try{auth=readAuthority(['activityAttendanceRegistry/current'],false);board=true}catch(e){if(!['AUTHORIZATION_DENIED','AUTH_BINDING_INVALID','ACCESS_REVOKED','M13_AUTHZ_DENIED'].includes(e.code))throw e;try{auth=readAuthority([],true,'ACTIVITY_AGENDA_READ_SELF');}catch(inner){if(!['AUTHORIZATION_DENIED','AUTH_BINDING_INVALID','ACCESS_REVOKED','M13_AUTHZ_DENIED'].includes(inner.code))throw inner;auth=readAuthority([],true,'PRESENTATION_MATERIAL_READ_SELF');materialOnly=true;}}
  const head=get('memberAggregateHeads/'+auth.personId),person=get('memberIdentities/'+auth.personId),profile=get('memberPortalProfiles/'+auth.personId);
  N(head&&person,'UNRESOLVED_DEPENDENCY');recheck([auth]);
  const hints=materialOnly?new Set():capabilityHints();
  // UI hints only. Every prepare/admission/commit reauthorizes the real command.
  return {schemaVersion:'act-portal-context-v1',idPessoa:auth.personId,nome:profile?.sourceRevision===head.revision?profile.displayLabel:person.displayName||'',accessProfile:auth.proof.kind==='DEV_ADMIN'?'DEV_ADMIN':null,board,materialOnly,activityAgendaDirect:settings.activityAgendaDirect===true,canManageActivity:board&&hints.has('ACTIVITY_MANAGE')&&settings.commands.includes('ACTIVITY_SAVE'),canRealizeActivity:board&&hints.has('ACTIVITY_REALIZATION_MANAGE')&&settings.commands.includes('ACTIVITY_REALIZE'),canConfigureActivity:board&&hints.has('ACTIVITY_CONFIG_MANAGE')&&settings.commands.includes('ACTIVITY_TYPE_SAVE'),canCreateActivity:board&&hints.has('ACTIVITY_MANAGE')&&settings.commands.includes('ACTIVITY_SAVE'),canRegisterAttendance:board&&hints.has('ACTIVITY_ATTENDANCE_MANAGE')&&settings.commands.includes('ATTENDANCE_FINALIZE'),validUntil:Math.min(identity.expiresAt,clock()+60000)};
 });}
 function authorizePrepared(reader,request){
  N(request?.commandType?.startsWith('ACT.')&&settings.commands.includes(request.commandType.slice(4)),'CONFIG_CAPABILITY_DISABLED');
  const q=f4ActivitiesPipeline_().authorization(reader,request,identity,clock(),identity.expiresAt)[0];
  return f4SharedAuthz_().authorize(reader,q,clock());
 }
 return {withSourceSnapshot:fn=>sourceSnapshot(fn,true),recheckAuthority:auth=>sourceSnapshot(()=>f4SharedAuthz_().recheck(get,auth.q,auth.proof,clock()),true),prepare,read,readAuthority,capabilityHints,commandAvailable,readWithAuthoritySnapshot,resume,context,authorizeUpload,authorizePrepared};
}



// Group source fetches inside this exact transaction; authorization still derives
// and validates every source, and the next transaction starts with fresh reads.
function f4ActivitiesPrefetchCommand_(tx,uid,key){
 const C=f4Core_(),N=C.need;N(C.id(uid),'AUTH_BINDING_INVALID');
 const batch=paths=>{const list=[...new Set(paths)];N(list.length<=720,'SCOPE_TOO_LARGE');for(let i=0;i<list.length;i+=180)tx.getAll(list.slice(i,i+180));};
 const get=p=>tx.get(p).data();
 batch(['coreDevAdmins/'+uid,'memberAuthorityControl/MEMBERS','memberAuthBindings/'+uid,'portalUsers/'+uid,'portalReaderControl/current',...(key?.operationId?['memberAdmissionCounters/'+(f4ActivitiesAdmissionLimits_(uid)?.counterId||uid)]:[]),...(key?.operationId?['commandRequests','commandProvenance','commandLedger','commandReceipts','commandTombstones','commandLeases','memberAudit'].map(c=>c+'/'+key.operationId):[])]);
 const b=get('memberAuthBindings/'+uid);if(b?.idPessoa){
  batch(['memberAggregateHeads/'+b.idPessoa,'memberCurrentStates/'+b.idPessoa,'memberAuthzSubjects/'+b.idPessoa]);
  const h=get('memberAggregateHeads/'+b.idPessoa);if(Array.isArray(h?.paths)){
   batch(h.paths);const functions=h.paths.filter(p=>p.startsWith('memberFunctions/')).map(get).filter(Boolean);
   batch(functions.flatMap(f=>['memberBoardHeads/'+f.idDiretoria,'memberRoleCatalog/'+f.idCargo,'memberGovernanceTerms/'+f.idTermo]));
   const boardPaths=functions.flatMap(f=>get('memberBoardHeads/'+f.idDiretoria)?.paths||[]);batch(boardPaths);
   batch(boardPaths.filter(p=>p.startsWith('memberGovernanceBoardTerms/')).map(get).filter(Boolean).map(t=>'memberGovernanceTerms/'+t.idTermo));
  }
 }
 const r=key?.operationId&&get('commandRequests/'+key.operationId),p=key?.operationId&&get('commandProvenance/'+key.operationId);
 if(r?.ticketId)batch(['memberAdmissionTickets/'+r.ticketId,'memberAdmissionUses/'+r.ticketId]);
 if(p?.authRequests)batch(p.authRequests.filter(q=>['ROLE','SELF','EXPLICIT'].includes(q.kind)).map(q=>(q.kind==='EXPLICIT'?'memberExplicitGrants/':'memberAuthzPolicies/')+q.sourceId));
}
