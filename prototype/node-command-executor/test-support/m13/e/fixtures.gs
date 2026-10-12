// E synthetic admission documents, never a family producer or institutional source.
function m13EFixture_(family='A',patchOverride,actorKind='EXPLICIT',operationId='e'.repeat(32)){
 const C=f4Core_(),F=f4Firestore_(),K=f4MemberCommandCatalog_(),T=F.Timestamp.fromMillis;
 const types={A:'members.self.update',B:'members.profile.correct',C:'members.ingress.execute',D:'members.discipline.suspend',E:'members.discipline.dismiss',F:'members.membership.voluntaryExit',G:'members.governance.appoint',H:'members.governance.endFunction',I:'members.governance.boardTerm',J:'members.binding.associate',K1:'members.binding.revoke'},type=types[family],cap=f4MemberCommandCatalog_().types[type][0];
 const membership=['D','E','F'].includes(family),governance=['G','H'].includes(family),binding=['J','K1'].includes(family),x=m13BFixture_(family==='A'?'SELF':actorKind,cap,family==='F'?'DIRETOR_RECURSOS_HUMANOS':'DIRETOR_ADMINISTRATIVO'),id=operationId;
 const target=x.req.target.personId,patch=family==='A'?{contactEmail:'updated@example.invalid'}:patchOverride||{displayName:'Corrected synthetic name'};
 const paths=membership?['memberEvents/EVT_'+id]:family==='A'?['memberPrivateProfiles/'+target]:[...(Object.hasOwn(patch,'displayName')?['memberIdentities/'+target]:[]),...(Object.keys(patch).some(k=>k!=='displayName')?['memberPrivateProfiles/'+target]:[])];
 x.d['memberAuthorityControl/MEMBERS'].activeWriterVersion='M13';
 x.req.target.resources=membership?[...paths,'memberLinks/VINC-002']:paths;x.req.target.fieldMask=Object.keys(patch).sort();x.d[x.grantPath].targets.allowedResources=x.req.target.resources;x.d[x.grantPath].fieldMasks[x.req.capability]=[...x.req.target.fieldMask];
 x.req.tokenExpiresAt=T(x.now+3600000);x.req.admissionExpiresAt=T(x.now+120000);
 if(family==='B')x.d['memberEvidenceAttestations/SYN_EVIDENCE'].targetResourceIds=paths;
 const r={schemaVersion:'geapa-members-command-v13',operationId:id,ownerUid:'SYN_UID',commandType:type,target:{personId:target},expectedRevisions:[{resourceType:'memberAggregateHeads',id:target,revision:1},...paths.map(path=>({resourceType:path.split('/')[0],id:target,revision:1}))],payload:family==='A'?{patch}:{corrections:patch,justification:'Synthetic correction'},evidenceRefs:x.req.evidenceRefs,decisionRef:family==='A'?null:'SYN_APPROVAL',ticketId:'E_TICKET',digest:'0'.repeat(64),receivedAt:T(x.now),state:'RECEIVED'};
 if(membership){r.target.linkId='VINC-002';r.payload=family==='D'?{effectiveDate:'2026-09-14',endDate:'2026-09-30'}:{effectiveDate:'2026-09-14'};if(family==='F')r.decisionRef=null;r.expectedRevisions[1]={resourceType:'memberEvents',id:'EVT_'+id,revision:0};x.req.target.fieldMask=[];x.d[x.grantPath].fieldMasks[cap]=[];x.req.target.effect=family==='D'?'SUSPENSAO':'DESLIGAMENTO';const att=x.d['memberEvidenceAttestations/SYN_EVIDENCE'];att.effect=x.req.target.effect;att.targetResourceIds=paths;}
 let facts={};
 if(family==='C'){
  for(const path of ['memberIdentities/'+target,'memberPrivateProfiles/'+target,'memberLinks/VINC-002','memberAggregateHeads/'+target])delete x.d[path];
  r.target={personId:target,linkId:'VINC-002'};r.payload={operation:'RECORD_INGRESS',identity:{displayName:'Approved synthetic entrant'},identifiers:[{id:'IDENT_E',kind:'EMAIL',value:'entrant@example.invalid'}],startDate:'2026-09-14',endDate:null,originalEventId:null};
  x.req.target.resources=['memberIdentities/'+target,'memberIdentifiers/IDENT_E','memberLinks/VINC-002','memberEvents/EVT_'+id];x.req.target.fieldMask=[];x.req.target.effect='INGRESSO';x.d[x.grantPath].targets.allowedResources=x.req.target.resources;x.d[x.grantPath].fieldMasks[cap]=[];x.d['memberEvidenceAttestations/SYN_EVIDENCE'].effect='INGRESSO';
  r.expectedRevisions=[{resourceType:'memberAggregateHeads',id:target,revision:0},...x.req.target.resources.map(p=>({resourceType:p.split('/')[0],id:p.split('/')[1],revision:0}))];
  facts={schemaVersion:'member-ingress-content-v13',personId:target,processRef:'PROCESS_E',cohortRef:'COHORT_E',approval:{kind:patchOverride?.exceptional?'EXCEPTIONAL':'REGULAR',reference:'SYN_APPROVAL'},invitation:{reference:'INVITATION_E',personId:target,processRef:'PROCESS_E',cohortRef:'COHORT_E'},acceptance:{reference:'ACCEPTANCE_E',personId:target,invitationRef:'INVITATION_E'},integration:patchOverride?.integrationRequired?{required:true,ruleRef:'RULE_E',completionRef:null}:{required:false,ruleRef:null,completionRef:null},ingress:null};
 }
 if(governance){
  const functionId=patchOverride?.functionId||'FN_E',council=patchOverride?.council===true;
  const add=(col,key,fields)=>x.d[col+'/'+key]=fixture.entity(col,key,fields);
  add('memberGovernanceBoards','DIR_E',{name:'Synthetic board',startDate:'2026-01-01',endDate:'2026-12-31'});add('memberGovernanceTerms','TERM_E',{name:'Synthetic term',startDate:'2026-01-01',endDate:'2026-12-31'});add('memberGovernanceBoardTerms','REL_E',{idDiretoria:'DIR_E',idTermo:'TERM_E',startDate:'2026-01-01',endDate:'2026-12-31'});add('memberRoleCatalog','ROLE_E',{roleKey:'DIRETOR_ADMINISTRATIVO',group:'DIRETORIA',active:true});
  x.d['memberBoardHeads/DIR_E']={schemaVersion:'m11-head-v1',id:'DIR_E',kind:'BOARD',revision:1,paths:['memberGovernanceBoards/DIR_E','memberGovernanceBoardTerms/REL_E']};
  if(council){x.d['memberRoleCatalog/ROLE_E'].roleKey='CONSELHEIRO_CONSULTIVO';x.d['memberRoleCatalog/ROLE_E'].group='CONSELHO';}
  const fields={idVinculo:'VINC-002',idCargo:'ROLE_E',scopeKey:council?null:'ADMINISTRACAO',idDiretoria:'DIR_E',idTermo:'TERM_E',startDate:'2026-01-01',endDatePlanned:'2026-12-31'},fp='memberFunctions/'+functionId;
  if(family==='H'){add('memberFunctions','FN_E',{idPessoa:target,...fields,endDateReal:null});x.d['memberAggregateHeads/'+target].paths.push(fp);}
  r.target={personId:target,functionId};r.payload=family==='G'?fields:{endDateReal:'2026-09-14'};r.expectedRevisions=[{resourceType:'memberAggregateHeads',id:target,revision:1},{resourceType:'memberFunctions',id:functionId,revision:family==='G'?0:1}];if(family==='H')r.expectedRevisions.push({resourceType:'memberEvents',id:'EVT_'+id,revision:0});
  const deps=['memberLinks/VINC-002','memberRoleCatalog/ROLE_E','memberGovernanceBoards/DIR_E','memberGovernanceTerms/TERM_E','memberBoardHeads/DIR_E','memberGovernanceBoardTerms/REL_E'];facts={dependencies:deps.map(path=>({path,revision:1}))};
  x.req.target.boardId='DIR_E';x.d[x.grantPath].targets.allowedBoards=['DIR_E'];x.req.target.resources=[fp,...(family==='H'?['memberEvents/EVT_'+id]:[]),...deps];x.req.target.fieldMask=[];x.req.target.linkId='VINC-002';x.req.target.effectiveDate=family==='G'?fields.startDate:r.payload.endDateReal;x.req.target.effect=family==='G'?'NOMEACAO':'TERMINO';x.d[x.grantPath].targets.allowedResources=x.req.target.resources;x.d[x.grantPath].fieldMasks[cap]=[];
  const att=x.d['memberEvidenceAttestations/SYN_EVIDENCE'];att.linkId='VINC-002';att.effect=x.req.target.effect;att.effectiveDate=x.req.target.effectiveDate;
  if(council){x.req.target.scopeKey=null;x.req.target.roleKey='CONSELHEIRO_CONSULTIVO';x.d[x.grantPath].targets.allowedTargetScopes=[];x.d[x.grantPath].targets.allowedRoles=['CONSELHEIRO_CONSULTIVO'];}
 }
 if(binding){const uid='TARGET_UID',path='memberAuthBindings/'+uid;
  r.target={personId:target,uid};r.payload=family==='J'?{}:{reason:'Synthetic revocation'};r.evidenceRefs=['SYN_EVIDENCE'];r.decisionRef=null;r.expectedRevisions=[{resourceType:'memberAggregateHeads',id:target,revision:1},{resourceType:'memberAuthBindings',id:uid,revision:family==='J'?0:1}];
  if(family==='K1')x.d[path]={schemaVersion:'binding-v1',id:uid,uid,idPessoa:target,revision:1,origin:'SYNTHETIC',active:true,bindingVersion:1,policyVersion:1,sourceRevision:1,createdAt:x.now-1000,updatedAt:x.now-1000,domainVersion:'m11'};
  x.req.target.resources=[path,'portalUsers/'+uid,'portalAccessStatus/'+uid];x.req.target.fieldMask=[];x.d[x.grantPath].targets.allowedResources=x.req.target.resources;x.d[x.grantPath].fieldMasks[cap]=[];
  x.d['memberEvidenceAttestations/SYN_EVIDENCE']={...m13BFixture_('EXPLICIT','MEMBERS_PROFILE_CORRECT').d['memberEvidenceAttestations/SYN_EVIDENCE'],holderId:target};
  facts=family==='J'?{schemaVersion:'member-association-content-v13',uid,idPessoa:target,projectId:fixture.config.projectId,verifiedAt:x.now-1000,expiresAt:x.now+60000,verificationRef:'AUTH_VERIFIED_SYNTHETIC'}:{};
 }
 if(family==='I'){
  const action=patchOverride?.action||'CREATE_BOARD',create=action.startsWith('CREATE_'),board='BOARD_E',term=action.endsWith('BOARD')?null:'TERM_E',relation=action.endsWith('RELATION')?'REL_E':null;
  r.target={boardId:board,termId:term,relationId:relation,scopeKey:'ADMINISTRACAO'};r.payload=create?{action,...(!relation?{name:'Synthetic governance'}:{}),startDate:'2026-01-01',endDate:'2026-12-31'}:{action,endDate:'2026-09-14'};
  const col=action.endsWith('BOARD')?'memberGovernanceBoards':action.endsWith('TERM')?'memberGovernanceTerms':'memberGovernanceBoardTerms',key=relation||term||board,path=col+'/'+key,headPath='memberBoardHeads/'+board;
  const resources=new Set([path,headPath]);
  if(action!=='CREATE_BOARD'){
   x.d['memberGovernanceBoards/'+board]=fixture.entity('memberGovernanceBoards',board,{name:'Synthetic board',startDate:'2026-01-01',endDate:'2026-12-31'});x.d[headPath]={schemaVersion:'m11-head-v1',id:board,kind:'BOARD',revision:1,paths:['memberGovernanceBoards/'+board]};resources.add('memberGovernanceBoards/'+board);
   const inv={schemaVersion:'member-governance-inventory-v13',boardId:board,boardIds:[board],revision:1,readAuthzEpoch:1,dependents:[]};x.d['memberGovernanceInventories/'+board]=inv;resources.add('memberGovernanceInventories/'+board);facts={inventoryHash:C.hash(inv)};
  }else facts={inventoryHash:null};
  if(term&&action!=='CREATE_TERM'){x.d['memberGovernanceTerms/'+term]=fixture.entity('memberGovernanceTerms',term,{name:'Synthetic term',startDate:'2026-01-01',endDate:'2026-12-31'});resources.add('memberGovernanceTerms/'+term);}
  if(action==='END_RELATION'){x.d[path]=fixture.entity(col,key,{idDiretoria:board,idTermo:term,startDate:'2026-01-01',endDate:'2026-12-31'});x.d[headPath].paths.push(path);}
  if(patchOverride?.dependent){
   C.need(action==='END_BOARD','SCHEMA_INVALID');
   const add=(col,key,fields)=>x.d[col+'/'+key]=fixture.entity(col,key,fields);
   add('memberGovernanceTerms','TERM_DEP',{name:'Synthetic dependent term',startDate:'2026-01-01',endDate:'2026-12-31'});
   add('memberGovernanceBoardTerms','REL_DEP',{idDiretoria:board,idTermo:'TERM_DEP',startDate:'2026-01-01',endDate:'2026-09-14'});
   add('memberRoleCatalog','ROLE_DEP',{roleKey:'DIRETOR_ADMINISTRATIVO',group:'DIRETORIA',active:true});
   add('memberFunctions','FN_DEP',{idPessoa:target,idVinculo:'VINC-002',idCargo:'ROLE_DEP',scopeKey:'ADMINISTRACAO',idDiretoria:board,idTermo:'TERM_DEP',startDate:'2026-01-01',endDatePlanned:'2026-12-31',endDateReal:'2026-09-14'});
   x.d[headPath].paths.push('memberGovernanceBoardTerms/REL_DEP');x.d['memberAggregateHeads/'+target].paths.push('memberFunctions/FN_DEP');
   resources.add('memberGovernanceBoardTerms/REL_DEP');resources.add('memberGovernanceTerms/TERM_DEP');
   const inv=x.d['memberGovernanceInventories/'+board];inv.dependents=[{personId:target,personRevision:1,functionIds:['FN_DEP']}];facts.inventoryHash=C.hash(inv);
  }
  r.expectedRevisions=[{resourceType:'memberBoardHeads',id:board,revision:action==='CREATE_BOARD'?0:1},{resourceType:col,id:key,revision:create?0:1}];
  x.req.schemaVersion='member-authz-governance-request-v13';x.req.target={...r.target,resources:[...resources].sort(),caseId:'CASE_GOV',purpose:'GOVERNANCE',effect:action,effectiveDate:create?r.payload.startDate:r.payload.endDate};
  Object.assign(x.d[x.grantPath].targets,{allowedBoards:[board],allowedRoles:[],allowedPersons:[],allowedResources:x.req.target.resources});x.d[x.grantPath].fieldMasks[cap]=[];
  x.d['memberEvidenceAttestations/SYN_EVIDENCE']={schemaVersion:'member-governance-attestation-v13',id:'SYN_EVIDENCE',revision:1,active:true,kind:'BOARD_DECISION',boardId:board,termId:term,relationId:relation,targetResourceIds:x.req.target.resources,effect:action,effectiveDate:x.req.target.effectiveDate,issuerRef:'SYN_ISSUER',approvalRef:'SYN_APPROVAL',contentHash:'a'.repeat(64),verifiedAt:T(x.now-1000),expiresAt:T(x.now+3600000)};
 }
 r.expectedRevisions.sort((a,b)=>(a.resourceType+'/'+a.id)<(b.resourceType+'/'+b.id)?-1:1);
 if(family!=='A'){x.d['memberEvidenceAttestations/SYN_EVIDENCE'].targetResourceIds=x.req.target.resources;const body={schemaVersion:'member-act-content-v13',...f4MemberActContent_().material(r),facts};const hash=f4MemberActContent_().digest(body);x.d['memberEvidenceContents/'+hash]=body;x.d['memberEvidenceAttestations/SYN_EVIDENCE'].contentHash=hash;}
 r.digest=K.digest(fixture.config.projectId,r);const proof=(family==='I'?f4MemberAuthz_().authorizeGovernance(x.get,x.req,r,x.now):family==='C'?f4MemberAuthz_().authorizeCommand(x.get,x.req,r,x.now):f4MemberAuthz_().authorize(x.get,x.req,x.now)).proof;
 x.d['commandRequests/'+id]=r;x.d['commandProvenance/'+id]={schemaVersion:'member-command-provenance-v13',operationId:id,ownerUid:r.ownerUid,ticketId:r.ticketId,digest:r.digest,authRequests:[x.req]};
 x.d['memberAdmissionTickets/'+r.ticketId]={schemaVersion:'member-admission-ticket-v13',id:r.ticketId,operationId:id,ownerUid:r.ownerUid,idPessoa:'PES-000001',requestEnvelope:K.template(r),digest:r.digest,authProofs:[proof],issuedAt:T(x.now),expiresAt:T(x.now+120000),tokenExpiresAt:T(x.now+3600000),authorityEpoch:1,status:'ISSUED'};
 x.d['memberAdmissionUses/'+r.ticketId]={schemaVersion:'member-admission-use-v13',ticketId:r.ticketId,operationId:id,ownerUid:r.ownerUid,reservedAt:r.receivedAt};
 x.d['commandLeases/'+id]={schemaVersion:'member-command-lease-v13',operationId:id,domain:'MEMBERS',workerId:'E_WORKER',fencingToken:1,deadline:T(x.now+120000)};
 const read=m13BFixture_('SELF','MEMBERS_SELF_READ');read.d['memberAuthzPolicies/SYN_POLICY'].id='E_READ';read.d['memberAuthzPolicies/SYN_POLICY'].targets.allowedResources=['memberCurrentStates/PES-000001'];
 x.d['memberAuthzPolicies/E_READ']=read.d['memberAuthzPolicies/SYN_POLICY'];x.d['memberAuthorityControl/MEMBERS'].selfReadPolicy={...read.d['memberAuthorityControl/MEMBERS'].selfReadPolicy,policyId:'E_READ'};
 x.key={operationId:id,digest:r.digest,commandType:type};x.handle={workerId:'E_WORKER',fencingToken:1};x.request=r;
 x.make=()=>{const producer=family==='I'?f4MemberBoardCommands_():family==='C'?f4MemberIngressCommands_():membership?f4MemberMembershipCommands_():governance?f4MemberGovernanceCommands_():binding?f4MemberBindingCommands_(fixture.config):f4MemberProfileCommands_();return f4Commands_().createMembersLocal({config:fixture.config,clock:()=>x.now,verifyIdentity:token=>{C.need(token==='E_SYNTHETIC_ONLY','AUTH_BINDING_INVALID');return {uid:'SYN_UID',expiresAt:x.now+3600000}},producers:{[type]:{validate:producer.validate,resolve:producer.resolve}}});};
 return x;
}

// Synthetic second admission after a real local C commit; retains its ledger/audit/evidence chain.
function m13EContinueIngress_(x,encoded,operation,integrationRequired=false){
 const C=f4Core_(),F=f4Firestore_(),K=f4MemberCommandCatalog_(),T=F.Timestamp.fromMillis;
 for(const key of Object.keys(x.d))delete x.d[key];for(const [path,fields]of Object.entries(encoded))x.d[path]=F.decode({mapValue:{fields}});
 const old=x.d['commandRequests/'+'e'.repeat(32)],oldAtt=x.d['memberEvidenceAttestations/SYN_EVIDENCE'],oldBody=x.d['memberEvidenceContents/'+oldAtt.contentHash],id='f'.repeat(32),r=F.decode(F.encode(old));
 r.operationId=id;r.ticketId='E_TICKET_SECOND';r.target.linkId=operation==='RECORD_EFFECTIVATION'?'VINC_EFFECTIVE':'VINC-002';r.payload={operation,identity:null,identifiers:[],startDate:operation==='RECORD_EFFECTIVATION'?'2026-09-15':'2026-09-13',endDate:null,originalEventId:operation==='RETIFICACAO'?'EVT_'+old.operationId:null};r.evidenceRefs=['EVIDENCE_SECOND'];
 const changes=[['memberLinks','VINC-002',1],['memberEvents','COR_'+id,0],['memberEvents','EVT_'+id,0]];if(operation==='RECORD_EFFECTIVATION')changes.push(['memberLinks','VINC_EFFECTIVE',0]);r.expectedRevisions=[{resourceType:'memberAggregateHeads',id:r.target.personId,revision:1},...changes.map(([resourceType,id,revision])=>({resourceType,id,revision}))].sort((a,b)=>(a.resourceType+'/'+a.id)<(b.resourceType+'/'+b.id)?-1:1);
 const q=F.decode(F.encode(x.req));q.target.linkId=r.target.linkId;q.target.effect=operation==='RECORD_EFFECTIVATION'?'EFETIVACAO':'INGRESSO';q.target.effectiveDate=r.payload.startDate;q.target.resources=changes.map(([col,id])=>col+'/'+id);q.evidenceRefs=r.evidenceRefs;
 x.d[x.grantPath].targets.allowedResources=q.target.resources;
 const facts={...F.decode(F.encode(oldBody.facts)),integration:integrationRequired?{required:true,ruleRef:'RULE_E',completionRef:'COMPLETION_E'}:{required:false,ruleRef:null,completionRef:null},ingress:{eventId:'EVT_'+old.operationId,operationId:old.operationId,contentHash:oldAtt.contentHash}};
 const body={schemaVersion:'member-act-content-v13',...f4MemberActContent_().material(r),facts},hash=f4MemberActContent_().digest(body);x.d['memberEvidenceContents/'+hash]=body;x.d['memberEvidenceAttestations/EVIDENCE_SECOND']={...oldAtt,id:'EVIDENCE_SECOND',linkId:r.target.linkId,effect:q.target.effect,effectiveDate:r.payload.startDate,targetResourceIds:q.target.resources,contentHash:hash};r.digest=K.digest(fixture.config.projectId,r);
 const proof=f4MemberAuthz_().authorizeCommand(x.get,q,r,x.now).proof;
 x.d['commandRequests/'+id]=r;x.d['commandProvenance/'+id]={...x.d['commandProvenance/'+old.operationId],operationId:id,ticketId:r.ticketId,digest:r.digest,authRequests:[q]};x.d['memberAdmissionTickets/'+r.ticketId]={...x.d['memberAdmissionTickets/'+old.ticketId],id:r.ticketId,operationId:id,requestEnvelope:K.template(r),digest:r.digest,authProofs:[proof]};x.d['memberAdmissionUses/'+r.ticketId]={schemaVersion:'member-admission-use-v13',ticketId:r.ticketId,operationId:id,ownerUid:r.ownerUid,reservedAt:r.receivedAt};x.d['commandLeases/'+id]={schemaVersion:'member-command-lease-v13',operationId:id,domain:'MEMBERS',workerId:'E_WORKER',fencingToken:1,deadline:T(x.now+120000)};
 x.request=r;x.req=q;x.key={operationId:id,digest:r.digest,commandType:r.commandType};return x.d;
}
