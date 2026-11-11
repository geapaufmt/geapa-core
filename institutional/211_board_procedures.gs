// Typed institutional procedures. No elections, generic effects or implicit grants.
// The governance executor owns admission, CAS, journals, receipts and audit.
function f4BoardProcedures_(){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),copy=v=>F.decode(F.encode(v)),hash=v=>C.hash(F.encode(v));
 const DAY=86400000,types={
  TEMPORARY_OFFICE:{label:'Criar cargo diretivo temporário',rule:'UNANIMITY',basis:'Art. 30',president:true},
  DIRECTOR_APPOINTMENT:{label:'Nomear diretor não eletivo',rule:'ADMINISTRATIVE',basis:'Art. 35, §6º',president:true},
  ADVISOR_APPOINTMENT:{label:'Nomear assessor',rule:'ADMINISTRATIVE',basis:'Art. 31, §6º',president:false},
  ORDINARY_EXTENSION:{label:'Prorrogação ordinária',rule:'CONSENT',basis:'Art. 33',president:true},
  EXTRAORDINARY_EXTENSION:{label:'Prorrogação extraordinária',rule:'ABSOLUTE_MAJORITY',basis:'Art. 33, §5º; exceção específica necessária',president:true},
  VOLUNTARY_EXIT:{label:'Desligamento voluntário da função',rule:'NOTICE',basis:'Art. 34, I',president:false},
  DIRECTOR_DISMISSAL:{label:'Destituição de diretor não eletivo',rule:'ABSOLUTE_MAJORITY',basis:'Art. 34, §3º',president:true},
  ADVISOR_END:{label:'Encerramento não disciplinar de assessoria',rule:'ADMINISTRATIVE',basis:'Art. 31, §6º',president:false}
 };
 const caps={BP_ACCESS:'INSTITUTIONAL_BOARD_MANIFEST',BP_ACK:'INSTITUTIONAL_BOARD_MANIFEST',BP_NOTICE:'INSTITUTIONAL_BOARD_PROPOSE',BP_ENACT:'INSTITUTIONAL_BOARD_ENACT',BP_CREATE:'INSTITUTIONAL_BOARD_PROPOSE',BP_OPEN:'INSTITUTIONAL_BOARD_PROPOSE',BP_CANCEL:'INSTITUTIONAL_BOARD_PROPOSE',BP_VOTE:'INSTITUTIONAL_BOARD_MANIFEST',BP_CONSENT:'INSTITUTIONAL_BOARD_MANIFEST',BP_DEFEND:'INSTITUTIONAL_BOARD_MANIFEST',BP_CLOSE:'INSTITUTIONAL_BOARD_PROPOSE',BP_CLEARANCE:'INSTITUTIONAL_BOARD_PROPOSE'};
 const text=(v,min,max)=>typeof v==='string'&&v===v.trim()&&v.length>=min&&v.length<=max&&!/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(v);
 const person=auth=>auth.actor.actor.h.id;
 const directors=auth=>auth.actor.derived.currentDirectors;
 const president=(auth,board)=>directors(auth).some(f=>f.idDiretoria===board&&f.roleKey==='PRESIDENTE');
 const collective=p=>['UNANIMITY','ABSOLUTE_MAJORITY'].includes(types[p.type].rule);
 function enabled(c,action,board,type){N(c.boardProcedures?.enabled===true&&c.boardProcedures.actions?.includes(action)&&c.boardProcedures.boardIds?.includes(board)&&c.boardProcedures.types?.includes(type),'PILOT_DISABLED');}
 function scope(auth,board){N(directors(auth).some(f=>f.idDiretoria===board),'AUTHORIZATION_DENIED');}
 function composition(get,o,boardId){
  // Civil dates in the canonical graph are inclusive. Same-day midnight/expiry
  // semantics remain in Member derive; this is not a second eligibility model.
  const date=new Date(o.clock()-4*3600000).toISOString().slice(0,10);
  const raw=f4InstitutionalMinutes_().composition(get,o.config,boardId,date);
  return {boardId:raw.boardId,boardRevision:raw.boardRevision,people:raw.people.map(p=>({personId:p.personId,name:p.name,functions:p.functions.map(f=>{
   const v=get('memberFunctions/'+f.id);N(v&&v.idPessoa===p.personId&&v.idDiretoria===boardId,'COMPOSITION_MISMATCH');
   return {...f,linkId:v.idVinculo,startDate:v.startDate,endDate:v.endDateReal||v.endDatePlanned};
  }).sort((a,b)=>a.id.localeCompare(b.id)),criterion:'CURRENT_CANONICAL_DIRECTOR'})).sort((a,b)=>a.personId.localeCompare(b.personId))};
 }
 function target(get,o,type,boardId,value){
  const fields={TEMPORARY_OFFICE:['personId','linkId','roleName','purpose','duties','startDate','endDate','advisoryAllowed'],DIRECTOR_APPOINTMENT:['personId','linkId','roleId','startDate','endDate'],ADVISOR_APPOINTMENT:['personId','linkId','responsibleFunctionId','purpose','duties','startDate','endDate'],ORDINARY_EXTENSION:['personId','functionId','endDate'],EXTRAORDINARY_EXTENSION:['personId','functionId','endDate'],VOLUNTARY_EXIT:['personId','functionId','endDate'],DIRECTOR_DISMISSAL:['personId','functionId','grounds','defenseDeadline'],ADVISOR_END:['personId','functionId','endDate','reason']};
  C.exact(value,fields[type]);N(C.id(value.personId),'COMMAND_PAYLOAD_INVALID');
  const authority=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),loaded=f4MemberAuthz_().load(get,value.personId,authority),d=loaded.M.derive(loaded.d,value.personId,loaded.h.revision,o.clock()),board=get('memberGovernanceBoards/'+boardId);
  N(board&&C.civilStart(board.startDate)<=o.clock()&&o.clock()<C.civilStart(C.nextDay(board.endDate)),'BOARD_NOT_CURRENT');
  const fn=value.functionId?loaded.d['memberFunctions/'+value.functionId]:null;
  if(value.functionId){N(fn&&fn.idPessoa===value.personId&&fn.idDiretoria===boardId,'RESOURCE_NOT_FOUND');const role=get('memberRoleCatalog/'+fn.idCargo);N(role,'RESOURCE_NOT_FOUND');
   if(type==='ADVISOR_END')N(role.group==='ASSESSORIA'&&d.currentFunctions.some(f=>f.idVigencia===fn.id),'TARGET_FUNCTION_INVALID');
   else N(d.currentDirectors.some(f=>f.idVigencia===fn.id),'TARGET_FUNCTION_INVALID');
   if(type==='DIRECTOR_DISMISSAL'||type.endsWith('EXTENSION'))N(!['PRESIDENTE','VICE_PRESIDENTE'].includes(role.roleKey),'ELECTIVE_ROLE_PROTECTED');
  }
  if(value.linkId){const link=loaded.d['memberLinks/'+value.linkId];N(link&&link.idPessoa===value.personId&&link.type==='MEMBRO_EFETIVO'&&d.state.state==='MEMBRO_EFETIVO'&&!d.state.blocked,'MEMBER_NOT_ELIGIBLE');N(C.civilStart(link.startDate)<=o.clock()&&(!link.endDate||o.clock()<C.civilStart(C.nextDay(link.endDate))),'MEMBER_NOT_ELIGIBLE');}
  if(value.roleId){const role=get('memberRoleCatalog/'+value.roleId);N(role?.active&&role.group==='DIRETORIA'&&!['PRESIDENTE','VICE_PRESIDENTE'].includes(role.roleKey),'ELECTIVE_ROLE_PROTECTED');}
  if(value.startDate){loaded.M.interval(value.startDate,value.endDate);N(C.civilStart(value.startDate)>=C.civilStart(board.startDate)&&value.endDate<=board.endDate,'TERM_OUTSIDE_BOARD');}
  if(value.endDate){C.civilStart(value.endDate);N(value.endDate<=board.endDate,'TERM_OUTSIDE_BOARD');}
  if(type.includes('EXTENSION'))N(value.endDate>(fn.endDateReal||fn.endDatePlanned),'EXTENSION_INVALID');
  if(type==='VOLUNTARY_EXIT')N(C.civilStart(value.endDate)>=o.clock()+7*DAY,'NOTICE_PERIOD_REQUIRED');
  if(type==='TEMPORARY_OFFICE')N(text(value.roleName,3,120)&&text(value.purpose,10,2000)&&text(value.duties,10,4000)&&typeof value.advisoryAllowed==='boolean','COMMAND_PAYLOAD_INVALID');
  if(type==='ADVISOR_APPOINTMENT'){N(text(value.purpose,10,2000)&&text(value.duties,10,4000),'COMMAND_PAYLOAD_INVALID');const responsible=get('memberFunctions/'+value.responsibleFunctionId);N(responsible&&composition(get,o,boardId).people.some(p=>p.functions.some(f=>f.id===responsible.id)),'RESPONSIBLE_DIRECTOR_INVALID');}
  if(type==='DIRECTOR_DISMISSAL')N(text(value.grounds,20,4000)&&Number.isSafeInteger(value.defenseDeadline)&&value.defenseDeadline>0,'DEFENSE_PERIOD_REQUIRED');
  if(type==='ADVISOR_END')N(['VOLUNTARY_REQUEST','ASSIGNMENT_COMPLETED','RESPONSIBILITY_ENDED'].includes(value.reason),'COMMAND_PAYLOAD_INVALID');
  return {personId:value.personId,personRevision:loaded.h.revision,functionRevision:fn?.revision||null,boardRevision:board.revision};
 }
 // Process availability, authenticated access, receipt and defense are distinct facts.
 function noticeReady(p,now){
  if(p.defense?.personId===p.target.personId&&p.defense.contentHash===p.contentHash)return true;
  const n=p.notice;
  return !!(n?.receivedAt&&n.contentHash===p.contentHash&&n.recipientPersonId===p.target.personId&&n.defenseUntil>=n.receivedAt+7*DAY&&now>=n.defenseUntil&&(n.acknowledgment||n.externalEvidence));
 }
 function validateDocuments(get,o,v,auth,refs){
  N(Array.isArray(refs)&&refs.length>0&&refs.length<=10,'DOCUMENT_FILE_SCOPE_INVALID');
  for(const ref of refs){C.exact(ref,['recordId','revision','sha256']);const r=v.records.find(r=>r.id===ref.recordId);N(r&&r.revision===ref.revision&&r.attachments?.some(f=>f.sha256===ref.sha256)&&f4InstitutionalDocumentManagement_().readable(get,o,r,auth),'DOCUMENT_FILE_SCOPE_INVALID');}
 }
 function content(p){return {type:p.type,title:p.title,object:p.object,basis:p.basis,reason:p.reason,boardId:p.boardId,target:p.target,documents:p.documents,contentVersion:p.contentVersion};}
 function outcome(p,now){
  if(p.state!=='OPEN')return p.state;
  const n=p.electorate.people.length,yes=p.votes.filter(v=>v.choice==='YES').length,no=p.votes.filter(v=>v.choice==='NO').length,abstain=p.votes.filter(v=>v.choice==='ABSTAIN').length;
  if(now<p.deadline&&!(p.allowEarlyClose&&p.votes.length===n))return 'OPEN';
  const required=p.rule==='UNANIMITY'?n:Math.floor(n/2)+1;
  if(yes>=required)return 'APPROVED';
  if(p.rule==='UNANIMITY'&&(no||abstain)||no+abstain>n-required)return 'REJECTED';
  return 'INSUFFICIENT_MANIFESTATIONS';
 }
 function effective(p,current,now){if(p.state==='OPEN'&&hash(p.electorate)!==hash(current))return 'SUSPENDED_COMPOSITION';return outcome(p,now);}
 function readiness(p){
  if(p.enactment?.state==='ENACTED')return copy(p.enactment);const blocks=[];if(!['APPROVED','RECORDED'].includes(p.state))blocks.push('PROCEDURE_NOT_APPROVED');
  if(collective(p))blocks.push('ELECTRONIC_PROCEDURE_NORMATIVE_ENABLEMENT_REQUIRED');
  if(p.type==='EXTRAORDINARY_EXTENSION')blocks.push('ART33_SPECIFIC_EXCEPTION_REQUIRED');
  if(p.type==='TEMPORARY_OFFICE')blocks.push('COMMUNICATION_TO_EFFECTIVE_MEMBERS_REQUIRED');
  if(['TEMPORARY_OFFICE','DIRECTOR_APPOINTMENT','ORDINARY_EXTENSION','EXTRAORDINARY_EXTENSION'].includes(p.type))blocks.push('CANONICAL_TENURE_REVALIDATION_REQUIRED');
  blocks.push('CANONICAL_EFFECT_PREVIEW_REQUIRED');
  return {state:'NOT_ENACTED',blocks};
 }
 function transition(get,o,input,v,auth,c){
  N(caps[input.action]&&v.revision===input.expectedRevision&&text(input.reason,10,1500),'REVISION_CONFLICT');
  const old=input.id?v.records.find(r=>r.id===input.id):null,now=o.clock(),actor=person(auth);
  if(input.action==='BP_CREATE'){
   C.exact(input,['action','id','expectedRevision','expectedRecordRevision','reason','proposal']);N(input.id===''&&input.expectedRecordRevision===0,'REVISION_CONFLICT');const p=input.proposal;
   C.exact(p,['type','title','object','basis','boardId','target','documents','days']);N(Object.hasOwn(types,p.type)&&text(p.title,5,240)&&text(p.object,10,4000)&&text(p.basis,5,1500)&&Number.isSafeInteger(p.days)&&p.days>=1&&p.days<=30,'COMMAND_PAYLOAD_INVALID');
   scope(auth,p.boardId);enabled(c,input.action,p.boardId,p.type);N(c.boardProcedures.targetPersonIds?.includes(p.target.personId),'SCOPE_INVALID');if(types[p.type].president)N(president(auth,p.boardId),'PRESIDENT_REQUIRED');
   const allowedDays=c.boardProcedures.votingDaysByType?.[p.type]||[7];N(Array.isArray(allowedDays)&&allowedDays.length>0&&allowedDays.every(n=>Number.isSafeInteger(n)&&n>=1&&n<=30)&&allowedDays.includes(p.days),'COMMAND_PAYLOAD_INVALID');
   if(p.type==='ADVISOR_APPOINTMENT')N(president(auth,p.boardId)||get('memberFunctions/'+p.target.responsibleFunctionId)?.idPessoa===actor,'AUTHORIZATION_DENIED');
   if(p.type==='VOLUNTARY_EXIT')N(p.target.personId===actor,'SELF_REQUIRED');
   if(p.type==='DIRECTOR_DISMISSAL')N(p.target.defenseDeadline>=now+7*DAY,'DEFENSE_PERIOD_REQUIRED');
   const targetSnapshot=target(get,o,p.type,p.boardId,p.target);N(Array.isArray(p.documents)&&p.documents.length<=10,'COMMAND_PAYLOAD_INVALID');
   for(const ref of p.documents){C.exact(ref,['recordId','revision','sha256']);const r=v.records.find(r=>r.id===ref.recordId);N(r&&r.revision===ref.revision&&r.attachments?.some(f=>f.sha256===ref.sha256)&&f4InstitutionalDocumentManagement_().readable(get,o,r,auth),'DOCUMENT_FILE_SCOPE_INVALID');}
   const state=p.type==='DIRECTOR_DISMISSAL'?'DEFENSE_PENDING':['ORDINARY_EXTENSION','EXTRAORDINARY_EXTENSION'].includes(p.type)?'CONSENT_PENDING':collective({type:p.type})?'DRAFT':'RECORDED';
   const procedure={...copy(p),schemaVersion:'board-procedure-v1',reason:input.reason,proposer:actor,createdAt:now,state,contentVersion:1,targetSnapshot,electorate:null,rule:types[p.type].rule,openedAt:null,deadline:null,closedAt:null,votes:[],consent:null,defense:null,clearance:null,allowEarlyClose:false,enactment:{state:'NOT_ENACTED'}};
   procedure.contentHash=hash(content(procedure));
   if(p.type==='DIRECTOR_DISMISSAL')procedure.notice={availableAt:now,recipientPersonId:p.target.personId,contentHash:procedure.contentHash,access:null,acknowledgment:null,externalEvidence:null,receivedAt:null,defenseUntil:p.target.defenseDeadline};
   return {old:null,next:{kind:'BOARD_PROCEDURE',title:p.title,meetingDate:null,body:p.object,documentaryState:state,publication:'CLOSED',formalApproval:null,effectiveFrom:null,sourceRecordId:null,references:[],createdAt:now,procedure}};
  }
  N(old?.kind==='BOARD_PROCEDURE'&&old.revision===input.expectedRecordRevision,'REVISION_CONFLICT');const next=copy(old),p=next.procedure;scope(auth,p.boardId);enabled(c,input.action,p.boardId,p.type);
  const base=['action','id','expectedRevision','expectedRecordRevision','reason'];
  C.exact(input,[...base,...(input.action==='BP_VOTE'?['choice','contentHash']:input.action==='BP_CONSENT'?['accepted','contentHash']:input.action==='BP_DEFEND'?['statement','contentHash']:input.action==='BP_CLEARANCE'?['pendingDuties']:input.action==='BP_ACCESS'||input.action==='BP_ACK'?['contentHash']:input.action==='BP_NOTICE'?['contentHash','receivedAt','documents']:[])]);
  N(p.enactment?.state!=='ENACTED','PROCEDURE_IMMUTABLE');
  if(input.action==='BP_ENACT'){
   const effect=f4MemberBoardEffects_().resolve(get,o,old,auth,c);
   p.enactment={documentaryAuthorization:effect.documentaryAuthorization,state:'ENACTED',at:now,actorPersonId:actor,operationId:o.operationId,functionId:effect.functionId,evidence:effect.evidence,tenure:effect.summary,responsibleFunctionId:effect.responsibleFunctionId};
  }else if(input.action==='BP_CANCEL'){N(actor===p.proposer||president(auth,p.boardId),'AUTHORIZATION_DENIED');N(!['APPROVED','REJECTED','INSUFFICIENT_MANIFESTATIONS','CANCELLED'].includes(outcome(p,now)),'PROCEDURE_IMMUTABLE');p.state='CANCELLED';p.cancelled={at:now,personId:actor,reason:input.reason};}
  else if(input.action==='BP_CONSENT'){N(p.state==='CONSENT_PENDING'&&p.target.personId===actor,'SELF_REQUIRED');N(input.contentHash===p.contentHash&&typeof input.accepted==='boolean','CONTENT_VERSION_CONFLICT');p.consent={personId:actor,at:now,accepted:input.accepted,contentHash:p.contentHash,operationId:o.operationId};p.state=input.accepted?(collective(p)?'DRAFT':'RECORDED'):'REJECTED';}
  else if(['BP_ACCESS','BP_ACK','BP_NOTICE'].includes(input.action)){
   N(p.type==='DIRECTOR_DISMISSAL'&&p.state==='DEFENSE_PENDING','PROCEDURE_STATE_INVALID');N(input.contentHash===p.contentHash,'CONTENT_VERSION_CONFLICT');
   p.notice=p.notice||{availableAt:p.createdAt,recipientPersonId:p.target.personId,contentHash:p.contentHash,access:null,acknowledgment:null,externalEvidence:null,receivedAt:null,defenseUntil:p.target.defenseDeadline};
   const n=p.notice;N(n.contentHash===p.contentHash&&n.recipientPersonId===p.target.personId,'CONTENT_VERSION_CONFLICT');
   if(input.action==='BP_ACCESS'){N(actor===p.target.personId,'SELF_REQUIRED');N(!n.access,'DUPLICATE_MANIFESTATION');n.access={personId:actor,at:now,operationId:o.operationId,meaning:'AUTHENTICATED_ACCESS_NOT_RECEIPT_OR_AGREEMENT'};}
   else if(input.action==='BP_ACK'){N(actor===p.target.personId,'SELF_REQUIRED');N(!n.acknowledgment,'DUPLICATE_MANIFESTATION');n.acknowledgment={personId:actor,at:now,operationId:o.operationId,meaning:'RECEIPT_NOT_AGREEMENT'};n.receivedAt=n.receivedAt||now;n.defenseUntil=Math.max(n.defenseUntil,now+7*DAY);}
   else{N(president(auth,p.boardId),'PRESIDENT_REQUIRED');N(!n.externalEvidence,'DUPLICATE_MANIFESTATION');N(Number.isSafeInteger(input.receivedAt)&&input.receivedAt>=p.createdAt&&input.receivedAt<=now,'NOTICE_EVIDENCE_INVALID');validateDocuments(get,o,v,auth,input.documents);n.externalEvidence={recordedBy:actor,recordedAt:now,receivedAt:input.receivedAt,documents:copy(input.documents),operationId:o.operationId,verification:'PENDING_INSTITUTIONAL_BASIS'};n.receivedAt=n.receivedAt||input.receivedAt;n.defenseUntil=Math.max(n.defenseUntil,input.receivedAt+7*DAY);}
  }
  else if(input.action==='BP_DEFEND'){N(p.state==='DEFENSE_PENDING'&&p.target.personId===actor&&now<(p.notice?.defenseUntil||p.target.defenseDeadline),'SELF_REQUIRED');N(!p.defense&&input.contentHash===p.contentHash&&text(input.statement,1,8000),'DEFENSE_INVALID');p.defense={personId:actor,at:now,statement:input.statement,contentHash:p.contentHash,operationId:o.operationId};}
  else if(input.action==='BP_CLEARANCE'){N(p.type==='VOLUNTARY_EXIT'&&president(auth,p.boardId),'PRESIDENT_REQUIRED');N(p.state==='RECORDED'&&text(input.pendingDuties,0,4000),'COMMAND_PAYLOAD_INVALID');p.clearance={at:now,personId:actor,pendingDuties:input.pendingDuties,meaning:'DUTIES_RECORD_NOT_RESIGNATION_VETO'};}
  else if(input.action==='BP_OPEN'){
   // A date passing is not proof the affected person had a real opportunity
   // to respond. Until documented delivery is supported, only their actual
   // authenticated response can complete this stage.
   N(president(auth,p.boardId)&&collective(p),'PRESIDENT_REQUIRED');N(p.state==='DRAFT'||p.state==='DEFENSE_PENDING'&&noticeReady(p,now),'DEFENSE_OR_CONSENT_PENDING');
   if(p.type==='DIRECTOR_DISMISSAL'&&!p.defense)f4MemberBoardEffects_().defenseBasis(get,c,p,now);
   p.targetSnapshot=target(get,o,p.type,p.boardId,p.target);p.electorate=composition(get,o,p.boardId);p.openedAt=now;p.deadline=now+p.days*DAY;p.allowEarlyClose=c.boardProcedures.earlyCloseTypes?.includes(p.type)===true;p.state='OPEN';p.absenceOfDefenseIsAdmission=false;
  }else if(input.action==='BP_VOTE'){
   N(p.state==='OPEN'&&now<p.deadline,'VOTING_CLOSED');N(effective(p,composition(get,o,p.boardId),now)!=='SUSPENDED_COMPOSITION','COMPOSITION_CHANGED');N(p.electorate.people.some(x=>x.personId===actor),'NOT_ELECTOR');N(!p.votes.some(x=>x.personId===actor),'DUPLICATE_MANIFESTATION');N(['YES','NO','ABSTAIN'].includes(input.choice)&&input.contentHash===p.contentHash,'CONTENT_VERSION_CONFLICT');
   p.votes.push({personId:actor,choice:input.choice,at:now,contentVersion:p.contentVersion,contentHash:p.contentHash,operationId:o.operationId,confirmationKind:'INTERNAL_ELECTRONIC_NOT_ICP_BRASIL'});
  }else if(input.action==='BP_CLOSE'){
   N(p.state==='OPEN','PROCEDURE_STATE_INVALID');const state=effective(p,composition(get,o,p.boardId),now);N(state!=='OPEN','VOTING_STILL_OPEN');p.state=state;p.closedAt=now;p.closedBy=actor;
  }else N(false,'COMMAND_TYPE_INVALID');
  p.enactment=readiness(p);next.documentaryState=p.state;return {old,next};
 }
 function view(get,o,r,auth,c){
  const p=copy(r.procedure),actor=person(auth);scope(auth,p.boardId);
  if(p.state==='OPEN')p.state=effective(p,composition(get,o,p.boardId),o.clock());
  const possible=[];if(president(auth,p.boardId)&&collective(p)&&(p.state==='DRAFT'||p.state==='DEFENSE_PENDING'&&noticeReady(p,o.clock())))possible.push('BP_OPEN');
  if(p.state==='OPEN'&&p.electorate.people.some(x=>x.personId===actor)&&!p.votes.some(x=>x.personId===actor))possible.push('BP_VOTE');
  if(p.target.personId===actor&&p.state==='CONSENT_PENDING')possible.push('BP_CONSENT');
  if(p.target.personId===actor&&p.state==='DEFENSE_PENDING'&&!p.defense&&o.clock()<(p.notice?.defenseUntil||p.target.defenseDeadline))possible.push('BP_DEFEND');
  if(p.type==='DIRECTOR_DISMISSAL'&&p.state==='DEFENSE_PENDING'){
   if(p.target.personId===actor){if(!p.notice?.access)possible.push('BP_ACCESS');if(!p.notice?.acknowledgment)possible.push('BP_ACK');}
   if(president(auth,p.boardId)&&!p.notice?.externalEvidence)possible.push('BP_NOTICE');
  }
  if(r.procedure.state==='OPEN'&&p.state!=='OPEN')possible.push('BP_CLOSE');
  if(p.type==='VOLUNTARY_EXIT'&&p.state==='RECORDED'&&president(auth,p.boardId))possible.push('BP_CLEARANCE');
  if((actor===p.proposer||president(auth,p.boardId))&&!['APPROVED','REJECTED','INSUFFICIENT_MANIFESTATIONS','CANCELLED'].includes(p.state))possible.push('BP_CANCEL');
  if(['RECORDED','APPROVED'].includes(p.state)&&p.enactment?.state!=='ENACTED'&&c?.boardProcedures?.actions?.includes('BP_ENACT'))possible.push('BP_ENACT');
  const actions=possible.filter(action=>{if(p.enactment?.state==='ENACTED')return false;try{enabled(c,action,p.boardId,p.type);if(action==='BP_OPEN'&&p.type==='DIRECTOR_DISMISSAL'&&!p.defense)f4MemberBoardEffects_().defenseBasis(get,c,p,o.clock());f4InstitutionalGovernance_().control(get,o.config,o.identity,o.clock(),true,true,caps[action]);return true;}catch(e){if(['AUTHORIZATION_DENIED','M13_AUTHZ_DENIED','ACCESS_REVOKED','PILOT_DISABLED','DEFENSE_PROCESS_BASIS_REQUIRED','NOTICE_EVIDENCE_UNVERIFIED','M13_ACT_INVALID','ELECTRONIC_PROCEDURE_NORMATIVE_ENABLEMENT_REQUIRED','ELECTRONIC_FORMALIZATION_REQUIRED','MEETING_REQUIRED'].includes(e.code))return false;throw e;}});
  let enactment=readiness(p);if(actions.includes('BP_ENACT')){try{const effect=f4MemberBoardEffects_().resolve(get,o,r,auth,c);enactment={state:'READY_FOR_ENACTMENT',blocks:[],tenure:effect.summary};}catch(e){actions.splice(actions.indexOf('BP_ENACT'),1);enactment={state:'NOT_ENACTED',blocks:[e.code||'RECONCILIATION_REQUIRED']};}}
  return {...r,procedure:{...p,canonicalState:r.procedure.state,enactment},documentAvailable:['APPROVED','REJECTED','INSUFFICIENT_MANIFESTATIONS','CANCELLED','SUSPENDED_COMPOSITION','RECORDED'].includes(r.procedure.state),actions};
 }
 function candidates(get,o,auth,c){
  if(!c?.boardProcedures?.enabled)return [];
  const ids=c.boardProcedures.targetPersonIds;N(Array.isArray(ids)&&ids.length<=50&&ids.every(C.id)&&new Set(ids).size===ids.length,'SCHEMA_INVALID');
  return ids.map(id=>{const authority=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),loaded=f4MemberAuthz_().load(get,id,authority),d=loaded.M.derive(loaded.d,id,loaded.h.revision,o.clock());
   return {personId:id,name:loaded.d['memberIdentities/'+id].displayName,eligible:d.state.state==='MEMBRO_EFETIVO'&&!d.state.blocked,links:loaded.h.paths.filter(p=>p.startsWith('memberLinks/')).map(p=>loaded.d[p]).filter(l=>l.type==='MEMBRO_EFETIVO'&&C.civilStart(l.startDate)<=o.clock()&&(!l.endDate||o.clock()<C.civilStart(C.nextDay(l.endDate)))).map(l=>({id:l.id,type:l.type})),functions:d.currentFunctions.filter(f=>directors(auth).some(a=>a.idDiretoria===f.idDiretoria)).map(f=>({id:f.idVigencia,roleKey:f.roleKey,boardId:f.idDiretoria,group:f.group}))};
  });
 }
 function options(get,o,auth,c){
  if(!c?.boardProcedures?.enabled)return {roles:[],responsibleFunctions:[]};
  const roles=(c.boardProcedures.roleIds||[]).map(id=>get('memberRoleCatalog/'+id)).filter(r=>r?.active&&r.group==='DIRETORIA'&&!['PRESIDENTE','VICE_PRESIDENTE'].includes(r.roleKey)).map(r=>({id:r.id,label:r.name||r.roleKey}));
  const responsibleFunctions=c.boardProcedures.boardIds.flatMap(boardId=>composition(get,o,boardId).people.flatMap(person=>person.functions.map(f=>({id:f.id,boardId,personId:person.personId,label:person.name+' · '+f.roleKey}))));
  return {roles,responsibleFunctions};
 }
 function officeSummary(get,o,fn,functions){
  let auth;try{const d=get('institutionalDirectoryControl/current');N(d?.enabled===true,'PILOT_DISABLED');auth=f4InstitutionalPolicies_(get,o.identity,o.clock(),d.policyIds,'INSTITUTIONAL_DIRECTORY_READ','institutionalDirectory/governance');N(auth.mask.includes('summary')&&auth.mask.includes('fields'),'AUTHORIZATION_DENIED');}
  catch(e){if(['PILOT_DISABLED','AUTHORIZATION_DENIED','M13_AUTHZ_DENIED','ACCESS_REVOKED'].includes(e.code))return null;throw e;}
  const v=f4InstitutionalGovernance_().state(get),rows=v.records.filter(r=>r.kind==='BOARD_PROCEDURE'&&r.procedure.boardId===fn.idDiretoria&&r.procedure.target.personId===fn.idPessoa&&(!r.procedure.target.functionId||r.procedure.target.functionId===fn.idVigencia)&&f4InstitutionalDocumentManagement_().readable(get,o,r,auth));
  const summaries=rows.map(r=>({id:r.id,title:r.title,state:outcome(r.procedure,o.clock()),effect:r.procedure.enactment?.state||'NOT_ENACTED',at:r.procedure.enactment?.at||r.procedure.closedAt||r.procedure.createdAt})).sort((a,b)=>b.at-a.at||a.id.localeCompare(b.id));
  let balance={state:'PENDING_EVIDENCE',reason:'Base documental de permanência ainda não validada nesta consulta.'};
  try{const E=f4MemberBoardEffects_(),c=get('institutionalGovernanceControl/current'),p={boardId:fn.idDiretoria,type:'ORDINARY_EXTENSION',rule:'ADMINISTRATIVE'},b=E.evidence(get,c,p,o.clock()).basis;balance=E.balance(get,functions,b.termUnits,new Date(o.clock()-4*3600000).toISOString().slice(0,10));}
  catch(e){if(!['M13_ACT_INVALID','TENURE_EVIDENCE_REQUIRED'].includes(e.code))throw e;}
  return {balance,pending:summaries.filter(s=>!['APPROVED','REJECTED','INSUFFICIENT_MANIFESTATIONS','CANCELLED'].includes(s.state)&&s.effect!=='ENACTED'),latest:summaries[0]||null};
 }
 function contextual(get,o,fn){
  let c,auth;try{({c,auth}=f4InstitutionalGovernance_().control(get,o.config,o.identity,o.clock(),true,true,caps.BP_CREATE));}catch(e){if(['PILOT_DISABLED','AUTHORIZATION_DENIED','M13_AUTHZ_DENIED','ACCESS_REVOKED'].includes(e.code))return [];throw e;}
  const control=c.boardProcedures;if(!control?.enabled||!control.actions.includes('BP_CREATE')||!control.boardIds.includes(fn.idDiretoria)||!control.targetPersonIds.includes(fn.idPessoa))return [];
  const elected=['PRESIDENTE','VICE_PRESIDENTE'].includes(fn.roleKey),self=person(auth)===fn.idPessoa,pres=president(auth,fn.idDiretoria);
  return ['ORDINARY_EXTENSION','EXTRAORDINARY_EXTENSION','VOLUNTARY_EXIT','DIRECTOR_DISMISSAL'].filter(type=>control.types.includes(type)&&(type==='VOLUNTARY_EXIT'?self:pres&&!elected)).map(type=>({type,label:types[type].label,personId:fn.idPessoa,functionId:fn.idVigencia,boardId:fn.idDiretoria}));
 }
 return {officeSummary,noticeReady,contextual,options,types,caps,composition,target,transition,view,outcome,effective,readiness,content,hash,candidates};
}

