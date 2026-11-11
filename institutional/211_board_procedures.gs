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
 const caps={BP_CREATE:'INSTITUTIONAL_BOARD_PROPOSE',BP_OPEN:'INSTITUTIONAL_BOARD_PROPOSE',BP_CANCEL:'INSTITUTIONAL_BOARD_PROPOSE',BP_VOTE:'INSTITUTIONAL_BOARD_MANIFEST',BP_CONSENT:'INSTITUTIONAL_BOARD_MANIFEST',BP_DEFEND:'INSTITUTIONAL_BOARD_MANIFEST',BP_CLOSE:'INSTITUTIONAL_BOARD_PROPOSE',BP_CLEARANCE:'INSTITUTIONAL_BOARD_PROPOSE'};
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
   if(type==='DIRECTOR_DISMISSAL')N(!['PRESIDENTE','VICE_PRESIDENTE'].includes(role.roleKey),'ELECTIVE_ROLE_PROTECTED');
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
  const blocks=[];if(!['APPROVED','RECORDED'].includes(p.state))blocks.push('PROCEDURE_NOT_APPROVED');
  if(collective(p))blocks.push('ELECTRONIC_PROCEDURE_NORMATIVE_ENABLEMENT_REQUIRED');
  if(p.type==='EXTRAORDINARY_EXTENSION')blocks.push('ART33_SPECIFIC_EXCEPTION_REQUIRED');
  if(p.type==='TEMPORARY_OFFICE')blocks.push('COMMUNICATION_TO_EFFECTIVE_MEMBERS_REQUIRED');
  if(['TEMPORARY_OFFICE','DIRECTOR_APPOINTMENT','ORDINARY_EXTENSION','EXTRAORDINARY_EXTENSION'].includes(p.type))blocks.push('CANONICAL_TENURE_REVALIDATION_REQUIRED');
  blocks.push('TYPED_MEMBER_EXECUTOR_REQUIRED');
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
   procedure.contentHash=hash(content(procedure));return {old:null,next:{kind:'BOARD_PROCEDURE',title:p.title,meetingDate:null,body:p.object,documentaryState:state,publication:'CLOSED',formalApproval:null,effectiveFrom:null,sourceRecordId:null,references:[],createdAt:now,procedure}};
  }
  N(old?.kind==='BOARD_PROCEDURE'&&old.revision===input.expectedRecordRevision,'REVISION_CONFLICT');const next=copy(old),p=next.procedure;scope(auth,p.boardId);enabled(c,input.action,p.boardId,p.type);
  const base=['action','id','expectedRevision','expectedRecordRevision','reason'];
  C.exact(input,[...base,...(input.action==='BP_VOTE'?['choice','contentHash']:input.action==='BP_CONSENT'?['accepted','contentHash']:input.action==='BP_DEFEND'?['statement','contentHash']:input.action==='BP_CLEARANCE'?['pendingDuties']:[])]);
  if(input.action==='BP_CANCEL'){N(actor===p.proposer||president(auth,p.boardId),'AUTHORIZATION_DENIED');N(!['APPROVED','REJECTED','INSUFFICIENT_MANIFESTATIONS','CANCELLED'].includes(outcome(p,now)),'PROCEDURE_IMMUTABLE');p.state='CANCELLED';p.cancelled={at:now,personId:actor,reason:input.reason};}
  else if(input.action==='BP_CONSENT'){N(p.state==='CONSENT_PENDING'&&p.target.personId===actor,'SELF_REQUIRED');N(input.contentHash===p.contentHash&&typeof input.accepted==='boolean','CONTENT_VERSION_CONFLICT');p.consent={personId:actor,at:now,accepted:input.accepted,contentHash:p.contentHash,operationId:o.operationId};p.state=input.accepted?(collective(p)?'DRAFT':'RECORDED'):'REJECTED';}
  else if(input.action==='BP_DEFEND'){N(p.state==='DEFENSE_PENDING'&&p.target.personId===actor&&now<p.target.defenseDeadline,'SELF_REQUIRED');N(!p.defense&&input.contentHash===p.contentHash&&text(input.statement,1,8000),'DEFENSE_INVALID');p.defense={personId:actor,at:now,statement:input.statement,contentHash:p.contentHash,operationId:o.operationId};}
  else if(input.action==='BP_CLEARANCE'){N(p.type==='VOLUNTARY_EXIT'&&president(auth,p.boardId),'PRESIDENT_REQUIRED');N(p.state==='RECORDED'&&text(input.pendingDuties,0,4000),'COMMAND_PAYLOAD_INVALID');p.clearance={at:now,personId:actor,pendingDuties:input.pendingDuties,meaning:'DUTIES_RECORD_NOT_RESIGNATION_VETO'};}
  else if(input.action==='BP_OPEN'){
   // A date passing is not proof the affected person had a real opportunity
   // to respond. Until documented delivery is supported, only their actual
   // authenticated response can complete this stage.
   N(president(auth,p.boardId)&&collective(p),'PRESIDENT_REQUIRED');N(p.state==='DRAFT'||p.state==='DEFENSE_PENDING'&&!!p.defense,'DEFENSE_OR_CONSENT_PENDING');
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
  const possible=[];if(president(auth,p.boardId)&&collective(p)&&(p.state==='DRAFT'||p.state==='DEFENSE_PENDING'&&p.defense))possible.push('BP_OPEN');
  if(p.state==='OPEN'&&p.electorate.people.some(x=>x.personId===actor)&&!p.votes.some(x=>x.personId===actor))possible.push('BP_VOTE');
  if(p.target.personId===actor&&p.state==='CONSENT_PENDING')possible.push('BP_CONSENT');
  if(p.target.personId===actor&&p.state==='DEFENSE_PENDING'&&!p.defense&&o.clock()<p.target.defenseDeadline)possible.push('BP_DEFEND');
  if(r.procedure.state==='OPEN'&&p.state!=='OPEN')possible.push('BP_CLOSE');
  if(p.type==='VOLUNTARY_EXIT'&&p.state==='RECORDED'&&president(auth,p.boardId))possible.push('BP_CLEARANCE');
  if((actor===p.proposer||president(auth,p.boardId))&&!['APPROVED','REJECTED','INSUFFICIENT_MANIFESTATIONS','CANCELLED'].includes(p.state))possible.push('BP_CANCEL');
  const actions=possible.filter(action=>{try{enabled(c,action,p.boardId,p.type);f4InstitutionalGovernance_().control(get,o.config,o.identity,o.clock(),true,true,caps[action]);return true;}catch(e){if(['AUTHORIZATION_DENIED','M13_AUTHZ_DENIED','ACCESS_REVOKED','PILOT_DISABLED'].includes(e.code))return false;throw e;}});
  return {...r,procedure:{...p,canonicalState:r.procedure.state,enactment:readiness(p)},documentAvailable:['APPROVED','REJECTED','INSUFFICIENT_MANIFESTATIONS','CANCELLED','SUSPENDED_COMPOSITION','RECORDED'].includes(r.procedure.state),actions};
 }
 function candidates(get,o,auth,c){
  if(!c?.boardProcedures?.enabled)return [];
  const ids=c.boardProcedures.targetPersonIds;N(Array.isArray(ids)&&ids.length<=50&&ids.every(C.id)&&new Set(ids).size===ids.length,'SCHEMA_INVALID');
  return ids.map(id=>{const authority=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),loaded=f4MemberAuthz_().load(get,id,authority),d=loaded.M.derive(loaded.d,id,loaded.h.revision,o.clock());
   return {personId:id,name:loaded.d['memberIdentities/'+id].displayName,eligible:d.state.state==='MEMBRO_EFETIVO'&&!d.state.blocked,links:loaded.h.paths.filter(p=>p.startsWith('memberLinks/')).map(p=>loaded.d[p]).filter(l=>l.type==='MEMBRO_EFETIVO'&&C.civilStart(l.startDate)<=o.clock()&&(!l.endDate||o.clock()<C.civilStart(C.nextDay(l.endDate)))).map(l=>({id:l.id,type:l.type})),functions:d.currentFunctions.filter(f=>directors(auth).some(a=>a.idDiretoria===f.idDiretoria)).map(f=>({id:f.idVigencia,roleKey:f.roleKey,boardId:f.idDiretoria,group:f.group}))};
  });
 }
 return {types,caps,composition,target,transition,view,outcome,effective,readiness,content,hash,candidates};
}
