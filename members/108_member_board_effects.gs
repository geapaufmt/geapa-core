// Typed bridge from a sealed institutional procedure to the existing Membros
// mutation engine. This module never writes, grants access or issues evidence.
function f4MemberBoardEffects_(){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),copy=v=>F.decode(F.encode(v)),hash=v=>C.hash(F.encode(v));
 const collective=p=>['UNANIMITY','ABSOLUTE_MAJORITY'].includes(p.rule);
 function evidence(get,c,p,now){
  const ref=c.boardProcedures?.effectEvidenceByType?.[p.type];N(C.id(ref),'M13_ACT_INVALID');
  const a=f4MemberEvidence_().validate(get('memberEvidenceAttestations/'+ref));
  N(a.active&&a.kind==='BOARD_DECISION'&&a.holderId===p.boardId&&a.effect==='BOARD_PROCEDURE_BASIS'&&a.targetResourceIds.includes('memberGovernanceBoards/'+p.boardId)&&C.ms(a.verifiedAt)<=now&&now<C.ms(a.expiresAt),'M13_ACT_INVALID');
  const b=f4MemberActContent_().read(get,a.contentHash);
  const v2=b.schemaVersion==='member-board-effect-basis-v2';
  C.exact(b,['schemaVersion','boardId','types','normativeReferences','termUnits','allowedRoleIds','advisorRoleByScope','electronicProcedureEnabled','exceptionProcedureHashes','communicationProcedureHashes','sourceHash',...(v2?['procedureRequirements','defenseProcess']:[])]);
  N(['member-board-effect-basis-v1','member-board-effect-basis-v2'].includes(b.schemaVersion)&&b.boardId===p.boardId&&Array.isArray(b.types)&&b.types.includes(p.type)&&/^[a-f0-9]{64}$/.test(b.sourceHash),'M13_ACT_INVALID');
  N(Array.isArray(b.normativeReferences)&&b.normativeReferences.length>0,'M13_ACT_INVALID');
  for(const ref of b.normativeReferences){C.exact(ref,['path','sha256','approvalHash']);N(/^institutionalNormVersions\/[A-Za-z0-9_-]+$/.test(ref.path),'M13_ACT_INVALID');const norm=get(ref.path),approval=norm&&get('institutionalNormApprovals/'+norm.approvalRef);N(norm?.status==='CURRENT'&&hash(norm)===ref.sha256&&C.civilStart(norm.effectiveFrom)<=now&&approval?.active===true&&approval.normId===norm.id&&approval.immutableVersionRef===norm.immutableVersionRef&&hash(approval)===ref.approvalHash,'M13_ACT_INVALID');}
  if(collective(p)){
   if(!v2)N(b.electronicProcedureEnabled===true,'ELECTRONIC_PROCEDURE_NORMATIVE_ENABLEMENT_REQUIRED');
   else{
    const q=b.procedureRequirements?.[p.type];N(q&&['MINUTES_REQUIRED','MEETING_REQUIRED','ADDITIONAL_EXCEPTION'].includes(q.category),'ELECTRONIC_FORMALIZATION_REQUIRED');
    N(q.category!=='MEETING_REQUIRED','MEETING_REQUIRED');
    const d=get('institutionalGovernanceDrafts/current')?.records.find(r=>r.id===q.conventionalMinuteId);
    N(d?.minutes?.state==='ARCHIVED'&&d.minutes.archive?.contentHash===d.minutes.submitted?.contentHash&&hash(d)===q.conventionalMinuteHash&&q.authorizedTypes?.includes(p.type)&&q.authorizesAsynchronous===true,'ELECTRONIC_FORMALIZATION_REQUIRED');
   }
  }
  if(p.type==='EXTRAORDINARY_EXTENSION')N(b.exceptionProcedureHashes.includes(p.contentHash),'ART33_SPECIFIC_EXCEPTION_REQUIRED');
  if(p.type==='TEMPORARY_OFFICE')N(b.communicationProcedureHashes.includes(p.contentHash),'COMMUNICATION_TO_EFFECTIVE_MEMBERS_REQUIRED');
  return {basis:b,ref,contentHash:a.contentHash};
 }
 function defenseBasis(get,c,p,now){
  const b=evidence(get,c,p,now).basis;
  N(b.schemaVersion==='member-board-effect-basis-v2'&&b.defenseProcess?.allowProceedWithoutStatement===true&&b.defenseProcess.minimumDays>=7,'DEFENSE_PROCESS_BASIS_REQUIRED');
  const n=p.notice;N(f4BoardProcedures_().noticeReady(p,now)&&now>=n.receivedAt+b.defenseProcess.minimumDays*86400000,'DEFENSE_PROCESS_BASIS_REQUIRED');
  if(!n.acknowledgment)N(n.externalEvidence&&b.defenseProcess.verifiedNoticeOperationIds?.includes(n.externalEvidence.operationId),'NOTICE_EVIDENCE_UNVERIFIED');
  return true;
 }
 // A semester is a documented canonical period, never a fixed 180 days.
 // Overlapping appointments count each civil day only once in the same period.
 function tenure(get,functions,termUnits){
  functions=functions.filter(f4MemberOccupations_().active);
  N(Array.isArray(termUnits)&&termUnits.every(t=>C.id(t.termId)&&[1,2].includes(t.semesterUnits))&&new Set(termUnits.map(t=>t.termId)).size===termUnits.length,'TENURE_EVIDENCE_REQUIRED');let total=0,afterElection=0;const groups={};
  const elected=functions.filter(f=>['PRESIDENTE','VICE_PRESIDENTE'].includes(get('memberRoleCatalog/'+f.idCargo)?.roleKey));
  const electedEnd=elected.length?Math.max(...elected.map(f=>C.civilStart(C.nextDay(f.endDateReal||f.endDatePlanned)))):null;
  for(const f of functions){const role=get('memberRoleCatalog/'+f.idCargo);if(role?.group!=='DIRETORIA')continue;const term=get('memberGovernanceTerms/'+f.idTermo),units=termUnits.find(t=>t.termId===f.idTermo)?.semesterUnits;N(term&&[1,2].includes(units),'TENURE_EVIDENCE_REQUIRED');
   const a=C.civilStart(f.startDate),z=C.civilStart(C.nextDay(f.endDateReal||f.endDatePlanned)),ta=C.civilStart(term.startDate),tz=C.civilStart(C.nextDay(term.endDate));N(a>=ta&&z<=tz,'TENURE_EVIDENCE_REQUIRED');(groups[f.idTermo]||(groups[f.idTermo]={ta,tz,units,spans:[]})).spans.push([a,z]);
  }
  // Different period IDs may not overlap: unresolved historical aliases need evidence.
  const gs=Object.values(groups).sort((a,b)=>a.ta-b.ta);for(let i=1;i<gs.length;i++)N(gs[i-1].tz<=gs[i].ta,'TENURE_EVIDENCE_REQUIRED');
  for(const g of gs){const merged=[];for(const span of g.spans.sort((a,b)=>a[0]-b[0])){const last=merged.at(-1);if(last&&span[0]<=last[1])last[1]=Math.max(last[1],span[1]);else merged.push([...span]);}for(const [a,z]of merged){total+=(z-a)/(g.tz-g.ta)*g.units;if(electedEnd!==null)afterElection+=Math.max(0,z-Math.max(a,electedEnd))/(g.tz-g.ta)*g.units;}}
  return {semesterEquivalents:total,postElectedSemesterEquivalents:afterElection,hasElectedHistory:electedEnd!==null,criterion:'UNION_OF_CIVIL_DAYS_OVER_DOCUMENTED_PERIODS'};
 }
 function balance(get,functions,termUnits,today){
  const applicable=functions.filter(f=>f.startDate<=today),accrued=tenure(get,applicable.map(f=>({...f,endDateReal:(f.endDateReal||f.endDatePlanned)>today?today:(f.endDateReal||f.endDatePlanned)})),termUnits),committed=tenure(get,functions,termUnits);
  const remaining=s=>Math.max(0,Math.min(3-s.semesterEquivalents,s.hasElectedHistory?1-s.postElectedSemesterEquivalents:3));
  return {state:'CALCULATED',asOf:today,accrued,committed,remainingAccrued:remaining(accrued),remainingAfterDocumentedTerms:remaining(committed),notice:'Saldo documental não autoriza prorrogação; fundamento, revisão, período e competência são revalidados.'};
 }
 function resolve(get,o,r,auth,c){
  const p=r.procedure,now=o.clock(),actor=auth.actor.actor.h.id,B=f4BoardProcedures_();
  N(r.kind==='BOARD_PROCEDURE'&&['RECORDED','APPROVED'].includes(p.state)&&p.enactment?.state!=='ENACTED','PROCEDURE_NOT_APPROVED');
  N(c.boardProcedures?.enabled&&c.boardProcedures.actions.includes('BP_ENACT')&&c.boardProcedures.types.includes(p.type)&&c.boardProcedures.boardIds.includes(p.boardId)&&c.boardProcedures.targetPersonIds.includes(p.target.personId),'SCOPE_INVALID');
  const own=auth.actor.derived.currentDirectors.filter(f=>f.idDiretoria===p.boardId),pres=own.some(f=>f.roleKey==='PRESIDENTE');N(own.length,'AUTHORIZATION_DENIED');
  if(p.type==='VOLUNTARY_EXIT')N(actor===p.target.personId||pres,'SELF_REQUIRED');
  else if(p.type==='ADVISOR_APPOINTMENT'){const responsible=get('memberFunctions/'+p.target.responsibleFunctionId);N(pres||responsible?.idPessoa===actor&&own.some(f=>f.idVigencia===responsible.id),'AUTHORIZATION_DENIED');}
  else if(p.type==='ADVISOR_END'){const source=get('institutionalGovernanceDrafts/current')?.records.find(x=>x.procedure?.type==='ADVISOR_APPOINTMENT'&&x.procedure.enactment?.state==='ENACTED'&&x.procedure.enactment.functionId===p.target.functionId),responsible=source&&get('memberFunctions/'+source.procedure.target.responsibleFunctionId);N(pres||responsible?.idPessoa===actor&&own.some(f=>f.idVigencia===responsible.id),'AUTHORIZATION_DENIED');}
  else N(pres,'PRESIDENT_REQUIRED');
  N(B.hash(B.content(p))===p.contentHash,'CONTENT_VERSION_CONFLICT');
  if(collective(p)){N(p.state==='APPROVED'&&B.hash(B.composition(get,o,p.boardId))===B.hash(p.electorate),'COMPOSITION_CHANGED');N(B.outcome({...p,state:'OPEN'},now)==='APPROVED','PROCEDURE_NOT_APPROVED');}
  if(p.type.includes('EXTENSION'))N(p.consent?.accepted===true&&p.consent.personId===p.target.personId&&p.consent.contentHash===p.contentHash,'DEFENSE_OR_CONSENT_PENDING');
  if(p.type==='DIRECTOR_DISMISSAL'){N(B.noticeReady(p,now),'DEFENSE_OR_CONSENT_PENDING');if(!p.defense)defenseBasis(get,c,p,now);}
  const {basis,ref,contentHash}=evidence(get,c,p,now),authority=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),loaded=f4MemberAuthz_().load(get,p.target.personId,authority),M=loaded.M,d=loaded.d,t=p.target,board=get('memberGovernanceBoards/'+p.boardId),today=new Date(now-4*3600000).toISOString().slice(0,10);
  N(board&&board.startDate<=today&&today<=board.endDate,'BOARD_NOT_CURRENT');N(loaded.h.revision===p.targetSnapshot.personRevision&&board.revision===p.targetSnapshot.boardRevision,'REVISION_CONFLICT');
  const state=M.derive(d,t.personId,loaded.h.revision,now);N(!state.state.blocked&&state.state.state==='MEMBRO_EFETIVO','MEMBER_NOT_ELIGIBLE');
  const O=f4MemberOccupations_(),originals=M.rows(d,'memberFunctions'),old=t.functionId?d['memberFunctions/'+t.functionId]:null;let fn,role,roleChange=null,segments=null;
  if(old){N(old.revision===p.targetSnapshot.functionRevision&&old.endDateReal===null&&state.currentFunctions.some(f=>O.members(originals,old).filter(O.active).some(s=>s.id===f.idVigencia)),'REVISION_CONFLICT');role=get('memberRoleCatalog/'+old.idCargo);N(role?.active,'TARGET_FUNCTION_INVALID');
   if(p.type!=='VOLUNTARY_EXIT')N(!['PRESIDENTE','VICE_PRESIDENTE'].includes(role.roleKey),'ELECTIVE_ROLE_PROTECTED');
   if(p.type==='ADVISOR_END')N(role.group==='ASSESSORIA','TARGET_FUNCTION_INVALID');else N(role.group==='DIRETORIA','TARGET_FUNCTION_INVALID');
   const end=p.type==='DIRECTOR_DISMISSAL'?today:t.endDate;N(end>=today&&end>=old.startDate&&end<=board.endDate,'TERM_OUTSIDE_BOARD');
   if(p.type.includes('EXTENSION')){segments=O.extend(get,originals,old,end);fn=segments.find(f=>f.id===old.id)||copy(old);}
   else{if(p.type==='VOLUNTARY_EXIT')N(C.civilStart(end)>=p.createdAt+7*86400000,'NOTICE_PERIOD_REQUIRED');segments=O.close(originals,old,end);fn=segments.find(f=>f.id===old.id)||copy(old);}
  }else{
   N(t.startDate>=today&&t.endDate<=board.endDate,'TERM_OUTSIDE_BOARD');const link=d['memberLinks/'+t.linkId];N(link?.type==='MEMBRO_EFETIVO'&&link.startDate<=t.startDate&&(!link.endDate||link.endDate>=t.endDate),'MEMBER_NOT_ELIGIBLE');
   let roleId=t.roleId,scope;
   if(p.type==='ADVISOR_APPOINTMENT'){const responsible=get('memberFunctions/'+t.responsibleFunctionId);N(responsible&&B.composition(get,o,p.boardId).people.some(x=>x.functions.some(f=>f.id===responsible.id))&&t.endDate<=(responsible.endDateReal||responsible.endDatePlanned),'RESPONSIBLE_DIRECTOR_INVALID');scope=responsible.scopeKey;roleId=basis.advisorRoleByScope[scope];}
   if(p.type==='TEMPORARY_OFFICE'){N(C.id(basis.advisorRoleByScope.TEMPORARY_SCOPE)&&f4MemberContracts_().scopes.includes(basis.advisorRoleByScope.TEMPORARY_SCOPE),'M13_ACT_INVALID');scope=basis.advisorRoleByScope.TEMPORARY_SCOPE;roleId='CARGO_'+r.id.slice(4);role={schemaVersion:'m1-memberRoleCatalog-v2',id:roleId,revision:1,origin:'BOARD_PROCEDURE',roleKey:'DIRETOR_TEMPORARIO',group:'DIRETORIA',active:true,boardId:p.boardId,scopeKey:scope,procedureId:r.id,name:t.roleName,startDate:t.startDate,endDate:t.endDate,advisoryAllowed:t.advisoryAllowed};N(!get('memberRoleCatalog/'+roleId),'REVISION_CONFLICT');roleChange={path:'memberRoleCatalog/'+roleId,expectedRevision:0,data:role};}
   else{N(basis.allowedRoleIds.includes(roleId),'SCOPE_INVALID');role=get('memberRoleCatalog/'+roleId);N(role?.active&&!['PRESIDENTE','VICE_PRESIDENTE'].includes(role.roleKey),'ELECTIVE_ROLE_PROTECTED');scope=f4MemberContracts_().roles[role.roleKey]?.[1];N(role.group===(p.type==='ADVISOR_APPOINTMENT'?'ASSESSORIA':'DIRETORIA'),'TARGET_FUNCTION_INVALID');}
   const periods=O.periods(get,p.boardId,t.startDate,t.endDate);
   const id='VIG_'+r.id.slice(4);N(!get('memberFunctions/'+id),'REVISION_CONFLICT');fn={schemaVersion:'m1-memberFunctions-v2',id,revision:1,origin:'BOARD_PROCEDURE',idPessoa:t.personId,idVinculo:t.linkId,idCargo:roleId,scopeKey:scope,idDiretoria:p.boardId,idTermo:periods[0].termId,startDate:t.startDate,endDatePlanned:periods[0].endDate,endDateReal:null};
   segments=periods.map((period,index)=>({...copy(fn),schemaVersion:'m1-memberFunctions-v3',occupationId:id,cancelled:false,id:index?'VIG_'+C.hash({occupationId:id,termId:period.termId,startDate:period.startDate}).slice(0,32):id,idTermo:period.termId,startDate:period.startDate,endDatePlanned:period.endDate}));fn=segments[0];
  }
  segments=segments||[fn];const ids=new Set(segments.map(f=>f.id)),all=originals.filter(f=>!ids.has(f.id)),nextFunctions=[...all,...segments];O.validate(nextFunctions);
  const nextGet=path=>path==='memberRoleCatalog/'+role.id?role:get(path),bounds=f=>M.interval(f.startDate,f.endDateReal||f.endDatePlanned),overlap=(a,b)=>a.a<b.z&&b.a<a.z;
  N(!all.filter(O.active).some(f=>get('memberRoleCatalog/'+f.idCargo)?.group==='DIRETORIA'&&role.group==='DIRETORIA'&&segments.filter(O.active).some(n=>overlap(bounds(f),bounds(n)))),'FUNCTION_EXCLUSIVITY_CONFLICT');
  // The authority epoch read/write serializes this bounded inventory with every
  // canonical Membros function mutation, including concurrent appointments.
  const db=F.connect(o.config);let cursor;for(let page=0;page<10;page++){const rows=db.query({collection:'memberFunctions',filters:[{field:'idDiretoria',op:'EQUAL',value:p.boardId}],limit:20,order:[{field:'__name__',direction:'ASCENDING'}],...(cursor?{cursor:[cursor]}:{})});for(const row of rows){const f=get(row.path);N(ids.has(f.id)||!O.active(f)||f.idCargo!==fn.idCargo||!segments.filter(O.active).some(n=>overlap(bounds(f),bounds(n))),'FUNCTION_EXCLUSIVITY_CONFLICT');}if(rows.length<20)break;N(page<9,'SCOPE_TOO_LARGE');cursor=new F.DocumentReference('projects/'+o.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+rows.at(-1).path);}
  const summary=role.group==='DIRETORIA'&&['DIRECTOR_APPOINTMENT','TEMPORARY_OFFICE','ORDINARY_EXTENSION','EXTRAORDINARY_EXTENSION'].includes(p.type)?tenure(nextGet,nextFunctions,basis.termUnits):null;
  if(summary&&p.type!=='EXTRAORDINARY_EXTENSION')N(summary.semesterEquivalents<=3+1e-10&&(!summary.hasElectedHistory||summary.postElectedSemesterEquivalents<=1+1e-10),'TENURE_LIMIT_EXCEEDED');
  for(const segment of segments)if(!O.modern(segment))segment.schemaVersion='m1-memberFunctions-v2';
  const changes=[...(roleChange?[roleChange]:[]),...segments.map(f=>({path:'memberFunctions/'+f.id,expectedRevision:originals.find(x=>x.id===f.id)?.revision||0,data:f}))];
  for(const changed of segments){const old=originals.find(f=>f.id===changed.id);if(!old)continue;const fn=changed,id='EVT_'+(old.id===t.functionId?o.operationId:C.hash({operationId:o.operationId,functionId:old.id}).slice(0,32));N(!get('memberEvents/'+id),'OPERATION_CONFLICT');changes.push({path:'memberEvents/'+id,expectedRevision:0,data:{schemaVersion:'m1-memberEvents-v1',id,revision:1,origin:'BOARD_PROCEDURE',idPessoa:t.personId,idVinculo:fn.idVinculo,type:'RETIFICACAO',effect:null,effectiveDate:today,endDate:null,supersedes:null,correctionPath:'memberFunctions/'+fn.id,prior:copy(old)}});}
  return {documentaryAuthorization:collective(p)&&basis.schemaVersion==='member-board-effect-basis-v2'?copy(basis.procedureRequirements[p.type]):null,id:t.personId,expected:loaded.h.revision,changes,now,boardExpected:{},mode:'M13',deadline:auth.validUntil,evidence:{ref,contentHash},summary,functionId:fn.id,responsibleFunctionId:t.responsibleFunctionId||null};
 }
 return {resolve,tenure,balance,evidence,defenseBasis};
}
