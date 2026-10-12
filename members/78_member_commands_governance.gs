// E G/H. Institutional role catalog and nullable approvals are read, never minted here.
function f4MemberGovernanceCommands_(){
 const C=f4Core_(),K=f4MemberCommandCatalog_(),N=C.need,U=f4MemberFamilyInputs_(),copy=v=>f4Firestore_().decode(f4Firestore_().encode(v));
 function validate(r){N(['members.governance.appoint','members.governance.endFunction'].includes(r.commandType),'COMMAND_TYPE_INVALID');C.exact(r.target,['personId','functionId']);N(C.id(r.target.personId)&&C.id(r.target.functionId)&&r.evidenceRefs.length>0&&C.id(r.decisionRef),'COMMAND_PAYLOAD_INVALID');
  if(r.commandType==='members.governance.appoint'){C.exact(r.payload,['idVinculo','idCargo','scopeKey','idDiretoria','idTermo','startDate','endDatePlanned']);for(const k of ['idVinculo','idCargo','idDiretoria','idTermo'])N(C.id(r.payload[k]),'COMMAND_PAYLOAD_INVALID');N(r.payload.endDatePlanned!==null,'COMMAND_PAYLOAD_INVALID');f4MemberModel_().interval(r.payload.startDate,r.payload.endDatePlanned);}
  else{C.exact(r.payload,['endDateReal']);C.civilStart(r.payload.endDateReal);}return true;
 }
 function resolve(get,r,now){validate(r);const {request:q,proof}=U.authorization(get,r),appoint=r.commandType==='members.governance.appoint',path='memberFunctions/'+r.target.functionId,old=get(path),revision=U.revision(r,path);
  N(q.capability==='MEMBERS_GOVERNANCE_EXECUTE'&&q.target.personId===r.target.personId,'AUTH_BINDING_INVALID');f4MemberAuthz_().recheck(get,q,proof,now);
  const body=f4MemberActContent_().act(get,r,now);C.exact(body.facts,['dependencies']);N(Array.isArray(body.facts.dependencies),'M13_ACT_INVALID');
  const data=appoint?{schemaVersion:'m1-memberFunctions-v2',id:r.target.functionId,idPessoa:r.target.personId,revision:1,origin:'M13_COMMAND',...r.payload,endDateReal:null}:{...copy(old),schemaVersion:'m1-memberFunctions-v2',revision:revision+1,endDateReal:r.payload.endDateReal};
  N(appoint?!old&&revision===0:old&&old.idPessoa===r.target.personId&&old.revision===revision,'REVISION_CONFLICT');
  const required=['memberLinks/'+data.idVinculo,'memberRoleCatalog/'+data.idCargo,'memberGovernanceBoards/'+data.idDiretoria,'memberGovernanceTerms/'+data.idTermo,'memberBoardHeads/'+data.idDiretoria];
  const bh=get('memberBoardHeads/'+data.idDiretoria);N(bh&&Array.isArray(bh.paths),'COMMAND_PAYLOAD_INVALID');const relations=bh.paths.filter(p=>p.startsWith('memberGovernanceBoardTerms/')).filter(p=>get(p)?.idTermo===data.idTermo);N(relations.length===1,'COMMAND_PAYLOAD_INVALID');required.push(...relations);
  N(K.equal(body.facts.dependencies.map(x=>x.path).sort(),required.sort()),'COMMAND_PAYLOAD_INVALID');for(const dep of body.facts.dependencies){C.exact(dep,['path','revision']);N(C.rev(dep.revision)&&get(dep.path)?.revision===dep.revision,'REVISION_CONFLICT');}
  const role=get('memberRoleCatalog/'+data.idCargo),control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS'));N(role&&role.active&&q.target.roleKey===role.roleKey&&q.target.scopeKey===data.scopeKey&&q.target.boardId===data.idDiretoria&&q.target.linkId===data.idVinculo,'AUTH_BINDING_INVALID');
  f4MemberContracts_().functionScope(data,role,control);N(q.target.effectiveDate===(appoint?data.startDate:data.endDateReal),'AUTH_BINDING_INVALID');
  const changes=[{path,expectedRevision:revision,data}];
  if(!appoint){N(old.endDateReal===null,'COMMAND_PAYLOAD_INVALID');const id='EVT_'+r.operationId,ep='memberEvents/'+id;N(U.revision(r,ep)===0&&!get(ep),'REVISION_CONFLICT');changes.push({path:ep,expectedRevision:0,data:{schemaVersion:'m1-memberEvents-v1',id,revision:1,origin:'M13_COMMAND',idPessoa:r.target.personId,idVinculo:data.idVinculo,type:'RETIFICACAO',effect:null,effectiveDate:data.endDateReal,endDate:null,supersedes:null,correctionPath:path,prior:copy(old)}});}
  N(K.equal([...q.target.resources].sort(),[...changes.map(c=>c.path),...required].sort()),'AUTH_BINDING_INVALID');
  return {id:r.target.personId,expected:U.revision(r,'memberAggregateHeads/'+r.target.personId),changes,now,boardExpected:{},mode:'M13',authorization:{request:q,proof},deadline:C.ms(proof.expiresAt)};
 }
 return {validate,resolve};
}

// E I, B4: no fictitious person/role target. Inventory is protected and epoch-bound.
function f4MemberBoardCommands_(){
 const C=f4Core_(),K=f4MemberCommandCatalog_(),N=C.need,U=f4MemberFamilyInputs_(),M=f4MemberModel_(),copy=v=>f4Firestore_().decode(f4Firestore_().encode(v));
 const actions={CREATE_BOARD:'memberGovernanceBoards',END_BOARD:'memberGovernanceBoards',CREATE_TERM:'memberGovernanceTerms',END_TERM:'memberGovernanceTerms',CREATE_RELATION:'memberGovernanceBoardTerms',END_RELATION:'memberGovernanceBoardTerms'};
 function validate(r){N(r.commandType==='members.governance.boardTerm','COMMAND_TYPE_INVALID');C.exact(r.target,['boardId','termId','relationId','scopeKey']);const t=r.target,a=r.payload.action;N(Object.hasOwn(actions,a)&&C.id(t.boardId)&&f4MemberContracts_().scopes.includes(t.scopeKey)&&C.id(r.decisionRef)&&r.evidenceRefs.length>0,'COMMAND_PAYLOAD_INVALID');N(a.endsWith('BOARD')?t.termId===null&&t.relationId===null:a.endsWith('TERM')?C.id(t.termId)&&t.relationId===null:C.id(t.termId)&&C.id(t.relationId),'COMMAND_PAYLOAD_INVALID');
  C.exact(r.payload,a.startsWith('END_')?['action','endDate']:a==='CREATE_RELATION'?['action','startDate','endDate']:['action','name','startDate','endDate']);N(r.payload.endDate!==null,'COMMAND_PAYLOAD_INVALID');C.civilStart(r.payload.endDate);if(a.startsWith('CREATE_'))M.interval(r.payload.startDate,r.payload.endDate);return true;
 }
 function resolve(get,r,now){validate(r);const {request:q,proof}=U.authorization(get,r);N(q.schemaVersion==='member-authz-governance-request-v13','AUTH_BINDING_INVALID');f4MemberAuthz_().recheckGovernance(get,q,proof,now,r);
  const t=r.target,a=r.payload.action,col=actions[a],id=col==='memberGovernanceBoards'?t.boardId:col==='memberGovernanceTerms'?t.termId:t.relationId,path=col+'/'+id,old=get(path),create=a.startsWith('CREATE_'),revision=U.revision(r,path),expected=U.revision(r,'memberBoardHeads/'+t.boardId),bh=get('memberBoardHeads/'+t.boardId);
  N((old?.revision||0)===revision&&(bh?.revision||0)===expected&&(create?!old:!!old),'REVISION_CONFLICT');N(q.target.effect===a&&q.target.effectiveDate===(create?r.payload.startDate:r.payload.endDate),'AUTH_BINDING_INVALID');
  const inventory=f4MemberGovernanceContent_().inventory(get,r,q,now),grant=get((q.kind==='EXPLICIT'?'memberExplicitGrants/':'memberAuthzPolicies/')+q.sourceId);N(inventory.boards.every(b=>grant.targets.allowedBoards.includes(b)),'AUTH_BINDING_INVALID');
  const data=create?{schemaVersion:'m1-'+col+'-v1',id,revision:1,origin:'M13_COMMAND',...(col==='memberGovernanceBoardTerms'?{idDiretoria:t.boardId,idTermo:t.termId}:{name:r.payload.name}),startDate:r.payload.startDate,endDate:r.payload.endDate}:{...copy(old),revision:revision+1,endDate:r.payload.endDate};M.validate(col,data);
  if(!create)N(C.civilStart(data.endDate)<=C.civilStart(old.endDate),'COMMAND_PAYLOAD_INVALID');
  const deps=new Set(['memberBoardHeads/'+t.boardId]),graph={};for(const board of inventory.boards){const head=get('memberBoardHeads/'+board);if(a==='CREATE_BOARD'&&board===t.boardId)continue;N(head&&head.kind==='BOARD'&&head.id===board&&Array.isArray(head.paths),'COMMAND_PAYLOAD_INVALID');deps.add('memberBoardHeads/'+board);for(const p of head.paths){N(p==='memberGovernanceBoards/'+board||p.startsWith('memberGovernanceBoardTerms/'),'COMMAND_PAYLOAD_INVALID');const v=get(p);N(v&&v.id===p.split('/')[1],'COMMAND_PAYLOAD_INVALID');deps.add(p);graph[p]=v;if(p.startsWith('memberGovernanceBoardTerms/')){const tp='memberGovernanceTerms/'+v.idTermo;deps.add(tp);graph[tp]=get(tp);}}}
  if(a!=='CREATE_BOARD')deps.add('memberGovernanceInventories/'+t.boardId);if(t.termId&&a.endsWith('RELATION')){const tp='memberGovernanceTerms/'+t.termId;deps.add(tp);graph[tp]=get(tp);}graph[path]=data;M.graph(graph);
  for(const personId of inventory.persons){const control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),dependent=f4MemberAuthz_().load(get,personId,control);dependent.M.graph({...dependent.d,[path]:data});}
  const resources=[...new Set([path,...deps])].sort();N(K.equal([...q.target.resources].sort(),resources),'AUTH_BINDING_INVALID');
  const outbox={schemaVersion:'member-outbox-v13',operationId:r.operationId,commandType:r.commandType,kind:'MEMBERS_REPROJECT',state:'PENDING',idempotencyKey:r.operationId,attempts:0,personIds:inventory.persons,sourceDigest:r.digest};
  return {id:t.boardId,expected,changes:[{path,expectedRevision:revision,data}],now,boardExpected:{[t.boardId]:expected},mode:'M13',authorization:{request:q,proof,command:r},deadline:C.ms(proof.expiresAt),outbox};
 }
 function outbox(v,r){C.exact(v,['schemaVersion','operationId','commandType','kind','state','idempotencyKey','attempts','personIds','sourceDigest']);N(v.schemaVersion==='member-outbox-v13'&&v.operationId===r.operationId&&v.commandType===r.commandType&&r.commandType==='members.governance.boardTerm'&&v.kind==='MEMBERS_REPROJECT'&&v.state==='PENDING'&&v.idempotencyKey===r.operationId&&v.attempts===0&&v.sourceDigest===r.digest&&Array.isArray(v.personIds)&&new Set(v.personIds).size===v.personIds.length&&v.personIds.every(C.id),'COMMAND_PAYLOAD_INVALID');return true;}
 return {validate,resolve,outbox};
}
