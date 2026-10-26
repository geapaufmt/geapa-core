// Request/review producer in the existing Members admission and transaction
// family. No browser-supplied authorization or independent writer.
function f4MemberVoluntaryCommands_(){
 const C=f4Core_(),N=C.need,U=f4MemberFamilyInputs_(),W=f4MemberVoluntaryWorkflow_();
 const types=['members.voluntary.submit','members.voluntary.review'];
 const selfActions=['CANCEL_SELF','COMPLEMENT'];
 function validate(r){
  N(types.includes(r.commandType),'COMMAND_TYPE_INVALID');C.exact(r.target,['personId','linkId','requestId']);
  N(C.id(r.target.personId)&&C.id(r.target.linkId)&&C.op(r.target.requestId)&&r.decisionRef===null,'COMMAND_PAYLOAD_INVALID');
  if(r.commandType==='members.voluntary.submit'){C.exact(r.payload,['modality','reason','semesterId',...(r.payload.modality==='SUSPENSAO_TEMPORARIA'?['startDate','endDate']:[]),...(Object.hasOwn(r.payload,'submission')?['submission']:[])]);N(r.target.requestId===r.operationId&&r.evidenceRefs.length===0,'COMMAND_PAYLOAD_INVALID');}
  else{C.exact(r.payload,['action','reason','minutesReference','confirmed',...(Object.hasOwn(r.payload,'graveException')?['graveException']:[]),...(Object.hasOwn(r.payload,'requestKind')?['requestKind']:[])]);N(typeof r.payload.action==='string'&&(!r.payload.requestKind||r.payload.requestKind==='SUSPENSAO_TEMPORARIA'),'COMMAND_PAYLOAD_INVALID');}
  return true;
 }
 function capability(r){return r.commandType==='members.voluntary.submit'||selfActions.includes(r.payload.action)?'MEMBERS_VOLUNTARY_REQUEST':r.payload.requestKind==='SUSPENSAO_TEMPORARIA'?'MEMBERS_VOLUNTARY_SUSPENSION_EXECUTE':'MEMBERS_VOLUNTARY_EXIT_EXECUTE';}
 function resolve(get,r,now){
  validate(r);const {request:q,proof}=U.authorization(get,r),cap=capability(r,get),id=r.target.personId,requestPath='memberVoluntaryRequests/'+r.target.requestId,indexPath='memberVoluntaryIndexes/'+id;
  N(q.capability===cap&&q.target.personId===id,'AUTH_BINDING_INVALID');
  f4MemberAuthz_().recheck(get,q,proof,now);
  const control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),loaded=f4MemberAuthz_().load(get,id,control),link=loaded.d['memberLinks/'+r.target.linkId],state=loaded.M.derive(loaded.d,id,loaded.h.revision,now).state;
  N(link&&link.idPessoa===id,'COMMAND_PAYLOAD_INVALID');
  const today=Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd'),old=get(requestPath),suspension=(old?.modality||r.payload.modality)==='SUSPENSAO_TEMPORARIA',policy=get('memberOperationalConfig/'+(suspension?'voluntarySuspension':'voluntaryExit')),priorIndex=W.index(get(indexPath),id);
  N(U.revision(r,'memberAggregateHeads/'+id)===loaded.h.revision&&U.revision(r,indexPath)===priorIndex.revision&&U.revision(r,requestPath)===(old?.revision||0),'REVISION_CONFLICT');
  const self=cap==='MEMBERS_VOLUNTARY_REQUEST';N(!self||proof.idPessoa===id,'ACCESS_REVOKED');
  let result;
  if(r.commandType==='members.voluntary.submit'){
   N(!old,'OPERATION_CONFLICT');const semester=r.payload.semesterId?get('academicSemesters/'+r.payload.semesterId):null;
   result=W.create({operationId:r.operationId,actorUid:r.ownerUid,personId:id,link,state,today,now,policy,semester,priorIndex,payload:r.payload});
   if(!suspension){const intent=get('memberPreparedIntents/'+r.operationId);N(intent?.ownerUid===r.ownerUid,'AUTH_BINDING_INVALID');f4MemberActivityChecks_().exit(get,intent.operationalChecks,id,today,policy.presentationBlockDays);}
  }else{
   W.validate(old);N(old.personId===id&&old.linkId===link.id&&priorIndex.activeRequestId===old.id&&(r.payload.requestKind==='SUSPENSAO_TEMPORARIA')===suspension,'COMMAND_PAYLOAD_INVALID');
   // Approval uses the existing voluntaryExit kernel so the institutional
   // event and workflow closure share exactly one canonical transaction.
   N(suspension||r.payload.action!=='APPROVE','COMMAND_TYPE_INVALID');
   const semester=old.semesterId?get('academicSemesters/'+old.semesterId):null;
   result=W.review({request:old,expectedRevision:U.revision(r,requestPath),actorUid:r.ownerUid,actorPersonId:proof.idPessoa,administrative:!self,today,now,policy,semester,payload:r.payload,operationId:r.operationId,link,priorSuspensions:loaded.M.rows(loaded.d,'memberEvents').filter(e=>e.effect==='SUSPENSAO'&&!loaded.M.rows(loaded.d,'memberEvents').some(c=>c.supersedes===e.id))});
   N(suspension||result.effect===null,'COMMAND_PAYLOAD_INVALID');result.index={...priorIndex,revision:priorIndex.revision+1,activeRequestId:result.terminal?null:old.id};
  }
  if(suspension&&(r.commandType==='members.voluntary.submit'||r.payload.action==='APPROVE')){const intent=get('memberPreparedIntents/'+r.operationId);N(intent?.ownerUid===r.ownerUid,'AUTH_BINDING_INVALID');f4MemberActivityChecks_().suspension(get,intent.operationalChecks,id,result.request.startDate,result.request.endDate);}
  const changes=[{path:requestPath,expectedRevision:old?.revision||0,data:result.request},{path:indexPath,expectedRevision:priorIndex.revision,data:result.index}];
  if(result.effect){const e=result.effect,path='memberEvents/'+e.id;N(!get(path)&&U.revision(r,path)===0&&q.target.effect===e.type&&q.target.effectiveDate===e.effectiveDate,'REVISION_CONFLICT');const data={schemaVersion:'m1-memberEvents-v1',id:e.id,revision:1,origin:'M13_COMMAND',idPessoa:id,idVinculo:link.id,type:e.type,effect:e.type,effectiveDate:e.effectiveDate,endDate:e.endDate,supersedes:null,correctionPath:null,prior:null};loaded.M.validate('memberEvents',data);changes.push({path,expectedRevision:0,data});}
  N(changes.every(c=>q.target.resources.includes(c.path)),'AUTH_BINDING_INVALID');
  return {id,expected:loaded.h.revision,changes,now,boardExpected:{},mode:'M13',authorization:{request:q,proof},deadline:C.ms(proof.expiresAt)};
 }
 return {validate,resolve,capability};
}

// A SELF policy authorizes the person's fixed index. The command producer must
// still bind any individual request to that person and the exact operation.
function f4MemberVoluntaryScope_(target,limits,personId){
 const K=f4MemberAuthzContracts_(),N=K.need;K.targets(limits);
 N(limits.ownPerson&&target.personId===personId&&target.resources.includes('memberVoluntaryIndexes/'+personId));
 N(limits.commandScopes?.includes('OWN_MEMBER_RESOURCES')||limits.allowedResources.includes('memberVoluntaryIndexes/'+personId));
 N(target.resources.every(path=>path==='memberVoluntaryIndexes/'+personId||/^memberVoluntaryRequests\/[a-f0-9]{32}$/.test(path)));
}
function f4MemberVoluntaryChange_(change,personId){
 const C=f4Core_(),N=C.need,W=f4MemberVoluntaryWorkflow_(),d=change.data;
 N(C.rev(change.expectedRevision)&&d.revision===change.expectedRevision+1,'REVISION_CONFLICT');
 if(change.path.startsWith('memberVoluntaryRequests/')){W.validate(d);N(change.path==='memberVoluntaryRequests/'+d.id&&d.personId===personId,'AUTH_BINDING_INVALID');}
 else {W.index(d,personId);N(change.path==='memberVoluntaryIndexes/'+personId,'AUTH_BINDING_INVALID');}
 return true;
}

// Final human decision is attached to the already-existing voluntaryExit
// command, so there is no independently successful workflow closure.
function f4MemberVoluntaryExecution_(get,r,now){
 const C=f4Core_(),N=C.need,U=f4MemberFamilyInputs_(),W=f4MemberVoluntaryWorkflow_();
 const facts=f4MemberActContent_().act(get,r,now).facts;
 C.exact(facts,['schemaVersion','requestId','requestRevision','policyRevision','review','externalReview']);
 N(facts.schemaVersion==='member-voluntary-decision-v1'&&C.op(facts.requestId)&&C.rev(facts.requestRevision)&&C.rev(facts.policyRevision),'COMMAND_PAYLOAD_INVALID');
 const path='memberVoluntaryRequests/'+facts.requestId,indexPath='memberVoluntaryIndexes/'+r.target.personId,request=W.validate(get(path)),index=W.index(get(indexPath),r.target.personId),policy=get('memberOperationalConfig/voluntaryExit');
 N(request.personId===r.target.personId&&request.linkId===r.target.linkId&&index.activeRequestId===request.id&&policy?.revision===facts.policyRevision,'AUTH_BINDING_INVALID');
 N(U.revision(r,path)===facts.requestRevision&&request.revision===facts.requestRevision&&U.revision(r,indexPath)===index.revision,'REVISION_CONFLICT');
 N(facts.review.action==='APPROVE','COMMAND_PAYLOAD_INVALID');
 C.exact(facts.externalReview,['presentations','materials','functions','basis']);
 N(['presentations','materials','functions'].every(k=>facts.externalReview[k]==='REVIEWED')&&typeof facts.externalReview.basis==='string'&&facts.externalReview.basis.trim().length>=10&&facts.externalReview.basis.length<=4000,'COMMAND_PAYLOAD_INVALID');
 const target=f4MemberAuthz_().load(get,r.target.personId,f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')));
 // A checkbox cannot contradict a known canonical active function. Its formal
 // closure/substitution must be recorded by the governance command first.
 N(target.M.derive(target.d,r.target.personId,target.h.revision,now).currentFunctions.length===0,'DOMAIN_STATE_INVALID');
 const {proof}=U.authorization(get,r),semester=request.semesterId?get('academicSemesters/'+request.semesterId):null,today=Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd');
 const intent=get('memberPreparedIntents/'+r.operationId);N(intent?.ownerUid===r.ownerUid,'AUTH_BINDING_INVALID');f4MemberActivityChecks_().exit(get,intent.operationalChecks,r.target.personId,today,policy.presentationBlockDays);
 const result=W.review({request,expectedRevision:facts.requestRevision,actorUid:r.ownerUid,actorPersonId:proof.idPessoa,administrative:true,today,now,policy,semester,payload:facts.review});
 N(result.effect&&result.effect.effectiveDate===r.payload.effectiveDate,'COMMAND_PAYLOAD_INVALID');
 return [{path,expectedRevision:request.revision,data:result.request},{path:indexPath,expectedRevision:index.revision,data:{...index,revision:index.revision+1,activeRequestId:null}}];
}
