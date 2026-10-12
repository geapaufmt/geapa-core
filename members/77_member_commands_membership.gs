// E D/E/F: derive one immutable institutional event from a protected, exact act.
function f4MemberMembershipCommands_(){
 const C=f4Core_(),K=f4MemberCommandCatalog_(),N=C.need,U=f4MemberFamilyInputs_();
 const types={'members.discipline.suspend':'SUSPENSAO','members.discipline.dismiss':'DESLIGAMENTO','members.membership.voluntaryExit':'DESLIGAMENTO'};
 function validate(r){
  N(Object.hasOwn(types,r.commandType),'COMMAND_TYPE_INVALID');C.exact(r.target,['personId','linkId']);N(C.id(r.target.personId)&&C.id(r.target.linkId),'COMMAND_PAYLOAD_INVALID');
  C.exact(r.payload,r.commandType==='members.discipline.suspend'?['effectiveDate','endDate']:['effectiveDate']);C.civilStart(r.payload.effectiveDate);
  if(r.commandType==='members.discipline.suspend')f4MemberModel_().interval(r.payload.effectiveDate,r.payload.endDate);
  N(r.evidenceRefs.length>0&&(r.commandType==='members.membership.voluntaryExit'?r.decisionRef===null:C.id(r.decisionRef)),'COMMAND_PAYLOAD_INVALID');return true;
 }
 function resolve(get,r,now){
  validate(r);const {request:q,proof}=U.authorization(get,r),effect=types[r.commandType],voluntary=r.commandType==='members.membership.voluntaryExit';
  N(q.capability===(voluntary?'MEMBERS_VOLUNTARY_EXIT_EXECUTE':'MEMBERS_DISCIPLINE_EXECUTE')&&(voluntary||q.kind==='EXPLICIT'),'AUTH_BINDING_INVALID');
  N(q.target.personId===r.target.personId&&q.target.linkId===r.target.linkId&&q.target.effect===effect&&q.target.effectiveDate===r.payload.effectiveDate,'AUTH_BINDING_INVALID');
  f4MemberActContent_().empty(get,r,now);f4MemberAuthz_().recheck(get,q,proof,now);
  const control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),loaded=f4MemberAuthz_().load(get,r.target.personId,control),M=loaded.M,link=loaded.d['memberLinks/'+r.target.linkId],at=C.civilStart(r.payload.effectiveDate);
  N(link&&link.idPessoa===r.target.personId&&['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE'].includes(link.type),'COMMAND_PAYLOAD_INVALID');const iv=M.interval(link.startDate,link.endDate);N(iv.a<=at&&at<iv.z,'COMMAND_PAYLOAD_INVALID');
  const state=M.derive(loaded.d,r.target.personId,loaded.h.revision,at).state;N(!['DESLIGADO','SEM_VINCULO'].includes(state.state)&&(effect!=='SUSPENSAO'||!state.blocked),'COMMAND_PAYLOAD_INVALID');
  const id='EVT_'+r.operationId,path='memberEvents/'+id;N(!get(path)&&U.revision(r,path)===0,'REVISION_CONFLICT');
  const data={schemaVersion:'m1-memberEvents-v1',id,revision:1,origin:'M13_COMMAND',idPessoa:r.target.personId,idVinculo:r.target.linkId,type:effect,effect,effectiveDate:r.payload.effectiveDate,endDate:effect==='SUSPENSAO'?r.payload.endDate:null,supersedes:null,correctionPath:null,prior:null};M.validate('memberEvents',data);
  N(K.equal([...q.target.resources].sort(),[path,'memberLinks/'+r.target.linkId].sort()),'AUTH_BINDING_INVALID');
  return {id:r.target.personId,expected:U.revision(r,'memberAggregateHeads/'+r.target.personId),changes:[{path,expectedRevision:0,data}],now,boardExpected:{},mode:'M13',authorization:{request:q,proof},deadline:C.ms(proof.expiresAt)};
 }
 return {validate,resolve};
}

// E C: the proposed target is derived from an exact protected act, not an actor.
function f4MemberIngressCommands_(){
 const C=f4Core_(),K=f4MemberCommandCatalog_(),U=f4MemberFamilyInputs_(),N=C.need,F=f4Firestore_(),copy=v=>F.decode(F.encode(v));
 function validate(r){N(r.commandType==='members.ingress.execute','COMMAND_TYPE_INVALID');C.exact(r.target,['personId','linkId']);N(C.id(r.target.personId)&&C.id(r.target.linkId)&&C.id(r.decisionRef)&&r.evidenceRefs.length>0,'COMMAND_PAYLOAD_INVALID');
  C.exact(r.payload,['operation','identity','identifiers','startDate','endDate','originalEventId']);const p=r.payload;N(['RECORD_INGRESS','RECORD_EFFECTIVATION','RETIFICACAO'].includes(p.operation)&&Array.isArray(p.identifiers),'COMMAND_PAYLOAD_INVALID');f4MemberModel_().interval(p.startDate,p.endDate);
  if(p.identity!==null){C.exact(p.identity,['displayName']);N(typeof p.identity.displayName==='string'&&p.identity.displayName.trim().length>0,'COMMAND_PAYLOAD_INVALID');}
  for(const v of p.identifiers){C.exact(v,['id','kind','value']);N(C.id(v.id),'COMMAND_PAYLOAD_INVALID');}
  N(p.operation==='RETIFICACAO'?C.id(p.originalEventId):p.originalEventId===null,'COMMAND_PAYLOAD_INVALID');return true;
 }
 function proposal(get,r,now){validate(r);const p=r.payload,id=r.target.personId,control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),M=f4MemberModel_(true,control),body=f4MemberIngressContent_().validate(get,r,now),h=get('memberAggregateHeads/'+id),expected=U.revision(r,'memberAggregateHeads/'+id),changes=[];
  N((h?.revision||0)===expected,'REVISION_CONFLICT');const identity=get('memberIdentities/'+id);N(identity? p.identity===null&&p.identifiers.length===0:p.operation==='RECORD_INGRESS'&&p.identity!==null&&!h,'COMMAND_PAYLOAD_INVALID');
  const entity=(col,key,fields,old)=>({schemaVersion:'m1-'+col+'-v1',id:key,revision:(old?.revision||0)+1,origin:old?.origin||'M13_COMMAND',...fields});
  function add(col,key,data){const path=col+'/'+key,old=get(path),revision=U.revision(r,path);N((old?.revision||0)===revision,'REVISION_CONFLICT');changes.push({path,expectedRevision:revision,data});}
  if(!identity){add('memberIdentities',id,entity('memberIdentities',id,{idPessoa:id,displayName:p.identity.displayName,technicalAccount:false,active:true}));for(const v of p.identifiers){N(!get('memberIdentifiers/'+v.id),'COMMAND_PAYLOAD_INVALID');add('memberIdentifiers',v.id,entity('memberIdentifiers',v.id,{idPessoa:id,kind:v.kind,value:v.value,active:true}));}}
  const old=get('memberLinks/'+r.target.linkId);N(p.operation==='RETIFICACAO'?old&&old.idPessoa===id:!old,'COMMAND_PAYLOAD_INVALID');
  const original=p.originalEventId&&get('memberEvents/'+p.originalEventId);if(p.operation==='RETIFICACAO')N(original&&original.idPessoa===id&&original.idVinculo===r.target.linkId&&['INGRESSO','EFETIVACAO'].includes(original.effect),'COMMAND_PAYLOAD_INVALID');
  const effect=p.operation==='RECORD_INGRESS'?'INGRESSO':p.operation==='RECORD_EFFECTIVATION'?'EFETIVACAO':original.effect,type=effect==='INGRESSO'?'MEMBRO_INGRESSANTE':'MEMBRO_EFETIVO';
  const event=(key,link,effect,date,extra={})=>entity('memberEvents',key,{idPessoa:id,idVinculo:link,type:effect,effect,effectiveDate:date,endDate:null,supersedes:null,correctionPath:null,prior:null,...extra});
  function correctLink(link,data){add('memberLinks',link.id,data);const eid='COR_'+r.operationId;N(!get('memberEvents/'+eid),'REVISION_CONFLICT');add('memberEvents',eid,event(eid,link.id,null,p.startDate,{type:'RETIFICACAO',correctionPath:'memberLinks/'+link.id,prior:copy(link)}));}
  if(p.operation==='RECORD_EFFECTIVATION'){const ingress=get('memberEvents/'+body.facts.ingress.eventId),link=get('memberLinks/'+ingress.idVinculo);N(link&&link.idPessoa===id&&link.type==='MEMBRO_INGRESSANTE'&&C.civilStart(link.startDate)<C.civilStart(p.startDate),'COMMAND_PAYLOAD_INVALID');const end=new Date(Date.parse(p.startDate+'T12:00:00Z')-86400000).toISOString().slice(0,10);N(link.endDate===null||link.endDate===end,'COMMAND_PAYLOAD_INVALID');correctLink(link,{...copy(link),revision:link.revision+1,endDate:end});}
  const link=entity('memberLinks',r.target.linkId,{idPessoa:id,type,startDate:p.startDate,endDate:p.endDate,exitDate:null},old);if(old)correctLink(old,link);else add('memberLinks',link.id,link);
  const eid='EVT_'+r.operationId;N(!get('memberEvents/'+eid),'REVISION_CONFLICT');add('memberEvents',eid,event(eid,link.id,effect,p.startDate,p.operation==='RETIFICACAO'?{type:'RETIFICACAO',supersedes:original.id}:{}));
  const overlay=Object.fromEntries(changes.map(c=>[c.path,c.data])),paths=[...new Set([...(h?.paths||[]),...changes.map(c=>c.path)])].sort(),future=p=>p==='memberAggregateHeads/'+id?{schemaVersion:'m11-head-v1',id,kind:'PERSON',revision:Math.max(1,expected+1),paths}:Object.hasOwn(overlay,p)?overlay[p]:get(p);
  const target=f4MemberAuthz_().load(future,id,control);M.graph(target.d);return {changes,target,effect,expected};
 }
 function resolve(get,r,now){const p=proposal(get,r,now),{request:q,proof}=U.authorization(get,r);f4MemberAuthz_().recheckCommand(get,q,proof,now,r);return {id:r.target.personId,expected:p.expected,changes:p.changes,now,boardExpected:{},mode:'M13',authorization:{request:q,proof,command:r},deadline:C.ms(proof.expiresAt)};}
 return {validate,resolve,proposal};
}
