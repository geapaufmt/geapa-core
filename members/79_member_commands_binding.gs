// E J/K1. Technical association content proves the target, never the operator's power.
function f4MemberBindingCommands_(config){
 const C=f4Core_(),K=f4MemberCommandCatalog_(),U=f4MemberFamilyInputs_(),N=C.need;
 C.destination(config);N((config.environment==='SYNTHETIC_LOCAL' || f4OperationalDestination_(config)),'DESTINATION_INVALID');const projectId=config.projectId;
 function validate(r){N(['members.binding.associate','members.binding.revoke'].includes(r.commandType),'COMMAND_TYPE_INVALID');C.exact(r.target,['personId','uid']);N(C.id(r.target.personId)&&C.id(r.target.uid)&&r.evidenceRefs.length>0,'COMMAND_PAYLOAD_INVALID');
  if(r.commandType==='members.binding.associate')C.exact(r.payload,[]);else{C.exact(r.payload,['reason']);N(typeof r.payload.reason==='string'&&r.payload.reason.trim().length>0&&r.payload.reason.length<=120&&!/[\x00-\x1f<>]/.test(r.payload.reason),'COMMAND_PAYLOAD_INVALID');}return true;
 }
 function resolve(get,r,now){validate(r);const {request:q,proof}=U.authorization(get,r),associate=r.commandType==='members.binding.associate';N(q.kind==='EXPLICIT'&&q.capability==='MEMBERS_BINDING_MANAGE'&&q.target.personId===r.target.personId,'AUTH_BINDING_INVALID');f4MemberAuthz_().recheck(get,q,proof,now);
  if(associate)f4MemberActContent_().association(get,r,now,projectId);else f4MemberActContent_().empty(get,r,now);
  const path='memberAuthBindings/'+r.target.uid,old=get(path),revision=U.revision(r,path);N((old?.revision||0)===revision,'REVISION_CONFLICT');N(!old||old.idPessoa===r.target.personId&&old.uid===r.target.uid,'AUTH_BINDING_INVALID');
  const control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),loaded=f4MemberAuthz_().load(get,r.target.personId,control),person=loaded.d['memberIdentities/'+r.target.personId];
  N(!person.technicalAccount,'AUTH_BINDING_INVALID');if(associate){const state=loaded.M.derive(loaded.d,r.target.personId,loaded.h.revision,now).state;N(person.active&&!state.blocked&&['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE','EGRESSO'].includes(state.state),'AUTH_BINDING_INVALID');}else N(old&&old.active,'AUTH_BINDING_INVALID');
  const resources=[path,'portalUsers/'+r.target.uid,'portalAccessStatus/'+r.target.uid];N(K.equal([...q.target.resources].sort(),resources.sort()),'AUTH_BINDING_INVALID');
  const data={schemaVersion:'binding-v1',id:r.target.uid,revision:revision+1,origin:old?.origin||'M13_COMMAND',uid:r.target.uid,idPessoa:r.target.personId,active:associate,bindingVersion:(old?.bindingVersion||0)+1,policyVersion:1,sourceRevision:revision+1,createdAt:old?.createdAt??now,updatedAt:now,domainVersion:'m11'};
  f4MemberModel_(true).validate('memberAuthBindings',data);
  return {id:r.target.personId,expected:U.revision(r,'memberAggregateHeads/'+r.target.personId),changes:[{path,expectedRevision:revision,data}],now,boardExpected:{},mode:'M13',authorization:{request:q,proof},deadline:C.ms(proof.expiresAt)};
 }
 return {validate,resolve};
}
