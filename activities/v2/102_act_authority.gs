// Core read facade. Identity and restrictions come from the current shared M13 authority.
function f4ActivitiesAuthority_(resolveRequest){
 const M=f4ActivitiesModel_(),N=M.need,A=f4MemberAuthz_();
 function identity(get,uid,now){const a=A.identify(get,uid,now,s=>s.state==='SUSPENSO'&&s.blocked);N(a.control.activeWriterVersion==='M13','AUTHORIZATION_DENIED');const b=a.binding;return {b,h:a.actor.h,s:a.state,i:a.person,actor:{kind:'PORTAL_AUTHENTICATED_MEMBER',uid,personRef:M.ref('person',b.idPessoa),personId:b.idPessoa,bindingVersion:b.bindingVersion},authorityBasis:{personRevision:a.actor.h.revision,bindingVersion:b.bindingVersion,state:a.state.state,authorityEpoch:a.control.epoch,policyEpoch:a.control.policyEpoch,subjectEpoch:a.subject.authzEpoch}};}
 function ordinary(get,uid,scope,cap,now){
  const base=identity(get,uid,now);N(base.s.state==='MEMBRO_EFETIVO'&&!base.s.blocked,'AUTHORIZATION_DENIED');
  if(scope==='SELF'&&cap==='PRESENTATION_SELF_READ')return {g:{capabilities:[cap],entries:{[cap]:base.authorityBasis}}};
  N(typeof resolveRequest==='function','AUTHORIZATION_DENIED','CURRENT_AUTHORIZATION_REQUIRED');const q=resolveRequest(get,{uid,scope,capability:cap},now);N(q?.schemaVersion==='act-authz-request-v13'&&q.uid===uid&&q.capability===cap&&q.target.scopeKey===scope,'AUTHORIZATION_DENIED');const v=f4SharedAuthz_().authorize(get,q,now);return {g:{capabilities:[cap],entries:{[cap]:v.proof}}};
 }
 function suspension(get,base,now){const a=identity(get,base.b.uid,now);N(a.s.state==='SUSPENSO'&&a.s.blocked&&a.h.revision===base.h.revision,'AUTHORIZATION_DENIED','RESTRICTION_UNPROVEN');return {kind:'MEMBROS_SUSPENSION_ONLY',personRevision:a.h.revision};}
 return {identity,ordinary,suspension};
}

// Called by the current M13 mutation/renew planner before its write phase.
// Global epoch makes an incomplete/stale capacity assessment unusable without roster fanout.
function f4ActivitiesMembersFence_(get,pending,personId,now){
 const M=f4ActivitiesModel_(),key='presentationCapacityHeads/registry',old=get(key);
 if(!old)return; // Domain has not been initialized. First ACT cycle creates its registry.
 pending[key]={...old,revision:old.revision+1,membersEpoch:old.membersEpoch+1,updatedAt:M.stamp(now)};
 if(personId){const key2='presentationMemberViews/'+personId,h=get(key2);if(h)pending[key2]={...h,state:'INVALIDATED',viewEpoch:h.viewEpoch+1,rebuildDueAt:M.stamp(now)};}
}
