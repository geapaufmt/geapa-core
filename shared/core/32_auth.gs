// F4A: native synchronous port of F3.5 auth/bindings/index.cjs. No external module loading.
function f4Auth_() {
  if (f4Auth_.value_) return f4Auth_.value_;
  return (f4Auth_.value_ = (function () {
'use strict';
const C=f4Core_();
const {validateProjection}=f4Validation_();
function authorize(get,uid,scope,cap,now){
 C.need(C.id(uid),'AUTH_BINDING_INVALID');const b=get('memberAuthBindings/'+uid);
 if(b?.domainVersion==='m1'||b?.domainVersion==='m11')return f4MemberRepository_().authorize(get,uid,scope,cap,now);
 if(b?.idPessoa&&get('memberIdentities/'+b.idPessoa))C.fail('AUTH_BINDING_INVALID');
 C.need(b&&b.schemaVersion==='binding-v1'&&b.uid===uid&&b.active===true&&b.policyVersion===1&&C.id(b.idPessoa)&&C.rev(b.bindingVersion)&&b.bindingVersion>0,'AUTH_BINDING_INVALID');
 const c=get('memberCurrentStates/'+b.idPessoa),p=get('portalUsers/'+uid);
 C.need(c&&c.schemaVersion==='member-state-v1'&&c.idPessoa===b.idPessoa&&c.policyVersion===1&&c.blocked===false&&c.stale===false&&C.rev(c.sourceRevision)&&C.rev(c.accessEpoch)&&c.accessEpoch>0&&['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE','EGRESSO','ADMIN_TECNICO'].includes(c.state)&&C.time(c.startsAt,c.validUntil,now),'ACCESS_REVOKED');
 try{validateProjection('portal',p)}catch{C.fail('ACCESS_REVOKED')}
 C.need(p.uid===uid&&p.idPessoa===b.idPessoa&&p.ativo===true&&p.podeAcessarPortal===true&&p.stale===false&&p.policyVersion===1&&p.bindingVersion===b.bindingVersion&&p.accessEpoch===c.accessEpoch&&p.sourceRevision===c.sourceRevision&&C.time(p.startsAt,p.validUntil,now)&&C.ms(p.startsAt)>=C.ms(c.startsAt)&&C.ms(p.validUntil)<=C.ms(c.validUntil),'ACCESS_REVOKED');
 const priv=['JUSTIFICATION_REVIEW','PUBLIC_CONTENT_MANAGE'].includes(cap),g=get(priv?'memberFunctions/'+uid+'__'+scope:'capabilityGrants/'+uid+'/scopes/'+scope);
 C.need(g,'CAPABILITY_REQUIRED');C.need(g.scopeKey===scope,'SCOPE_INVALID');
 C.need((g.schemaVersion==='capability-grant-v1'||(priv&&g.schemaVersion==='member-function-v1'))&&g.uid===uid&&g.idPessoa===b.idPessoa&&g.bindingVersion===b.bindingVersion&&g.accessEpoch===c.accessEpoch&&g.policyVersion===1&&g.stateRevision===c.sourceRevision,'ACCESS_REVOKED');
 C.need(Array.isArray(g.capabilities)&&g.capabilities.length<=8&&g.capabilities.includes('PORTAL_ACCESS')&&g.capabilities.includes(cap),'CAPABILITY_REQUIRED');
 C.need(C.time(g.startsAt,g.endsAtExclusive,now)&&C.ms(g.endsAtExclusive)<=C.ms(p.validUntil),'FUNCTION_EXPIRED');
 const allowed=p.technicalAccount?['PORTAL_ACCESS','TECH_HEALTH_READ']:c.state==='MEMBRO_EFETIVO'?null:c.state==='MEMBRO_INGRESSANTE'?['PORTAL_ACCESS','SELF_STATUS_READ']:c.state==='EGRESSO'?['PORTAL_ACCESS','SELF_STATUS_READ','EGRESS_HISTORY_READ']:['PORTAL_ACCESS','TECH_HEALTH_READ','PUBLIC_CONTENT_MANAGE'];
 C.need(!allowed||allowed.includes(cap),'CAPABILITY_REQUIRED');
 if(priv){C.need(g.sourceFunctionId===uid+'__'+scope&&g.active===true&&C.rev(g.revision)&&g.sourceRevision===g.revision,'CAPABILITY_REQUIRED')}
 else if(g.sourceFunctionId===null){C.need(g.sourceRevision===c.sourceRevision,'ACCESS_REVOKED')}
 else{C.need(cap!=='JUSTIFICATION_SUBMIT_SELF'&&C.id(g.sourceFunctionId),'CAPABILITY_REQUIRED');const f=get('memberFunctions/'+g.sourceFunctionId);C.need(f&&f.schemaVersion==='member-function-v1'&&f.active===true&&f.uid===uid&&f.idPessoa===b.idPessoa&&f.policyVersion===1&&f.sourceRevision===g.sourceRevision&&f.scopeKey===scope&&Array.isArray(f.capabilities)&&g.capabilities.every(x=>f.capabilities.includes(x))&&C.time(f.startsAt,f.endsAtExclusive,now)&&C.ms(g.startsAt)>=C.ms(f.startsAt)&&C.ms(g.endsAtExclusive)<=C.ms(f.endsAtExclusive),'CAPABILITY_REQUIRED')}
 const account=get('institutionalAccounts/'+b.idPessoa);C.need(account&&typeof account.human==='boolean'&&account.sourceRevision===c.sourceRevision&&account.human===!p.technicalAccount,'ACCESS_REVOKED');return {b,c,p,g,account};
}
// Explicit internal M13 entry. Legacy authorize is byte-preserved; errors never fall back.
function authorizeM13(get,request,now){return f4MemberAuthz_().authorize(get,request,now)}
return {authorize,authorizeM13};




  })());
}
