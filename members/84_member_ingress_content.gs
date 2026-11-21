// E protected ingress chain. Facts attest evidence; they do not authorize an operator.
function f4MemberIngressContent_(){
 const C=f4Core_(),K=f4MemberCommandCatalog_(),N=f4MemberModel_(true).need;
 function incoming(get,id,now){
  const loaded=f4MemberAuthz_().load(get,id,f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')));
  const state=loaded.M.derive(loaded.d,id,loaded.h.revision,now).state;
  N(state.state==='MEMBRO_INGRESSANTE'&&!state.blocked,'DOMAIN_STATE_INVALID');
  const links=loaded.M.rows(loaded.d,'memberLinks').filter(link=>{
   const iv=loaded.M.interval(link.startDate,link.endDate);
   return link.type==='MEMBRO_INGRESSANTE'&&iv.a<=now&&now<iv.z;
  });
  N(links.length===1,'RECONCILIATION_REQUIRED');
  const events=loaded.M.rows(loaded.d,'memberEvents'),superseded=new Set(events.map(e=>e.supersedes));
  const entries=events.filter(e=>e.idVinculo===links[0].id&&e.effect==='INGRESSO'&&!superseded.has(e.id));
  N(entries.length===1,'RECONCILIATION_REQUIRED');
  return {link:links[0],event:entries[0]};
 }
 function effectivation(get,r,v,now){
  C.exact(v,['schemaVersion','personId','ingress','recognizedOn','reference']);
  C.exact(v.ingress,['eventId']);
  const today=Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd');
  const actor=get('memberAuthBindings/'+r.ownerUid),old=incoming(get,r.target.personId,now);
  N(actor?.active&&actor.idPessoa!==r.target.personId,'ACCESS_REVOKED');
  N(r.payload.operation==='RECORD_EFFECTIVATION'&&r.payload.startDate===today&&r.payload.endDate===null&&
    v.personId===r.target.personId&&old.event.id===v.ingress.eventId&&v.recognizedOn>=old.link.startDate&&
    C.civilStart(v.recognizedOn)<=C.civilStart(today),'COMMAND_PAYLOAD_INVALID');
  N(typeof v.reference==='string'&&v.reference.length<=500&&!/[<>\x00-\x1f]/.test(v.reference),'COMMAND_PAYLOAD_INVALID');
  if(v.recognizedOn!==today)N(v.reference.trim()&&get('memberOperationalConfig/registration')?.retroactiveEffectivation===true,'HUMAN_DECISION_REQUIRED');
 }
 function validate(get,r,now){const body=f4MemberActContent_().act(get,r,now),v=body.facts;
  if(v.schemaVersion==='member-effectivation-v1'){effectivation(get,r,v,now);return body;}
  if(v.schemaVersion==='member-administrative-ingress-v1')return f4MemberAdministrativeIngress_().validate(get,r,body,now);
  C.exact(v,['schemaVersion','personId','processRef','cohortRef','approval','invitation','acceptance','integration','ingress']);
  N(v.schemaVersion==='member-ingress-content-v13'&&v.personId===r.target.personId&&C.id(v.processRef)&&C.id(v.cohortRef),'M13_ACT_INVALID');
  C.exact(v.approval,['kind','reference']);N(['REGULAR','EXCEPTIONAL'].includes(v.approval.kind)&&C.id(v.approval.reference)&&v.approval.reference===r.decisionRef,'M13_ACT_INVALID');
  C.exact(v.invitation,['reference','personId','processRef','cohortRef']);N(C.id(v.invitation.reference)&&v.invitation.personId===v.personId&&v.invitation.processRef===v.processRef&&v.invitation.cohortRef===v.cohortRef,'M13_ACT_INVALID');
  C.exact(v.acceptance,['reference','personId','invitationRef']);N(C.id(v.acceptance.reference)&&v.acceptance.personId===v.personId&&v.acceptance.invitationRef===v.invitation.reference,'M13_ACT_INVALID');
  C.exact(v.integration,['required','ruleRef','completionRef']);N(typeof v.integration.required==='boolean','M13_ACT_INVALID');
  if(v.integration.required){N(C.id(v.integration.ruleRef)&&(v.integration.completionRef===null||C.id(v.integration.completionRef)),'M13_ACT_INVALID');if(r.payload.operation==='RECORD_EFFECTIVATION')N(C.id(v.integration.completionRef),'M13_ACT_INVALID');}
  else N(v.integration.ruleRef===null&&v.integration.completionRef===null,'M13_ACT_INVALID');
  if(r.payload.operation==='RECORD_INGRESS')N(v.ingress===null,'M13_ACT_INVALID');
  else{C.exact(v.ingress,['eventId','operationId','contentHash']);N(C.id(v.ingress.eventId)&&C.op(v.ingress.operationId)&&K.hex(v.ingress.contentHash),'M13_ACT_INVALID');
   const event=get('memberEvents/'+v.ingress.eventId),ledger=get('commandLedger/'+v.ingress.operationId);N(event&&event.idPessoa===v.personId&&event.effect==='INGRESSO'&&ledger,'M13_ACT_INVALID');
   f4MemberLedgerM13_().ledger(ledger);N(ledger.commandType==='members.ingress.execute'&&ledger.state==='COMPLETED'&&ledger.revisionAfter.some(v=>v.resourceType==='memberEvents'&&v.id===event.id),'M13_ACT_INVALID');
   const original=f4MemberActContent_().read(get,v.ingress.contentHash);N(original.schemaVersion==='member-act-content-v13'&&original.commandType==='members.ingress.execute'&&original.target.personId===v.personId&&original.payload.operation==='RECORD_INGRESS','M13_ACT_INVALID');
   // Link proof back to the protected historical audit, not a client-selected content object.
   const audit=get('memberAudit/'+v.ingress.operationId);N(audit&&audit.resultCode==='OK'&&audit.evidenceRefs.some(ref=>get('memberEvidenceAttestations/'+ref)?.contentHash===v.ingress.contentHash),'M13_ACT_INVALID');
   N(K.equal(original.facts.invitation,v.invitation)&&K.equal(original.facts.acceptance,v.acceptance)&&original.facts.processRef===v.processRef&&original.facts.cohortRef===v.cohortRef,'M13_ACT_INVALID');
   N(original.facts.integration&&(!original.facts.integration.required||v.integration.required),'M13_ACT_INVALID');
   if(r.payload.operation==='RETIFICACAO'&&get('memberEvents/'+r.payload.originalEventId)?.effect==='EFETIVACAO'&&v.integration.required)N(C.id(v.integration.completionRef),'M13_ACT_INVALID');
  }
  return body;
 }
 return {validate,incoming};
}
