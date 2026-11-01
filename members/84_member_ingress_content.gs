// E protected ingress chain. Facts attest evidence; they do not authorize an operator.
function f4MemberIngressContent_(){
 const C=f4Core_(),K=f4MemberCommandCatalog_(),N=f4MemberModel_(true).need;
 function validate(get,r,now){const body=f4MemberActContent_().act(get,r,now),v=body.facts;
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
 return {validate};
}
