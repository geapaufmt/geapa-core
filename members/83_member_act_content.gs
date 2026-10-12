// User-approved E evidence addendum. Read-only protected content; never an approval issuer.
function f4MemberActContent_(){
 const C=f4Core_(),K=f4MemberCommandCatalog_(),N=f4MemberModel_(true).need;
 const digest=v=>C.hash(v);
 function read(get,hash){N(K.hex(hash),'M13_ACT_INVALID');const v=get('memberEvidenceContents/'+hash);N(v&&digest(v)===hash,'M13_ACT_INVALID');return v}
 function material(r){return {commandType:r.commandType,target:r.target,payload:r.payload,expectedRevisions:r.expectedRevisions,evidenceRefs:r.evidenceRefs,decisionRef:r.decisionRef}}
 function act(get,r,now){
  N(r.evidenceRefs.length>0,'M13_ACT_INVALID');let body;
  for(const ref of r.evidenceRefs){const a=f4MemberEvidence_().validate(get('memberEvidenceAttestations/'+ref));N(a.id===ref&&a.active&&a.holderId===r.target.personId&&C.ms(a.verifiedAt)<=now&&now<C.ms(a.expiresAt),'M13_ACT_INVALID');
   const b=read(get,a.contentHash);C.exact(b,['schemaVersion',...Object.keys(material(r)),'facts']);N(b.schemaVersion==='member-act-content-v13'&&K.equal(material(b),material(r)),'M13_ACT_INVALID');
   if(body)N(K.equal(body,b),'M13_ACT_INVALID');body=b;
  }
  return body;
 }
 function association(get,r,now,projectId){const b=act(get,r,now);C.exact(b.facts,['schemaVersion','uid','idPessoa','projectId','verifiedAt','expiresAt','verificationRef']);const v=b.facts;
  N(v.schemaVersion==='member-association-content-v13'&&v.uid===r.target.uid&&v.idPessoa===r.target.personId&&v.projectId===projectId&&(/^demo-geapa-m(04|13)-local$/.test(projectId)||(projectId==='geapa-convergencia-dev'&&f4DevBoundary_().permit({environment:'DEV',projectId,database:'(default)'})))&&C.id(v.verificationRef)&&Number.isSafeInteger(v.verifiedAt)&&Number.isSafeInteger(v.expiresAt)&&v.verifiedAt<=now&&now<v.expiresAt,'M13_ACT_INVALID');return b;
 }
 function empty(get,r,now){const b=act(get,r,now);C.exact(b.facts,[]);return b}
 return {read,digest,material,act,empty,association};
}
