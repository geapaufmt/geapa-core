// Registration adapter: reuse canonical identifier claims, never a name match.
// This read-only helper neither authorizes nor creates a person. The existing
// ingress command must re-read these claims in its atomic mutation plan.
// Parity source: geapa-membros 16_members_ingress_v2.gs (local 47f3689).
function f4MemberRegistrationIdentity_(){
 const C=f4Core_(),M=f4MemberModel_(true),N=C.need;
 function resolve(get,input){
  C.exact(input,['email','rga']);
  N(typeof get==='function'&&typeof input.email==='string'&&typeof input.rga==='string','COMMAND_PAYLOAD_INVALID');
  const email=input.email.trim().toLowerCase(),rga=input.rga.replace(/[.\s-]/g,'');
  N(email.length<=180&&/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)&&/^[0-9]{5,30}$/.test(rga),'COMMAND_PAYLOAD_INVALID');
  const refs=[],people=new Set();
  for(const [kind,value] of [['EMAIL',email],['RGA',rga]]){
   const path='memberIdentifierClaims/'+C.hash({kind,value}),claim=get(path);
   refs.push(path);
   if(!claim)continue;
   N(claim.active===true&&C.id(claim.idPessoa)&&C.id(claim.identifierId),'RECONCILIATION_REQUIRED');
   const sourcePath='memberIdentifiers/'+claim.identifierId,source=get(sourcePath);
   N(source,'RECONCILIATION_REQUIRED');M.validate('memberIdentifiers',source);
   N(source.active&&source.idPessoa===claim.idPessoa&&source.kind===kind&&source.value===value,'RECONCILIATION_REQUIRED');
   people.add(claim.idPessoa);refs.push(sourcePath);
  }
  N(people.size<=1,'RECONCILIATION_REQUIRED');
  const personId=people.size?[...people][0]:null;
  if(personId){
   const path='memberIdentities/'+personId,person=get(path),headPath='memberAggregateHeads/'+personId,head=get(headPath);
   N(person&&head,'RECONCILIATION_REQUIRED');M.validate('memberIdentities',person);
   N(person.active&&!person.technicalAccount&&head.id===personId&&Array.isArray(head.paths)&&head.paths.includes(path)&&refs.filter(p=>p.startsWith('memberIdentifiers/')).every(p=>head.paths.includes(p)),'RECONCILIATION_REQUIRED');
   refs.push(path,headPath);
  }
  // NEW_CANDIDATE is not permission to write: allocate server-side and let
  // memberIdentifierClaims serialize concurrent admissions at command commit.
  return {kind:personId?'REUSE':'NEW_CANDIDATE',personId,identifiers:[{kind:'EMAIL',value:email},{kind:'RGA',value:rga}],readPaths:refs};
 }
 return {resolve};
}
