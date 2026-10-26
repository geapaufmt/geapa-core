// Private profile extension. Identifiers remain in memberIdentifiers; editing
// contact information never changes the Firebase binding or login identity.
function f4MemberProfileContract_(){
 const C=f4Core_(),N=C.need,E=f4MemberEnrollmentContract_();
 const editable=['contactEmail','phone','preferredName','instagram','lattes','orcid','linkedin','website','photoUrl','academicBio','interestAxisIds','optionalCommunications','stateCode','city','municipalityCode'];
 const sensitive=['birthDate','countryCode','courseId','birthCity','sex'];
 const extra=[...editable.filter(k=>k!=='contactEmail'),...sensitive];
 const readFields=['idPessoa','displayLabel','profileDetails'];
 function patch(raw,self,today){
  N(raw&&Object.getPrototypeOf(raw)===Object.prototype&&Object.keys(raw).length>0,'COMMAND_PAYLOAD_INVALID');
  const allowed=self?editable:[...editable,...sensitive,'academicReference','displayName'];
  N(Object.keys(raw).every(k=>allowed.includes(k)),'COMMAND_PAYLOAD_INVALID');
  const out={},enrollment={};
  for(const [k,v]of Object.entries(raw)){
   if(['contactEmail','phone','preferredName','academicReference','displayName'].includes(k)){
    N(typeof v==='string'&&v.length<=(k==='contactEmail'?180:120)&&!/[\x00-\x1f<>]/.test(v),'COMMAND_PAYLOAD_INVALID');
    out[k]=v.trim();
    if(k==='contactEmail')N(/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v),'COMMAND_PAYLOAD_INVALID');
    if(k==='phone'&&v)N(/^\+[1-9][0-9]{9,14}$/.test(v),'COMMAND_PAYLOAD_INVALID');
    if(k==='displayName')N(v.trim().length>0,'COMMAND_PAYLOAD_INVALID');
   }else enrollment[k]=v;
  }
  Object.assign(out,E.profile(enrollment,today));return out;
 }
 function validate(v){
  const keys=['schemaVersion','id','revision','origin','idPessoa','contactEmail','academicReference',...extra.filter(k=>Object.hasOwn(v,k))];
  C.exact(v,keys);N(v.schemaVersion==='m1-memberPrivateProfiles-v2'&&C.id(v.id)&&v.id===v.idPessoa&&C.rev(v.revision)&&v.revision>0&&C.id(v.origin),'M1_SCHEMA_INVALID');
  N(typeof v.contactEmail==='string'&&v.contactEmail.length<=180&&typeof v.academicReference==='string'&&v.academicReference.length<=120,'M1_SCHEMA_INVALID');
  const p=Object.fromEntries(extra.filter(k=>Object.hasOwn(v,k)).map(k=>[k,v[k]]));
  if(Object.keys(p).length)patch(p,false,'9999-12-31');return v;
 }
 function detail(d,id,get=()=>null){
  const identity=d['memberIdentities/'+id],p=d['memberPrivateProfiles/'+id];N(identity,'AUTH_BINDING_INVALID');
  const ids=Object.entries(d).filter(([path,v])=>path.startsWith('memberIdentifiers/')&&v.idPessoa===id&&v.active).map(([,v])=>v);
  const one=kind=>{const values=[...new Set(ids.filter(v=>v.kind===kind).map(v=>v.value))];N(values.length<=1,'RECONCILIATION_REQUIRED');return values[0]||null;};
  const fields={};for(const k of ['contactEmail','academicReference',...extra])fields[k]=p&&Object.hasOwn(p,k)?p[k]:null;
  const missing=['cpf','birthDate','countryCode','city','courseId'].filter(k=>k==='cpf'?!one('CPF'):!fields[k]);if(fields.countryCode==='BR'&&!fields.stateCode)missing.push('stateCode');
  const course=fields.courseId?get('memberCourseCatalog/'+fields.courseId):null;
  return {schemaVersion:'member-own-profile-details-v1',profileRevision:p?.revision||0,personRevision:null,cpf:one('CPF'),rga:one('RGA'),fields,missing,courseLabel:course?.id===fields.courseId&&typeof course.name==='string'?course.name:null};
 }
 return {editable,sensitive,extra,readFields,patch,validate,detail};
}
