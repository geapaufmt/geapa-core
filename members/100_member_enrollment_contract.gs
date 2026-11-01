// Pure enrollment rules for the existing Members command family. No writer,
// permission grant, mail send, binding creation or membership activation here.
// Ordinary registration is a draft; the existing ingress kernel remains the
// sole producer of a membership event after these facts have been validated.
function f4MemberEnrollmentContract_(){
 const C=f4Core_(),N=C.need,K=f4MemberCommandCatalog_();
 const required=['cpf','birthDate','countryCode','city','courseId'];
 const fields=[...required,'stateCode','municipalityCode','displayName','instagram','lattes','orcid','linkedin','website','photoUrl','academicBio','interestAxisIds','optionalCommunications','birthCity','sex'];
 const acknowledgement='Declaro que tive acesso ao Regimento Interno e às Normas Complementares vigentes do GEAPA, estou ciente de suas disposições e comprometo-me a cumpri-las.';
 function text(v,max,empty=true){N(typeof v==='string'&&v.length<=max&&!/[\x00-\x08\x0b\x0c\x0e-\x1f<>]/.test(v),'COMMAND_PAYLOAD_INVALID');const s=v.trim();N(empty||s.length>0,'COMMAND_PAYLOAD_INVALID');return s;}
 function cpf(value){
  const n=text(value,20,false).replace(/[.\s-]/g,'');N(/^\d{11}$/.test(n)&&!(/^(\d)\1{10}$/).test(n),'COMMAND_PAYLOAD_INVALID');
  // Same check-digit algorithm as corePerfilNormalizeCpf_ in the existing Core.
  function digit(base,weight){let sum=0;for(let i=0;i<base.length;i++)sum+=Number(base[i])*(weight-i);const rest=(sum*10)%11;return rest===10?0:rest;}
  N(digit(n.slice(0,9),10)===Number(n[9])&&digit(n.slice(0,10),11)===Number(n[10]),'COMMAND_PAYLOAD_INVALID');return n;
 }
 function quick(value){
  C.exact(value,['fullName','email','rga','phone']);
  const fullName=text(value.fullName,180,false).replace(/\s+/g,' '),email=text(value.email,180,false).toLowerCase(),rga=text(value.rga,40,false).replace(/[.\s-]/g,'');
  N(fullName.length>=5&&/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)&&/^\d{5,30}$/.test(rga),'COMMAND_PAYLOAD_INVALID');
  let phone=text(value.phone,40,false).replace(/[\s()+.-]/g,'');N(/^\d{10,15}$/.test(phone),'COMMAND_PAYLOAD_INVALID');
  // Local Brazilian numbers follow the pre-existing registration convention.
  if(phone.length===10||phone.length===11)phone='55'+phone;
  return {fullName,email,rga,phone:'+'+phone};
 }
 function profile(p,today){
  N(p&&Object.getPrototypeOf(p)===Object.prototype&&Object.keys(p).every(k=>fields.includes(k)),'COMMAND_PAYLOAD_INVALID');
  const out={};for(const [key,v]of Object.entries(p)){
   if(key==='interestAxisIds'){N(Array.isArray(v)&&v.length<=20&&v.every(C.id)&&new Set(v).size===v.length,'COMMAND_PAYLOAD_INVALID');out[key]=v.slice();continue;}
   if(key==='optionalCommunications'){N(typeof v==='boolean','COMMAND_PAYLOAD_INVALID');out[key]=v;continue;}
   const s=text(v,key==='academicBio'?3000:key.endsWith('Url')||['lattes','linkedin','website'].includes(key)?2000:180);
   if(key==='cpf'&&s)out[key]=cpf(s);else out[key]=s;
   if(key==='birthDate'&&s){C.civilStart(s);N(s>='1900-01-01'&&s<=today,'COMMAND_PAYLOAD_INVALID');}
   if(key==='countryCode'&&s)N(/^[A-Z]{2}$/.test(s),'COMMAND_PAYLOAD_INVALID');
   // Optional private historical facts. Birthplace is distinct from the broader city-of-origin field;
   // country of origin is not inferred from a Brazilian state or nationality.
   if(key==='sex'&&s)N(s.length<=60,'COMMAND_PAYLOAD_INVALID');
   if(key==='stateCode'&&s)N(['AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS','MG','PA','PB','PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO'].includes(s),'COMMAND_PAYLOAD_INVALID');
   if(key==='municipalityCode'&&s)N(/^\d{7}$/.test(s),'COMMAND_PAYLOAD_INVALID');
   if(key==='courseId'&&s)N(C.id(s),'COMMAND_PAYLOAD_INVALID');
   if(['lattes','linkedin','website','photoUrl'].includes(key)&&s)N(/^https:\/\/[^\s/@]+(?:\/[^\s]*)?$/.test(s),'COMMAND_PAYLOAD_INVALID');
   if(key==='orcid'&&s)N(/^\d{4}-\d{4}-\d{4}-[\dX]{4}$/.test(s),'COMMAND_PAYLOAD_INVALID');
   if(key==='instagram'&&s)N(/^@?[a-zA-Z0-9._]{1,30}$/.test(s),'COMMAND_PAYLOAD_INVALID');
  }
  return out;
 }
 function completeness(get,p,today){
  const v=profile(p,today),missing=required.filter(k=>!v[k]);
  if(v.countryCode==='BR'&&!v.stateCode)missing.push('stateCode');
  if(v.countryCode==='BR'&&v.stateCode){N(['AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS','MG','PA','PB','PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO'].includes(v.stateCode),'COMMAND_PAYLOAD_INVALID');}
  if(v.municipalityCode){const city=get('memberMunicipalityCatalog/'+v.municipalityCode);N(city?.active&&city.countryCode===v.countryCode&&city.stateCode===v.stateCode&&city.name===v.city,'COMMAND_PAYLOAD_INVALID');}
  let course=null;if(v.courseId){course=get('memberCourseCatalog/'+v.courseId);N(course?.schemaVersion==='member-course-v1'&&course.id===v.courseId&&course.active===true&&C.rev(course.revision)&&course.revision>0,'COMMAND_PAYLOAD_INVALID');}
  // Academic institution/campus/level come from the course catalog. Semester
  // derived from RGA may be unknown and never blocks enrollment.
  return {complete:missing.length===0,missing,profile:v,course};
 }
 function identity(invite,verified,now,projectId){
  N(projectId===f4ExecutionEnvironment_().project&&verified?.projectId===projectId&&verified.emailVerified===true&&C.id(verified.uid)&&Number.isSafeInteger(verified.expiresAt)&&now<verified.expiresAt,'AUTH_BINDING_INVALID');
  N(typeof verified.email==='string'&&verified.email.trim().toLowerCase()===invite.registration.email,'AUTH_BINDING_INVALID');
  N(invite.uid===null||invite.uid===verified.uid,'AUTH_BINDING_INVALID');
  N(invite.status!=='CANCELLED'&&now<invite.expiresAt,'ACCESS_REVOKED');return verified.uid;
 }
 function create({id,personId,registration,approval,actorUid,now,expiresAt}){
  N(C.op(id)&&C.id(personId)&&C.id(actorUid)&&Number.isSafeInteger(now)&&Number.isSafeInteger(expiresAt)&&expiresAt>now,'COMMAND_PAYLOAD_INVALID');
  C.exact(approval,['reference','approvedBy','approvedOn','membershipType']);
  N(C.id(approval.reference)&&C.id(approval.approvedBy)&&['MEMBRO_INGRESSANTE','MEMBRO_EFETIVO'].includes(approval.membershipType),'COMMAND_PAYLOAD_INVALID');
  C.civilStart(approval.approvedOn);N(C.civilStart(approval.approvedOn)<=now,'COMMAND_PAYLOAD_INVALID');
  return {schemaVersion:'member-enrollment-v1',id,personId,revision:1,registration:quick(registration),approval:{...approval},createdBy:actorUid,createdAt:now,updatedAt:now,expiresAt,uid:null,status:'PRE_REGISTERED',invitedAt:null,respondedAt:null,completedAt:null,profile:{},normAcknowledgement:null,imageManifestation:null};
 }
 function validate(v){
  C.exact(v,['schemaVersion','id','personId','revision','registration','approval','createdBy','createdAt','updatedAt','expiresAt','uid','status','invitedAt','respondedAt','completedAt','profile','normAcknowledgement','imageManifestation']);
  N(v.schemaVersion==='member-enrollment-v1'&&C.op(v.id)&&C.id(v.personId)&&C.id(v.createdBy)&&C.rev(v.revision)&&v.revision>0&&['PRE_REGISTERED','INVITED','ACCEPTED','DECLINED','CANCELLED','COMPLETED'].includes(v.status),'SCHEMA_INVALID');
  N(K.equal(quick(v.registration),v.registration)&&Number.isSafeInteger(v.createdAt)&&Number.isSafeInteger(v.updatedAt)&&v.createdAt<=v.updatedAt&&Number.isSafeInteger(v.expiresAt)&&v.expiresAt>v.createdAt,'SCHEMA_INVALID');
  C.exact(v.approval,['reference','approvedBy','approvedOn','membershipType']);N(C.id(v.approval.reference)&&C.id(v.approval.approvedBy)&&['MEMBRO_INGRESSANTE','MEMBRO_EFETIVO'].includes(v.approval.membershipType)&&C.civilStart(v.approval.approvedOn)<=v.createdAt,'SCHEMA_INVALID');
  N(v.uid===null||C.id(v.uid),'SCHEMA_INVALID');for(const k of ['invitedAt','respondedAt','completedAt'])N(v[k]===null||Number.isSafeInteger(v[k])&&v.createdAt<=v[k]&&v[k]<=v.updatedAt,'SCHEMA_INVALID');
  N(!['ACCEPTED','DECLINED','COMPLETED'].includes(v.status)||v.uid!==null&&v.respondedAt!==null&&v.invitedAt!==null,'SCHEMA_INVALID');
  N(v.status!=='PRE_REGISTERED'||v.invitedAt===null&&v.uid===null&&v.respondedAt===null&&v.completedAt===null,'SCHEMA_INVALID');
  N(v.status!=='INVITED'||v.invitedAt!==null&&v.uid===null&&v.respondedAt===null,'SCHEMA_INVALID');
  N(v.status==='COMPLETED'?v.completedAt!==null:v.completedAt===null,'SCHEMA_INVALID');
  profile(v.profile,'9999-12-31');return v;
 }
 function response(old,expectedRevision,action,verified,now,projectId){
  N(old.revision===expectedRevision,'REVISION_CONFLICT');N(['ACCEPT','DECLINE'].includes(action)&&['INVITED','ACCEPTED'].includes(old.status)&&old.invitedAt!==null,'DOMAIN_STATE_INVALID');
  const uid=identity(old,verified,now,projectId);N(old.status!=='ACCEPTED'||action==='ACCEPT','DOMAIN_STATE_INVALID');
  return {...old,uid,status:action==='ACCEPT'?'ACCEPTED':'DECLINED',respondedAt:old.respondedAt||now,revision:old.revision+1,updatedAt:now};
 }
 function save(old,expectedRevision,patch,verified,now,projectId,today){
  N(old.revision===expectedRevision,'REVISION_CONFLICT');const uid=identity(old,verified,now,projectId);N(old.uid===uid&&old.status==='ACCEPTED','ACCESS_REVOKED');
  // Identifiers provided at pre-registration are deliberately absent from the
  // editable patch. Corrections go through the existing administrative flow.
  return {...old,revision:old.revision+1,profile:{...old.profile,...profile(patch,today)},updatedAt:now};
 }
 function normReceipt({personId,enrollmentId,uid,catalog,accepted,now,termVersion}){
  N(accepted===true&&C.id(personId)&&C.op(enrollmentId)&&C.id(uid)&&C.id(termVersion)&&Array.isArray(catalog)&&catalog.length>0,'COMMAND_PAYLOAD_INVALID');
  const documents=catalog.map(v=>{
   N(v?.status==='CURRENT'&&v.appliesToEnrollment===true&&C.id(v.id)&&C.id(v.version)&&C.id(v.approvalRef)&&typeof v.immutableVersionRef==='string'&&v.immutableVersionRef.length>0&&/^https:\/\//.test(v.pdfUrl),'DOMAIN_STATE_INVALID');
   N(C.civilStart(v.effectiveFrom)<=now,'DOMAIN_STATE_INVALID');return {id:v.id,version:v.version,immutableVersionRef:v.immutableVersionRef,contentHash:v.contentHash||null};
  });N(new Set(documents.map(v=>v.id)).size===documents.length,'DOMAIN_STATE_INVALID');
  return {schemaVersion:'member-norm-acknowledgement-v1',personId,enrollmentId,uid,at:now,termVersion,text:acknowledgement,documents};
 }
 function imageReceipt({personId,enrollmentId,uid,choice,term,now,previous}){
  N(C.id(personId)&&C.op(enrollmentId)&&C.id(uid)&&['GRANTED','DECLINED','REVOKED'].includes(choice)&&C.id(term?.version),'COMMAND_PAYLOAD_INVALID');
  N(choice!=='REVOKED'||previous?.state==='GRANTED','DOMAIN_STATE_INVALID');
  if(previous)N(previous.personId===personId&&previous.uid===uid&&previous.enrollmentId===enrollmentId,'AUTH_BINDING_INVALID');
  return {schemaVersion:'member-image-manifestation-v1',personId,enrollmentId,uid,at:now,state:choice,termVersion:term.version,text:text(term.text,3000,false),purpose:text(term.purpose,1000,false),previousHash:previous?C.hash(previous):null};
 }
 function ready(get,old,expectedRevision,verified,now,projectId,today,norms){
  N(old.revision===expectedRevision,'REVISION_CONFLICT');const uid=identity(old,verified,now,projectId);N(old.uid===uid&&old.status==='ACCEPTED'&&old.respondedAt!==null,'ACCESS_REVOKED');
  const state=completeness(get,old.profile,today);N(state.complete,'COMMAND_PAYLOAD_INVALID');
  const claim=get('memberIdentifierClaims/'+C.hash({kind:'CPF',value:state.profile.cpf}));N(!claim||claim.active===true&&claim.idPessoa===old.personId,'RECONCILIATION_REQUIRED');
  const receipt=old.normAcknowledgement;N(receipt&&receipt.personId===old.personId&&receipt.enrollmentId===old.id&&receipt.uid===uid,'AUTH_BINDING_INVALID');
  const expected=normReceipt({personId:old.personId,enrollmentId:old.id,uid,catalog:norms,accepted:true,now:receipt.at,termVersion:receipt.termVersion});
  N(K.equal(expected,receipt)&&receipt.at>=old.respondedAt&&receipt.at<=now,'AUTH_BINDING_INVALID');
  // Returns protected ingress facts, never active membership. The caller must
  // commit them together with the existing ingress/binding kernel and receipt.
  return {personId:old.personId,uid,enrollmentId:old.id,membershipType:old.approval.membershipType,effectiveDate:today,completedAt:now,courseId:state.course.id,courseRevision:state.course.revision,profile:state.profile,normAcknowledgement:receipt};
 }
 return {quick,cpf,profile,completeness,identity,create,validate,response,save,normReceipt,imageReceipt,ready,acknowledgement};
}
