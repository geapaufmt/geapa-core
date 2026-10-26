// Administrative admission of a person already approved in a completed
// selection process. An invitation/acceptance is never fabricated. Those facts
// remain mandatory on the existing onboarding/effectivation path.
function f4MemberAdministrativeIngress_(){
 const C=f4Core_(),N=C.need,K=f4MemberCommandCatalog_();
 function value(v,max=300){N(typeof v==='string'&&v.length<=max,'COMMAND_PAYLOAD_INVALID');return v.trim();}
 function normalize(p,today){
  C.exact(p,['modalidadeCadastro','nomeCompleto','nomeExibicao','emailPrincipal','telefone','rga','cursoId','cursoNomeOutro','dataIngresso','semestreEntrada','formaIngresso','origemCadastro','documentoReferencia','observacaoAdministrativa','dadosComplementares']);
  N(['RAPIDO','COMPLETO'].includes(p.modalidadeCadastro)&&p.formaIngresso==='PROCESSO_SELETIVO'&&['CADASTRO_INTERESSE','CADASTRO_MANUAL','SELETIVO_FECHADO'].includes(p.origemCadastro),'COMMAND_PAYLOAD_INVALID');
  const nome=value(p.nomeCompleto,180).replace(/\s+/g,' '),email=value(p.emailPrincipal,180).toLowerCase(),rga=value(p.rga,60).replace(/[.\s-]/g,'');
  N(nome.length>=5&&/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)&&/^[0-9]{5,30}$/.test(rga),'COMMAND_PAYLOAD_INVALID');
  const reference=value(p.documentoReferencia),note=value(p.observacaoAdministrativa,1500);N(reference||note.length>=20,'COMMAND_PAYLOAD_INVALID');
  C.civilStart(p.dataIngresso);N(p.dataIngresso<=today&&C.id(p.cursoId)&&value(p.semestreEntrada,100).length>0,'COMMAND_PAYLOAD_INVALID');
  let phone=value(p.telefone,50).replace(/\D/g,'');if(phone){if((phone.length===12||phone.length===13)&&phone.startsWith('55'))phone=phone.slice(2);N(/^\d{10}$|^\d{2}9\d{8}$/.test(phone),'COMMAND_PAYLOAD_INVALID');phone='+55'+phone;}
  const d=p.dadosComplementares;C.exact(d,['dataNascimento','instagram','paisOrigemCodigo','municipioOrigemCodigo','cidadeNatal','ufOrigem','regiaoOrigem']);
  const complementary=Object.fromEntries(Object.entries(d).map(([k,v])=>[k,value(v,120)]));
  complementary.instagram=complementary.instagram.replace(/^@+/,'').replace(/\s/g,'');N(!complementary.instagram||/^[A-Za-z0-9._]{1,30}$/.test(complementary.instagram)&&!complementary.instagram.includes('..'),'COMMAND_PAYLOAD_INVALID');
  if(complementary.dataNascimento){C.civilStart(complementary.dataNascimento);N(complementary.dataNascimento>='1900-01-01'&&complementary.dataNascimento<=today,'COMMAND_PAYLOAD_INVALID');}
  N(!complementary.paisOrigemCodigo||/^[A-Z]{2}$/.test(complementary.paisOrigemCodigo),'COMMAND_PAYLOAD_INVALID');N(!complementary.municipioOrigemCodigo||/^\d{7}$/.test(complementary.municipioOrigemCodigo),'COMMAND_PAYLOAD_INVALID');N(!complementary.ufOrigem||/^[A-Z]{2}$/.test(complementary.ufOrigem),'COMMAND_PAYLOAD_INVALID');
  if(p.modalidadeCadastro==='RAPIDO')N(Object.values(complementary).every(v=>v===''),'COMMAND_PAYLOAD_INVALID');
  return {...p,nomeCompleto:nome,nomeExibicao:value(p.nomeExibicao,100),emailPrincipal:email,rga,telefone:phone,cursoNomeOutro:value(p.cursoNomeOutro,180),semestreEntrada:value(p.semestreEntrada,100),documentoReferencia:reference,observacaoAdministrativa:note,dadosComplementares:complementary};
 }
 function validate(get,r,body,now){
  const v=body.facts;C.exact(v,['schemaVersion','personId','policyRevision','courseRevision','semesterId','semesterRevision','registration','approval']);
  N(v.schemaVersion==='member-administrative-ingress-v1'&&v.personId===r.target.personId&&r.payload.operation==='RECORD_INGRESS'&&r.payload.originalEventId===null,'COMMAND_PAYLOAD_INVALID');
  C.exact(v.approval,['reference','confirmedByUid','previouslyApproved']);N(v.approval.reference===r.decisionRef&&v.approval.confirmedByUid===r.ownerUid&&v.approval.previouslyApproved===true,'AUTH_BINDING_INVALID');
  const policy=get('memberOperationalConfig/registration');N(policy?.schemaVersion==='member-registration-policy-v1'&&policy.enabled===true&&policy.revision===v.policyRevision&&policy.allowedForms?.includes('PROCESSO_SELETIVO')&&policy.approvalRef,'ACCESS_REVOKED');
  const today=Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd'),data=normalize(v.registration,today);N(K.equal(data,v.registration)&&data.dataIngresso===r.payload.startDate&&r.payload.endDate===null,'COMMAND_PAYLOAD_INVALID');
  const course=get('memberCourseCatalog/'+data.cursoId);N(course?.schemaVersion==='member-course-v1'&&course.id===data.cursoId&&course.active===true&&course.revision===v.courseRevision,'REVISION_CONFLICT');N(!course.requiresOtherName||data.cursoNomeOutro.length>0,'COMMAND_PAYLOAD_INVALID');
  const semester=get('academicSemesters/'+v.semesterId);N(semester&&semester.revision===v.semesterRevision&&semester.label===data.semestreEntrada,'REVISION_CONFLICT');
  const identity=f4MemberRegistrationIdentity_().resolve(get,{email:data.emailPrincipal,rga:data.rga});N(identity.personId===null||identity.personId===r.target.personId,'AUTH_BINDING_INVALID');
  if(r.payload.identity)N(r.payload.identity.displayName===(data.nomeExibicao||data.nomeCompleto),'COMMAND_PAYLOAD_INVALID');
  for(const row of r.payload.identifiers)N(identity.identifiers.some(v=>v.kind===row.kind&&v.value===row.value),'COMMAND_PAYLOAD_INVALID');
  for(const row of identity.identifiers){const claim=get('memberIdentifierClaims/'+C.hash(row));N(claim?.active&&claim.idPessoa===r.target.personId||r.payload.identifiers.some(v=>v.kind===row.kind&&v.value===row.value),'COMMAND_PAYLOAD_INVALID');}
  return body;
 }
 return {normalize,validate};
}
