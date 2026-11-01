// Occurrence details for the existing activity model form. A named external
// speaker is display/contact context; it never creates a Membros identity or an
// attendance fact. Known people are stored only by their canonical reference.
function f4ActivityContent_(model,value,get,at,memberCreation=false){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),v=value||{};N(v&&Object.getPrototypeOf(v)===Object.prototype&&Object.keys(v).every(k=>['principal','principalAxisRef','secondaryAxisRef'].includes(k)),'VALIDATION_ERROR');
 N(!model.memberPresentation||memberCreation,'UNRESOLVED_DEPENDENCY','MEMBER_ACTIVITY_CALENDAR_REQUIRED');if(model.memberPresentation)model={...model,requiresAxis:false};
 const text=(s,max=180)=>typeof s==='string'&&s.length<=max&&!/[<>\x00-\x08]/.test(s);
 for(const [key,ref]of [['principalAxisRef',v.principalAxisRef],['secondaryAxisRef',v.secondaryAxisRef]])if(ref){N(ref.kind==='policy','VALIDATION_ERROR');const axis=get(M.path(ref)),cycle=axis?.cycleRef&&get(M.path(axis.cycleRef)),config=cycle?.currentConfigRef&&get(M.path(cycle.currentConfigRef));N(axis?.approved&&axis.active&&axis.usage==='THEMATIC_AXIS'&&config?.approved&&config.parameters.axisRefs?.some(r=>M.path(r)===M.path(ref)),'CONFIG_CAPABILITY_DISABLED','ACTIVITY_AXIS_NOT_CONFIGURED');N(Number.isSafeInteger(at)&&M.gap(cycle.civilStart,M.day(at))>=0&&M.gap(M.day(at),cycle.civilEnd)>=0,'CONFIG_CAPABILITY_DISABLED','ACTIVITY_AXIS_DATE');if(key==='secondaryAxisRef')N(model.allowsSecondaryAxis&&v.principalAxisRef&&v.principalAxisRef.id!==ref.id,'VALIDATION_ERROR','SECONDARY_AXIS');}
 N(!model.requiresAxis||v.principalAxisRef,'INCOMPLETE_REQUEST','PRINCIPAL_AXIS_REQUIRED');const principal=v.principal;N(!model.requiresPrincipal||principal,'INCOMPLETE_REQUEST','PRINCIPAL_REQUIRED');
 if(principal){
  N(['CORE_PERSON','EXTERNAL_DISPLAY'].includes(principal.kind),'VALIDATION_ERROR');C.exact(principal,['kind','participantType','participationRole','institutionLabel',...(principal.kind==='CORE_PERSON'?['personRef']:['displayName','contactEmail'])]);
  N(model.allowedPrincipalTypes.includes(principal.participantType)&&text(principal.participationRole,80)&&text(principal.institutionLabel),'VALIDATION_ERROR','PRINCIPAL_TYPE');N(!model.principalRole||principal.participationRole===model.principalRole,'VALIDATION_ERROR','PRINCIPAL_ROLE');N(!model.requiresPrincipalInstitution||principal.institutionLabel.trim(),'INCOMPLETE_REQUEST','PRINCIPAL_INSTITUTION');
  if(principal.kind==='CORE_PERSON'){
   N(principal.personRef?.kind==='person','VALIDATION_ERROR');const identity=get(M.path(principal.personRef)),head=get('memberAggregateHeads/'+principal.personRef.id);N(identity?.idPessoa===principal.personRef.id&&!identity.technicalAccount&&head?.paths?.includes(M.path(principal.personRef)),'UNRESOLVED_DEPENDENCY','PRINCIPAL_CORE_IDENTITY');
   N(principal.participantType!=='MEMBRO'||model.allowsMemberPrincipal,'CONFIG_CAPABILITY_DISABLED');N(principal.participantType!=='PROFESSOR'||model.allowsProfessorPrincipal,'CONFIG_CAPABILITY_DISABLED');
   if(model.requiresPrincipalEmail){const profile=get('memberPrivateProfiles/'+principal.personRef.id);N(head.paths.includes('memberPrivateProfiles/'+principal.personRef.id)&&profile?.idPessoa===principal.personRef.id&&profile.contactEmail,'INCOMPLETE_REQUEST','PRINCIPAL_EMAIL_REQUIRED');}
  }else{N(model.allowsExternalPrincipal&&principal.participantType!=='MEMBRO','CONFIG_CAPABILITY_DISABLED');N(text(principal.displayName)&&principal.displayName.trim()&&text(principal.contactEmail,240),'INCOMPLETE_REQUEST');N(!model.requiresPrincipalEmail||principal.contactEmail,'INCOMPLETE_REQUEST','PRINCIPAL_EMAIL_REQUIRED');N(!principal.contactEmail||/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(principal.contactEmail),'VALIDATION_ERROR');}
 }
 return v;
}
function f4ActivityContentFromPortal_(model,v,get,at){
 const M=f4ActivitiesModel_(),personId=v.idPessoaPrincipal||'',name=v.nomePessoaPrincipalPublico||'',type=v.tipoPessoaPrincipal||(model.allowedPrincipalTypes.length===1?model.allowedPrincipalTypes[0]:''),principal=personId||name?{kind:personId?'CORE_PERSON':'EXTERNAL_DISPLAY',participantType:type,participationRole:model.principalRole||v.papelPessoaPrincipal||'',institutionLabel:v.instituicaoPessoaPrincipal||'',...(personId?{personRef:M.ref('person',personId)}:{displayName:name,contactEmail:v.emailPessoaPrincipal||''})}:null;
 return f4ActivityContent_(model,{...(principal?{principal}:{}),...(v.eixoTematicoPrincipal?{principalAxisRef:M.ref('policy',v.eixoTematicoPrincipal)}:{}),...(v.eixoTematicoSecundario?{secondaryAxisRef:M.ref('policy',v.eixoTematicoSecundario)}:{})},get,at);
}
function f4ActivityContentDto_(activity,admin,get){
 const v=activity.modelContent;if(!v||!get)return {};const M=f4ActivitiesModel_(),p=v.principal,identity=p?.kind==='CORE_PERSON'?get(M.path(p.personRef)):null,name=p?.kind==='EXTERNAL_DISPLAY'?p.displayName:identity?.displayName||'',axis=r=>r?get(M.path(r))?.label||'':'';
 return {nomePessoaPrincipalPublico:name,tipoPessoaPrincipal:p?.participantType||'',papelPessoaPrincipal:p?.participationRole||'',instituicaoPessoaPrincipal:p?.institutionLabel||'',eixoTematicoPrincipal:axis(v.principalAxisRef),eixoTematicoSecundario:axis(v.secondaryAxisRef),...(admin?{idPessoaPrincipal:p?.personRef?.id||'',emailPessoaPrincipal:p?.kind==='EXTERNAL_DISPLAY'?p.contactEmail:p?.personRef?get('memberPrivateProfiles/'+p.personRef.id)?.contactEmail||'':''}:{})};
}
