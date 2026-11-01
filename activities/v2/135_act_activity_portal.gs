// Fixed adapter for the existing "Criar atividade por modelo" form.
function f4ActivityPortalModel_(v){
 const C=f4Core_(),N=f4ActivitiesModel_().need;
 const strings=['type','subtype','group','description','principalRole','materialInstructions'];
 const booleans=['generatesCertificate','requiresTitle','requiresAxis','allowsSecondaryAxis','requiresPrincipal','allowsExternalPrincipal','allowsMemberPrincipal','allowsProfessorPrincipal','requiresPrincipalEmail','requiresPrincipalInstitution','requiresMaterial','showCalendar','showUpcoming','reviewBeforePublish','memberPresentation'];
 C.exact(v,[...strings,...booleans,'allowedPrincipalTypes','materialExtensions','maxMaterialMb','displayOrder']);
 strings.forEach(k=>N(typeof v[k]==='string'&&v[k].length<=800&&!/[<>\x00-\x08]/.test(v[k]),'VALIDATION_ERROR'));
 booleans.forEach(k=>N(typeof v[k]==='boolean','VALIDATION_ERROR'));
 N(v.type&&v.subtype&&Number.isInteger(v.displayOrder)&&v.displayOrder>=0&&Number.isInteger(v.maxMaterialMb)&&v.maxMaterialMb>=0&&v.maxMaterialMb<=30,'VALIDATION_ERROR');
 N(Array.isArray(v.allowedPrincipalTypes)&&v.allowedPrincipalTypes.length<=5&&v.allowedPrincipalTypes.every(t=>['MEMBRO','PROFESSOR','CONVIDADO','EXTERNO','EGRESSO'].includes(t)),'VALIDATION_ERROR');
 N(Array.isArray(v.materialExtensions)&&v.materialExtensions.length<=8&&v.materialExtensions.every(t=>['pdf','ppt','pptx','odp','docx','jpg','jpeg','png'].includes(t)),'VALIDATION_ERROR');
 return {...v};
}
function f4ActivityPortalDetails_(v){
 const N=f4ActivitiesModel_().need,allowed=['descricaoPublica','descricaoInterna','formato','responsavelInterno','responsavelEmail','publicoAlvo','observacoes'];
 N(v&&Object.getPrototypeOf(v)===Object.prototype&&Object.keys(v).every(k=>allowed.includes(k)),'VALIDATION_ERROR','PROTECTED_MODEL_FIELDS');
 for(const [key,value] of Object.entries(v))N(typeof value==='string'&&value.length<=(key.startsWith('descricao')||key==='observacoes'?2000:240)&&!/[<>\x00-\x08]/.test(value),'VALIDATION_ERROR');
 if(v.responsavelEmail)N(/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v.responsavelEmail),'VALIDATION_ERROR');
 if(v.formato)N(['PRESENCIAL','REMOTO','HIBRIDO'].includes(v.formato),'VALIDATION_ERROR');return {...v};
}
function f4ActivityPortalModelDto_(type){
 const N=f4ActivitiesModel_().need,m=f4ActivityPortalModel_(type.portalModel);N(typeof type.countsAbsence==='boolean','CONFIG_CAPABILITY_DISABLED','ATTENDANCE_POLICY_REQUIRED');
 return {idConfig:type.id,nomeModeloPortal:type.name,descricaoModeloPortal:m.description,grupoModelo:m.group,ordemExibicao:m.displayOrder,tipoAtividade:m.type,subtipoAtividade:m.subtype,classificacaoReuniao:'',tipoPublico:m.type,cargaHorariaPadrao:type.durationMinutes==null?null:type.durationMinutes/60,exigeTituloPublico:m.requiresTitle,exigeEixoTematico:m.requiresAxis,permiteEixoSecundario:m.allowsSecondaryAxis,exigePessoaPrincipal:m.requiresPrincipal,papelPadraoPessoaPrincipal:m.principalRole,tipoPessoaPrincipalPadrao:m.allowedPrincipalTypes,permitePessoaExternaPrincipal:m.allowsExternalPrincipal,permiteMembroComoPrincipal:m.allowsMemberPrincipal,permiteProfessorComoPrincipal:m.allowsProfessorPrincipal,exigeEmailPessoaPrincipal:m.requiresPrincipalEmail,exigeInstituicaoPessoaPrincipal:m.requiresPrincipalInstitution,contaPresencaPadrao:type.countsAttendance,contaFaltaPadrao:type.countsAbsence,geraCertificadoPadrao:m.generatesCertificate,permiteJustificativa:type.allowsJustification,exigeMaterialPadrao:m.requiresMaterial,tiposArquivoMaterialPermitidos:m.materialExtensions,tamanhoMaxMaterialMb:m.maxMaterialMb,instrucaoUploadMaterial:m.materialInstructions,visibilidadePortalPadrao:'DIRETORIA',statusPublicacaoPortalPadrao:'RASCUNHO',exibeNoCalendario:m.showCalendar,exibeEmProximasAtividades:m.showUpcoming,exigeRevisaoAntesPublicar:m.reviewBeforePublish,tituloEixoPosterior:m.memberPresentation,materialPosterior:m.memberPresentation,responsavelInternoAutomatico:m.memberPresentation?'Secretaria GEAPA':'',versaoConfigModelo:String(type.revision)};
}
