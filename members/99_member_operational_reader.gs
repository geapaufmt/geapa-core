// Small adapter for the existing Members screens. SELF/catalog/context become
// fenced direct projections. Private decision dossiers stay in the Core.
function f4MemberOperationalReader_({config,identity,clock=Date.now}){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),A=f4MemberAuthz_(),K=f4MemberAuthzContracts_(),W=f4MemberVoluntaryWorkflow_(),db=F.connect(config);
 const routes=['/admin/frequencia-rh/providencias','/members/operational-context','/admin/membros','/admin/ingressos-membros','/admin/ingressos-membros/catalogos','/meu-vinculo/acompanhamento','/meu-vinculo/opcoes','/meu-vinculo/solicitacoes','/admin/solicitacoes-vinculo','/admin/solicitacoes-vinculo/detalhe','/meu-vinculo/notificacoes'];
 const iso=v=>new Date(C.ms(v)).toISOString(),today=n=>Utilities.formatDate(new Date(n),'America/Cuiaba','yyyy-MM-dd');
 function access(get,now){
  C.exact(identity,['uid','expiresAt']);
  N(C.id(identity.uid)&&Number.isSafeInteger(identity.expiresAt),'AUTH_BINDING_INVALID');
  N(Number.isSafeInteger(now)&&now<identity.expiresAt,'ACCESS_REVOKED');
  const technical=f4DevAdmin_().profile(get,identity.uid,now),actor=technical?f4DevAdmin_().identify(get,identity.uid,now):A.identify(get,identity.uid,now),settings=f4MemberPortalPrepare_({config,identity,clock}).configuration(get);
  const policies=settings.policyIds.map(id=>get('memberAuthzPolicies/'+id)).filter(Boolean);
  function scope(capability,commandScope){
   if(capability==='MEMBERS_INGRESS_EXECUTE'&&commandScope==='ADMINISTRATIVE_INGRESS'){const dev=f4DevAdministrativeAccess_().source(get,identity,now,capability);if(dev&&dev.source.admissionApprovalIds?.length)return {...dev,kind:'DEV_TECHNICAL',expiresAt:dev.validUntil};}
   if(technical)return {source:technical,kind:'DEV_ADMIN',expiresAt:Math.min(identity.expiresAt,C.ms(technical.expiresAt))};
   for(const p of policies){K.policy(p);if(!p.active||!p.capabilities.includes(capability)||(commandScope&&!p.targets.commandScopes?.includes(commandScope))||p.proofKind!=='ROLE')continue;
    if(commandScope&&(p.targets.ownPerson||!p.targets.allowedTargetScopes.includes(['ADMINISTRATIVE_INGRESS','ENROLLMENT_SUPERVISION'].includes(commandScope)?'ADMINISTRACAO':'RECURSOS_HUMANOS')))continue;
    if(!(C.ms(p.issuedAt)<=now&&now<C.ms(p.expiresAt)))continue;
    const fn=actor.derived.currentFunctions.find(v=>v.roleKey===p.actorSelector.roleKey&&v.group===p.actorSelector.group&&v.scopeKey===p.actorSelector.actorScope);
    if(fn)return {source:p,kind:'ROLE',functionId:fn.idVigencia,actorScope:fn.scopeKey,expiresAt:Math.min(identity.expiresAt,C.ms(p.expiresAt),C.ms(actor.state.validUntil))};
   }return null;
  }
  const reviewExit=settings.commands.includes('members.voluntary.review')&&scope('MEMBERS_VOLUNTARY_EXIT_EXECUTE','VOLUNTARY_REVIEW'),reviewSuspension=settings.commands.includes('members.voluntary.review')&&scope('MEMBERS_VOLUNTARY_SUSPENSION_EXECUTE','VOLUNTARY_REVIEW');
  const own=['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE'].includes(actor.state.state)&&!actor.state.blocked;
  const supervision=settings.commands.includes('members.ingress.execute')&&f4MemberInvitation_().reviewer(actor,get('memberOperationalConfig/registration'))&&scope('MEMBERS_INGRESS_EXECUTE','ENROLLMENT_SUPERVISION');
  return {actor,technical,settings,own,supervision,directory:scope('MEMBERS_CASE_READ',null),ingress:settings.commands.includes('members.ingress.execute')&&scope('MEMBERS_INGRESS_EXECUTE','ADMINISTRATIVE_INGRESS'),review:reviewExit||reviewSuspension,reviewExit,reviewSuspension};
 }
 function item(v){W.validate(v);return {idSolicitacao:v.id,revision:v.revision,tipo:v.modality==='SUSPENSAO_TEMPORARIA'?'SUSPENSAO_VOLUNTARIA':'DESLIGAMENTO_VOLUNTARIO',modalidade:v.modality,status:v.status,idSemestre:v.semesterId,dataSolicitacao:iso(v.createdAt),mensagemDiretoria:v.history.filter(h=>!['SUBMIT','CANCEL_SELF','COMPLEMENT'].includes(h.action)).at(-1)?.note||'',podeCancelar:!W.terminal.includes(v.status)&&!['HOMOLOGADO_AGUARDANDO_INICIO','SUSPENSAO_ATIVA'].includes(v.status),podeComplementar:v.status==='AGUARDANDO_COMPLEMENTO'};}
 function tracking(params){
  // Restricted SELF tracking is a read of a person's existing voluntary cases.
  // It cannot restore ordinary membership, roles or command authorization.
  C.exact(params,[]);C.exact(identity,['uid','expiresAt']);
  N(C.id(identity.uid)&&Number.isSafeInteger(identity.expiresAt),'AUTH_BINDING_INVALID');
  return db.transaction(tx=>{
   const get=p=>tx.get(p).data(),now=clock();N(now<identity.expiresAt,'ACCESS_REVOKED');
   const actor=A.identify(get,identity.uid,now,s=>s.blocked&&['SUSPENSO','DESLIGADO'].includes(s.state)),id=actor.binding.idPessoa;
   N(actor.state.blocked&&['SUSPENSO','DESLIGADO'].includes(actor.state.state),'ACCESS_REVOKED');
   const settings=f4MemberPortalPrepare_({config,identity,clock}).configuration(get),index=W.index(get('memberVoluntaryIndexes/'+id),id);
   N(index.requestIds.length>0,'ACCESS_REVOKED');
   const source=settings.policyIds.map(pid=>get('memberAuthzPolicies/'+pid)).find(p=>{
    if(!p?.active||p.proofKind!=='SELF'||!p.capabilities?.includes('MEMBERS_VOLUNTARY_REQUEST'))return false;
    K.policy(p);return C.ms(p.issuedAt)<=now&&now<C.ms(p.expiresAt)&&p.targets.ownPerson&&(p.targets.commandScopes?.includes('OWN_MEMBER_RESOURCES')||p.targets.allowedResources.includes('memberVoluntaryIndexes/'+id));
   });N(source,'ACCESS_REVOKED');
   const cases=index.requestIds.map(key=>W.validate(get('memberVoluntaryRequests/'+key)));
   N(cases.every(v=>v.personId===id&&actor.actor.d['memberLinks/'+v.linkId]?.idPessoa===id),'RECONCILIATION_REQUIRED');
   const validUntil=Math.min(identity.expiresAt,C.ms(source.expiresAt),C.ms(actor.state.validUntil));
   N(now<validUntil,'ACCESS_REVOKED');
   return {ok:true,data:{restrictedTracking:true,personId:id,validUntil,state:actor.state.state,sourceRevision:actor.actor.h.revision,
    options:{parametroSuspensao:null,vinculo:{tipo:actor.state.state,status:'INATIVO',dataInicio:null},semestre:null,solicitacaoAberta:index.activeRequestId?{idSolicitacao:index.activeRequestId}:null,opcoes:{suspensaoDisponivel:false,desligamentoDisponivel:false}},
    requests:{items:cases.map(v=>({...item(v),podeCancelar:false,podeComplementar:false}))}}};
  });
 }
 function selfData(get,a,now){
  const id=a.actor.binding.idPessoa,i=W.index(get('memberVoluntaryIndexes/'+id),id),requests=i.requestIds.map(id=>W.validate(get('memberVoluntaryRequests/'+id)));
  N(requests.every(v=>v.personId===id),'RECONCILIATION_REQUIRED');
  const link=a.actor.actor.M.rows(a.actor.actor.d,'memberLinks').find(v=>v.type===a.actor.state.state&&a.actor.actor.M.interval(v.startDate,v.endDate).a<=now&&now<a.actor.actor.M.interval(v.startDate,v.endDate).z);
  const registration=get('memberOperationalConfig/registration'),semesterIds=registration?.semesterIds||[];N(Array.isArray(semesterIds)&&semesterIds.length<=32&&semesterIds.every(C.id),'SCHEMA_INVALID');
  const current=semesterIds.map(id=>get('academicSemesters/'+id)).filter(s=>s&&s.civilStart<=today(now)&&today(now)<=s.civilEnd),semester=current.length===1?current[0]:null;
  const policy=get('memberOperationalConfig/voluntaryExit'),enabled=policy?.enabled===true&&a.settings.commands.includes('members.voluntary.submit');
  if(enabled)W.parameter(policy);
  // The command must still prove its separate SELF request capability.
  const canSubmit=a.own&&a.settings.commands.includes('members.voluntary.submit')&&(a.technical||a.settings.policyIds.some(pid=>{const p=get('memberAuthzPolicies/'+pid);if(!p?.active||p.proofKind!=='SELF'||!p.capabilities.includes('MEMBERS_VOLUNTARY_REQUEST')||!(C.ms(p.issuedAt)<=now&&now<C.ms(p.expiresAt)))return false;K.policy(p);return p.targets.ownPerson&&(p.targets.commandScopes?.includes('OWN_MEMBER_RESOURCES')||p.targets.allowedResources.includes('memberVoluntaryIndexes/'+id));}));
  const suspension=get('memberOperationalConfig/voluntarySuspension'),canSuspend=canSubmit&&a.actor.state.state==='MEMBRO_EFETIVO'&&!!semester&&suspension?.enabled===true;if(canSuspend)f4MemberVoluntarySuspension_().policy(suspension);
  return {options:{ok:true,data:{parametroSuspensao:canSuspend?{valor:suspension.minimumDays,unidade:'dias',baseLegal:'NC01/2025, art. 15'}:null,vinculo:{tipo:a.actor.state.state,status:a.actor.state.blocked?'INATIVO':'ATIVO',dataInicio:link?.startDate||null},semestre:semester?{canonicalId:semester.id,idSemestre:semester.label,dataFim:semester.civilEnd}:null,solicitacaoAberta:i.activeRequestId?{idSolicitacao:i.activeRequestId}:null,opcoes:{suspensaoDisponivel:!!canSuspend,desligamentoDisponivel:!!(canSubmit&&enabled)}}},requests:{ok:true,data:{items:requests.map(item)}}};
 }
 function catalog(get){
  const p=get('memberOperationalConfig/registration');N(p?.schemaVersion==='member-registration-policy-v1'&&p.enabled&&Array.isArray(p.courseIds)&&p.courseIds.length<=100&&p.courseIds.every(C.id),'ACCESS_REVOKED');
  return {ok:true,data:{cursos:p.courseIds.map(id=>get('memberCourseCatalog/'+id)).filter(v=>v?.active).map(v=>({cursoId:v.id,nomeCurso:v.name,instituicao:v.institution||'',campus:v.campus||''})),semestres:(p.semesterIds||[]).map(id=>get('academicSemesters/'+id)).filter(Boolean).map(s=>({id:s.id,label:s.label})),emailEnabled:false}};
 }
 function project(tx,view,now){
  // Reuse the canonical identity prefetch in this same transaction, before access().
  const get=p=>tx.get(p).data(),status=f4MemberOwnReader_().readTransaction(tx,identity,{view:'status'},now),a=access(get,now),actor=a.actor,gate=get('portalReaderControl/current');
  N(gate?.enabled&&gate.operationalWrites&&f4ServiceAvailability_().live(gate,now),'ACCESS_REVOKED');
  let data,expires=f4ServiceAvailability_().bound(gate,Math.min(status.validUntil,identity.expiresAt));
  if(view==='members_self')data=selfData(get,a,now);
  else if(view==='members_catalog')data=catalog(get);
  else {N(view==='members_context','COMMAND_PAYLOAD_INVALID');data={canRegister:!!a.ingress,canSuperviseEnrollment:!!a.supervision,canReview:!!a.review,canReadDirectory:!!a.directory,canReadPeople:false,canReviewCorrections:false};try{const people=f4PeopleDirectory_({config,identity,clock}).access(get,now);data.canReadPeople=true;data.canReadDirectory=true;expires=Math.min(expires,people.expiresAt);}catch(e){if(!['M13_AUTHZ_DENIED','ACCESS_REVOKED','SCOPE_TOO_LARGE'].includes(e.code||e.message))throw e;}try{const correction=f4MemberProfileCorrectionReader_({config,identity,clock}).access(get,now,false);data.canReviewCorrections=correction.mask.length>0;expires=Math.min(expires,correction.expires);}catch{};for(const v of [a.ingress,a.supervision,a.review,a.directory])if(v)expires=Math.min(expires,v.expiresAt);}
  const sp=actor.control.selfReadPolicy,p=a.technical||get('memberAuthzPolicies/'+sp.policyId),fence=get('portalReaderResourceFences/members_operations');
  const basis={generation:gate.generation,ownerUid:identity.uid,personId:actor.binding.idPessoa,bindingVersion:actor.binding.bindingVersion,personRevision:actor.actor.h.revision,authorityEpoch:actor.control.epoch,kind:a.technical?'DEV_ADMIN':'SELF',capability:'MEMBERS_SELF_READ',issuedAt:F.Timestamp.fromMillis(now),expiresAt:F.Timestamp.fromMillis(expires),dataFence:{id:'members_operations',revision:fence?.revision||0}};
  if(a.technical)basis.profileRevision=p.revision;else Object.assign(basis,{subjectEpoch:actor.subject.authzEpoch,policyEpoch:actor.control.policyEpoch,readAuthzEpoch:actor.control.readAuthzEpoch,policyId:p.id,policyRevision:p.revision,state:actor.state.state});
  const path='portalReaderViews/'+identity.uid+'/items/'+view,v={schemaVersion:'portal-reader-view-v1',viewId:view,basis,data};N(JSON.stringify(F.encode(v)).length<700000,'SCOPE_TOO_LARGE');get(path);tx.set(path,v);return data;
 }
 function dossier(get,id,a,now){
  N(a.review&&C.op(id),'ACCESS_REVOKED');const v=W.validate(get('memberVoluntaryRequests/'+id));if(v.modality==='SUSPENSAO_TEMPORARIA')return suspensionDossier(get,v,a,now);N(a.reviewExit,'ACCESS_REVOKED');const policy=get('memberOperationalConfig/voluntaryExit');W.parameter(policy);
  const target=A.load(get,v.personId,a.actor.control),state=target.M.derive(target.d,v.personId,target.h.revision,now),fn=state.currentFunctions.length>0;
  const changed=C.hash(v.minutesSnapshot)!==C.hash(f4MemberVoluntaryRules_().minutes(policy.minutes));
  return {ok:true,data:{revision:v.revision,solicitacao:{...W.legacy(v),ID_SOLICITACAO:v.id,MOTIVO_CATEGORIA:v.submission?.reasonCategory||'',OBSERVACOES_MEMBRO:v.reason,DOCUMENTO_REFERENCIA:v.submission?.documentReference||'',VALIDACAO_VINCULO_ATIVO:state.state.blocked?'INATIVO':'ATIVO',VALIDACAO_SEMESTRE:v.semesterId?'CONFERIR':'NAO_APLICAVEL',VALIDACAO_APRESENTACAO:'NAO_VERIFICADO',VALIDACAO_ARQUIVOS_PENDENTES:'NAO_VERIFICADO',VALIDACAO_OBRIGACOES:'NAO_VERIFICADO',VALIDACAO_FUNCAO_ATIVA:fn?'ATIVA':'SEM_FUNCAO_ATIVA',RESULTADO_VALIDACAO:'PENDENTE_ANALISE_MANUAL',MENSAGEM_VALIDACAO:'Conferir apresentações, materiais e obrigações antes da decisão. Uma função ativa deve ser encerrada formalmente. Antecedência normativa: '+policy.presentationBlockDays+' dias.',ANALISE_FINAL_DISPONIVEL:v.modality==='DESLIGAMENTO_FIM_SEMESTRE'&&v.semesterId&&today(now)>=(get('academicSemesters/'+v.semesterId)?.civilEnd||'9999-12-31')},pessoa:{nome:target.d['memberIdentities/'+v.personId].displayName,rgaMascarado:'',emailMascarado:''},parametrosNormativos:{divergencias:changed?['ATA_RULE_CHANGED']:[]},requisitosDecisaoFinal:{exigeAta:policy.minutes.valor,baseLegal:policy.minutes.baseLegal,possuiDivergencia:changed,exigeTratamentoTransicao:changed},historico:f4MemberOwnReader_().wire(v.history)}};
 }
 function suspensionDossier(get,v,a,now){
  N(a.reviewSuspension,'ACCESS_REVOKED');const policy=f4MemberVoluntarySuspension_().policy(get('memberOperationalConfig/voluntarySuspension')),target=A.load(get,v.personId,a.actor.control),todayDate=today(now);
  return {ok:true,data:{revision:v.revision,solicitacao:{...W.legacy(v),ID_SOLICITACAO:v.id,MOTIVO_CATEGORIA:v.submission.reasonCategory,OBSERVACOES_MEMBRO:v.reason,DOCUMENTO_REFERENCIA:v.submission.documentReference,VALIDACAO_VINCULO_ATIVO:'CONFERIDO_NO_COMANDO',VALIDACAO_SEMESTRE:'CONFERIDO_NO_COMANDO',VALIDACAO_APRESENTACAO:'CONFERIDO_NO_COMANDO',VALIDACAO_ARQUIVOS_PENDENTES:'NAO_APLICAVEL',VALIDACAO_OBRIGACOES:'NAO_APLICAVEL',VALIDACAO_FUNCAO_ATIVA:'CONFERIR',RESULTADO_VALIDACAO:'DECISAO_HUMANA',MENSAGEM_VALIDACAO:'A agenda e o intervalo são revalidados no envio. Suspensão disponível somente para membros efetivos.',PODE_INICIAR_SUSPENSAO:v.status==='HOMOLOGADO_AGUARDANDO_INICIO'&&todayDate>=v.startDate&&todayDate<=v.endDate,PODE_REGISTRAR_RETORNO:['SUSPENSAO_ATIVA','HOMOLOGADO_AGUARDANDO_INICIO'].includes(v.status)&&todayDate>v.endDate},pessoa:{nome:target.d['memberIdentities/'+v.personId].displayName,rgaMascarado:'',emailMascarado:''},parametrosNormativos:{divergencias:policy.revision===v.policyRevision?[]:['SUSPENSION_POLICY_CHANGED']},requisitosDecisaoFinal:{exigeAta:false,baseLegal:'NC01/2025, art. 15',possuiDivergencia:false,exigeTratamentoTransicao:false},historico:f4MemberOwnReader_().wire(v.history)}};
 }
 function read(input){
  if(input?.route==='/admin/frequencia-rh/providencias'){C.exact(input,['route','params']);return attendanceCases(input.params);}
  if(f4MemberProfileCorrectionReader_({config,identity,clock}).routes.includes(input?.route))return f4MemberProfileCorrectionReader_({config,identity,clock}).read(input);
  C.exact(input,['route','params']);N(routes.includes(input.route),'COMMAND_TYPE_INVALID');const {route,params}=input;
  if(route==='/meu-vinculo/acompanhamento')return tracking(params);
  if(route==='/meu-vinculo/notificacoes')return notifications(params);
  if(route==='/admin/membros')return directory(params);
  if(route==='/admin/ingressos-membros')return enrollments(params);
  if(!route.startsWith('/admin/solicitacoes-vinculo')){C.exact(params,[]);const view=route==='/members/operational-context'?'members_context':route==='/admin/ingressos-membros/catalogos'?'members_catalog':'members_self';const v=db.transaction(tx=>project(tx,view,clock()));return view==='members_self'?route.endsWith('/opcoes')?v.options:v.requests:view==='members_context'?{ok:true,data:v}:v;}
  if(route.endsWith('/detalhe')){C.exact(params,['payload']);let p;try{p=JSON.parse(params.payload);}catch{N(false,'COMMAND_PAYLOAD_INVALID');}C.exact(p,['idSolicitacao']);return db.transaction(tx=>{const get=p=>tx.get(p).data();return dossier(get,p.idSolicitacao,access(get,clock()),clock());});}
  return requestPage(params);
 }
 function notifications(params){
  C.exact(params,[...('notificationId'in params?['notificationId']:[]),...('action'in params?['action']:[]),...('personId'in params?['personId']:[])]);
  return db.transaction(tx=>{
   const get=p=>tx.get(p).data(),now=clock(),M=f4MemberNotifications_();
   if(params.action==='MARK_READ'){
    N(params.notificationId,'COMMAND_PAYLOAD_INVALID');
    return M.markRead(tx,identity,params.notificationId,now);
   }
   if(params.action==='MARK_ALL_READ'){
    return M.markAllRead(tx,identity,now);
   }
   const view=M.readAuthorized(get,identity,now,params);
   return {ok:true,data:view};
  });
 }
 function attendanceCases(params){
  C.exact(params,['cycleId',...('personId'in params?['personId']:[]),...('cursor'in params?['cursor']:[])]);N(C.id(params.cycleId)&&(!params.personId||C.id(params.personId)),'COMMAND_PAYLOAD_INVALID');
  const get=p=>db.get(p),before=access(get,clock()),W=f4MemberAttendanceCases_();
  if(!before.directory)return syntheticAttendanceCase(params);
  if(!before.technical)K.subset(['attendanceCases'],before.directory.source.fieldMasks.MEMBERS_CASE_READ);
  const policy=W.rules(get,params.cycleId),fence=get('portalReaderResourceFences/members_operations')?.revision||0;
  const allowed=id=>before.technical||before.directory.source.targets.allowedPersons.includes(id);
  if(params.cursor){C.exact(params.cursor,['cycleId','afterId']);N(!params.personId&&params.cursor.cycleId===params.cycleId&&C.id(params.cursor.afterId),'COMMAND_PAYLOAD_INVALID');}
  if(params.personId)N(allowed(params.personId),'ACCESS_REVOKED');
  const rows=params.personId?[get('memberAttendanceCases/'+W.caseId(params.personId,params.cycleId))].filter(Boolean):db.query({collection:'memberAttendanceCases',filters:[{field:'cycleId',op:'EQUAL',value:params.cycleId}],order:[{field:'__name__',direction:'ASCENDING'}],limit:20,...(params.cursor?{cursor:[new F.DocumentReference('projects/'+config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+'memberAttendanceCases/'+params.cursor.afterId)]}:{})}).map(v=>v.data);
  const items=rows.filter(d=>allowed(d.personId)).map(W.record),pagination={hasMore:!params.personId&&rows.length===20,nextCursor:!params.personId&&rows.length===20?{cycleId:params.cycleId,afterId:rows.at(-1).id}:null};
  const person=params.personId&&get('memberAggregateHeads/'+params.personId),links=person?.paths?.filter(p=>p.startsWith('memberLinks/')).map(get).filter(l=>l&&['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE'].includes(l.type)&&l.startDate<=policy.cycle.civilEnd&&(!l.endDate||policy.cycle.civilStart<=l.endDate))||[];
  const grants=(before.settings.disciplineGrantIds||[]).map(id=>get('memberExplicitGrants/'+id)).filter(g=>g?.active&&g.uid===identity.uid&&g.capabilities?.includes('MEMBERS_DISCIPLINE_EXECUTE')&&C.ms(g.issuedAt)<=clock()&&clock()<C.ms(g.expiresAt));
  const canRecord=!!person&&before.settings.commands.includes('members.attendance.record')&&(policy.established||items[0]?.disciplineSnapshot)&&policy.cycle.status==='OPEN'&&grants.some(g=>g.targets.allowedPersons.includes(params.personId)&&g.targets.allowedResources.includes('memberAttendanceCases/'+W.caseId(params.personId,params.cycleId)));
  const institutionHistory=person?.paths?.filter(p=>p.startsWith('memberEvents/')).map(get).filter(e=>['SUSPENSAO','DESLIGAMENTO','RETORNO'].includes(e?.effect)).map(e=>({id:e.id,effect:e.effect,effectiveDate:e.effectiveDate,endDate:e.endDate,supersedes:e.supersedes}))||[];
  return db.transaction(tx=>{const fresh=p=>tx.get(p).data(),current=access(fresh,clock());N(current.directory&&K.digest(current.directory.source)===K.digest(before.directory.source)&&(fresh('portalReaderResourceFences/members_operations')?.revision||0)===fence,'ACCESS_REVOKED');for(const d of items)N(fresh('memberAttendanceCases/'+d.id)?.revision===d.revision,'REVISION_CONFLICT');
   const allowedActions=canRecord?W.allowedActions(items[0],before.settings.commands.includes('members.discipline.dismiss'),clock()):[];
   const documentRequiredByAction=Object.fromEntries(allowedActions.map(action=>[action,W.requiresDocument(action)]));
   return {ok:true,data:{cycleId:params.cycleId,pagination,items:f4MemberOwnReader_().wire(items),policy:policy.established?{established:true,policyId:policy.policyId,policyRevision:policy.policyRevision,rules:policy.rules}:{established:false},allowedActions,documentRequiredByAction,person:person?{id:person.id,revision:person.revision,links:links.map(l=>({id:l.id,type:l.type,startDate:l.startDate,endDate:l.endDate}))}:null,canRecord,canDismiss:canRecord&&before.settings.commands.includes('members.discipline.dismiss'),institutionalHistory:institutionHistory,historyScope:'PERMANENT_INSTITUTIONAL',notice:!policy.established?'O ciclo ainda não possui política de providências reconhecida na configuração aprovada.':!canRecord?'O registro exige autorização institucional específica; consultar frequência não concede essa competência.':''}};});
 }
 // A DEV fixture may inspect one specifically granted case. Directory and
 // discipline authority continue to be decided by their existing policies.
 function syntheticAttendanceCase(params){
  N(config.projectId==='geapa-convergencia-dev'&&config.environment==='DEV'&&params.personId&&!params.cursor,'ACCESS_REVOKED');
  return db.transaction(tx=>{
   const get=path=>tx.get(path).data(),now=clock(),current=access(get,now);
   const W=f4MemberAttendanceCases_(),caseId=W.caseId(params.personId,params.cycleId),path='memberAttendanceCases/'+caseId;
   requireSyntheticAttendanceGrant(get,current,params.personId,path,now);
   const row=get(path);
   N(row&&row.id===caseId&&row.personId===params.personId&&row.cycleId===params.cycleId,'ACCESS_REVOKED');
   const policy=W.rules(get,params.cycleId);
   return {ok:true,data:{cycleId:params.cycleId,pagination:{hasMore:false,nextCursor:null},items:f4MemberOwnReader_().wire([W.record(row)]),
    policy:policy.established?{established:true,policyId:policy.policyId,policyRevision:policy.policyRevision,rules:policy.rules}:{established:false},
    allowedActions:[],documentRequiredByAction:{},person:{id:params.personId,links:[]},canRecord:false,canDismiss:false,institutionalHistory:[],
    historyScope:'AUTHORIZED_CASE',notice:'Consulta autorizada para este caso de teste.'}};
  });
 }
 function requireSyntheticAttendanceGrant(get,current,target,path,now){
  const ids=current.settings.attendanceReadGrantIds,actor=current.actor;
  N(Array.isArray(ids)&&ids.length>0&&ids.length<=16&&ids.every(C.id)&&actor.state.state==='MEMBRO_EFETIVO','ACCESS_REVOKED');
  const control=get('memberExecutorControl/current');
  N(control?.projectId===config.projectId&&control.enabled===true&&control.browserEnabled===true&&control.appCheck?.mode==='REQUIRED'&&
   f4ExecutorWindow_().live(control,now)&&control.actorUids?.includes(identity.uid)&&Array.isArray(control.scopeIds)&&control.scopeIds.length<=4,'ACCESS_REVOKED');
  const scopes=control.scopeIds.map(id=>get('memberExecutorScopes/'+id)).filter((s,i)=>s?.schemaVersion==='member-executor-scope-v1'&&
   s.id===control.scopeIds[i]&&s.projectId===config.projectId&&s.synthetic===true&&Array.isArray(s.personIds));
  N(scopes.some(s=>s.personIds.includes(actor.binding.idPessoa)),'ACCESS_REVOKED');
  const allowed=ids.some(id=>{
   const g=get('memberExplicitGrants/'+id);
   if(!g?.active||g.id!==id||g.uid!==identity.uid||g.origin!=='SYNTHETIC_TEST'||g.environment!=='DEV')return false;
   K.grant(g);
   return g.idPessoa===actor.binding.idPessoa&&g.bindingVersion===actor.binding.bindingVersion&&g.subjectEpoch===actor.subject.authzEpoch&&
    C.ms(g.issuedAt)<=now&&now<C.ms(g.expiresAt)&&g.capabilities.length===1&&g.capabilities[0]==='MEMBERS_CASE_READ'&&
    g.fieldMasks.MEMBERS_CASE_READ?.includes('attendanceCases')&&g.actorScopes.includes('RECURSOS_HUMANOS')&&
    g.targets.commandScopes?.includes('ATTENDANCE_RH')&&g.targets.allowedTargetScopes.includes('RECURSOS_HUMANOS')&&
    !g.targets.ownPerson&&g.targets.allowedPersons.includes(target)&&g.targets.allowedResources.includes(path)&&
    scopes.some(s=>s.id===g.runId&&s.personIds.includes(target));
  });
  N(allowed,'ACCESS_REVOKED');
 }
 function requestPage(params){
  C.exact(params,['filtros',...(Object.hasOwn(params,'cursor')?['cursor']:[])]);let filters;try{filters=JSON.parse(params.filtros);}catch{N(false,'COMMAND_PAYLOAD_INVALID');}
  N(filters&&Object.getPrototypeOf(filters)===Object.prototype&&Object.keys(filters).every(k=>['tipo','modalidade','status','idSemestre','dataInicio','dataFim','membro','responsavel'].includes(k)),'COMMAND_PAYLOAD_INVALID');
  for(const v of Object.values(filters))N(typeof v==='string'&&v.length<=160,'COMMAND_PAYLOAD_INVALID');
  filters=Object.fromEntries(Object.entries(filters).filter(([,v])=>v));
  N(!filters.responsavel&&(!filters.membro||C.id(filters.membro))&&(!filters.idSemestre||C.id(filters.idSemestre)),'COMMAND_PAYLOAD_INVALID');
  const allModes=['SUSPENSAO_TEMPORARIA','DESLIGAMENTO_APOS_HOMOLOGACAO','DESLIGAMENTO_FIM_SEMESTRE'];
  N(!filters.tipo||['SUSPENSAO_VOLUNTARIA','DESLIGAMENTO_VOLUNTARIO'].includes(filters.tipo),'COMMAND_PAYLOAD_INVALID');
  N(!filters.modalidade||allModes.includes(filters.modalidade),'COMMAND_PAYLOAD_INVALID');
  if(filters.dataInicio)C.civilStart(filters.dataInicio);if(filters.dataFim)C.civilStart(filters.dataFim);N(!filters.dataInicio||!filters.dataFim||filters.dataInicio<=filters.dataFim,'COMMAND_PAYLOAD_INVALID');
  const get=p=>db.get(p),before=access(get,clock());N(before.review,'ACCESS_REVOKED');
  const fence=get('portalReaderResourceFences/members_operations')?.revision||0;
  const source=v=>({exit:v.reviewExit?.source||null,suspension:v.reviewSuspension?.source||null});
  const paging=f4MemberListPage_({uid:identity.uid,route:'/admin/solicitacoes-vinculo',source:source(before),actor:before.actor,fence,filters});
  let cursor=null;if(params.cursor){N(typeof params.cursor==='string'&&/^\d{13}:[a-f0-9]{32}\.[a-f0-9]{64}$/.test(params.cursor),'COMMAND_PAYLOAD_INVALID');const split=params.cursor.indexOf(':');cursor=[Number(params.cursor.slice(0,split)),paging.after(params.cursor.slice(split+1))];N(Number.isSafeInteger(cursor[0]),'COMMAND_PAYLOAD_INVALID');}
  const allowed=allModes.filter(m=>m==='SUSPENSAO_TEMPORARIA'?before.reviewSuspension:before.reviewExit).filter(m=>(!filters.modalidade||filters.modalidade===m)&&(!filters.tipo||(m==='SUSPENSAO_TEMPORARIA'?'SUSPENSAO_VOLUNTARIA':'DESLIGAMENTO_VOLUNTARIO')===filters.tipo));
  const queryFilters=[];for(const [input,field]of [['status','status'],['idSemestre','semesterId'],['membro','personId']])if(filters[input])queryFilters.push({field,op:'EQUAL',value:filters[input]});
  if(filters.dataInicio)queryFilters.push({field:'createdAt',op:'GREATER_THAN_OR_EQUAL',value:F.Timestamp.fromMillis(C.civilStart(filters.dataInicio))});
  if(filters.dataFim)queryFilters.push({field:'createdAt',op:'LESS_THAN',value:F.Timestamp.fromMillis(C.civilStart(filters.dataFim)+86400000)});
  // The two exit modalities need two bounded queries for an exit-only actor.
  // No collection scan, no hidden next-page loading, no client authorization.
  const groups=allowed.length===3?[null]:allowed;
  const rows=groups.flatMap(mode=>db.query({collection:'memberVoluntaryRequests',filters:queryFilters.concat(mode?[{field:'modality',op:'EQUAL',value:mode}]:[]),order:[{field:'createdAt',direction:'DESCENDING'},{field:'id',direction:'ASCENDING'}],limit:10,...(cursor?{cursor:[F.Timestamp.fromMillis(cursor[0]),cursor[1]]}:{})})).sort((a,b)=>C.ms(b.data.createdAt)-C.ms(a.data.createdAt)||a.data.id.localeCompare(b.data.id)).slice(0,10);
  return db.transaction(tx=>{const get=p=>tx.get(p).data(),current=access(get,clock());N(current.review&&K.digest(source(current))===K.digest(source(before))&&current.actor.actor.h.revision===before.actor.actor.h.revision&&(get('portalReaderResourceFences/members_operations')?.revision||0)===fence,'ACCESS_REVOKED');
   const items=rows.map(row=>{const v=W.validate(row.data);N(allowed.includes(v.modality),'ACCESS_REVOKED');return item(v);});
   return {ok:true,data:{items,pageSize:10,nextCursor:rows.length===10?String(C.ms(rows.at(-1).data.createdAt))+':'+paging.next(rows.at(-1).data.id):null}};
  });
 }

 function enrollments(params){
  C.exact(params,Object.hasOwn(params,'cursor')?['cursor']:[]);
  const get=p=>db.get(p),before=access(get,clock()),permission=before.ingress||before.supervision;N(permission,'ACCESS_REVOKED');const fence=get('portalReaderResourceFences/members_operations')?.revision||0;
  const paging=f4MemberListPage_({uid:identity.uid,route:'/admin/ingressos-membros',source:permission.source,actor:before.actor,fence,filters:{}}),after=paging.after(params.cursor);
  const rows=permission.kind==='DEV_TECHNICAL'?permission.source.syntheticPersonIds.flatMap(id=>{const index=get('memberEnrollmentIndexes/'+id),v=index?.activeEnrollmentId&&get('memberEnrollments/'+index.activeEnrollmentId);return v&&v.createdBy===identity.uid&&permission.source.admissionApprovalIds.includes(v.approvalId||f4MemberPreRegistration_().key(v.registration))&&(!after||v.id>after)?[{data:v}]:[];}).sort((a,b)=>a.data.id.localeCompare(b.data.id)).slice(0,10):db.query({collection:'memberEnrollments',filters:after?[{field:'id',op:'GREATER_THAN',value:after}]:[],order:[{field:'id',direction:'ASCENDING'}],limit:10});
  return db.transaction(tx=>{
   const get=p=>tx.get(p).data(),current=access(get,clock()),live=current.ingress||current.supervision;N(live&&live.kind===permission.kind&&K.digest(live.source)===K.digest(permission.source)&&(get('portalReaderResourceFences/members_operations')?.revision||0)===fence,'ACCESS_REVOKED');
   // Every enrollment command advances this fence atomically. Checking it in
   // the authorization transaction lets us reuse the query snapshot safely.
   const policy=get('memberOperationalConfig/registration'),invitation=f4MemberInvitation_(),canReview=live.kind!=='DEV_TECHNICAL'&&invitation.reviewer(current.actor,policy);
   const items=rows.map(row=>{const v=f4MemberEnrollmentContract_().validate(row.data),open=['PRE_REGISTERED','INVITED','ACCEPTED'].includes(v.status),
    releasable=(live.kind!=='DEV_TECHNICAL'||f4DevAdministrativeAccess_().invitationAllowed(get,live,v,clock()))&&(!invitation.requiresReview(policy)||v.review?.state==='ACKNOWLEDGED'),email=v.registration.email,at=email.indexOf('@');return {enrollmentId:v.id,personId:v.personId,revision:v.revision,name:v.registration.fullName,membershipType:v.approval.membershipType,approvalReference:v.approval.reference,recognizedOn:v.approval.recognizedEffects?.on||null,acceptedAt:v.respondedAt&&v.status!=='DECLINED'?new Date(v.respondedAt).toISOString():null,completedAt:v.completedAt?new Date(v.completedAt).toISOString():null,reviewState:v.review?.state||'PENDING',reviewNote:v.review?.note||'',canReview:canReview&&open,canAmend:live.kind!=='DEV_TECHNICAL'&&!!current.ingress&&open,canResolve:live.kind!=='DEV_TECHNICAL'&&!!current.ingress&&open&&v.review?.state==='CORRECTION_REQUESTED',emailMasked:current.ingress?email.slice(0,Math.min(2,at))+'***'+email.slice(at):null,status:v.status,expired:clock()>=v.expiresAt,createdAt:new Date(v.createdAt).toISOString(),canSend:!!current.ingress&&releasable&&v.status==='PRE_REGISTERED',canResend:live.kind!=='DEV_TECHNICAL'&&!!current.ingress&&releasable&&['INVITED','ACCEPTED'].includes(v.status),canCancel:live.kind!=='DEV_TECHNICAL'&&!!current.ingress&&['PRE_REGISTERED','INVITED','ACCEPTED'].includes(v.status)};});
   return {ok:true,data:{items,nextCursor:rows.length===10?paging.next(rows.at(-1).data.id):null,pageSize:10,notificationMode:get('memberOperationalConfig/notifications')?.mode==='DRY_RUN'?'DRY_RUN':'DISABLED'}};
  });
 }
 function directory(params){
  C.exact(params,['filtros']);let filters;try{filters=JSON.parse(params.filtros);}catch{N(false,'COMMAND_PAYLOAD_INVALID');}
  N(filters&&Object.keys(filters).every(k=>['texto','tipoVinculo','statusVinculo','perfilPortal','portalAtivo','comPendencias','situacaoFrequencia','pagina','pageSize','ordem'].includes(k)),'COMMAND_PAYLOAD_INVALID');
  N(filters.ordem===undefined||['title-asc','title-desc'].includes(filters.ordem),'COMMAND_PAYLOAD_INVALID');
  N(Number.isSafeInteger(filters.pagina)&&filters.pagina>=1&&filters.pageSize===25,'COMMAND_PAYLOAD_INVALID');
  const now=clock(),get=p=>db.get(p),a=access(get,now),fence=get('portalReaderResourceFences/members_operations')?.revision||0;N(a.directory,'ACCESS_REVOKED');if(!a.technical)K.subset(['idPessoa','displayLabel','state'],a.directory.source.fieldMasks.MEMBERS_CASE_READ);let ids;
  if(a.technical){ids=[];let cursor=null;for(let page=0;page<25;page++){const rows=db.query({collection:'memberAggregateHeads',filters:cursor?[{field:'id',op:'GREATER_THAN',value:cursor}]:[],order:[{field:'id',direction:'ASCENDING'}],limit:20});ids.push(...rows.map(v=>v.data.id));if(rows.length<20)break;N(page<24,'SCOPE_TOO_LARGE');cursor=rows.at(-1).data.id;}}
  else ids=a.directory.source.targets.allowedPersons;
  N(ids.length<=500,'SCOPE_TOO_LARGE');
  // Keep the existing bounded policy resource set. A new admission does not
  // silently grant its operator private access to unrelated people.
  const approved=a.technical?ids:ids.filter(id=>['memberIdentities/'+id,'memberCurrentStates/'+id].every(p=>a.directory.source.targets.allowedResources.includes(p)));
  const paths=approved.flatMap(id=>['memberAggregateHeads/'+id,'memberIdentities/'+id,'memberCurrentStates/'+id]),docs={};for(let i=0;i<paths.length;i+=180){const part=paths.slice(i,i+180),values=db.getMany(part);part.forEach((p,j)=>{docs[p]=values[j];});}
  // A missing/expired projection is not a missing member. Derive only those
  // states from the same canonical aggregate, without repairing facts or grants.
  const unresolved=approved.filter(id=>{const s=docs['memberCurrentStates/'+id],h=docs['memberAggregateHeads/'+id];return !s||s.sourceRevision!==h?.revision||!(C.ms(s.startsAt)<=now&&now<C.ms(s.validUntil));});
  const dependencyPaths=[...new Set(unresolved.flatMap(id=>docs['memberAggregateHeads/'+id]?.paths||[]))];
  for(let i=0;i<dependencyPaths.length;i+=180)db.getMany(dependencyPaths.slice(i,i+180));
  const repository=f4MemberRepository_(a.actor.control);
  const items=approved.filter(id=>docs['memberIdentities/'+id]?.technicalAccount!==true).map(id=>{const h=docs['memberAggregateHeads/'+id],person=docs['memberIdentities/'+id];let state=docs['memberCurrentStates/'+id];N(h&&person,'RECONCILIATION_REQUIRED');if(unresolved.includes(id)){const loaded=repository.load(get,id);N(loaded.reg.revision===h.revision,'REVISION_CONFLICT');state=f4MemberModel_(true,a.actor.control).derive(loaded.docs,id,h.revision,now).state;}N(state&&state.sourceRevision===h.revision&&now<C.ms(state.validUntil),'REVISION_CONFLICT');N(person.idPessoa===id&&!person.technicalAccount,'RECONCILIATION_REQUIRED');return {idPessoa:id,nomeExibicao:person.displayName,rga:null,email:null,tipoVinculoAtual:state.state,statusVinculoAtual:state.blocked?'INATIVO':'ATIVO',cargoFuncaoAtual:null,perfilPortalCalculado:null,portalAtivo:null,tempoEfetivoNoGrupo:null,qtdSemestresNoGrupo:null,qtdApresentacoesRealizadas:null,cicloUltimaApresentacao:null,frequenciaResumida:null,pendenciasAbertas:null,flagJaFoiSuspenso:null,statusElegibilidadeDiretoria:null,ultimaAtualizacao:new Date(now).toISOString()};});
  // Indicators outside this sanitized contract stay unknown, never fabricated.
  const filtered=items.filter(v=>(!filters.texto||(v.nomeExibicao+' '+v.idPessoa+' '+v.tipoVinculoAtual).toLocaleLowerCase().includes(String(filters.texto).toLocaleLowerCase()))&&(!filters.tipoVinculo||v.tipoVinculoAtual===filters.tipoVinculo)&&(!filters.statusVinculo||v.statusVinculoAtual===filters.statusVinculo));
  N(!filters.perfilPortal&&!filters.portalAtivo&&!filters.comPendencias&&!filters.situacaoFrequencia,'COMMAND_PAYLOAD_INVALID');
  // Order the authorized, filtered set before slicing. Old callers retain their order.
  if(filters.ordem)f4MemberDirectorySort_(filtered,filters.ordem);
  const current=access(get,clock());N(current.directory&&current.directory.kind===a.directory.kind&&C.hash(F.encode(current.directory.source))===C.hash(F.encode(a.directory.source))&&current.actor.actor.h.revision===a.actor.actor.h.revision,'ACCESS_REVOKED');N((get('portalReaderResourceFences/members_operations')?.revision||0)===fence,'STALE_REVISION');const pages=Math.max(1,Math.ceil(filtered.length/25)),page=Math.min(filters.pagina,pages);
  return {ok:true,data:{itens:filtered.slice((page-1)*25,page*25),paginacao:{pagina:page,pageSize:25,totalItens:filtered.length,totalPaginas:pages},opcoesFiltros:{tiposVinculo:[...new Set(items.map(v=>v.tipoVinculoAtual))],statusVinculo:['ATIVO','INATIVO'],perfisPortal:[],frequencias:[]},ultimaAtualizacaoResumo:new Date(now).toISOString()}};
 }
 function registrationApproval(params){
  const pair=f4MemberRegistrationIdentity_().normalize(params);
  return db.transaction(tx=>{
   const get=p=>tx.get(p).data(),now=clock(),a=access(get,now),permission=a.ingress;
   N(permission&&['ROLE','DEV_TECHNICAL'].includes(permission.kind),'AUTHORIZATION_DENIED');
   const approval=f4MemberPreRegistration_().locate(get,pair,null,now);
   if(permission.kind==='DEV_TECHNICAL')N(permission.mask.includes('preRegister')&&permission.source.admissionApprovalIds.includes(approval.id)&&approval.synthetic===true&&approval.devTechnicalOwnerUid===identity.uid&&permission.source.syntheticPersonIds.includes(approval.devPersonId),'AUTHORIZATION_DENIED');
   f4MemberRegistrationIdentity_().resolve(get,pair);
   return {ok:true,data:{status:'APPROVED',active:true,membershipType:approval.membershipType,approvedOn:approval.approvedOn,revision:approval.revision}};
  });
 }
 return {read,access,routes,registrationApproval};
}

function f4MemberDirectorySort_(items,order){
 const collator=new Intl.Collator('pt-BR',{sensitivity:'base',numeric:true}),direction=order==='title-desc'?-1:1;
 return items.sort((a,b)=>direction*collator.compare(a.nomeExibicao,b.nomeExibicao)||collator.compare(a.idPessoa,b.idPessoa));
}

// Cursor is navigation only, never an authorization proof. Binding it to the
// current actor, filters and mutation fence requires a restart after changes.
function f4MemberListPage_({uid,route,source,actor,fence,filters}){
 const C=f4Core_(),F=f4Firestore_(),scope=C.hash(F.encode({uid,route,source,personRevision:actor.actor.h.revision,authorityEpoch:actor.control.epoch,policyEpoch:actor.control.policyEpoch,readAuthzEpoch:actor.control.readAuthzEpoch,subjectEpoch:actor.subject?.authzEpoch||0,fence,filters}));
 return {after(cursor){if(!cursor)return null;C.need(typeof cursor==='string'&&/^[a-f0-9]{32}\.[a-f0-9]{64}$/.test(cursor),'COMMAND_PAYLOAD_INVALID');const [id,binding]=cursor.split('.');C.need(binding===scope,'REVISION_CONFLICT');return id;},next(id){C.need(C.op(id),'SCHEMA_INVALID');return id+'.'+scope;}};
}
