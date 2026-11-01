// Bounded private dossiers. These are kept behind the Core because analysis
// reveals sensitive proposed values, not a general member-directory projection.
function f4MemberProfileCorrectionReader_({config,identity,clock=Date.now,auditClock=clock}){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),K=f4MemberAuthzContracts_(),W=f4MemberProfileCorrections_(),db=F.connect(config);
 const routes=['/meu-perfil/correcoes','/admin/correcoes-cadastrais','/admin/correcoes-cadastrais/detalhe'];
 const fields={displayName:'NOME_COMPLETO',cpf:'CPF',rga:'RGA',birthDate:'DATA_NASCIMENTO',countryCode:'PAIS_ORIGEM',courseId:'CURSO_ID'};
 function access(get,now,self){
  N(C.id(identity.uid)&&Number.isSafeInteger(identity.expiresAt)&&now<identity.expiresAt,'ACCESS_REVOKED');
  const technical=f4DevAdmin_().profile(get,identity.uid,now),actor=technical?f4DevAdmin_().identify(get,identity.uid,now):f4MemberAuthz_().identify(get,identity.uid,now),settings=f4MemberPortalPrepare_({config,identity,clock}).configuration(get);
  N(settings.commands.includes(self?'members.self.update':'members.profile.correct'),'ACCESS_REVOKED');
  if(technical)return {actor,source:technical,mask:W.fields,persons:null,expires:Math.min(identity.expiresAt,C.ms(technical.expiresAt))};
  const cap=self?'MEMBERS_SELF_UPDATE':'MEMBERS_PROFILE_CORRECT';
  const source=settings.policyIds.map(id=>get('memberAuthzPolicies/'+id)).find(p=>{
   if(!p?.active||!p.capabilities?.includes(cap))return false;K.policy(p);if(!(C.ms(p.issuedAt)<=now&&now<C.ms(p.expiresAt)))return false;
   if(self)return p.proofKind==='SELF'&&p.targets.ownPerson&&p.targets.commandScopes?.includes('OWN_MEMBER_RESOURCES')&&['MEMBRO_INGRESSANTE','MEMBRO_EFETIVO'].includes(actor.state.state);
   if(p.proofKind!=='ROLE'||p.targets.ownPerson||!p.targets.commandScopes?.includes('PROFILE_CORRECTION')||!p.targets.allowedTargetScopes.includes('ADMINISTRACAO'))return false;
   return actor.derived.currentFunctions.some(f=>f.roleKey===p.actorSelector.roleKey&&f.group===p.actorSelector.group&&f.scopeKey===p.actorSelector.actorScope);
  });N(source,'ACCESS_REVOKED');
  // Empty legacy DEV role scopes remain compatible. PROD must name its targets.
  const persons=self?[actor.binding.idPessoa]:source.targets.allowedPersons.length?source.targets.allowedPersons:f4ExecutionEnvironment_().production?[]:null;
  return {actor,source,persons,mask:self?W.fields:W.fields.filter(f=>source.fieldMasks[cap].includes(f)),expires:Math.min(identity.expiresAt,C.ms(source.expiresAt),C.ms(actor.state.validUntil))};
 }
 function dto(get,v,reveal){const privateField=['cpf','rga','birthDate'].includes(v.field),mask=x=>x==null?'':privateField?'••••'+String(x).slice(-2):String(x),identity=get('memberIdentities/'+v.personId);N(identity?.idPessoa===v.personId,'RECONCILIATION_REQUIRED');
  return {id:v.id,revision:v.revision,campo:fields[v.field],field:v.field,status:v.status,solicitadoEm:new Date(v.createdAt).toISOString(),valorSolicitadoMascarado:mask(v.proposedValue),pessoa:{nomeExibicao:identity.displayName,rgaMascarado:'',emailMascarado:''},requerRevelacao:true,dadosRevelados:reveal,valorAtual:reveal?W.source(get,v.personId,v.field).value:mask(W.source(get,v.personId,v.field).value),valorSolicitado:reveal?v.proposedValue:mask(v.proposedValue),...(reveal?{valorAnterior:v.previousValue,justificativa:v.history[0].reason,historicoDecisoes:v.history.map(h=>({status:h.action,motivoPublico:h.reason,analisadoEm:new Date(h.at).toISOString()})),canAnalyze:!W.terminal.includes(v.status),canApply:['PENDENTE','EM_ANALISE'].includes(v.status)}:{canAnalyze:false,canApply:false})};
 }
 function read(input){
  C.exact(input,['route','params']);const {route,params}=input;
  N(routes.includes(route),'COMMAND_TYPE_INVALID');
  if(route==='/meu-perfil/correcoes')return db.transaction(tx=>{
   C.exact(params,[]);const get=p=>tx.get(p).data(),a=access(get,clock(),true),id=a.actor.binding.idPessoa,ix=W.index(get('memberProfileCorrectionIndexes/'+id),id);
   const items=ix.requestIds.map(key=>W.record(get('memberProfileCorrections/'+key))).map(v=>{N(v.personId===id,'RECONCILIATION_REQUIRED');return {...dto(get,v,true),canCancel:!W.terminal.includes(v.status),canComplement:v.status==='COMPLEMENTO_SOLICITADO'};});
   const courseIds=get('memberOperationalConfig/registration')?.courseIds||[];N(Array.isArray(courseIds)&&courseIds.length<=100&&courseIds.every(C.id),'SCHEMA_INVALID');
   const courses=courseIds.map(id=>get('memberCourseCatalog/'+id)).filter(c=>c?.active).map(c=>({id:c.id,name:c.name}));
   return {ok:true,data:{items,fields:W.fields,courses,personRevision:a.actor.actor.h.revision,validUntil:a.expires}};
  });
  const detail=route.endsWith('/detalhe');C.exact(params,detail?['idSolicitacao','revelarDados']:['cursor','status','field']);N(detail?C.op(params.idSolicitacao)&&typeof params.revelarDados==='boolean':(!params.status||W.statuses.includes(params.status))&&(!params.field||W.fields.includes(params.field)),'COMMAND_PAYLOAD_INVALID');
  const before=access(p=>db.get(p),clock(),false),fence=db.get('portalReaderResourceFences/members_operations')?.revision||0;
  const paging=detail?null:f4MemberListPage_({uid:identity.uid,route,source:before.source,actor:before.actor,fence,filters:{status:params.status,field:params.field}}),after=paging?.after(params.cursor);
  const filters=[];if(after)filters.push({field:'id',op:'GREATER_THAN',value:after});if(params.status)filters.push({field:'status',op:'EQUAL',value:params.status});if(params.field){N(before.mask.includes(params.field),'ACCESS_REVOKED');filters.push({field:'field',op:'EQUAL',value:params.field});}
  // Named targets use their canonical indexes, never a scan of private dossiers.
  // Fetch at most one page; cursor/fence still bind actor, policy and generation.
  let rows;
  if(detail)rows=[{path:'memberProfileCorrections/'+params.idSolicitacao}];
  else if(before.persons!==null){
   N(before.persons.length<=100,'SCOPE_TOO_LARGE');
   const ids=[...new Set(before.persons.flatMap(id=>W.index(db.get('memberProfileCorrectionIndexes/'+id),id).requestIds))].sort().filter(id=>!after||id>after);
   rows=ids.slice(0,10).map(id=>({path:'memberProfileCorrections/'+id}));
  }else rows=db.query({collection:'memberProfileCorrections',filters,order:[{field:'id',direction:'ASCENDING'}],limit:10});
  return db.transaction(tx=>{
   const get=p=>tx.get(p).data(),a=access(get,clock(),false);N(K.digest(a.source)===K.digest(before.source)&&a.actor.actor.h.revision===before.actor.actor.h.revision&&(get('portalReaderResourceFences/members_operations')?.revision||0)===fence,'ACCESS_REVOKED');
   const values=rows.map(v=>W.record(v.data||get(v.path))),permitted=v=>a.persons===null||a.persons.includes(v.personId);
   if(detail)N(permitted(values[0])&&a.mask.includes(values[0].field),'ACCESS_REVOKED');
   else N(values.every(permitted),'ACCESS_REVOKED');
   const items=values.filter(v=>a.mask.includes(v.field)&&(detail||(!params.status||v.status===params.status)&&(!params.field||v.field===params.field))).map(v=>dto(get,v,detail&&params.revelarDados));
   if(detail){if(params.revelarDados){const id=C.hash({uid:identity.uid,request:values[0].id,at:auditClock(),purpose:'PROFILE_CORRECTION_READ'}).slice(0,32),p='memberAudit/'+id;get(p);tx.set(p,{schemaVersion:'member-profile-correction-read-v1',actorUid:identity.uid,requestId:values[0].id,requestRevision:values[0].revision,createdAt:F.Timestamp.fromMillis(auditClock()),authorizationSource:a.source.id||'DEV_ADMIN',authorizationRevision:a.source.revision});}return {ok:true,data:items[0]};}
   return {ok:true,data:{solicitacoes:items,nextCursor:rows.length===10?paging.next(values.at(-1).id):null,pageSize:10}};
  });
 }
 return {read,routes,access};
}
