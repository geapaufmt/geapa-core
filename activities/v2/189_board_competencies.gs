// Read-only transparency. This DTO is never accepted as an authorization proof.
// Occupancy derives from the same canonical Member graph used by ROLE commands.
function f4BoardCompetencies_(options){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),db=F.connect(options.config),A=f4MemberAuthz_(),now=(options.clock||Date.now)();
 return options.ports.readWithAuthoritySnapshot((check,get)=>{
  const proof=options.ports.readAuthority(['activityAttendanceRegistry/current'],false,'ACTIVITY_BOARD_READ');
  const caller=A.identify(get,options.identity.uid,now),boards=[...new Set(caller.derived.currentFunctions.filter(f=>f.group==='DIRETORIA').map(f=>f.idDiretoria))];
  N(boards.length>0&&boards.length<=4,'AUTHORIZATION_DENIED');
  const memberConfig=get('memberOperationalConfig/portal'),recovery=get('projectionRecoveryControl/current');
  const policyIds=[...new Set([...(options.settings?.policyIds||[]),...(memberConfig?.policyIds||[])])];N(policyIds.length<=64,'SCOPE_TOO_LARGE');
  const policies=policyIds.map(id=>get('memberAuthzPolicies/'+id)).filter(Boolean),functions=[];
  for(const boardId of boards){let cursor;
   for(let page=0;page<5;page++){const rows=db.query({collection:'memberFunctions',filters:[{field:'idDiretoria',op:'EQUAL',value:boardId}],limit:20,order:[{field:'__name__',direction:'ASCENDING'}],...(cursor?{cursor:[cursor]}:{})});functions.push(...rows.map(r=>r.data));if(rows.length<20)break;N(page<4,'SCOPE_TOO_LARGE');cursor=new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+rows.at(-1).path);}
  }
  const ids=[...new Set(functions.map(f=>f.idPessoa))];N(ids.length<=40,'SCOPE_TOO_LARGE');
  // Discover each dependency layer in bulk; loading one aggregate field at a
  // time would waste SDK discovery rounds. All reads stay in one snapshot.
  const prefetch=paths=>{const unique=[...new Set(paths)];N(unique.length<=720,'SCOPE_TOO_LARGE');for(let i=0;i<unique.length;i+=180)db.getMany(unique.slice(i,i+180));};
  prefetch(ids.flatMap(id=>['memberAggregateHeads/'+id,'memberAuthzSubjects/'+id]));
  prefetch(ids.flatMap(id=>get('memberAggregateHeads/'+id)?.paths||[]));
  const allFunctions=ids.flatMap(id=>(get('memberAggregateHeads/'+id)?.paths||[]).filter(p=>p.startsWith('memberFunctions/')).map(get)).filter(Boolean);
  prefetch(allFunctions.flatMap(f=>['memberBoardHeads/'+f.idDiretoria,'memberRoleCatalog/'+f.idCargo,'memberGovernanceTerms/'+f.idTermo]));
  const boardPaths=[...new Set(allFunctions.flatMap(f=>get('memberBoardHeads/'+f.idDiretoria)?.paths||[]))];prefetch(boardPaths);
  prefetch(boardPaths.filter(p=>p.startsWith('memberGovernanceBoardTerms/')).map(get).filter(Boolean).map(r=>'memberGovernanceTerms/'+r.idTermo));
  const contracts={...Object.fromEntries(Object.entries(f4ActivitiesContracts_()).map(([k,v])=>['ACT.'+k,v.capabilities])),...f4MemberCommandCatalog_().types};
  // These are the existing protected route documents, not a new permission matrix.
  const routes=[];let cursor;
  for(let page=0;page<8;page++){const rows=db.query({collection:'commandExecutorRoutes',filters:[],limit:20,order:[{field:'__name__',direction:'ASCENDING'}],...(cursor?{cursor:[cursor]}:{})});routes.push(...rows.map(r=>r.data));if(rows.length<20)break;N(page<7,'SCOPE_TOO_LARGE');cursor=new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+rows.at(-1).path);}
  const bindings=(recovery?.actorUids||[]).map(uid=>get('memberAuthBindings/'+uid)).filter(b=>b?.active&&b.sourceRevision===b.revision),items=[];
  for(const id of ids){const loaded=A.load(get,id,caller.control),derived=loaded.M.derive(loaded.d,id,loaded.h.revision,now),person=loaded.d['memberIdentities/'+id],subject=get('memberAuthzSubjects/'+id);
   for(const fn of derived.currentFunctions.filter(f=>f.group==='DIRETORIA'&&boards.includes(f.idDiretoria))){
    const role=loaded.d['memberRoleCatalog/'+loaded.d['memberFunctions/'+fn.idVigencia].idCargo];
    const eligible=subject?.active===true&&!derived.state.blocked&&loaded.M.functionEligible(derived.state.state,role);
    const matching=policies.filter(p=>p.active&&p.proofKind==='ROLE'&&p.actorSelector.roleKey===fn.roleKey&&p.actorSelector.actorScope===fn.scopeKey&&C.ms(p.issuedAt)<=now&&now<C.ms(p.expiresAt));
    const capabilities=[...new Set(matching.flatMap(p=>p.capabilities))].sort();
    const permissions=capabilities.map(cap=>{
     const sources=matching.filter(p=>p.capabilities.includes(cap)),act=cap.startsWith('ACTIVITY_')||cap.startsWith('PRESENTATION_');
     const permitted=sources.filter(p=>act?f4ActivitiesAuthzDomain_().contracts().roleCapability(cap,p.actorSelector)&&p.targets.allowedPersons.includes(id):f4MemberAuthzContracts_().roleCapability(cap,p.actorSelector));
     const uids=bindings.filter(b=>b.idPessoa===id).map(b=>b.uid),read=cap.endsWith('_READ')||cap==='MEMBERS_CASE_READ';
     const commands=Object.keys(contracts).filter(c=>contracts[c].includes(cap));
     const activeRoutes=routes.filter(r=>commands.includes(r.commandType)&&r.state==='ACTIVE'&&r.projectId===options.config.projectId&&f4ExecutorWindow_().live(r,now)&&r.actorUids?.some(uid=>uids.includes(uid)));
     const enabled=eligible&&uids.length>0&&permitted.length>0&&(read||activeRoutes.length>0);
     return {capability:cap,label:f4BoardCapabilityLabel_(cap),module:cap.startsWith('PEOPLE_')?'Pessoas':cap.startsWith('MEMBERS_')?'Membros':cap.startsWith('PRESENTATION_')?'Apresentações':'Atividades e frequência',action:read?'Consulta':/DECIDE|REVIEW|RECOGNIZE|EXECUTE|CORRECT/.test(cap)?'Decisão administrativa':'Gerenciamento',status:enabled?'CONDITIONAL':'NOT_ENABLED',limits:{cycleIds:[...new Set(permitted.flatMap(p=>p.targets.resourceScope?.cycleIds||[]))],nominalResources:permitted.some(p=>p.targets.allowedResources.length>0),nominalPersons:permitted.some(p=>p.targets.allowedPersons.length>0),policyValidUntil:permitted.length?new Date(Math.min(...permitted.map(p=>C.ms(p.expiresAt)))).toISOString():null},reason:!eligible?'Vínculo ou função sem elegibilidade atual':!uids.length?'Acesso pessoal ainda não validado':!permitted.length?'Sem autorização aplicável ao titular':!enabled?'Operação não habilitada neste ambiente':'Sujeita ao recurso, ciclo, estado, campos e procedimento autorizado'};
    });
    const tenure=f4MemberTenurePresentation_(loaded.d,id,derived,now).tenures.find(t=>t.segments.some(s=>s.idVigencia===fn.idVigencia));
    items.push({procedureSummary:f4BoardProcedures_().officeSummary(get,{config:options.config,identity:options.identity,clock:()=>now},fn,loaded.M.rows(loaded.d,'memberFunctions')),procedureActions:f4BoardProcedures_().contextual(get,{config:options.config,identity:options.identity,clock:()=>now},fn),personId:id,name:person.displayName,roleKey:fn.roleKey,boardId:fn.idDiretoria,startDate:tenure.startDate,endDate:tenure.endDate,tenure,currentRecordEnd:fn.endDateReal||fn.endDatePlanned||null,sourceRevision:loaded.h.revision,duties:f4BoardDuties_(fn.roleKey),permissions});
   }
  }
  let history=null;try{history=f4BoardHistory_(get,{config:options.config,identity:options.identity,clock:()=>now});}catch(e){if(!['PILOT_DISABLED','AUTHORIZATION_DENIED','M13_AUTHZ_DENIED','ACCESS_REVOKED'].includes(e.code))throw e;}check([proof]);const result={ok:true,data:{schemaVersion:'board-competencies-v1',asOf:new Date(now).toISOString(),validUntil:Math.min(options.identity.expiresAt,now+60000),items,history,notice:'Consulta informativa. Cada operação revalida vínculo, função, policy, recurso e precondições. Atribuição institucional não concede acesso técnico automaticamente.'}};return options.accessCatalog?f4MemberAccessCatalog_(result,caller.control,policies,now):result;
 });
}
function f4BoardDuties_(role){
 // Documentary summaries of the verified Regimento vigente (14/05/2026), art.29.
 // Temporary offices require their own appointment act; never infer its content.
 const duties={PRESIDENTE:'Representar o grupo; presidir reuniões; supervisionar planejamento e atividades; coordenar a transição de gestão.',VICE_PRESIDENTE:'Auxiliar a Presidência e substituí-la nos impedimentos e ausências.',SECRETARIO_GERAL:'Coordenar atas, documentos e comunicações internas; supervisionar o armazenamento de materiais; substituir a Vice-Presidência nas condições regimentais.',SECRETARIO_EXECUTIVO:'Apoiar a Secretaria Geral, organizar documentos e reuniões e responder pela organização e movimentação dos recursos financeiros.',DIRETOR_COMUNICACAO:'Planejar e executar a comunicação externa, redes sociais e divulgação de eventos e projetos.'};
 return duties[role]?{source:'Regimento GEAPA 2026, art. 29',summary:duties[role],verified:true}:{source:'Regimento GEAPA 2026, art. 30; ato específico de criação e nomeação',summary:'As atribuições específicas dependem do ato do cargo. Sua síntese documental ainda não foi vinculada a esta consulta.',verified:false};
}
function f4BoardCapabilityLabel_(cap){const labels={PEOPLE_DIRECTORY_READ:'Consultar diretório institucional de pessoas',ACTIVITY_BOARD_READ:'Consultar atividades administrativas',ACTIVITY_MANAGE:'Gerenciar atividades',ACTIVITY_REALIZATION_MANAGE:'Registrar realização',ACTIVITY_ATTENDANCE_MANAGE:'Gerenciar chamadas e justificativas',ACTIVITY_CONFIG_MANAGE:'Configurar modelos de atividades',MEMBERS_CASE_READ:'Consultar casos de membros',MEMBERS_PROFILE_CORRECT:'Aplicar correções cadastrais',MEMBERS_VOLUNTARY_EXIT_EXECUTE:'Processar desligamento voluntário',MEMBERS_VOLUNTARY_SUSPENSION_EXECUTE:'Processar suspensão voluntária',MEMBERS_INGRESS_EXECUTE:'Cadastrar ingresso aprovado',PRESENTATION_BOARD_READ:'Consultar apresentações administrativas',PRESENTATION_PLAN_MANAGE:'Gerenciar planejamento',PRESENTATION_ACTIVITY_CREATE:'Vincular atividade à apresentação',PRESENTATION_COMPOSITION_MANAGE:'Gerenciar composição',PRESENTATION_THEME_REVIEW:'Analisar temas',PRESENTATION_COMPATIBILITY_REVIEW:'Analisar compatibilidade',PRESENTATION_MATERIAL_REVIEW:'Analisar materiais',PRESENTATION_REALIZATION_RECOGNIZE:'Reconhecer realização',PRESENTATION_PHOTO_MANAGE:'Gerenciar fotos',PRESENTATION_PUBLICATION_MANAGE:'Gerenciar publicação',PRESENTATION_DYNAMIC_DECIDE:'Decidir apresentações dinâmicas'};return labels[cap]||cap.replaceAll('_',' ').toLowerCase();}
