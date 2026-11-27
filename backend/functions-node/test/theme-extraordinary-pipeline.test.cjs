const {test}=require('node:test'),a=require('node:assert/strict'),{scenario}=require('./theme-self-preparation.test.cjs');
const {encode,decode}=require('../src/firestore.cjs'),clone=v=>decode(encode(v));
const ref=(kind,id)=>({kind,id}),commands=['THEME_EXTRAORDINARY_AUTHORIZE','THEME_EXTRAORDINARY_REVOKE','THEME_REGULARIZE','THEME_REGULARIZE_ADMIN','THEME_VALIDATE','THEME_COMPLEMENT'];
function fixture(realized=false){
 const s=scenario(),self=clone(s.docs['memberAuthzPolicies/CF_DEV_ACT_POLICY']);let serial=10;
 const board={...clone(self),proofKind:'ROLE',actorSelector:{kind:'ROLE',roleKey:'DIRETOR_ADMINISTRATIVO',group:'DIRETORIA',actorScope:'ADMINISTRACAO'},capabilities:['PRESENTATION_BOARD_READ','PRESENTATION_THEME_REVIEW'],fieldMasks:{PRESENTATION_BOARD_READ:[],PRESENTATION_THEME_REVIEW:[]},targets:{...self.targets,ownPerson:false,allowedPersons:['PES-000001','PES-000002'],allowedTargetScopes:['PRESC_SYN_THEME_CYCLE'],allowedResources:[],resourceScope:{kind:'PRESENTATIONS'}}};
 const unit=()=>s.docs['presentationUnitHeads/SYN_THEME_UNIT'];
 // No ordinary deadline exception: extraordinary command must stand on its own.
 s.docs['presentationPolicyReferences/SYN_THEME_POLICY'].items[0].deadlineDecisions=[];
 s.docs['presentationFlowCases/SYN_THEME_BASIS'].itemDigest=JSON.parse(s.domain.run(x=>f4ActivitiesModel_().hash(x),s.docs['presentationPolicyReferences/SYN_THEME_POLICY'].items[0]));
 if(realized){unit().presentationRef=ref('presentation','SYN_PR');unit().quotaOccupancyRef=ref('quota','SYN_QUOTA');unit().obligationRef=ref('obligation','SYN_OB');
  s.docs['canonicalPresentations/SYN_PR']={id:'SYN_PR',revision:1,recordMode:'NEW_V2',unitRef:ref('unit',unit().id),personRef:unit().personRef,cycleRef:unit().cycleRef,realizationRef:ref('realization','SYN_RZ'),title:'Título aprovado'};
  s.docs['presentationRealizations/SYN_RZ']={id:'SYN_RZ',revision:1,recognized:true,occurred:true,presentationRef:ref('presentation','SYN_PR')};
  s.docs['presentationQuotaOccupancies/SYN_QUOTA']={id:'SYN_QUOTA',revision:4,phase:'R'};
  s.docs['presentationObligations/SYN_OB']={id:'SYN_OB',revision:3,state:'FULFILLED'};
  s.docs['presentationSessions/SYN_THEME_SESSION'].civilDate='2026-01-01';
 }
 function policy(admin){s.docs['memberAuthzPolicies/CF_DEV_ACT_POLICY']=clone(admin?board:self);}
 function prepare(action,extra={},admin=action!=='regularizar'){
  policy(admin);const u=unit(),t=u.thematicRequestRef&&s.docs['presentationThematicRequests/'+u.thematicRequestRef.id];
  const input={operationId:(++serial).toString(16).padStart(32,'0'),route:'/v2/apresentacoes/titulo-eixo/'+action,basis:{unitRevision:u.revision,thematicRevision:t?.revision||0},payload:{idApresentacao:'APU_'+u.id,...extra}};
  return JSON.parse(s.domain.run(x=>{const o={config:nodeConfig,identity:{uid:'CF_ADMIN',expiresAt:x.now+60000},settings:{commands:x.commands,policyIds:['CF_DEV_ACT_POLICY']},clock:()=>x.now};o.ports=f4ActivitiesFunctionalPorts_(o);return f4ActivitiesPresentationCompat_(o).prepare(x.input);},{input,now:s.now,commands}));
 }
 const content={tituloApresentacao:'Tema regularizado',observacoes:'Resumo recebido hoje',eixoTematicoPrincipal:'SYN_AXIS',eixoTematicoSecundario:''};
 function apply(action,payload,admin){const r=prepare(action,payload,admin);a(r.target.expected.length<30);a(!r.target.expected.some(x=>/^SYN_UNUSED_\d|SYN_UNRELATED|SYN_OTHER/.test(x.id)));const out=s.plan(r,true);a.equal(out.response.businessState,'COMMIT_APPLIED',JSON.stringify(out.response));return {r,out};}
 return {...s,unit,prepare,apply,policy,content,grant:()=>apply('autorizar',{confirmed:true,motivo:'Falha operacional comprovada'})};
}
test('full ROLE grant / SELF scheduled registration / stale repeat; no automatic approval',()=>{
 const s=fixture();s.grant();const {r,out}=s.apply('regularizar',s.content,false);a.equal(out.response.result.state,'ANALYSIS');a.equal(s.unit().extraordinaryTheme.status,'CONSUMED');a.throws(()=>s.plan(r),/STALE_REVISION/);
 const t=s.docs['presentationThematicRequests/'+s.unit().thematicRequestRef.id];a.equal(t.clockState,'MANUAL_ONLY');a.equal(t.nextEvaluationAt,undefined);
 a(out.writes.some(w=>w.path.includes('/revisions/')));a(out.writes.some(w=>w.path.startsWith('presentationFlowCases/')));
});
test('realized registration is append-only: approved content, quota, obligation and realization unchanged',()=>{
 const s=fixture(true),paths=['canonicalPresentations/SYN_PR','presentationRealizations/SYN_RZ','presentationQuotaOccupancies/SYN_QUOTA','presentationObligations/SYN_OB','presentationSessions/SYN_THEME_SESSION'];
 const before=paths.map(p=>structuredClone(s.docs[p]));s.grant();const {out}=s.apply('regularizar',s.content,false);a.equal(out.response.result.state,'RECORDED_NOT_APPROVED');paths.forEach((p,i)=>a.deepEqual(structuredClone(s.docs[p]),before[i]));a(!out.writes.some(w=>w.path.startsWith('presentationThematicRequests/')));const log=s.docs['presentationFlowCases/'+out.response.result.recordRef.id];a.equal(log.content.titleOrObject,s.content.tituloApresentacao);a.equal(log.actor.personId,'PES-000001');a.equal(log.recordedAt.toMillis(),s.now);
});
test('revocation, unauthorized actor and concurrent grant do not write',()=>{
 const s=fixture();const grant=s.prepare('autorizar',{confirmed:true,motivo:'Motivo'});s.policy(false);a.throws(()=>s.plan(grant),/AUTHORIZATION|AUTHZ|STALE/);s.grant();s.apply('revogar',{confirmed:true});const out=s.plan(s.prepare('regularizar',s.content,false));a.equal(out.response.businessState,'REJECTED');a.equal(out.writes.length,0);
});
test('administrative registration needs explicit attestation and dated institutional evidence',()=>{
 const s=fixture();s.grant();const p={...s.content,confirmed:true,institutionalEvidence:{reference:'Protocolo sintético DEV',receivedAt:s.now-1000}};const {out}=s.apply('registrar',p);const log=s.docs['presentationFlowCases/'+out.response.result.requestRef.id];const flow=out.writes.find(w=>w.path.startsWith('presentationFlowCases/'));a.equal(flow.data.submissionChannel,'ADMIN_INSTITUTIONAL');a.equal(flow.data.institutionalEvidence.attestedBy.id,'PES-000001');
});
test('request complement and SELF answer keep extraordinary analysis manual',()=>{
 const s=fixture();s.grant();s.apply('regularizar',s.content,false);s.apply('revisar',{decisao:'SOLICITAR_AJUSTE',observacaoPublica:'Detalhar o resumo',observacaoInterna:''});const t=s.docs['presentationThematicRequests/'+s.unit().thematicRequestRef.id];a.equal(t.analysisClock.state,'MANUAL_PAUSED');
 // Existing /enviar route maps the canonical correction case to THEME_COMPLEMENT.
 const {out}=s.apply('enviar',{...s.content,observacoes:'Resumo complementado'},false);a.equal(out.response.businessState,'COMMIT_APPLIED');const next=s.docs['presentationThematicRequests/'+s.unit().thematicRequestRef.id];a.equal(next.clockState,'MANUAL_ONLY');a.equal(next.nextEvaluationAt,undefined);
});
test('realized addendum preserves an approved thematic head and its immutable revision',()=>{
 const s=fixture(true),u=s.unit();u.thematicRequestRef=ref('thematic','SYN_APPROVED');
 const theme={id:'SYN_APPROVED',revision:7,unitRef:ref('unit',u.id),personRef:u.personRef,cycleRef:u.cycleRef,workflowState:'APPROVED',principalAxisRef:ref('policy','SYN_AXIS_2'),acceptedRevisionRef:{kind:'thematicRevision',id:'SYN_APPROVED',number:6},candidate:{titleOrObject:'Texto aprovado e imutável',summary:'Resumo aprovado',principalAxisRef:ref('policy','SYN_AXIS_2')}};
 s.docs['presentationThematicRequests/SYN_APPROVED']=theme;s.docs['presentationThematicRequests/SYN_APPROVED/revisions/6']={snapshot:clone(theme)};
 const before=clone(theme),snapshot=clone(s.docs['presentationThematicRequests/SYN_APPROVED/revisions/6']);s.grant();s.apply('regularizar',s.content,false);
 a.deepEqual(s.docs['presentationThematicRequests/SYN_APPROVED'],before);a.deepEqual(s.docs['presentationThematicRequests/SYN_APPROVED/revisions/6'],snapshot);
});
test('the same review capability held by another board role cannot grant regularization',()=>{
 const s=fixture();s.docs['memberRoleCatalog/CARGO-001'].roleKey='PRESIDENTE';a.throws(()=>s.prepare('autorizar',{confirmed:true,motivo:'Não autorizado'}),/AUTHORIZATION|AUTHZ/);
});
test('future institutional receipt date or missing explicit confirmation never consumes the grant',()=>{
 for(const extra of [{confirmed:false,institutionalEvidence:{reference:'Protocolo',receivedAt:Date.now()-1000}},{confirmed:true,institutionalEvidence:{reference:'Protocolo',receivedAt:Date.now()+86400000}}]){
  const s=fixture();s.grant();const out=s.plan(s.prepare('registrar',{...s.content,...extra}));a.equal(out.response.businessState,'REJECTED');a.equal(out.writes.length,0);a.equal(s.unit().extraordinaryTheme.status,'ACTIVE');
 }
});
module.exports={fixture};

test('direct scheduled registration respects the approved canonical analysis clock',()=>{
 const s=fixture(),cfg=s.docs['presentationCycles/SYN_THEME_CYCLE/configRevisions/CFG'];
 cfg.parameters.clockPolicyRef=ref('policy','SYN_CLOCK');s.docs['presentationPolicyReferences/SYN_CLOCK']={id:'SYN_CLOCK',revision:1,approved:true,active:true,analysisHours:48};
 const {out}=s.apply('registrar',{...s.content,motivo:'Registro sintético',confirmed:true});
 const t=s.docs['presentationThematicRequests/'+out.response.result.requestRef.id];
 a.equal(t.analysisClock.state,'RUNNING');a.equal(t.analysisClock.dueAt.toMillis(),s.now+48*3600000);a(!t.extraordinarySubmission);
});

test('director registers directly after the deadline, and only once against the displayed revisions',()=>{
 const s=fixture(),p={...s.content,motivo:'Falha de registro',confirmed:true};
 const first=s.prepare('registrar',p),concurrent=s.prepare('registrar',p);
 const out=s.plan(first,true);
 a.equal(out.response.result.state,'ANALYSIS',JSON.stringify(out.response));
 const log=out.writes.find(w=>w.path.startsWith('presentationFlowCases/')).data;
 a.equal(log.reason,p.motivo);a.equal(log.registeredBy.id,'PES-000001');a.equal(log.submissionChannel,'ADMIN_INSTITUTIONAL');
 a(!log.authorizationRef);a(!s.unit().extraordinaryTheme);a.throws(()=>s.plan(concurrent),/STALE_REVISION/);
});
test('direct registration refuses open deadlines, missing reasons, revoked access, and an approved scheduled theme',()=>{
 const p={...fixture().content,confirmed:true,motivo:'Motivo sintético'};
 for(const scenario of ['in-time','missing-reason','approved','revoked']){
  const s=fixture();
  if(scenario==='in-time')s.docs['presentationCycles/SYN_THEME_CYCLE/configRevisions/CFG'].parameters.thematicMinimumDays=1;
  if(scenario==='approved'){s.unit().thematicRequestRef=ref('thematic','SYN_APPROVED');s.docs['presentationThematicRequests/SYN_APPROVED']={id:'SYN_APPROVED',revision:2,workflowState:'APPROVED',candidate:{titleOrObject:'T?tulo aprovado'},unitRef:ref('unit',s.unit().id),cycleRef:s.unit().cycleRef,personRef:s.unit().personRef};}
  const r=s.prepare('registrar',{...p,...(scenario==='missing-reason'?{motivo:'   '}:{})});
  if(scenario==='revoked'){s.policy(false);a.throws(()=>s.plan(r),/AUTHORIZATION|AUTHZ|STALE/);continue;}
  const out=s.plan(r);a.equal(out.response.businessState,'REJECTED',scenario);a.equal(out.writes.length,0);
 }
});
function fileFixture(section){
 const s=fixture(true),pr=s.docs['canonicalPresentations/SYN_PR'],photo=section==='PHOTO';
 const speaker=ref('person','PES-000002');s.unit().personRef=speaker;pr.personRef=speaker;
 pr.materialRequirementRef=ref('requirement','SYN_M');pr.photoPolicyRef=ref('policy','SYN_PHOTO');
 const rq=s.docs['presentationMaterialRequirements/SYN_M']={id:'SYN_M',revision:1,presentationRef:s.unit().presentationRef,personRef:speaker,cycleRef:s.unit().cycleRef,currentRevisionRef:{kind:'requirementRevision',id:'SYN_M',number:1},uploadPolicyRef:ref('policy','SYN_FILE'),openCorrectionCount:0,reviewEpoch:0,opportunitiesEpoch:0,materialStateRevision:1};
 const desc={name:photo?'foto.png':'slides.pdf',mime:photo?'image/png':'application/pdf',size:10,sha256:'a'.repeat(64)};
 s.docs['presentationPolicyReferences/'+(photo?'SYN_PHOTO':'SYN_FILE')]={id:photo?'SYN_PHOTO':'SYN_FILE',revision:1,approved:true,active:true,usage:photo?'PRESENTATION_PHOTO':'MATERIAL',allowBoardUpload:true,cycleRef:s.unit().cycleRef,workspace:{enabled:true,maximumBytes:2048,mimeTypes:[desc.mime]}};
 const p={section,fileRevision:photo?0:1,motivo:'Arquivo recebido pela secretaria',confirmed:true,institutionalEvidence:{reference:'Protocolo sintético',receivedAt:s.now-3600000}};
 return {s,rq,pr,desc,p};
}
test('file registration does not inherit the separate thematic deadline',()=>{
 const {s,pr,desc,p}=fileFixture('MATERIAL');delete pr.realizationRef;
 s.docs['presentationSessions/SYN_THEME_SESSION'].civilDate=new Date(s.now+5*86400000).toISOString().slice(0,10);
 s.docs['presentationCycles/SYN_THEME_CYCLE/configRevisions/CFG'].parameters.thematicMinimumDays=1;
 const {out}=s.apply('registrar',{...p,fileDescriptor:desc});a(out.response.result.effectRef);
});
for(const section of ['PHOTO','MATERIAL'])test(section+' uses a confirmed canonical artifact and records the actual director and presenter separately',()=>{
 const {s,desc,p}=fileFixture(section);
 const initial=s.apply('registrar',{...p,fileDescriptor:desc}),effectRef=initial.out.response.result.effectRef,out=s.docs['jobOutbox/'+effectRef.id];
 a.equal(out.actor.personId,'PES-000001');a.equal(initial.out.response.result.provider,'GOOGLE_DRIVE_V3');
 const pending=s.plan(s.prepare('registrar',{...p,effectRef}));a.equal(pending.response.businessState,'REJECTED');a.equal(pending.writes.length,0);
 // Provider confirmation is a synthetic fixture for this domain-level test.
 Object.assign(out,{status:'CONFIRMED',revision:2,artifactRef:ref('artifact','SYN_FILE_VER'),workspaceReservation:{fileId:'SYN_OBJECT'}});
 s.docs['presentationArtifacts/SYN_FILE_VER']={id:'SYN_FILE_VER',revision:1,provider:'GOOGLE_DRIVE_V3',accessClass:'PRIVATE',presentationRef:out.targetRef,objectId:'SYN_OBJECT',targetKind:'FILE',versionId:'42',sha256:desc.sha256};
 const changed=s.plan(s.prepare('registrar',{...p,effectRef,motivo:'Outra justificativa'}));a.equal(changed.response.businessState,'REJECTED');a.equal(changed.writes.length,0);
 const {out:final}=s.apply('registrar',{...p,effectRef});
 const record=section==='PHOTO'?s.docs['canonicalPresentations/SYN_PR'].photo:final.writes.find(w=>w.path.startsWith('presentationMaterialSubmissions/')).data;
 a.equal(record.administrativeRegistration.registeredBy.id,'PES-000001');a.equal(record.administrativeRegistration.presenterRef.id,'PES-000002');a.equal(record.administrativeRegistration.institutionalEvidence.receivedAt.toMillis(),p.institutionalEvidence.receivedAt);
 a(!final.writes.some(w=>/^presentation(Realizations|MaterialReceipts|Obligations|QuotaOccupancies)/.test(w.path)));
 const repeat=s.plan(s.prepare('registrar',{...p,effectRef}));a.equal(repeat.response.businessState,'REJECTED');a.equal(repeat.writes.length,0);
});

test('effective extension restores direct registration only within the canonical function period',()=>{
 const s=fixture(true),f=s.docs['memberFunctions/VIG-001'],end=f.endDatePlanned;
 f.endDatePlanned=new Date(s.now-2*86400000).toISOString().slice(0,10);
 const p={observacoes:s.content.observacoes,motivo:'Registro em mandato prorrogado',confirmed:true};
 a.throws(()=>s.prepare('registrar',p),/AUTHORIZATION|AUTHZ/);
 const updates=JSON.parse(s.domain.run(x=>f4MemberOccupations_().extend(nodeRead,[x.f],x.f,x.end),{f,end}));
 for(const row of updates)s.docs['memberFunctions/'+row.id]=row;
 const before=clone(s.docs['canonicalPresentations/SYN_PR']);
 const {out}=s.apply('registrar',p);
 a.equal(out.response.result.state,'RECORDED_NOT_APPROVED');a.deepEqual(s.docs['canonicalPresentations/SYN_PR'],before);
});
test('upload rechecks the live function after reservation',()=>{
 const {s,p,desc}=fileFixture('PHOTO'),{r,out}=s.apply('registrar',{...p,fileDescriptor:desc});
 const effect=s.docs['jobOutbox/'+out.response.result.effectRef.id],uid=effect.actor.uid;
 const authRequest=JSON.parse(s.domain.run(x=>f4ActivitiesPipeline_().authorization(nodeRead,x.r,{uid:x.uid,expiresAt:x.now+60000},x.now,x.now+60000)[0],{r,uid,now:s.now}));
 const digest='a'.repeat(64);
 s.docs['commandProvenance/'+r.operationId]={operationId:r.operationId,ownerUid:uid,digest,ticketId:'SYN_ADMIN_TICKET',authRequests:[authRequest]};
 s.docs['memberAdmissionTickets/SYN_ADMIN_TICKET']={requestEnvelope:r};
 s.docs['commandRequests/'+r.operationId]={...r,ownerUid:uid,digest};
 const check=()=>s.domain.run(x=>{const ports=f4ActivitiesFunctionalPorts_({config:nodeConfig,identity:{uid:x.uid,expiresAt:x.now+60000},settings:{commands:['THEME_REGULARIZE_ADMIN'],policyIds:['CF_DEV_ACT_POLICY']},clock:()=>x.now});ports.authorizeUpload(nodeRead,nodeRead('jobOutbox/'+x.id));return true;},{uid,now:s.now,id:effect.id});
 a.equal(JSON.parse(check()),true);
 s.docs['memberAuthzPolicies/CF_DEV_ACT_POLICY'].active=false;
 a.throws(check,/AUTHORIZATION|AUTHZ|ACCESS_REVOKED/);
});

test('recognized presentation accepts a documentary addendum after the assignment and membership end',()=>{
 const s=fixture(true),u=s.unit();
 delete u.assignmentRef;u.membershipExit={recorded:true};
 const paths=['canonicalPresentations/SYN_PR','presentationRealizations/SYN_RZ','presentationQuotaOccupancies/SYN_QUOTA','presentationObligations/SYN_OB'];
 const before=paths.map(path=>clone(s.docs[path]));
 const {out}=s.apply('registrar',{observacoes:s.content.observacoes,motivo:'Arquivo histórico conferido',confirmed:true});
 a.equal(out.response.result.state,'RECORDED_NOT_APPROVED');paths.forEach((path,i)=>a.deepEqual(s.docs[path],before[i]));
});


test('partial documentary fields accumulate, remain readable, and reject material replacement',()=>{
 const s=fixture(true),pr=clone(s.docs['canonicalPresentations/SYN_PR']);
 const first=s.apply('registrar',{observacoes:'Resumo hist?rico',motivo:'Complemento parcial',confirmed:true});
 a.equal(first.out.response.result.state,'RECORDED_NOT_APPROVED');
 s.apply('registrar',{eixoTematicoPrincipal:'SYN_AXIS',motivo:'Eixo ausente',confirmed:true});
 const u=s.unit(),content=JSON.parse(s.domain.run(x=>f4ActivitiesDocumentaryTheme_(x.u,null,x.pr),{u,pr}));
 a.equal(content.titleOrObject,pr.title);a.equal(content.summary,'Resumo hist?rico');a.equal(content.principalAxisRef.id,'SYN_AXIS');
 a.deepEqual(s.docs['canonicalPresentations/SYN_PR'],pr);
 for(const p of [{observacoes:'Substitui??o silenciosa'},{tituloApresentacao:'T?tulo diferente'},{eixoTematicoPrincipal:'SYN_AXIS_2'}]){
  const result=s.plan(s.prepare('registrar',{...p,motivo:'Tentativa de substituir',confirmed:true}));
  a.equal(result.response.error.reasonCode,'EXISTING_THEME_REQUIRES_CHANGE_PROCEDURE');a.equal(result.writes.length,0);
 }
 const log=s.docs['presentationFlowCases/'+u.lastThemeRegularizationRef.id];a.equal(log.predecessorRef.id,first.out.response.result.recordRef.id);
 const projection=JSON.parse(s.domain.run(x=>f4ActivitiesViews_().member(x.pr,{kind:'presentation',id:'SYN_PR'},{generationRef:{id:'GEN'},viewEpoch:1,sourceMemberRevision:1,digest:'SYNTHETIC'},x.u.personRef.id,x.now,r=>nodeRead(f4ActivitiesModel_().path(r))),{u,pr,now:s.now}));
 a.equal(projection.documentaryTheme.summary,content.summary);a.equal(projection.documentaryRecordRef.id,u.lastThemeRegularizationRef.id);
 s.docs['memberAuthzPolicies/CF_DEV_ACT_POLICY'].targets.allowedTargetScopes.push('PRESENTATIONS_PLAN');
 const read=JSON.parse(s.domain.run(x=>{const o={config:nodeConfig,identity:{uid:'CF_ADMIN',expiresAt:x.now+60000},settings:{commands:['THEME_REGULARIZE_ADMIN'],policyIds:['CF_DEV_ACT_POLICY']},clock:()=>x.now};o.ports=f4ActivitiesFunctionalPorts_(o);return f4ActivitiesPresentationCompat_(o).read('/v2/apresentacoes/pendencias',{incluirSemPendencias:true});},{now:s.now}));
 a.equal(read.data.pendencias[0].resumoTema,content.summary);a.equal(read.data.pendencias[0].eixoTematicoPrincipal,'SYN_AXIS');
});

test('scheduled partial theme enters the canonical initial flow only when complete',()=>{
 const s=fixture();
 s.apply('registrar',{tituloApresentacao:'Tema parcial',motivo:'T?tulo recebido',confirmed:true});
 a(!s.unit().thematicRequestRef);
 const out=s.apply('registrar',{observacoes:'Resumo complementado',eixoTematicoPrincipal:'SYN_AXIS',motivo:'Conclus?o dos campos',confirmed:true}).out;
 const t=s.docs['presentationThematicRequests/'+out.response.result.requestRef.id];
 a.equal(t.candidate.titleOrObject,'Tema parcial');a.equal(t.candidate.summary,'Resumo complementado');a.equal(t.analysisClock.state,'UNRESOLVED');a(!t.extraordinarySubmission);
});

test('approved scheduled content accepts absent summary without a second approval',()=>{
 const s=fixture(),u=s.unit();u.thematicRequestRef=ref('thematic','APPROVED');
 const t={id:'APPROVED',revision:3,workflowState:'APPROVED',unitRef:ref('unit',u.id),personRef:u.personRef,cycleRef:u.cycleRef,principalAxisRef:ref('policy','SYN_AXIS'),candidate:{titleOrObject:'Tema aprovado',principalAxisRef:ref('policy','SYN_AXIS')},acceptedRevisionRef:{kind:'thematicRevision',id:'APPROVED',number:3}};
 s.docs['presentationThematicRequests/APPROVED']=clone(t);
 const {out}=s.apply('registrar',{observacoes:'Resumo ausente',motivo:'Complementa??o',confirmed:true});
 a.equal(out.response.result.state,'RECORDED_NOT_APPROVED');a.deepEqual(s.docs['presentationThematicRequests/APPROVED'],clone(t));
});
