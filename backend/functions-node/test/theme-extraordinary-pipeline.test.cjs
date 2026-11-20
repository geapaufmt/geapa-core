const {test}=require('node:test'),a=require('node:assert/strict'),{scenario}=require('./theme-self-preparation.test.cjs');
const {encode,decode}=require('../src/firestore.cjs'),clone=v=>decode(encode(v));
const ref=(kind,id)=>({kind,id}),commands=['THEME_EXTRAORDINARY_AUTHORIZE','THEME_EXTRAORDINARY_REVOKE','THEME_REGULARIZE','THEME_REGULARIZE_ADMIN','THEME_VALIDATE','THEME_COMPLEMENT'];
function fixture(realized=false){
 const s=scenario(),self=clone(s.docs['memberAuthzPolicies/CF_DEV_ACT_POLICY']);let serial=10;
 const board={...clone(self),proofKind:'ROLE',actorSelector:{kind:'ROLE',roleKey:'DIRETOR_ADMINISTRATIVO',group:'DIRETORIA',actorScope:'ADMINISTRACAO'},capabilities:['PRESENTATION_BOARD_READ','PRESENTATION_THEME_REVIEW'],fieldMasks:{PRESENTATION_BOARD_READ:[],PRESENTATION_THEME_REVIEW:[]},targets:{...self.targets,ownPerson:false,allowedPersons:['PES-000001'],allowedTargetScopes:['PRESC_SYN_THEME_CYCLE'],allowedResources:[],resourceScope:{kind:'PRESENTATIONS'}}};
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
