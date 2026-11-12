const {test}=require('node:test'),a=require('node:assert/strict'),{scenario}=require('./theme-self-preparation.test.cjs'),{Timestamp}=require('firebase-admin/firestore');
const ref=(kind,id)=>({kind,id});
function review(decision='APROVAR'){
 const s=scenario(),cfg=s.docs['presentationCycles/SYN_THEME_CYCLE/configRevisions/CFG'];
 cfg.parameters.clockPolicyRef=ref('policy','SYN_CLOCK');
 s.docs['presentationPolicyReferences/SYN_CLOCK']={id:'SYN_CLOCK',revision:1,approved:true,active:true,unit:'HOURS',amount:72};
 const sub=s.plan(s.prepare(),true);a.equal(sub.response.businessState,'COMMIT_APPLIED');
 const u=s.docs['presentationUnitHeads/SYN_THEME_UNIT'],t=s.docs['presentationThematicRequests/'+u.thematicRequestRef.id];
 const p=s.docs['memberAuthzPolicies/CF_DEV_ACT_POLICY'];Object.assign(p,{proofKind:'ROLE',actorSelector:{kind:'ROLE',roleKey:'DIRETOR_ADMINISTRATIVO',group:'DIRETORIA',actorScope:'ADMINISTRACAO'},capabilities:['PRESENTATION_BOARD_READ','PRESENTATION_THEME_REVIEW'],fieldMasks:{PRESENTATION_BOARD_READ:[],PRESENTATION_THEME_REVIEW:[]},targets:{...p.targets,ownPerson:false,allowedPersons:['PES-000001'],allowedTargetScopes:['PRESC_SYN_THEME_CYCLE'],allowedResources:[],resourceScope:{kind:'PRESENTATIONS'}}});
 const hash=v=>JSON.parse(s.domain.run(x=>f4ActivitiesModel_().hash(x),v));
 s.docs['presentationPolicyReferences/SYN_UNUSED_MATERIAL']={id:'SYN_UNUSED_MATERIAL',revision:1,state:'SEALED',complete:true,items:[],digest:hash([]),finalEndExclusive:Timestamp.fromMillis(s.now+86400000),uploadPolicyRef:ref('policy','SYN_UPLOAD')};
 s.docs['presentationPolicyReferences/SYN_UPLOAD']={id:'SYN_UPLOAD',revision:1,approved:true,active:true};
 const input={operationId:'b'.repeat(32),route:'/v2/apresentacoes/titulo-eixo/revisar',basis:{unitRevision:u.revision,thematicRevision:t.revision},payload:{idApresentacao:'APU_SYN_THEME_UNIT',decisao:decision,observacaoPublica:'Revisão sintética',observacaoInterna:'Motivo sintético',...(decision==='EDITAR_E_APROVAR'?{tituloApresentacao:'Título revisado sintético',eixoTematicoPrincipal:'SYN_AXIS_2',eixoTematicoSecundario:'',materialidade:'SUBSTANTIAL'}:{})}};
 const prepare=()=>JSON.parse(s.domain.run(x=>{const o={config:nodeConfig,identity:{uid:'CF_ADMIN',expiresAt:x.now+60000},settings:{commands:['THEME_VALIDATE'],policyIds:['CF_DEV_ACT_POLICY']},clock:()=>x.now};o.ports=f4ActivitiesFunctionalPorts_(o);return f4ActivitiesPresentationCompat_(o).prepare(x.input);},{input,now:s.now}));
 return {...s,input,prepare,unit:u,theme:t};
}
for(const [decision,state] of [['APROVAR','APPROVED'],['REPROVAR','REJECTED'],['SOLICITAR_AJUSTE','PAUSED'],['EDITAR_E_APROVAR','APPROVED']])test('bounded administrative review '+decision+' with 160 historical references',()=>{
 const s=review(decision),r=s.prepare();a.equal(r.target.authorization.kind,'ROLE');a(r.target.expected.length<32);a(!r.target.expected.some(v=>v.id.startsWith('SYN_UNRELATED')||/^SYN_UNUSED_\d/.test(v.id)||v.id==='SYN_OTHER_UNIT'));
 const out=s.plan(r,true);a.equal(out.response.businessState,'COMMIT_APPLIED',JSON.stringify(out.response));const t=s.docs['presentationThematicRequests/'+s.unit.thematicRequestRef.id];a.equal(t.workflowState,state);a(out.writes.some(w=>w.path.includes('/validations/')));a(out.writes.some(w=>w.path.includes('/revisions/')));
 if(state==='APPROVED')a(out.writes.some(w=>w.path.startsWith('presentationMaterialRequirements/')));
 if(decision==='EDITAR_E_APROVAR')a.equal(t.candidate.titleOrObject,'Título revisado sintético');
 a.throws(()=>s.plan(r),/STALE_REVISION/);
});
test('review displayed revision and authorization remain mandatory',()=>{const s=review();s.input.basis.thematicRevision--;a.throws(s.prepare,/STALE_REVISION/);const revoked=review();revoked.docs['memberAuthzPolicies/CF_DEV_ACT_POLICY'].active=false;a.throws(revoked.prepare,/AUTHORIZATION_DENIED/);});
module.exports={review};

test('quota edit fences both axes and preserves counters and stale preconditions',()=>{
 const s=review('EDITAR_E_APROVAR'),hash=v=>JSON.parse(s.domain.run(x=>f4ActivitiesModel_().hash(x),v));
 s.docs['presentationCycles/SYN_THEME_CYCLE/configRevisions/CFG'].parameters.quota=true;
 s.docs['presentationCapacityHeads/SYN_THEME_CYCLE'].capacityEpoch=1;s.unit.quotaOccupancyRef=ref('quota','SYN_OCCUPANCY');s.docs['presentationUnitHeads/SYN_THEME_UNIT']=s.unit;
 s.docs['presentationQuotaOccupancies/SYN_OCCUPANCY']={id:'SYN_OCCUPANCY',revision:1,unitRef:ref('unit',s.unit.id),cycleRef:s.unit.cycleRef,axisRef:ref('policy','SYN_AXIS'),phase:'H'};
 for(const axis of ['SYN_AXIS','SYN_AXIS_2']){
  const id=hash({cycle:'SYN_THEME_CYCLE',axis}).slice(0,32),queue='QUEUE_'+axis;
  s.docs['presentationQuotaBuckets/'+id]={id,revision:1,cycleRef:s.unit.cycleRef,axisRef:ref('policy',axis),reconciliationState:'CLEAN',limit:5,heldCount:axis==='SYN_AXIS'?1:0,confirmedCount:0,realizedCount:0,epoch:1,pendingQueueEpoch:1,queueManifestRef:ref('policy',queue)};
  s.docs['presentationPolicyReferences/'+queue]={id:queue,revision:1,schemaVersion:'act-quota-queue-v1',state:'SEALED',complete:true,queueEpoch:1,axisRef:ref('policy',axis),items:[],digest:hash([]),minimumPriorityAt:0,unresolvedPriority:false};
 }
 const r=s.prepare();a.equal(r.target.expected.filter(v=>v.kind==='bucket').length,2);const out=s.plan(r,true);a.equal(out.response.businessState,'COMMIT_APPLIED',JSON.stringify(out.response));a.equal(s.docs['presentationQuotaOccupancies/SYN_OCCUPANCY'].axisRef.id,'SYN_AXIS_2');a.equal(s.docs['presentationQuotaOccupancies/SYN_OCCUPANCY'].phase,'C');
});
