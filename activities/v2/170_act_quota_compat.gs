function f4ActivitiesQuotaCompat_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),db=F.connect(options.config),ports=options.ports,clock=options.clock||Date.now;
 function context(id){N(M.id(id),'VALIDATION_ERROR');const ref=M.ref('cycle',id),proof=ports.readAuthority([M.path(ref)],false,'PRESENTATION_BOARD_READ','PRESC_'+id),cycle=db.get(M.path(ref));N(cycle,'NOT_FOUND');const cfg=cycle.currentConfigRef&&db.get(M.path(cycle.currentConfigRef));N(cfg?.approved&&cfg.parameters.quota&&cfg.parameters.axisRefs?.length,'CONFIG_CAPABILITY_DISABLED','QUOTA_OFF');return {ref,proof,cycle,cfg};}
 function recheck(z){f4SharedAuthz_().recheck(p=>db.get(p),z.proof.q,z.proof.proof,clock());N(db.get(M.path(z.ref))?.revision===z.cycle.revision,'STALE_REVISION');}
 function read(input){
  C.exact(input,['cycleId',...('cursor'in input?['cursor']:[])]);const z=context(input.cycleId),head=db.get('presentationCapacityHeads/'+z.cycle.id),config=z.cycle.thematicQuotaConfigRef&&db.get(M.path(z.cycle.thematicQuotaConfigRef)),axes=z.cfg.parameters.axisRefs;
  const limits=axes.map(ref=>{const id=M.hash({cycle:z.cycle.id,axis:ref.id}).slice(0,32),b=db.get('presentationQuotaBuckets/'+id),axis=db.get(M.path(ref));return {axisId:ref.id,label:axis?.label||axis?.name||ref.id,bucketId:id,revision:b?.revision||0,limit:b?.limit??null,held:b?.heldCount||0,confirmed:b?.confirmedCount||0,realized:b?.realizedCount||0,queueEpoch:b?.pendingQueueEpoch||0};});
  if(input.cursor){C.exact(input.cursor,['id','epoch']);N(M.id(input.cursor.id)&&input.cursor.epoch===head.capacityEpoch,'STALE_REVISION','QUEUE_LIST_CHANGED');}
  const rows=db.query({collection:'presentationThematicRequests',filters:[{field:'cycleRef',op:'EQUAL',value:z.ref},{field:'workflowState',op:'EQUAL',value:'WAITING'}],order:[{field:'__name__',direction:'ASCENDING'}],limit:20,...(input.cursor?{cursor:[new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+'presentationThematicRequests/'+input.cursor.id)]}:{})});
  const waiting=rows.map(({data:t})=>{const u=db.get(M.path(t.unitRef)),person=u&&db.get('memberIdentities/'+u.personRef.id);N(u,'UNRESOLVED_DEPENDENCY');return {id:t.id,revision:t.revision,title:t.candidate?.titleOrObject||'',person:person?.displayName||'',axisId:(t.candidate?.principalAxisRef||t.principalAxisRef).id,priorityAt:M.ms(t.priorityBasis.completeAt),materialityPending:t.candidate?.materiality==='UNDECIDED'};});
  const fresh=db.get('presentationCapacityHeads/'+z.cycle.id);N(fresh?.capacityEpoch===head.capacityEpoch,'STALE_REVISION','QUEUE_LIST_CHANGED');for(const item of limits)N((db.get('presentationQuotaBuckets/'+item.bucketId)?.revision||0)===item.revision,'STALE_REVISION','QUOTA_READ_CHANGED');recheck(z);
  return {ok:true,data:{cycleId:z.cycle.id,cycleRevision:z.cycle.revision,label:z.cycle.cycleLabel,limits,waitAdmissionMode:config?.waitAdmissionMode||'DISABLED',configured:!!config,waiting,nextCursor:rows.length===20?{id:rows[19].data.id,epoch:head.capacityEpoch}:null}};
 }
 function prepare(input){
  C.exact(input,['operationId','route','payload']);N(input.route==='/admin/apresentacoes/cotas/salvar','VALIDATION_ERROR');const p=input.payload;N(['CONFIGURE','ADMIT','RESOLVE','MATERIALITY'].includes(p.action),'VALIDATION_ERROR');C.exact(p,['action','cycleId','cycleRevision','reason',...(p.action==='CONFIGURE'?['limits','waitAdmissionMode']:['requestId','requestRevision']),...(p.action==='MATERIALITY'?['materiality']:[])]);
  const z=context(p.cycleId);N(z.cycle.revision===p.cycleRevision,'STALE_REVISION','DISPLAYED_CYCLE_CHANGED');let command,basis=[];
  if(p.action==='CONFIGURE'){
   N(Array.isArray(p.limits)&&p.limits.length<=12,'VALIDATION_ERROR');const axisLimits=p.limits.map(i=>{C.exact(i,['axisId','revision','limit']);N(M.id(i.axisId)&&Number.isSafeInteger(i.revision)&&i.revision>=0,'VALIDATION_ERROR');const ref=M.ref('bucket',M.hash({cycle:z.cycle.id,axis:i.axisId}).slice(0,32));N((db.get(M.path(ref))?.revision||0)===i.revision,'STALE_REVISION','DISPLAYED_QUOTA_CHANGED');basis.push({ref,revision:i.revision});return {axisRef:M.ref('policy',i.axisId),limit:i.limit};});
   command={commandType:'ACT.QUOTA_CONFIGURE',payload:{cycleRef:z.ref,axisLimits,waitAdmissionMode:p.waitAdmissionMode,reason:p.reason}};
  }else{
   const ref=M.ref('thematic',p.requestId),t=db.get(M.path(ref));N(t?.cycleRef.id===z.cycle.id&&t.revision===p.requestRevision,'STALE_REVISION','DISPLAYED_REQUEST_CHANGED');basis.push({ref,revision:t.revision});
   if(p.action==='MATERIALITY'){N(t.workflowState==='WAITING'&&t.candidate?.materiality==='UNDECIDED'&&typeof p.reason==='string'&&p.reason.trim(),'INVALID_STATE_TRANSITION');command={commandType:'ACT.THEME_MATERIALITY_DECIDE',payload:{requestRef:ref,submittedRevisionRef:t.submittedRevisionRef,materiality:p.materiality,reason:p.reason,decisionBasis:{kind:'BOARD_WAITING_CLASSIFICATION',reason:p.reason}}};}
   else command={commandType:p.action==='ADMIT'?'ACT.THEME_WAIT_ADMIT':'ACT.QUOTA_PRIORITY_RESOLVE',payload:{requestRef:ref,reason:p.reason,...(p.action==='RESOLVE'?{bucketRef:M.ref('bucket',M.hash({cycle:z.cycle.id,axis:(t.candidate?.principalAxisRef||t.principalAxisRef).id}).slice(0,32))}:{})}};
  }
  recheck(z);const request=ports.prepare({operationId:input.operationId,...command});for(const b of basis)N(request.target.expected.some(e=>e.kind===b.ref.kind&&e.id===b.ref.id&&(b.revision?e.revision===b.revision:e.exists===false)),'STALE_REVISION');return request;
 }
 return {read,prepare};
}
