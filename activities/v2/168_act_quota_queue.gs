// Core produces a bounded, complete queue proof. A proof never grants a quota.
// Every thematic change fences the relevant axes with pendingQueueEpoch.
function f4ActivitiesQuotaQueue_(options){
 const M=f4ActivitiesModel_(),N=M.need,F=f4Firestore_(),db=options.db,clock=options.clock||Date.now;
 function refresh(bucketRef,authorize){
  const bucket=db.get(M.path(bucketRef));N(bucket?.cycleRef&&bucket.axisRef,'CONFIG_CAPABILITY_DISABLED','QUOTA_NOT_CONFIGURED');
  const current=bucket.queueManifestRef&&db.get(M.path(bucket.queueManifestRef));
  if(current?.complete&&current.state==='SEALED'&&current.queueEpoch===bucket.pendingQueueEpoch&&current.axisRef?.id===bucket.axisRef.id&&(!current.validUntil||M.ms(current.validUntil)>clock()))return false;
  authorize(p=>db.get(p));let cursor,themes=[];
  for(let page=0;page<10;page++){
   const rows=db.query({collection:'presentationThematicRequests',filters:[{field:'cycleRef',op:'EQUAL',value:bucket.cycleRef}],order:[{field:'__name__',direction:'ASCENDING'}],limit:20,...(cursor?{cursor:[cursor]}:{})});
   themes.push(...rows.map(v=>v.data));N(themes.length<=180,'SCOPE_TOO_LARGE','THEMATIC_QUEUE_BOUND');if(rows.length<20)break;
   cursor=new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+f4DevBoundary_().root+'/presentationThematicRequests/'+rows[rows.length-1].data.id);
  }
  const candidates=[],uncertain=[];
  for(const t of themes){
   if(!['ANALYSIS','PAUSED','WAITING'].includes(t.workflowState))continue;
   const axis=t.candidate?.principalAxisRef||t.principalAxisRef;if(axis?.id!==bucket.axisRef.id)continue;
   const u=db.get(M.path(t.unitRef)),occupancy=u?.quotaOccupancyRef&&db.get(M.path(u.quotaOccupancyRef));
   N(u?.cycleRef.id===bucket.cycleRef.id,'UNRESOLVED_DEPENDENCY','THEMATIC_UNIT_MISSING');
   if(occupancy&&occupancy.phase!=='RELEASED'&&occupancy.axisRef.id===axis.id)continue;
   if(t.candidate?.materiality==='UNDECIDED'){uncertain.push(t.id);continue;}
   const priority=M.ms(t.priorityBasis?.completeAt);N(Number.isFinite(priority),'UNRESOLVED_DEPENDENCY','PRIORITY_MISSING');
   candidates.push({ref:M.ref('thematic',t.id),revision:t.revision,priorityCompleteAt:priority});
  }
  const minimum=candidates.length?Math.min(...candidates.map(t=>t.priorityCompleteAt)):null,items=candidates.filter(t=>t.priorityCompleteAt===minimum).sort((a,b)=>a.ref.id.localeCompare(b.ref.id));
  N(items.length<=32,'SCOPE_TOO_LARGE','PRIORITY_TIE_BOUND');
  return db.transaction(tx=>{
   const get=p=>tx.get(p).data();authorize(get);const b=get(M.path(bucketRef));N(b?.revision===bucket.revision&&b.pendingQueueEpoch===bucket.pendingQueueEpoch,'STALE_REVISION','QUEUE_CHANGED');
   const id=M.hash({bucket:bucketRef.id,queueEpoch:b.pendingQueueEpoch,items,uncertain}).slice(0,32),ref=M.ref('assessment',id),old=get(M.path(ref));
   if(old)N(old.digest===M.hash(items)&&old.queueEpoch===b.pendingQueueEpoch,'CONFLICT');
   else tx.set(M.path(ref),{id,revision:1,schemaVersion:'act-quota-queue-v1',cycleRef:b.cycleRef,axisRef:b.axisRef,state:'SEALED',complete:true,sourceCoverageComplete:true,queueEpoch:b.pendingQueueEpoch,items,digest:M.hash(items),minimumPriorityAt:minimum,unresolvedPriority:uncertain.length>0,unresolvedCount:uncertain.length,waitingCount:candidates.length,recordedAt:M.stamp(clock())});
   tx.set(M.path(bucketRef),{...b,revision:b.revision+1,queueManifestRef:ref});return true;
  });
 }
 return {refresh};
}
