// Derived reader freshness participates in the same Core transaction as facts.
// No client writer and no scheduler are needed for these bounded projections.
function f4PortalProjectionFenceId_(view,personId){
 if(view==='self'||view==='context')return null;
 // Frequency already has a canonical registry epoch. Reuse it to keep Rules
 // within their access-call budget instead of consulting a second fence.
 if(view==='frequency')return null;
 f4Core_().need(view==='activities'||/^activity_[A-Za-z0-9_-]{1,128}$/.test(view),'COMMAND_PAYLOAD_INVALID');return view;
}

function f4PortalProjectionInvalidation_(tx,writes,operationId,now){
 const gate=tx.get('portalReaderControl/current').data();if(!gate?.operationalWrites)return [];
 const M=f4ActivitiesModel_(),F=f4Firestore_(),changed=new Map(writes.map(w=>[w.path,w.data])),fences=new Set(),seen=new Set();
 const read=(path,prior)=>prior?tx.get(path).data():changed.has(path)?changed.get(path):tx.get(path).data();
 function related(d,depth=0,prior=false){
  if(!d||depth>4)return;
  if(d.activityRef?.id)fences.add('activity_'+d.activityRef.id);
  for(const key of ['activityRef','sessionRef','presentationSessionRef','presentationRef','unitRef','assignmentRef','calendarAssignmentRef','submissionRef','requirementRef','reviewRef','opportunityRef','attendanceRef','justificationRef']){
   const ref=d[key];if(!ref?.kind||!ref.id)continue;
   const path=M.path(ref),visitKey=(prior?'before:':'after:')+path;if(seen.has(visitKey))continue;seen.add(visitKey);M.need(seen.size<=180,'SCOPE_TOO_LARGE','READER_FENCE_SCOPE');
   if(ref.kind==='activity')fences.add('activity_'+ref.id);else related(read(path,prior),depth+1,prior);
  }
 }
 for(const w of writes){
  if(w.path.split('/').length!==2)continue;const [collection,id]=w.path.split('/'),old=tx.get(w.path).data(),d=w.data||old;
  if(collection==='canonicalActivities')fences.add('activity_'+id);
  if(collection==='activityAttendance'||collection==='activityJustifications'){
   for(const row of [old,d]){const person=row?.personRef?.id||row?.ownerPersonRef;if(person)fences.add('frequency_'+person);}
  }
  if(/^(canonical|presentation|activity|publicPresentations|calendarAssignments)/.test(collection)){fences.add('activities');related(d);related(old,0,true);}
 }
 if(!fences.size)return [];
 const fenceIds=[...fences].filter(id=>!id.startsWith('frequency_')).sort();if(fenceIds.length)tx.getAll(fenceIds.map(id=>'portalReaderResourceFences/'+id));
 const effects=fenceIds.map(id=>{const path='portalReaderResourceFences/'+id,old=tx.get(path).data(),revision=(old?.revision||0)+1;return {kind:'set',path,data:{schemaVersion:'portal-reader-resource-fence-v1',id,revision,updatedAt:F.Timestamp.fromMillis(now),operationId,...(id==='activities'?{listChanges:f4ActivityListChanges_(tx,writes,old,revision,fenceIds.filter(k=>k.startsWith('activity_')).map(k=>k.slice(9)))}:{})}}});
 effects.push({kind:'set',path:'commandProjectionInvalidations/'+operationId,data:{schemaVersion:'portal-command-invalidation-v1',operationId,fences:[...fences].sort(),createdAt:F.Timestamp.fromMillis(now)}});
 return effects;
}

// Bounded invalidation hints, committed atomically with facts and fence. This
// is not a second data source. Unknown/shared dependencies invalidate the whole
// list; old executors without these hints naturally force a full rebuild.
function f4ActivityListChanges_(tx,writes,old,revision,relatedActivityIds=[]){
 const ids=new Set();let complete=true,capacityOnly=true,capacityChanges=0;
 for(const w of writes){const [col,id,...tail]=w.path.split('/');if(tail.length||! /^(canonical|presentation|activity|publicPresentations|calendarAssignments)/.test(col))continue;
  if(col==='activityAttendanceRegistry'||col==='activityHistory')continue;
  // Capacity proofs/heads do not feed activity-list rows. Keep the fence and
  // fresh authorization/CAS, but allow an explicit empty row delta. Registry,
  // configuration, unknown changes and historic partials remain full barriers.
  if(['presentationCapacityHeads','presentationCapacityAssessments'].includes(col)&&id!=='registry'){
   const versions=[tx.get(w.path).data(),w.data].filter(Boolean);
   if(versions.length&&versions.every(d=>d.cycleRef?.kind==='cycle'&&d.cycleRef.id&&d.recordMode!=='HISTORICAL_PARTIAL')){capacityChanges++;continue;}
  }
  capacityOnly=false;
  // File/material changes use the same bounded before/after dependency walk as
  // detail invalidation. Shared configuration and unlinked/partial history keep
  // the conservative full rebuild; no second index or authority is introduced.
  if(['canonicalPresentations','presentationMaterialRequirements','presentationMaterialSubmissions','presentationMaterialReviews','presentationMaterialReceipts','presentationCorrectionOpportunities','presentationArtifacts'].includes(col)){
   const before=tx.get(w.path).data(),after=w.data;
   if(!relatedActivityIds.length||[before,after].some(d=>d?.recordMode==='HISTORICAL_PARTIAL'))complete=false;
   else relatedActivityIds.forEach(id=>ids.add(id));continue;
  }
  if(!['canonicalActivities','activityAgendaViews','activityAttendanceCollections','activityRealizations','activityAttendance','activityJustifications'].includes(col)){complete=false;continue;}
  const before=tx.get(w.path).data(),after=w.data;
  if(col==='canonicalActivities'||col==='activityAgendaViews'||col==='activityAttendanceCollections')ids.add(id);
  else for(const d of [before,after].filter(Boolean)){if(d.activityRef?.kind==='activity'&&d.activityRef.id)ids.add(d.activityRef.id);else complete=false;}
  if([before,after].some(d=>d?.recordMode==='HISTORICAL_PARTIAL')&&!(col==='canonicalActivities'&&f4ActivityReviewListDelta_(before,after)))complete=false;
 }
 const valid=f4ActivitiesModel_().id;
 const noRows=complete&&capacityOnly&&capacityChanges>0&&ids.size===0;
 const entry={revision,complete:complete&&(ids.size>0||noRows)&&ids.size<=32&&[...ids].every(valid),ids:complete&&ids.size<=32?[...ids].sort():[],...(noRows?{noRows:true}:{})};
 const prior=Array.isArray(old?.listChanges)&&old.listChanges.every(e=>Number.isSafeInteger(e.revision)&&e.revision<=revision-1&&typeof e.complete==='boolean'&&Array.isArray(e.ids)&&e.ids.length<=32&&e.ids.every(valid))?old.listChanges:[];
 return [...prior.slice(-15),entry];
}

// A historical administrative review only changes the row's display fields and
// its independent review lock. Prove that boundary from both canonical versions;
// unknown historical changes still invalidate the entire list.
function f4ActivityReviewListDelta_(before,after){
 if(before?.recordMode!=='HISTORICAL_PARTIAL'||after?.recordMode!=='HISTORICAL_PARTIAL'||!before.administrativeReview||!after.administrativeReview||after.revision!==before.revision+1)return false;
 const C=f4Core_(),F=f4Firestore_(),R=f4ActivityReview_();
 try{R.validate(before.administrativeReview);R.validate(after.administrativeReview);}catch(e){return false;}
 const b=before.administrativeReview,n=after.administrativeReview;
 if(b.state!=='PENDING'||n.revision!==b.revision+1||!C.op(after.operationId))return false;
 const eq=(a,b)=>C.hash(F.encode(a??null))===C.hash(F.encode(b??null));
 for(const k of ['schemaVersion','cycleRef','initializationRef','initializedAt'])if(!eq(b[k],n[k]))return false;
 if(n.state==='CONFIRMED'&&n.confirmationOperationId!==after.operationId)return false;
 const allowed=['title','location','details','administrativeReview','revision','updatedAt','operationId'];
 for(const k of new Set([...Object.keys(before),...Object.keys(after)]))if(!allowed.includes(k)&&!eq(before[k],after[k]))return false;
 for(const k of new Set([...Object.keys(before.details||{}),...Object.keys(after.details||{})]))if(!R.fields.includes(k)&&!eq(before.details?.[k],after.details?.[k]))return false;
 return true;
}

function f4ActivityListPatchPlan_(prior,basis,fence,signature,now){
 const C=f4Core_(),M=f4ActivitiesModel_(),memo=prior?.data?.listMaintenance,rows=prior?.data?.bundle?.data?.calendario,b=prior?.basis;
 if(prior?.schemaVersion!=='portal-reader-view-v1'||prior.viewId!=='activities'||memo?.version!==1||memo.signature!==signature||!Array.isArray(rows)||rows.length>=400||!Array.isArray(memo.rows)||memo.rows.length!==rows.length||!b||now>=C.ms(b.expiresAt))return null;
 if(Object.keys(basis).filter(k=>!['issuedAt','expiresAt'].includes(k)).some(k=>JSON.stringify(b[k])!==JSON.stringify(basis[k])))return null;
 if(b.dataFence?.id!=='activities'||!Number.isSafeInteger(b.dataFence.revision)||!Number.isSafeInteger(fence?.revision)||b.dataFence.revision>fence.revision)return null;
 const rowIds=rows.map(r=>r.idAtividade),memoIds=memo.rows.map(r=>r.id);
 if(new Set(rowIds).size!==rowIds.length||rowIds.some(id=>!M.id(id))||JSON.stringify(rowIds)!==JSON.stringify(memoIds)||memo.rows.some(r=>!Number.isSafeInteger(r.nextAt)||r.nextAt<0))return null;
 const admin=prior.data.admin?.data?.atividades;if(admin&&JSON.stringify(admin.map(r=>r.idAtividade))!==JSON.stringify(rowIds))return null;
 const ids=new Set(memo.rows.filter(r=>r.nextAt<=now).map(r=>r.id)),entries=fence.listChanges;
 for(let revision=b.dataFence.revision+1;revision<=fence.revision;revision++){
  if(fence.revision-b.dataFence.revision>16||!Array.isArray(entries))return null;
  const matches=entries.filter(e=>e.revision===revision);if(matches.length!==1)return null;const e=matches[0];
  if(e.complete!==true||!Array.isArray(e.ids)||(!e.ids.length&&e.noRows!==true)||e.ids.length>32||e.ids.some(id=>!M.id(id)))return null;
  e.ids.forEach(id=>ids.add(id));
 }
 return ids.size<=64?{ids:[...ids].sort(),expiresAt:Math.min(C.ms(b.expiresAt),C.ms(basis.expiresAt))}:null;
}

function f4PortalProjectionRebuild_(options,viewIds){
 const F=f4Firestore_(),C=f4Core_(),N=f4ActivitiesModel_().need,db=F.connect(options.config),proofs=[],authorityPaths=[];
 N(Array.isArray(viewIds)&&viewIds.length>0&&viewIds.length<=4,'COMMAND_PAYLOAD_INVALID');
 const docs=f4DevReadTiming_().stage('projectionRebuild',()=>f4PortalReadModels_({...options,viewIds,proofs,authorityPaths}));
 db.transaction(tx=>{
  // Fetch the known source set inside this fresh transaction. Rechecks below
  // still derive authority from its current values, including changed heads.
  const paths=[...new Set(authorityPaths)];for(let i=0;i<paths.length;i+=180)tx.getAll(paths.slice(i,i+180));
  const get=p=>tx.get(p).data(),now=options.clock(),gate=get('portalReaderControl/current');
  N(gate?.enabled&&gate.operationalWrites&&C.ms(gate.startsAt)<=now&&now<C.ms(gate.expiresAt),'ACCESS_REVOKED');
  for(const {q,proof}of proofs)f4SharedAuthz_().recheck(get,q,proof,now);
  for(const [path,view]of Object.entries(docs)){
   const b=view.basis;N(b.generation===gate.generation&&now<C.ms(b.expiresAt),'STALE_REVISION');
   if(b.dataFence){const current=get('portalReaderResourceFences/'+b.dataFence.id);N((current?.revision||0)===b.dataFence.revision,'STALE_REVISION');}
   if(view.viewId==='frequency')N(view.data.epoch===(get('activityAttendanceRegistry/current')?.epoch||0),'STALE_REVISION');
   get(path);
  }
  // All reads precede the write phase; facts and fence revisions form a CAS.
  for(const [path,view]of Object.entries(docs))tx.set(path,view);
 });
 return docs;
}

function f4PortalProjectionRoute_(route,params){
 if(['/atividades/listar','/atividades/bundle','/atividades/detalhes-preload','/admin/atividades'].includes(route))return 'activities';
 if(['/atividades/detalhe','/admin/atividades/detalhe'].includes(route)){f4Core_().need(/^[A-Za-z0-9_-]{1,128}$/.test(params?.idAtividade),'COMMAND_PAYLOAD_INVALID');return 'activity_'+params.idAtividade;}
 if(route==='/v2/minha-frequencia')return 'frequency';return null;
}
function f4PortalProjectedRead_(options,request){
 const id=f4PortalProjectionRoute_(request.route,request.params);if(!id)return null;
 const docs=f4PortalProjectionRebuild_(options,[id]),v=docs['portalReaderViews/'+options.identity.uid+'/items/'+id].data;
 if(id==='frequency')return v.result;
 if(id.startsWith('activity_'))return request.route.startsWith('/admin/')?v.admin:v.normal;
 if(request.route==='/admin/atividades')return v.admin;
 if(request.route==='/atividades/listar')return {ok:true,data:v.bundle.data.calendario};return v.bundle;
}
function f4PortalAfterCommand_(options,result,key){
 // Canonical facts, receipt and invalidation have already committed together.
 // Never delay the command response for derived views. The next authorized
 // reader rebuilds only its requested view through f4PortalProjectedRead_.
 // Rules reject the old fence immediately, even if reconstruction fails.
 return result;
}

// Recovery freshness uses a newly authorized Core basis, never the browser's
// capability claims or an old projection as authority. No writes in this probe.
function f4PortalProjectionRecoveryInspect_(options,id){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),db=F.connect(options.config);
 N(id==='frequency'||id==='activities'||/^activity_[A-Za-z0-9_-]{1,128}$/.test(id),'COMMAND_PAYLOAD_INVALID');
 const basisId=id==='frequency'?'frequency':'context',current=f4PortalReadModels_({...options,basisOnly:true,viewIds:[basisId]})['portalReaderViews/'+options.identity.uid+'/items/'+basisId].basis;
 const gate=db.get('portalReaderControl/current'),now=options.clock();
 N(gate?.enabled&&gate.operationalWrites&&C.ms(gate.startsAt)<=now&&now<C.ms(gate.expiresAt),'ACCESS_REVOKED');
 if(id.startsWith('activity_'))N(db.get('canonicalActivities/'+id.slice(9)),'PILOT_RESOURCE_DENIED');
 const view=db.get('portalReaderViews/'+options.identity.uid+'/items/'+id);
 if(!view)return {state:'MISSING'};
 const b=view.basis,f=db.get('portalReaderResourceFences/'+id);
 const match=b&&Object.keys(current).filter(k=>!['issuedAt','expiresAt'].includes(k)).every(k=>JSON.stringify(b[k])===JSON.stringify(current[k]));
 return {state:match&&view.schemaVersion==='portal-reader-view-v1'&&view.viewId===id&&now<C.ms(b.expiresAt)&&(id==='frequency'?view.data?.epoch===(db.get('activityAttendanceRegistry/current')?.epoch||0):b.dataFence?.id===id&&b.dataFence.revision===(f?.revision||0))?'CURRENT':'STALE'};
}

// A generation change must recover the direct context document as well as its
// response. Otherwise every subsequent session check falls back to Core.
function f4PortalContextRead_(options){
 const context=f4ActivitiesFunctionalPorts_(options).context();
 // Suspended material access is contextual and remains a privileged reader.
 if(context.materialOnly)return context;
 const docs=f4PortalProjectionRebuild_(options,['context']);
 return docs['portalReaderViews/'+options.identity.uid+'/items/context'].data;
}
