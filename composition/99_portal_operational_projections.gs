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
  if(/^(canonical|presentation|activity|publicPresentations)/.test(collection)){fences.add('activities');related(d);related(old,0,true);}
 }
 if(!fences.size)return [];
 const fenceIds=[...fences].filter(id=>!id.startsWith('frequency_')).sort();if(fenceIds.length)tx.getAll(fenceIds.map(id=>'portalReaderResourceFences/'+id));
 const effects=fenceIds.map(id=>{const path='portalReaderResourceFences/'+id,old=tx.get(path).data();return {kind:'set',path,data:{schemaVersion:'portal-reader-resource-fence-v1',id,revision:(old?.revision||0)+1,updatedAt:F.Timestamp.fromMillis(now),operationId}}});
 effects.push({kind:'set',path:'commandProjectionInvalidations/'+operationId,data:{schemaVersion:'portal-command-invalidation-v1',operationId,fences:[...fences].sort(),createdAt:F.Timestamp.fromMillis(now)}});
 return effects;
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
