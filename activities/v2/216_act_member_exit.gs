// One consequence producer for canonical member exit and the existing ACT job.
// Past/realized presentations and their materials are never released here.
function f4ActivitiesMemberExitEffects_(x,ev,unitRefs){
 const {M,N,load,revision,save,now}=x,D=f4ActivitiesDomain_(x),today=M.day(now),out={released:[],preserved:[],review:[]};
 N(ev?.effect==='DESLIGAMENTO'&&M.id(ev.idPessoa),'UNRESOLVED_DEPENDENCY','CANONICAL_EXIT_REQUIRED');M.civil(ev.effectiveDate);
 N(ev.effectiveDate<=today,'INVALID_STATE_TRANSITION','EXIT_NOT_EFFECTIVE');
 for(const r of unitRefs){
  const u=load(r,true);N(r.kind==='unit'&&u.personRef?.id===ev.idPessoa&&u.cycleRef?.id===x.cycleRef.id,'AUTHORIZATION_DENIED');
  if(u.membershipExit?.eventId===ev.id)continue;
  const cycle=load(u.cycleRef),pr=u.presentationRef&&load(u.presentationRef),assignment=u.assignmentRef&&load(u.assignmentRef),dynamic=u.dynamicOccupancyRef&&load(u.dynamicOccupancyRef),session=assignment?.sessionRef?load(assignment.sessionRef):dynamic?.sessionRef?load(dynamic.sessionRef):null;
  const immutable=cycle.status==='CLOSED'||cycle.civilEnd<ev.effectiveDate||!!pr?.realizationRef||dynamic?.state==='REALIZED_HISTORY'||session?.civilDate<ev.effectiveDate;
  if(immutable){out.preserved.push(r.id);continue;}
  if(session&&session.civilDate<=today){out.review.push(r.id);continue;}
  N(cycle.status==='OPEN','INVALID_STATE_TRANSITION','EXIT_CYCLE_NOT_OPEN');
  const basis={membershipExit:{eventId:ev.id,effectiveDate:ev.effectiveDate,personId:ev.idPessoa,recordedAt:M.stamp(now)}};
  if(assignment&&['ACTIVE','CONFIRMED'].includes(assignment.status))D.release(u.assignmentRef,basis);
  if(dynamic?.state==='ACTIVE')D.releaseDynamic(u.dynamicOccupancyRef,basis);
  if(u.quotaOccupancyRef)D.releaseQuota(u.quotaOccupancyRef,true);
  if(u.obligationRef){const ob=load(u.obligationRef,true);N(ob.personRef?.id===ev.idPessoa&&ob.cycleRef?.id===u.cycleRef.id,'AUTHORIZATION_DENIED');if(ob.result!=='SATISFIED')revision(u.obligationRef,{...ob,incidenceState:'NOT_INCIDENT',...basis});}
  save(r,{...load(r),...basis});out.released.push(r.id);x.bumpCapacity();
 }
 return out;
}

// Called only while assembling the canonical, explicitly authorized dismissal.
// The source set was captured server-side and is fenced/rechecked in this transaction.
function f4MemberAttendanceExitPlan_(get,check,ev,now,operationId,authorization){
 const M=f4ActivitiesModel_(),N=M.need,F=f4Firestore_(),C=f4Core_();
 f4MemberActivityChecks_().read(get,check,ev.idPessoa);
 const paths=new Set(check.sources.map(s=>s.path)),writes=new Map(),original=p=>get(p),current=p=>writes.get(p)||get(p),actor={kind:'MEMBERS_DECISION',uid:authorization.proof.uid,personRef:M.ref('person',authorization.proof.idPessoa)},authorityBasis=authorization.proof;
 function load(r,expected,optional){const path=M.path(r),d=current(path);if(optional&&!d)return null;N(paths.has(path),'UNRESOLVED_DEPENDENCY','EXIT_SOURCE_SCOPE');N(optional||d,'UNRESOLVED_DEPENDENCY');N(!d||!['HISTORICAL_PARTIAL','HISTORICAL_RECOGNIZED'].includes(d.recordMode),'UNRESOLVED_DEPENDENCY','EXIT_HISTORICAL_REVIEW_REQUIRED');return M.clone(d);}
 function save(r,d){const path=M.path(r),old=original(path);N(paths.has(path)&&old&&!M.immutable.has(r.kind),'AUTHORIZATION_DENIED');const value={...d,id:r.id,revision:old.revision+1,operationId,updatedAt:M.stamp(now)};M.validate(r,value);writes.set(path,value);return value;}
 function revision(r,d){const v=save(r,d),path=M.path(r),rp=path+'/revisions/'+v.revision;N(!get(rp)||writes.has(rp),'CONFLICT','REVISION_EXISTS');writes.set(rp,{number:v.revision,predecessorRevision:original(path).revision,snapshot:v,recordedAt:M.stamp(now),actor,actorAuthorityBasis:authorityBasis,operationId});return v;}
 const groups=new Map();for(const id of check.unitIds){const r=M.ref('unit',id),u=load(r);N(u.cycleRef?.kind==='cycle','UNRESOLVED_DEPENDENCY');if(!groups.has(u.cycleRef.id))groups.set(u.cycleRef.id,[]);groups.get(u.cycleRef.id).push(r);}
 const summary={released:[],preserved:[],review:[]};
 for(const [cycleId,refs]of groups){const cycleRef=M.ref('cycle',cycleId),x={M,N,load,save,revision,now,cycleRef,get:current,bumpCapacity(){const r=M.ref('capacity',cycleId),h=load(r);save(r,{...h,state:'INVALIDATED',capacityEpoch:h.capacityEpoch+1});}},r=f4ActivitiesMemberExitEffects_(x,ev,refs);for(const k of Object.keys(summary))summary[k].push(...r[k]);}
 for(const [path,d]of [...writes])if(path.startsWith('presentationUnitHeads/')){const pr=d.presentationRef&&current(M.path(d.presentationRef));if(pr){const pubPath='publicPresentations/'+pr.id,pub=get(pubPath);if(pub?.published)writes.set(pubPath,{...pub,published:false,revision:(pub.revision||0)+1,updatedAt:M.stamp(now)});}}
 if(writes.size){const path='portalReaderResourceFences/activities',old=get(path);writes.set(path,{schemaVersion:'portal-reader-resource-fence-v1',id:'activities',revision:(old?.revision||0)+1,updatedAt:M.stamp(now)});}
 N(writes.size<=120,'SCOPE_TOO_LARGE');return {writes:[...writes].map(([path,data])=>({path,data,kind:'set'})),summary};
}
