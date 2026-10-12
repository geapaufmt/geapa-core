// Read models for historical and operational facts. This is a pure
// producer: the trusted caller commits its output with canonical read fences.
// No browser writer, new authority, or implicit historical migration.
function f4PortalReadModels_(options) {
 const F=f4Firestore_(),C=f4Core_(),N=C.need,db=F.connect(options.config),now=options.clock(),identity=options.identity;
 const ports=f4ActivitiesFunctionalPorts_(options),compat=f4ActivitiesActivityCompat_({...options,ports});
 return ports.readWithAuthoritySnapshot((check,get)=>{
 const technical=f4DevAdmin_().profile(get,identity.uid,now),actor=technical?f4DevAdmin_().identify(get,identity.uid,now):f4MemberAuthz_().identify(get,identity.uid,now),docs={};
 const requested=options.viewIds?new Set(options.viewIds):null,wants=id=>!requested||requested.has(id),gate=get('portalReaderControl/current');
 const context=ports.context();
 const auth=ports.readAuthority(context.board?['activityAttendanceRegistry/current']:[],!context.board,context.board?'ACTIVITY_BOARD_READ':'ACTIVITY_AGENDA_READ_SELF');
 function basis(policyId,capability,kind,expiry) {
  if(technical)return {generation:get('portalReaderControl/current').generation,ownerUid:identity.uid,personId:actor.binding.idPessoa,bindingVersion:actor.binding.bindingVersion,personRevision:actor.actor.h.revision,authorityEpoch:actor.control.epoch,kind:'DEV_ADMIN',profileRevision:technical.revision,capability,issuedAt:F.Timestamp.fromMillis(now),expiresAt:F.Timestamp.fromMillis(Math.min(expiry,C.ms(technical.expiresAt)))};
  const p=get('memberAuthzPolicies/'+policyId),c=actor.control;
  N(p?.active&&p.capabilities.includes(capability),'AUTHORIZATION_DENIED');
  return {generation:get('portalReaderControl/current').generation,ownerUid:identity.uid,personId:actor.binding.idPessoa,bindingVersion:actor.binding.bindingVersion,
   personRevision:actor.actor.h.revision,subjectEpoch:actor.subject.authzEpoch,authorityEpoch:c.epoch,
   policyEpoch:c.policyEpoch,readAuthzEpoch:c.readAuthzEpoch,policyId,policyRevision:p.revision,
   capability,kind,state:actor.state.state,issuedAt:F.Timestamp.fromMillis(now),
   expiresAt:F.Timestamp.fromMillis(Math.min(expiry,C.ms(p.expiresAt),C.ms(actor.state.validUntil)))};
 }
 const fenceBases=new Map();
 function start(id){const fenceId=f4PortalProjectionFenceId_(id,actor.binding.idPessoa);if(gate.operationalWrites&&fenceId)fenceBases.set(id,{id:fenceId,revision:get('portalReaderResourceFences/'+fenceId)?.revision||0});}
 function add(id,b,data){N(JSON.stringify(F.encode(data)).length<700000,'SCOPE_TOO_LARGE');docs['portalReaderViews/'+identity.uid+'/items/'+id]={schemaVersion:'portal-reader-view-v1',viewId:id,basis:{...b,...(fenceBases.has(id)?{dataFence:fenceBases.get(id)}:{})},data};}
 options.proofs?.push(auth);check([auth]);
 if(wants('self')){
 const self={};let selfEnd=identity.expiresAt;
 for(const view of ['profile','status','functions','directors']){const v=f4MemberOwnReader_().read(get,identity,{view},now);self[view]=f4MemberOwnReader_().wire(v);selfEnd=Math.min(selfEnd,v.validUntil);}
 add('self',basis(actor.control.selfReadPolicy.policyId,'MEMBERS_SELF_READ',technical?'DEV_ADMIN':'SELF',selfEnd),self);
 }
 const b=basis(auth.q.sourceId,auth.q.capability,auth.q.kind,C.ms(auth.proof.expiresAt));
 if(wants('context'))add('context',b,{...context,validUntil:C.ms(b.expiresAt),activityAgendaDirect:false});
 let rows=[];
 if(wants('activities')){
  start('activities');const bundle=compat.read('/atividades/bundle',{});rows=bundle.data.calendario;delete bundle.data.detalhesPorId;
  if(!gate.operationalWrites)for(const row of rows)N(get('canonicalActivities/'+row.idAtividade)?.recordMode==='HISTORICAL_PARTIAL','UNRESOLVED_DEPENDENCY','MUTABLE_ACTIVITY_NEEDS_COMMAND_PROJECTION');
  add('activities',b,{bundle,...(context.board?{admin:compat.read('/admin/atividades',{})}:{})});
 }
 const detailIds=requested?[...requested].filter(id=>id.startsWith('activity_')):rows.map(row=>'activity_'+row.idAtividade);
 for(const id of detailIds){start(id);const params={idAtividade:id.slice(9)};add(id,b,{normal:compat.read('/atividades/detalhe',params),...(context.board?{admin:compat.read('/admin/atividades/detalhe',params)}:{})});}
 // Empty frequency is preserved as an empty factual result, never invented
 // attendance or a percentage. Its epoch is checked independently by Rules.
 if(wants('frequency')){start('frequency');const fa=ports.readAuthority([],true,'ACTIVITY_ATTENDANCE_READ_SELF'),fb=basis(fa.q.sourceId,fa.q.capability,fa.q.kind,C.ms(fa.proof.expiresAt));options.proofs?.push(fa);
 check([fa]);add('frequency',fb,{epoch:get('activityAttendanceRegistry/current')?.epoch||0,result:f4ActivitiesFrequencyCompat_({...options,ports}).read('/v2/minha-frequencia')});}
 // The enclosing source snapshot rechecks all proofs against fresh time and
 // fresh batched authority sources, never the earlier projection timestamp.
 return docs;
 });
}
