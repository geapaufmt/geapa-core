// Administrative directory. This is not the public members directory or a grant
// of MEMBERS_CASE_READ. All classification and field selection stay in Core.
function f4PeopleDirectory_({config,identity,clock=Date.now}) {
 const C=f4Core_(),K=f4MemberAuthzContracts_(),N=K.need,A=f4MemberAuthz_(),F=f4Firestore_(),db=F.connect(config);
 const routes=['/admin/pessoas','/admin/pessoas/detalhe'],cap='PEOPLE_DIRECTORY_READ';
 const fields=['idPessoa','displayLabel','state','categories','functions','relations'];
 const generalPath=p=>/^member(Identities|Links|Events|Functions)\/[A-Za-z0-9_-]{1,48}$/.test(p);
 const norm=s=>String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
 function access(get,now){
  C.exact(identity,['uid','expiresAt']);N(C.id(identity.uid)&&now<identity.expiresAt,'ACCESS_REVOKED');
  // Never fall back to DEV_ADMIN or infer a grant from menu visibility.
  const actor=A.identify(get,identity.uid,now),settings=f4MemberPortalPrepare_({config,identity,clock}).configuration(get);
  const sources=settings.policyIds.map(id=>get('memberAuthzPolicies/'+id)).filter(p=>p?.capabilities?.includes(cap)).filter(p=>{
   K.policy(p);return p.active&&p.proofKind==='ROLE'&&C.ms(p.issuedAt)<=now&&now<C.ms(p.expiresAt)&&actor.derived.currentFunctions.some(f=>f.roleKey===p.actorSelector.roleKey&&f.group===p.actorSelector.group&&f.scopeKey===p.actorSelector.actorScope);
  });
  N(sources.length>0,'ACCESS_REVOKED');
  const ids=[...new Set(sources.flatMap(p=>p.targets.allowedPersons))].sort();N(ids.length>0&&ids.length<=250,'SCOPE_TOO_LARGE');
  const masks=id=>[...new Set(sources.filter(p=>!p.targets.ownPerson&&p.targets.allowedPersons.includes(id)&&p.targets.allowedResources.includes('memberIdentities/'+id)).flatMap(p=>p.fieldMasks[cap]))];
  const visible=ids.filter(id=>['idPessoa','displayLabel'].every(f=>masks(id).includes(f)));
  N(visible.length>0,'ACCESS_REVOKED');
  return {actor,sources,ids:visible,masks,expiresAt:Math.min(identity.expiresAt,C.ms(actor.state.validUntil),C.ms(actor.derived.state.validUntil),...sources.map(p=>C.ms(p.expiresAt)))};
 }
 function relation(get,id){
  const v=get('memberInstitutionalRelations/'+id);if(!v)return [];
  C.exact(v,['schemaVersion','idPessoa','revision','approvalRef','records']);N(v.schemaVersion==='people-relations-v1'&&v.idPessoa===id&&C.rev(v.revision)&&C.id(v.approvalRef)&&Array.isArray(v.records)&&v.records.length<=20,'RECONCILIATION_REQUIRED');
  const types=['ACADEMIC_COLLABORATOR','PROFESSOR','ADVISOR','COORDINATOR','EXTERNAL_PARTICIPANT','SPEAKER'];
  return v.records.map(r=>{C.exact(r,['id','type','label','sourceRef','startDate','endDate']);N(C.id(r.id)&&types.includes(r.type)&&typeof r.label==='string'&&r.label.length<=120&&C.id(r.sourceRef),'RECONCILIATION_REQUIRED');for(const d of [r.startDate,r.endDate])if(d!==null)C.civilStart(d);N(!r.startDate||!r.endDate||r.startDate<=r.endDate,'RECONCILIATION_REQUIRED');return {id:r.id,type:r.type,label:r.label,startDate:r.startDate,endDate:r.endDate};});
 }
 function general(get,id,a,now){
  // Reuse the canonical aggregate/head/dependency/model contracts, without
  // loading CPF, contact, birth data or identifiers for a directory query.
  const M=f4MemberModel_(true,a.actor.control),R=f4MemberRepository_(a.actor.control).planningHelpers(),h=R.head(get('memberAggregateHeads/'+id),id,'PERSON'),docs={};
  for(const path of h.paths.filter(generalPath)){const v=get(path);N(v&&v.idPessoa===id,'RECONCILIATION_REQUIRED');docs[path]=v;}
  R.dependencies(docs,get);const target={h,d:docs,M},person=docs['memberIdentities/'+id];N(person,'RECONCILIATION_REQUIRED');
  if(person.technicalAccount)return null;
  const d=target.M.derive(target.d,id,target.h.revision,now),mask=a.masks(id),relations=(mask.includes('relations')||mask.includes('categories'))?relation(get,id):[],categories=[];
  if(mask.includes('categories')){
   if(['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE','MEMBRO_EM_ESPERA','SUSPENSO'].includes(d.state.state))categories.push('MEMBER');
   if(d.state.state==='EGRESSO')categories.push('EGRESS');
   if(relations.some(r=>['ACADEMIC_COLLABORATOR','PROFESSOR','ADVISOR','COORDINATOR'].includes(r.type)))categories.push('ACADEMIC');
   if(relations.some(r=>r.type==='EXTERNAL_PARTICIPANT'))categories.push('EXTERNAL');
   if(!categories.length)categories.push('UNCLASSIFIED');
  }
  const dto={idPessoa:id,displayLabel:person.displayName};
  if(mask.includes('state'))dto.state=d.state.state;
  if(mask.includes('categories'))dto.categories=categories;
  if(mask.includes('functions'))dto.functions=d.currentFunctions.map(f=>({roleKey:f.roleKey,group:f.group,startDate:f.startDate,endDate:f.endDateReal||f.endDatePlanned}));
  if(mask.includes('relations'))dto.relations=relations;
  return {dto,target,derived:d};
 }
 function prefetch(tx,ids){
  // Bounded server-side lookup of the approved cohort, never a private browser
  // dump or an unbounded collection scan. 250-person cap is explicit.
  const get=p=>tx.get(p).data(),batch=paths=>{for(let i=0;i<paths.length;i+=150)tx.getAll(paths.slice(i,i+150));};
  batch(ids.map(id=>'memberAggregateHeads/'+id));
  const paths=[...new Set(ids.flatMap(id=>{const h=get('memberAggregateHeads/'+id);N(h&&h.id===id&&h.kind==='PERSON'&&Array.isArray(h.paths),'RECONCILIATION_REQUIRED');return h.paths.filter(generalPath);} ))];
  N(paths.length<=2500,'SCOPE_TOO_LARGE');batch(paths);batch(ids.map(id=>'memberInstitutionalRelations/'+id));
 }
 function read(input){
  C.exact(input,['route','params']);N(routes.includes(input.route),'COMMAND_TYPE_INVALID');const p=input.params;
  if(input.route===routes[1])C.exact(p,['personId']);else C.exact(p,['query','category','state','role','order','page','pageSize',...(Object.hasOwn(p,'snapshot')?['snapshot']:[])]);
  if(input.route===routes[0]){
   N(typeof p.query==='string'&&p.query.length<=80&&['','MEMBER','EGRESS','ACADEMIC','EXTERNAL','UNCLASSIFIED'].includes(p.category)&&typeof p.state==='string'&&p.state.length<=40&&typeof p.role==='string'&&p.role.length<=48&&['name-asc','name-desc'].includes(p.order)&&Number.isInteger(p.page)&&p.page>=1&&p.page<=250&&[10,25,50].includes(p.pageSize),'COMMAND_PAYLOAD_INVALID');
   N(p.snapshot===undefined||typeof p.snapshot==='string'&&/^[a-f0-9]{64}$/.test(p.snapshot),'COMMAND_PAYLOAD_INVALID');
  }else N(C.id(p.personId),'COMMAND_PAYLOAD_INVALID');
  return db.transaction(tx=>{
   const now=clock(),get=q=>tx.get(q).data(),a=access(get,now),fence=get('portalReaderResourceFences/members_operations')?.revision||0;
   if(input.route===routes[1]){
    N(a.ids.includes(p.personId),'ACCESS_REVOKED');const x=general(get,p.personId,a,now);N(x,'RESOURCE_NOT_FOUND');
    // History is institutional data too; it needs its own explicit field mask.
    const mask=a.masks(p.personId),detail={...x.dto};
    if(mask.includes('history'))detail.history={links:x.target.M.rows(x.target.d,'memberLinks').map(v=>({id:v.id,type:v.type,startDate:v.startDate,endDate:v.endDate})),functions:x.target.M.rows(x.target.d,'memberFunctions').map(v=>({id:v.id,roleKey:x.target.d['memberRoleCatalog/'+v.idCargo].roleKey,startDate:v.startDate,endDate:v.endDateReal||v.endDatePlanned}))};
    if(mask.includes('history'))detail.presentation=f4MemberTenurePresentation_(x.target.d,p.personId,x.derived,now);
    detail.privateFieldsAvailable=false;
    return {ok:true,data:detail,validUntil:a.expiresAt};
   }
   prefetch(tx,a.ids);const all=a.ids.map(id=>general(get,id,a,now)).filter(Boolean).map(x=>x.dto);
   const maskEvery=field=>a.ids.every(id=>a.masks(id).includes(field));
   N((!p.category||maskEvery('categories'))&&(!p.state||maskEvery('state'))&&(!p.role||maskEvery('functions')),'ACCESS_REVOKED');
   const snapshot=K.digest({uid:identity.uid,sources:a.sources,actorRevision:a.actor.actor.h.revision,subject:a.actor.subject.authzEpoch,authority:a.actor.control,rows:all,fence});
   if(p.snapshot)N(p.snapshot===snapshot,'DIRECTORY_CHANGED');
   const rows=all.filter(v=>(!p.query||norm(v.displayLabel+' '+v.idPessoa).includes(norm(p.query)))&&(!p.category||v.categories?.includes(p.category))&&(!p.state||v.state===p.state)&&(!p.role||v.functions?.some(f=>f.roleKey===p.role)));
   rows.sort((a,b)=>{const x=norm(a.displayLabel),y=norm(b.displayLabel),cmp=x<y?-1:x>y?1:a.idPessoa.localeCompare(b.idPessoa);return p.order==='name-desc'?-cmp:cmp;});
   return {ok:true,data:{items:rows.slice((p.page-1)*p.pageSize,p.page*p.pageSize),page:p.page,pageSize:p.pageSize,total:rows.length,pages:Math.max(1,Math.ceil(rows.length/p.pageSize)),snapshot,scope:'POLICY_PERSONS',options:{states:[...new Set(all.map(v=>v.state).filter(Boolean))].sort(),roles:[...new Set(all.flatMap(v=>(v.functions||[]).map(f=>f.roleKey)))].sort()},privateFieldsIncluded:false},validUntil:a.expiresAt};
  });
 }
 return {routes,read,access,fields};
}
