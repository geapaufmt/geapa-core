// Successor public projection, prepared for explicit per-person editorial approval.
// Not a public writer. v1 remains served by its existing reader and Rules.
function f4MemberTeamProjection_(get,id,now){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),p=get('memberDirectoryPreferences/'+id);
 N(p?.schemaVersion==='member-team-preferences-v2','RECONCILIATION_REQUIRED');
 C.exact(p,['schemaVersion','revision','publicKey','hidden','photoUrl','photoRevoked','approvalRef','approvedCategories','biography','links']);
 N(C.rev(p.revision)&&p.revision>0&&C.id(p.approvalRef)&&/^[a-f0-9]{32}$/.test(p.publicKey)&&typeof p.hidden==='boolean'&&typeof p.photoRevoked==='boolean','SCHEMA_INVALID');
 N(Array.isArray(p.approvedCategories)&&new Set(p.approvedCategories).size===p.approvedCategories.length&&p.approvedCategories.every(k=>['BOARD','MEMBER','ACADEMIC','EGRESS'].includes(k)),'SCHEMA_INVALID');
 const url=v=>typeof v==='string'&&v.length<=2048&&/^https:\/\/[^\s@]+$/.test(v);
 N((p.photoUrl===null||url(p.photoUrl))&&typeof p.biography==='string'&&p.biography.length<=2000&&Array.isArray(p.links)&&p.links.length<=5&&p.links.every(l=>Object.keys(l).sort().join() ==='label,url'&&typeof l.label==='string'&&l.label.length<=100&&url(l.url)),'SCHEMA_INVALID');
 const control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),R=f4MemberRepository_(control),M=f4MemberModel_(true,control),h=R.planningHelpers().head(get('memberAggregateHeads/'+id),id,'PERSON'),docs={};
 for(const path of h.paths.filter(q=>/^member(Identities|Links|Events|Functions)\//.test(q)))docs[path]=get(path);
 R.planningHelpers().dependencies(docs,get);const d=M.derive(docs,id,h.revision,now),person=docs['memberIdentities/'+id],relation=get('memberInstitutionalRelations/'+id),eligible=[];
 if(!person.active||person.technicalAccount||p.hidden)return {state:'HIDDEN',dto:null};
 if(['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE'].includes(d.state.state)&&!d.state.blocked)eligible.push('MEMBER');
 if(d.currentDirectors.length)eligible.push('BOARD');
 if(d.state.state==='EGRESSO')eligible.push('EGRESS');
 if(relation){N(relation.schemaVersion==='people-relations-v1'&&relation.idPessoa===id&&C.rev(relation.revision)&&Array.isArray(relation.records),'RECONCILIATION_REQUIRED');if(relation.records.some(r=>['ACADEMIC_COLLABORATOR','PROFESSOR','ADVISOR','COORDINATOR'].includes(r.type)&&(!r.startDate||C.civilStart(r.startDate)<=now)&&(!r.endDate||now<C.civilStart(C.nextDay(r.endDate)))))eligible.push('ACADEMIC');}
 const categories=p.approvedCategories.filter(k=>eligible.includes(k));if(!categories.length)return {state:'HIDDEN',dto:null};
 // No editorial name, role or membership field exists in the preference schema.
 const roles=d.currentFunctions.filter(f=>categories.includes('BOARD')&&f.group==='DIRETORIA'||categories.includes('MEMBER')&&f.group==='ASSESSORIA').map(f=>({roleKey:f.roleKey,startDate:f.startDate,endDate:f.endDateReal||f.endDatePlanned}));
 const dto={schemaVersion:'public-team-person-v2',publicKey:p.publicKey,displayName:person.displayName,photoUrl:p.photoRevoked?null:p.photoUrl,categories,roles,biography:p.biography,links:p.links.map(l=>({...l}))};
 const boundaries=Object.values(docs).filter(v=>v.startDate).flatMap(v=>{const i=M.interval(v.startDate,v.endDateReal||v.endDatePlanned||v.endDate);return [i.a,i.z].filter(t=>t>now);});
 if(relation)for(const r of relation.records)for(const time of [r.startDate?C.civilStart(r.startDate):0,r.endDate?C.civilStart(C.nextDay(r.endDate)):0])if(time>now)boundaries.push(time);
 return {state:'PREPARED_REQUIRES_PUBLIC_RULES_V2',dto,proof:{personId:id,sourceRevision:h.revision,preferencesRevision:p.revision,relationRevision:relation?.revision||0,authorityEpoch:control.epoch,readAuthzEpoch:control.readAuthzEpoch,asOf:F.Timestamp.fromMillis(now),validUntil:F.Timestamp.fromMillis(Math.min(C.ms(d.state.validUntil),...boundaries))}};
}
