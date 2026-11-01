// Public presentation only. Canonical membership remains the sole source of identity and roles.
// Preferences contain only visibility/photo controls, never a replacement name, link or role.
function f4MemberDirectory_(){
 const C=f4Core_(),F=f4Firestore_(),N=C.need;
 function preferences(p){C.exact(p,['schemaVersion','revision','publicKey','hidden','photoUrl','photoRevoked']);N(p.schemaVersion==='member-directory-preferences-v1'&&C.rev(p.revision)&&p.revision>0&&/^[a-f0-9]{32}$/.test(p.publicKey)&&typeof p.hidden==='boolean'&&typeof p.photoRevoked==='boolean','SCHEMA_INVALID');N(p.photoUrl===null||typeof p.photoUrl==='string'&&p.photoUrl.length<=2048&&/^https:\/\/[^\s@]+$/.test(p.photoUrl),'SCHEMA_INVALID');return p;}
 function project(get,id,now){
  N(C.id(id)&&Number.isSafeInteger(now),'SCHEMA_INVALID');
  const control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS'));N(control.mode==='CANONICAL'&&control.activeWriterVersion==='M13','ACCESS_REVOKED');
  const p=preferences(get('memberDirectoryPreferences/'+id)),repo=f4MemberRepository_(control),loaded=repo.load(get,id),M=f4MemberModel_(true,control),derived=M.derive(loaded.docs,id,loaded.reg.revision,now),person=loaded.docs['memberIdentities/'+id];
  const path='publicMemberDirectory/'+p.publicKey,proofPath='memberDirectoryProofs/'+p.publicKey;
  const oldProof=get(proofPath);N(!oldProof||oldProof.idPessoa===id,'OPERATION_CONFLICT');
  const eligible=person.active&&!person.technicalAccount&&derived.state.state==='MEMBRO_EFETIVO'&&!derived.state.blocked&&!p.hidden&&!derived.currentFunctions.some(f=>f.group==='CONSELHO');
  if(!eligible)return {documents:{},deletePaths:[path,proofPath]};
  const functions=derived.currentFunctions,category=derived.currentDirectors.length?'DIRETOR':'MEMBRO';
  const dto={schemaVersion:'public-member-directory-v1',displayName:person.displayName,photoUrl:p.photoRevoked?null:p.photoUrl,category,functions:[...new Set(functions.map(f=>f.roleKey.replaceAll('_',' ')))]};
  N(typeof dto.displayName==='string'&&dto.displayName.length>0&&dto.displayName.length<=200&&dto.functions.length<=23,'SCHEMA_INVALID');
  const bounds=Object.values(loaded.docs).filter(v=>v.startDate&&/^m1-member(Functions|Governance)/.test(v.schemaVersion)).flatMap(v=>{const i=M.interval(v.startDate,v.endDateReal||v.endDatePlanned||v.endDate);return [i.a,i.z].filter(t=>t>now);});
  const proof={schemaVersion:'member-directory-proof-v1',idPessoa:id,sourceRevision:loaded.reg.revision,authorityEpoch:control.epoch,readAuthzEpoch:control.readAuthzEpoch,preferencesRevision:p.revision,asOf:F.Timestamp.fromMillis(now),validUntil:F.Timestamp.fromMillis(Math.min(C.ms(derived.state.validUntil),...bounds)),dto};
  return {documents:{[path]:dto,[proofPath]:proof},deletePaths:[]};
 }
 function indexPlan(get,id,plan){const pref=preferences(get('memberDirectoryPreferences/'+id)),path='publicMemberDirectoryIndex/current',old=get(path);if(old){C.exact(old,['schemaVersion','keys']);N(old.schemaVersion==='public-member-directory-index-v1'&&Array.isArray(old.keys)&&old.keys.length<=500&&new Set(old.keys).size===old.keys.length&&old.keys.every(k=>/^[a-f0-9]{32}$/.test(k)),'SCHEMA_INVALID');}const keys=new Set(old?.keys||[]);if(plan.deletePaths.length)keys.delete(pref.publicKey);else keys.add(pref.publicKey);N(keys.size<=500,'AGGREGATE_LIMIT_EXCEEDED');plan.documents[path]={schemaVersion:'public-member-directory-index-v1',keys:[...keys].sort()};return plan;}
 function refresh(tx,id,now){const info=F.inspectTransaction(tx);N(info.environment==='SYNTHETIC_LOCAL'||f4OperationalDestination_(info),'DESTINATION_INVALID');const get=p=>tx.get(p).data(),plan=indexPlan(get,id,project(get,id,now)),writes=Object.entries(plan.documents).map(([path,data])=>({path,kind:'set',data}));for(const path of plan.deletePaths)if(tx.get(path).data())writes.push({path,kind:'delete',data:null});for(const w of writes)tx.get(w.path);F.previewTransaction(tx,writes);for(const w of writes)if(w.kind==='delete')tx.delete(w.path);else tx.set(w.path,w.data);return {state:plan.deletePaths.length?'HIDDEN':'PUBLISHED',writes:writes.length};}
 // Trusted operator port; not a browser command. Own removal requests use the existing
 // administrative channel. Changing visibility never edits a canonical person or function.
 function setHidden(tx,id,expected,hidden,now){N(C.id(id)&&C.rev(expected)&&typeof hidden==='boolean','SCHEMA_INVALID');const old=preferences(tx.get('memberDirectoryPreferences/'+id).data());N(old.revision===expected,'REVISION_CONFLICT');const info=F.inspectTransaction(tx);N(info.environment==='SYNTHETIC_LOCAL'||f4OperationalDestination_(info),'DESTINATION_INVALID');const control=f4MemberAuthority_().validate(tx.get('memberAuthorityControl/MEMBERS').data());N(control.mode==='CANONICAL'&&control.activeWriterVersion==='M13','ACCESS_REVOKED');const next={...old,revision:expected+1,hidden},get=p=>p==='memberDirectoryPreferences/'+id?next:tx.get(p).data(),plan=indexPlan(get,id,project(get,id,now));for(const path of [...Object.keys(plan.documents),...plan.deletePaths])tx.get(path);tx.set('memberDirectoryPreferences/'+id,next);for(const[path,value]of Object.entries(plan.documents))tx.set(path,value);for(const path of plan.deletePaths)tx.delete(path);return {hidden};}
 return Object.freeze({preferences,project,refresh,setHidden});
}
