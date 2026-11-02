// Documentary consultation. No Sheets I/O, normative execution or write endpoint.
function f4InstitutionalDirectoryModel_(){
 const C=f4Core_(),N=C.need,sections=['identity','parameters','norms','governance'];
 const text=(x,max=4000)=>typeof x==='string'&&x.length<=max;
 function validate(v){
  C.exact(v,['schemaVersion','id','revision','section','nature','records','sourceHash']);
  N(v.schemaVersion==='institutional-directory-v1'&&sections.includes(v.section)&&v.id===v.section&&C.rev(v.revision)&&v.nature==='DOCUMENTARY_ONLY'&&/^[a-f0-9]{64}$/.test(v.sourceHash),'SCHEMA_INVALID');
  N(Array.isArray(v.records)&&v.records.length<=150,'SCOPE_TOO_LARGE');
  N(new Set(v.records.map(x=>x.id)).size===v.records.length,'SCHEMA_INVALID');
  for(const r of v.records){
   C.exact(r,['id','title','kind','status','classification','visibility','fields','sources','issues','canonical']);
   N(C.id(r.id)&&text(r.title,240)&&text(r.kind,64)&&text(r.status,80)&&['EDITORIAL','CONFIGURABLE','NORMATIVE_CONDITIONAL','TECHNICAL_RESTRICTED'].includes(r.classification)&&['BOARD','RESTRICTED'].includes(r.visibility),'SCHEMA_INVALID');
   N(Array.isArray(r.fields)&&r.fields.length<=40&&r.fields.every(f=>{C.exact(f,['label','value']);return text(f.label,100)&&text(f.value,8000)}),'SCHEMA_INVALID');
   N(Array.isArray(r.sources)&&r.sources.length>0&&r.sources.length<=10&&r.sources.every(s=>{C.exact(s,['document','sheet','row','hash']);return text(s.document,120)&&text(s.sheet,80)&&Number.isInteger(s.row)&&s.row>0&&/^[a-f0-9]{64}$/.test(s.hash)}),'SCHEMA_INVALID');
   N(Array.isArray(r.issues)&&r.issues.length<=12&&r.issues.every(i=>text(i,600)),'SCHEMA_INVALID');
   N(r.canonical===null||typeof r.canonical==='object','SCHEMA_INVALID');
   if(r.canonical){C.exact(r.canonical,['path','status','evidence']);N(/^(institutionalNormVersions|institutionalNormApprovals|activityPolicies|presentationCycles|memberOperationalConfig|publicPages)\/[A-Za-z0-9_-]{1,80}$/.test(r.canonical.path)&&text(r.canonical.status,80)&&text(r.canonical.evidence,600),'SCHEMA_INVALID');}
  }
  return v;
 }
 function dto(v,mask){
  validate(v);return v.records.filter(r=>r.visibility==='BOARD').map(r=>({id:r.id,title:r.title,kind:r.kind,status:r.status,classification:r.classification,
   fields:mask.includes('fields')?r.fields:[],sources:mask.includes('provenance')?r.sources:[],issues:mask.includes('issues')?r.issues:[],canonical:mask.includes('canonical')?r.canonical:null}));
 }
 // Pure preparation only. No command, no admission and no write permission are
 // installed by this contract. The approved executor must revalidate at commit.
 const editable={EMAIL_OFICIAL:'EMAIL',PERFIL_GEAPA_INSTAGRAM:'HTTPS',CANAL_YOUTUBE_GEAPA:'HTTPS',LINKEDIN_GEAPA:'HTTPS',LINKTREE_GEAPA:'HTTPS',LOCAL_PADRAO_REUNIOES:'TEXT'};
 function prepareContact(v,input){
  validate(v);C.exact(input,['fieldId','value','expectedRevision','reason']);N(v.section==='identity'&&v.revision===input.expectedRevision,'REVISION_CONFLICT');
  N(Object.hasOwn(editable,input.fieldId)&&text(input.reason,1000)&&input.reason.trim().length>=10,'COMMAND_PAYLOAD_INVALID');
  N(text(input.value,300)&&input.value===input.value.trim()&&!/[\r\n\x00-\x1f]/.test(input.value),'COMMAND_PAYLOAD_INVALID');
  const type=editable[input.fieldId];N(type==='EMAIL'?/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(input.value):type==='HTTPS'?/^https:\/\/[A-Za-z0-9.-]+(?::443)?(?:\/[^\s]*)?$/.test(input.value)&&!/^https:\/\/(?:localhost|127\.|169\.254\.|10\.|192\.168\.)/i.test(input.value):input.value.length>=3,'COMMAND_PAYLOAD_INVALID');
  const old=v.records.find(r=>r.id===input.fieldId);N(old&&old.visibility==='BOARD','RESOURCE_NOT_FOUND');
  N(old.classification==='EDITORIAL'&&old.canonical===null&&old.fields.filter(f=>f.label==='Valor').length===1,'RECONCILIATION_REQUIRED');
  return {kind:'CONTACT_CHANGE_PROPOSAL',admissionEnabled:false,expectedRevision:v.revision,beforeHash:C.hash(v),fieldId:input.fieldId,value:input.value,reason:input.reason,approvalRequired:true,publicUpdate:false};
 }
 function publicFields(v,approval){
  validate(v);C.exact(approval,['sourceHash','revision','fields','approvalRef']);N(approval.sourceHash===C.hash(v)&&approval.revision===v.revision&&C.id(approval.approvalRef),'REVISION_CONFLICT');
  const allowed=['EMAIL_OFICIAL','PERFIL_GEAPA_INSTAGRAM','CANAL_YOUTUBE_GEAPA','LINKEDIN_GEAPA','LINKTREE_GEAPA','NOME_OFICIAL_GRUPO','SIGLA_OFICIAL_GRUPO','CIDADE_SEDE','ESTADO_SEDE'];
  N(v.section==='identity'&&Array.isArray(approval.fields)&&approval.fields.length>0&&new Set(approval.fields).size===approval.fields.length&&approval.fields.every(x=>allowed.includes(x)),'AUTHORIZATION_DENIED');
  return approval.fields.map(id=>{const r=v.records.find(x=>x.id===id);N(r?.visibility==='BOARD'&&!r.issues.length,'RECONCILIATION_REQUIRED');const value=r.fields.find(f=>f.label==='Valor')?.value;N(text(value,300)&&!/(?:chat\.whatsapp\.com|joinchat|token=)/i.test(value),'AUTHORIZATION_DENIED');return {id,value};});
 }
 return {sections,validate,dto,prepareContact,publicFields,editable};
}

// Shared ROLE resolution: a directory read never implies permission to edit.
function f4InstitutionalPolicies_(get,identity,now,ids,cap,path){
 const C=f4Core_(),K=f4MemberAuthzContracts_(),N=K.need;
 N(now<identity.expiresAt,'ACCESS_REVOKED');
 const actor=f4MemberAuthz_().identify(get,identity.uid,now);
 N(Array.isArray(ids)&&ids.length>0&&ids.length<=16&&new Set(ids).size===ids.length&&ids.every(C.id),'AUTHORIZATION_DENIED');
 const policies=ids.map(id=>get('memberAuthzPolicies/'+id)).filter(Boolean).filter(p=>{
  K.policy(p);return p.active&&p.proofKind==='ROLE'&&p.capabilities.includes(cap)&&C.ms(p.issuedAt)<=now&&now<C.ms(p.expiresAt)&&!p.targets.ownPerson&&p.targets.allowedPersons.includes(actor.actor.h.id)&&p.targets.allowedResources.includes(path)&&actor.derived.currentFunctions.some(f=>f.roleKey===p.actorSelector.roleKey&&f.group===p.actorSelector.group&&f.scopeKey===p.actorSelector.actorScope);
 });N(policies.length>0,'AUTHORIZATION_DENIED');
 return {actor,policies,mask:[...new Set(policies.flatMap(p=>p.fieldMasks[cap]))],validUntil:Math.min(now+60000,identity.expiresAt,C.ms(actor.state.validUntil),C.ms(actor.derived.state.validUntil),...policies.map(p=>C.ms(p.expiresAt)))};
}

function f4InstitutionalDirectoryRead_(options,params){
 const C=f4Core_(),K=f4MemberAuthzContracts_(),A=f4MemberAuthz_(),N=K.need,M=f4InstitutionalDirectoryModel_(),cap='INSTITUTIONAL_DIRECTORY_READ';
 C.exact(params,['section','query','kind','page']);N(M.sections.includes(params.section)&&typeof params.query==='string'&&params.query.length<=80&&typeof params.kind==='string'&&params.kind.length<=64&&Number.isInteger(params.page)&&params.page>=1&&params.page<=15,'COMMAND_PAYLOAD_INVALID');
 return f4Firestore_().connect(options.config).transaction(tx=>{
  const get=p=>tx.get(p).data(),now=(options.clock||Date.now)();N(now<options.identity.expiresAt,'ACCESS_REVOKED');
  const actor=A.identify(get,options.identity.uid,now),control=get('institutionalDirectoryControl/current');
  N(control?.schemaVersion==='institutional-directory-control-v1'&&control.enabled===true&&Array.isArray(control.policyIds)&&control.policyIds.length>0&&control.policyIds.length<=16,'FEATURE_DISABLED');
  const path='institutionalDirectory/'+params.section;
  const resolved=f4InstitutionalPolicies_(get,options.identity,now,control.policyIds,cap,path),policies=resolved.policies,mask=resolved.mask;N(mask.includes('summary'),'AUTHORIZATION_DENIED');
  const v=get(path);N(v,'PROJECTION_NOT_READY');const rows=M.dto(v,mask).map(r=>{if(!r.canonical)return r;const live=get(r.canonical.path);return {...r,canonical:{path:r.canonical.path,status:live?String(live.status||'REGISTRO CANÔNICO EXISTENTE'):'REFERÊNCIA AUSENTE',evidence:live?'Consulta canônica atual'+(live.revision?' · revisão '+live.revision:'')+(live.effectiveFrom?' · vigência '+live.effectiveFrom:'')+'. Mapeamento documental: '+r.canonical.evidence:'Referência não localizada; a fonte documental não assume sua autoridade.'}};}),normalize=s=>s.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
  let editableFields=[];const edit=get('institutionalEditControl/current');if(v.section==='identity'&&edit?.enabled===true&&typeof f4InstitutionalContactCommands_==='function'){try{const resolved=f4InstitutionalContactCommands_().control(get,options.config,options.identity,now,true,true);editableFields=v.records.filter(r=>r.visibility==='BOARD'&&r.classification==='EDITORIAL'&&r.canonical===null&&r.fields.filter(f=>f.label==='Valor').length===1&&resolved.c.fields.includes(r.id)&&resolved.auth.mask.includes(r.id)).map(r=>r.id);}catch(e){if(!['AUTHORIZATION_DENIED','ACCESS_REVOKED','PILOT_DISABLED','M13_AUTHZ_DENIED'].includes(e.code))throw e;}}
  const filtered=rows.filter(r=>(!params.kind||r.kind===params.kind)&&(!params.query||normalize(r.id+' '+r.title).includes(normalize(params.query))));
  filtered.sort((a,b)=>a.title.localeCompare(b.title,'pt-BR')||a.id.localeCompare(b.id));
  return {ok:true,data:{schemaVersion:'institutional-consultation-v1',section:v.section,revision:v.revision,sourceHash:v.sourceHash,nature:v.nature,items:filtered.slice((params.page-1)*10,params.page*10),total:filtered.length,page:params.page,pages:Math.max(1,Math.ceil(filtered.length/10)),kinds:[...new Set(rows.map(r=>r.kind))].sort(),editingEnabled:editableFields.length>0,editableFields,validUntil:Math.min(now+60000,options.identity.expiresAt,C.ms(actor.state.validUntil),C.ms(actor.derived.state.validUntil),...policies.map(p=>C.ms(p.expiresAt)))}};
 });
}
