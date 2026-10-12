// Trusted registry for the one Core pipeline. A domain discriminator is never a grant.
function f4SharedCommandCatalog_(){
 const K=f4MemberCommandCatalog_(),C=f4Core_(),N=C.need;
 const act=type=>typeof type==='string'&&type.startsWith('ACT.');
 const types=Object.freeze({...K.types,...Object.fromEntries(Object.entries(f4ActivitiesContracts_()).map(([k,v])=>['ACT.'+k,v.capabilities]))});
 function selector(v){C.exact(v,['operationId','digest','commandType']);N(C.op(v.operationId)&&K.hex(v.digest)&&Object.hasOwn(types,v.commandType),'COMMAND_PAYLOAD_INVALID');return v}
 function capabilities(r){return act(r.commandType)?[f4ActivitiesPipeline_().capability(r)]:K.types[r.commandType]}
 function base(r){if(!act(r?.commandType))return K.base(r);C.exact(r,['schemaVersion','operationId','ownerUid','commandType','target','expectedRevisions','payload','evidenceRefs','decisionRef','ticketId','digest','receivedAt','state']);N(r.schemaVersion==='geapa-act-command-v13'&&C.op(r.operationId)&&C.id(r.ownerUid)&&C.id(r.ticketId)&&K.hex(r.digest)&&Number.isSafeInteger(C.ms(r.receivedAt))&&r.state==='RECEIVED','COMMAND_PAYLOAD_INVALID');N(Array.isArray(r.expectedRevisions)&&r.expectedRevisions.length===0&&Array.isArray(r.evidenceRefs)&&r.evidenceRefs.length===0&&r.decisionRef===null,'COMMAND_PAYLOAD_INVALID');f4ActivitiesPipeline_().validate(r);return r}
 function envelope(project,r,validate){base(r);N(typeof validate==='function'&&validate(r)===true,'COMMAND_PAYLOAD_INVALID');N(r.digest===digest(project,r),'OPERATION_CONFLICT');return r}
 function revisionList(rows){N(Array.isArray(rows),'SCHEMA_INVALID');let previous='';for(const row of rows){C.exact(row,['resourceType','id','revision']);const path=row.resourceType+'/'+row.id;N(C.id(row.resourceType)&&((C.id(row.id))||f4ActivitiesModel_().validPath(path))&&C.rev(row.revision)&&path>previous,'SCHEMA_INVALID');previous=path;}return rows}
 function digest(project,r){if(act(r.commandType)&&f4ActivitiesSystem_().is(r))return K.digest(project,{...r,target:{...r.target,authorization:{...r.target.authorization,runId:'CORE_ATTEMPT'}}});return K.digest(project,r)}
 return {...K,types,selector,capabilities,act,base,envelope,revisionList,digest};
}
function f4SharedFamilyCatalog_(config){
 const a=f4ActivitiesPipeline_();return Object.freeze({...f4MemberFamilyCatalog_(config),...Object.fromEntries(Object.keys(f4ActivitiesContracts_()).map(k=>['ACT.'+k,Object.freeze({validate:a.validate,resolve:a.resolve})]))});
}
function f4SharedAuthz_(){
 const m=f4MemberAuthz_(),a=f4MemberAuthz_(f4ActivitiesAuthzDomain_());
 return {...m,authorize:(get,q,n)=>f4DevReadTiming_().stage('authorization',()=>q?.kind==='DEV_ADMIN'&&q.schemaVersion==='act-authz-request-v13'?f4DevAdmin_().authorize(get,q,n):q?.schemaVersion==='act-system-authz-request-v13'?f4ActivitiesSystem_().authorize(get,q,n):q?.schemaVersion==='act-authz-request-v13'?a.authorize(get,q,n):m.authorize(get,q,n)),recheck:(get,q,p,n)=>f4DevReadTiming_().stage('authorization',()=>q?.kind==='DEV_ADMIN'&&q.schemaVersion==='act-authz-request-v13'?f4DevAdmin_().recheck(get,q,p,n):q?.schemaVersion==='act-system-authz-request-v13'?f4ActivitiesSystem_().recheck(get,q,p,n):q?.schemaVersion==='act-authz-request-v13'?a.recheck(get,q,p,n):m.recheck(get,q,p,n))};
}


function f4CoreDomainInvalidation_(get,pending,person,now){f4ActivitiesMembersFence_(get,pending,person,now);}
function f4CoreDomainPaths_(path){return /^coreDevAdmins\/[A-Za-z0-9_-]{1,48}$/.test(path)||/^(portalReaderControl\/current|portalReaderResourceFences\/[A-Za-z0-9_-]{1,160}|commandProjectionInvalidations\/[a-f0-9]{32}|portalReaderViews\/[A-Za-z0-9_-]{1,48}\/items\/[A-Za-z0-9_-]{1,160})$/.test(path)||f4ActivitiesModel_().validPath(path);}
