// M63-A only: pure, closed contracts. No authorization, services or writer registration.
function f4MemberContracts_() {
  const C = f4Core_();
  const need = (ok) => { if (!ok) C.fail('M13_SCHEMA_INVALID'); };
  const id = x => typeof x === 'string' && /^[A-Za-z0-9_-]{1,48}$/.test(x);
  const positive = x => Number.isSafeInteger(x) && x > 0;
  const scopes = ['PRESIDENCIA','ADMINISTRACAO','RECURSOS_HUMANOS','FINANCAS','COMUNICACAO','EXTENSAO','EVENTOS'];
  // Explicit M611 associations. No inference for future roles.
  const roles = {
    PRESIDENTE:['DIRETORIA','PRESIDENCIA'], VICE_PRESIDENTE:['DIRETORIA','PRESIDENCIA'],
    SECRETARIO_GERAL:['DIRETORIA','ADMINISTRACAO'], SECRETARIO_EXECUTIVO:['DIRETORIA','ADMINISTRACAO'],
    DIRETOR_ADMINISTRATIVO:['DIRETORIA','ADMINISTRACAO'], DIRETOR_RECURSOS_HUMANOS:['DIRETORIA','RECURSOS_HUMANOS'],
    TESOUREIRO:['DIRETORIA','FINANCAS'], DIRETOR_COMUNICACAO:['DIRETORIA','COMUNICACAO'],
    ASSESSOR_COMUNICACAO_1:['ASSESSORIA','COMUNICACAO'], ASSESSOR_COMUNICACAO_2:['ASSESSORIA','COMUNICACAO'],
    DIRETOR_EXTENSAO:['DIRETORIA','EXTENSAO'], DIRETOR_EVENTOS:['DIRETORIA','EVENTOS'],
    ASSESSOR_ADMINISTRATIVO_1:['ASSESSORIA','ADMINISTRACAO']
  };
  const cases = [
    'add1094af15daa82fa0b6edd249175be1a4edcb397aa9b989de00403e7b1e8e4',
    'aeaff54772b1d4e5643d381fc0edca431dd92abe160bbe464ac60f064ad7b40b',
    '642917a378623dfbf7aca41cc36458b7d9839ba89f1a2c2fbec6079d4f3f51df'
  ];
  const dtoFields = ['idPessoa','idVigencia','roleKey','group','scopeKey','idDiretoria','idTermo','startDate','endDatePlanned','endDateReal'];
  const indexFields = ['idPessoa','kind','ids','sourceRevision','asOf','validUntil'];
  const proofFields = ['personRevision','functionRevision','boardId','boardEpoch','termId','termRevision','roleId','roleRevision','startsAt','endsAtExclusive','dto'];
  const controlFields = ['id','revision','mode','epoch','policyEpoch','readAuthzEpoch','activeWriterVersion','approvalRef','updatedAt','nullableFunctionApprovals','selfReadPolicy'];
  const versions = {
    'm1-memberFunctions-v2':['id','revision','origin','idPessoa','idVinculo','idCargo','scopeKey','idDiretoria','idTermo','startDate','endDatePlanned','endDateReal'],
    'safe-member-function-v13':dtoFields,
    'member-function-index-v13':indexFields.concat('authorityEpoch'),
    'member-function-proof-v13':proofFields.concat('authorityEpoch','nullableApprovalRef'),
    'member-authority-control-v13':controlFields
  };
  function exact(v, fields) { C.exact(v, fields); return v; }
  function timestamp(x) { need(!!x && typeof x.toMillis === 'function' && Number.isSafeInteger(x.toMillis())); return x.toMillis(); }
  function functional(roleKey, group, scope) {
    need(Object.prototype.hasOwnProperty.call(roles,roleKey) && roles[roleKey][0]===group && roles[roleKey][1]===scope && scopes.includes(scope));
  }
  function council(roleKey, group) { need(roleKey==='CONSELHEIRO_CONSULTIVO' && group==='CONSELHO'); }
  function approval(control, functionId) {
    need(control && Object.prototype.hasOwnProperty.call(control.nullableFunctionApprovals,functionId));
    return control.nullableFunctionApprovals[functionId];
  }
  function functionScope(source, role, control) {
    f4MemberAuthority_().validate(control);
    need(role.id===source.idCargo);
    if (source.scopeKey===null) { council(role.roleKey,role.group); approval(control,source.id); }
    else functional(role.roleKey,role.group,source.scopeKey);
    return source;
  }
  function dto(v) {
    need(v && ['safe-member-function-v12','safe-member-function-v13'].includes(v.schemaVersion));
    exact(v,['schemaVersion'].concat(dtoFields));
    if(v.schemaVersion==='safe-member-function-v12') {
      // Preserve the original decoder's string contract, including historical scopes.
      for(const k of dtoFields) need(k==='endDatePlanned'||k==='endDateReal' ? v[k]===null||typeof v[k]==='string' : typeof v[k]==='string');
    } else {
      for(const k of ['idPessoa','idVigencia','idDiretoria','idTermo']) need(id(v[k]));
      if(v.scopeKey===null) council(v.roleKey,v.group);
      // A v1 source may legitimately be reprojected with a historical non-M611 role/area.
      else need(id(v.scopeKey) && !['null','NO_FUNCTIONAL_SCOPE','CONSELHO'].includes(v.scopeKey) && id(v.roleKey) && ['DIRETORIA','ASSESSORIA','CONSELHO'].includes(v.group));
      const M=f4MemberModel_(true); need(v.endDatePlanned!==null||v.endDateReal!==null);
      M.interval(v.startDate,v.endDateReal||v.endDatePlanned);
      if(v.endDatePlanned!==null) M.interval(v.startDate,v.endDatePlanned);
    }
    return v;
  }
  function index(v) {
    const newer=v?.schemaVersion==='member-function-index-v13';
    need(newer||v?.schemaVersion==='member-function-index-v12');
    exact(v,['schemaVersion'].concat(indexFields,newer?['authorityEpoch']:[]));
    need(id(v.idPessoa)&&['memberCurrentFunctions','memberCurrentDirectors'].includes(v.kind)&&positive(v.sourceRevision));
    need(Array.isArray(v.ids)&&v.ids.every(id)&&new Set(v.ids).size===v.ids.length);
    need(timestamp(v.asOf)<timestamp(v.validUntil)); if(newer) need(positive(v.authorityEpoch)); return v;
  }
  function proof(v, control) {
    const newer=v?.schemaVersion==='member-function-proof-v13';
    need(newer||v?.schemaVersion==='member-function-proof-v12');
    exact(v,['schemaVersion'].concat(proofFields,newer?['authorityEpoch','nullableApprovalRef']:[]));
    for(const k of ['personRevision','functionRevision','boardEpoch','termRevision','roleRevision']) need(positive(v[k]));
    for(const k of ['boardId','termId','roleId']) need(id(v[k]));
    need(timestamp(v.startsAt)<timestamp(v.endsAtExclusive)); dto(v.dto);
    need(v.dto.schemaVersion===(newer?'safe-member-function-v13':'safe-member-function-v12')&&v.boardId===v.dto.idDiretoria&&v.termId===v.dto.idTermo);
    if(newer){f4MemberAuthority_().validate(control);need(v.authorityEpoch===control.epoch);
      need(v.nullableApprovalRef===(v.dto.scopeKey===null?approval(control,v.dto.idVigencia).approvalRef:null));}
    return v;
  }
  return {need,id,positive,scopes,roles,cases,dtoFields,indexFields,proofFields,controlFields,versions,exact,timestamp,functionScope,dto,index,proof};
}
// M64-B contracts: separate validator, no legacy capability aliases or public entry point.
function f4MemberAuthzContracts_() {
 const C=f4Core_(),A=f4MemberContracts_();
 const need=(ok,code='M13_AUTHZ_DENIED')=>{if(!ok){const e=new Error(code);e.code=code;throw e}};
 const exact=(v,keys)=>{need(v&&Object.getPrototypeOf(v)===Object.prototype);need(Object.keys(v).length===keys.length&&keys.every(k=>Object.prototype.hasOwnProperty.call(v,k)))};
 const id=A.id,positive=A.positive,scopes=A.scopes;
 const caps=['MEMBERS_SELF_READ','MEMBERS_SELF_UPDATE','MEMBERS_CASE_READ','MEMBERS_PROFILE_CORRECT','MEMBERS_INGRESS_EXECUTE','MEMBERS_DISCIPLINE_EXECUTE','MEMBERS_GOVERNANCE_EXECUTE','MEMBERS_BINDING_MANAGE','MEMBERS_PRIVATE_AUDIT','MEMBERS_VOLUNTARY_EXIT_EXECUTE'];
 const selfFields=[...A.dtoFields,'displayLabel','state','egressEndsAtExclusive','contactEmail'];
 const fields={MEMBERS_SELF_READ:selfFields,MEMBERS_SELF_UPDATE:['contactEmail'],MEMBERS_CASE_READ:['idPessoa','displayLabel','state','currentFunctions'],MEMBERS_PROFILE_CORRECT:['displayName','contactEmail','academicReference'],MEMBERS_INGRESS_EXECUTE:[],MEMBERS_DISCIPLINE_EXECUTE:[],MEMBERS_GOVERNANCE_EXECUTE:[],MEMBERS_BINDING_MANAGE:[],MEMBERS_PRIVATE_AUDIT:['idPessoa','displayLabel','contactEmail','academicReference'],MEMBERS_VOLUNTARY_EXIT_EXECUTE:[]};
 const list=(v,pred,nonempty=false)=>{need(Array.isArray(v)&&(!nonempty||v.length>0)&&new Set(v).size===v.length&&v.every(pred));return v};
 const resource=x=>typeof x==='string'&&/^[A-Za-z][A-Za-z0-9]*\/[A-Za-z0-9_-]{1,48}$/.test(x);
 const timestamp=x=>{need(x&&typeof x.toMillis==='function'&&Number.isSafeInteger(x.toMillis()));return x.toMillis()};
 const live=(v,now)=>need(timestamp(v.issuedAt)<=now&&now<timestamp(v.expiresAt));
 function targets(t){exact(t,['ownPerson','allowedTargetScopes','allowedBoards','allowedRoles','allowedPersons','allowedResources']);need(typeof t.ownPerson==='boolean');list(t.allowedTargetScopes,s=>scopes.includes(s));for(const k of ['allowedBoards','allowedRoles','allowedPersons'])list(t[k],id);list(t.allowedResources,resource);return t}
 function masks(m,capabilities){exact(m,capabilities);for(const cap of capabilities)list(m[cap],x=>fields[cap].includes(x),fields[cap].length>0);return m}
 function policy(p){exact(p,['schemaVersion','id','revision','active','origin','approvalRef','issuedAt','expiresAt','proofKind','capabilities','actorSelector','targets','fieldMasks']);need(p.schemaVersion==='member-authz-policy-v13'&&id(p.id)&&positive(p.revision)&&typeof p.active==='boolean'&&p.origin==='F4_M621_INSTITUTIONAL'&&id(p.approvalRef)&&['SELF','ROLE'].includes(p.proofKind));need(timestamp(p.issuedAt)<timestamp(p.expiresAt));list(p.capabilities,c=>caps.includes(c),true);const s=p.actorSelector;
 if(p.proofKind==='SELF'){exact(s,['kind']);need(s.kind==='SELF'&&p.capabilities.every(c=>['MEMBERS_SELF_READ','MEMBERS_SELF_UPDATE'].includes(c))&&p.targets.ownPerson===true)}else{exact(s,['kind','roleKey','group','actorScope']);need(s.kind==='ROLE'&&A.roles[s.roleKey]&&A.roles[s.roleKey][0]===s.group&&A.roles[s.roleKey][1]===s.actorScope);need(p.capabilities.every(c=>roleCapability(c,s)))}
 targets(p.targets);masks(p.fieldMasks,p.capabilities);return p;
 }
 function roleCapability(cap,s){if(s.group!=='DIRETORIA')return false;return cap==='MEMBERS_CASE_READ'||(['MEMBERS_PROFILE_CORRECT','MEMBERS_INGRESS_EXECUTE'].includes(cap)&&s.actorScope==='ADMINISTRACAO')||(cap==='MEMBERS_GOVERNANCE_EXECUTE'&&['PRESIDENCIA','ADMINISTRACAO'].includes(s.actorScope))||(cap==='MEMBERS_VOLUNTARY_EXIT_EXECUTE'&&s.roleKey==='DIRETOR_RECURSOS_HUMANOS'&&s.actorScope==='RECURSOS_HUMANOS')}
 function grant(g){exact(g,['schemaVersion','id','revision','active','origin','approvalRef','issuerRef','uid','idPessoa','bindingVersion','subjectEpoch','issuedAt','expiresAt','capabilities','actorScopes','targets','fieldMasks','environment','runId']);need(g.schemaVersion==='member-explicit-grant-v13'&&id(g.id)&&positive(g.revision)&&typeof g.active==='boolean'&&id(g.origin)&&g.origin!=='PORTAL_PERMISSOES'&&id(g.approvalRef)&&id(g.issuerRef)&&id(g.uid)&&id(g.idPessoa)&&positive(g.bindingVersion)&&positive(g.subjectEpoch)&&g.environment==='DEV'&&(g.runId===null||id(g.runId)));need(timestamp(g.issuedAt)<timestamp(g.expiresAt));list(g.capabilities,c=>caps.includes(c)&&!['MEMBERS_SELF_READ','MEMBERS_SELF_UPDATE'].includes(c),true);list(g.actorScopes,s=>scopes.includes(s),true);targets(g.targets);masks(g.fieldMasks,g.capabilities);return g}
 function subject(s){exact(s,['schemaVersion','idPessoa','revision','authzEpoch','active','approvalRef','updatedAt']);need(s.schemaVersion==='member-authz-subject-v13'&&id(s.idPessoa)&&positive(s.revision)&&positive(s.authzEpoch)&&typeof s.active==='boolean'&&id(s.approvalRef));timestamp(s.updatedAt);return s}
 function target(t,cap){exact(t,['personId','scopeKey','boardId','roleKey','resources','caseId','purpose','fieldMask','linkId','effect','effectiveDate']);need(id(t.personId));for(const k of ['boardId','roleKey','caseId','purpose','linkId','effect'])need(t[k]===null||id(t[k]));need(t.scopeKey===null||scopes.includes(t.scopeKey));list(t.resources,resource,true);list(t.fieldMask,x=>fields[cap].includes(x),fields[cap].length>0);if(t.effectiveDate!==null)C.civilStart(t.effectiveDate);
 const own=['MEMBERS_SELF_READ','MEMBERS_SELF_UPDATE'].includes(cap);if(own)need(t.scopeKey===null&&t.boardId===null&&t.roleKey===null&&t.caseId===null&&t.purpose===null&&t.linkId===null&&t.effect===null&&t.effectiveDate===null);
 else{need(id(t.caseId)&&id(t.purpose));need(scopes.includes(t.scopeKey)||(cap==='MEMBERS_GOVERNANCE_EXECUTE'&&t.roleKey==='CONSELHEIRO_CONSULTIVO'&&id(t.boardId)));}
 if(cap==='MEMBERS_GOVERNANCE_EXECUTE')need(id(t.boardId)&&id(t.roleKey));
 if(['MEMBERS_INGRESS_EXECUTE','MEMBERS_DISCIPLINE_EXECUTE','MEMBERS_VOLUNTARY_EXIT_EXECUTE'].includes(cap))need(id(t.linkId)&&id(t.effect)&&t.effectiveDate!==null&&t.resources.includes('memberLinks/'+t.linkId));
 return t;
 }
 function request(r){exact(r,['schemaVersion','uid','kind','capability','sourceId','functionId','actorScope','expected','tokenExpiresAt','admissionExpiresAt','target','evidenceRefs']);need(r.schemaVersion==='member-authz-request-v13'&&id(r.uid)&&['SELF','ROLE','EXPLICIT','DEV_ADMIN'].includes(r.kind)&&caps.includes(r.capability)&&id(r.sourceId));need(r.kind==='ROLE'?id(r.functionId):r.functionId===null);need(r.kind==='DEV_ADMIN'?r.actorScope===null:r.kind==='SELF'?r.actorScope===null:scopes.includes(r.actorScope));exact(r.expected,['authorityEpoch','policyEpoch','readAuthzEpoch','subjectEpoch','bindingVersion','personRevision','authorizationRevision']);Object.values(r.expected).forEach(v=>need(positive(v)));timestamp(r.tokenExpiresAt);timestamp(r.admissionExpiresAt);target(r.target,r.capability);list(r.evidenceRefs,id);return r}
 function subset(actual,allowed){need(actual.every(x=>allowed.includes(x)))}
 function match(t,limits,person){targets(limits);if(limits.ownPerson)need(t.personId===person);else need(limits.allowedPersons.includes(t.personId));if(t.scopeKey!==null)need(limits.allowedTargetScopes.includes(t.scopeKey));if(t.boardId!==null)need(limits.allowedBoards.includes(t.boardId));if(t.roleKey!==null)need(limits.allowedRoles.includes(t.roleKey));subset(t.resources,limits.allowedResources)}
 const normalized=v=>v&&typeof v.toMillis==='function'?v.toMillis():Array.isArray(v)?v.map(normalized):v&&typeof v==='object'?Object.fromEntries(Object.entries(v).map(([k,x])=>[k,normalized(x)])):v;
 const digest=v=>C.hash(normalized(v));
 function rehearsal(v){exact(v,['schemaVersion','projectId','environment','database','scriptId','uid','idPessoa','technicalAccount','runId','approvalRef','capabilities','actorScopes','targets','resources','paths','issuedAt','expiresAt','revisions','planHash','rollbackRef']);need(v.schemaVersion==='member-dev-rehearsal-v13'&&v.projectId==='geapa-convergencia-dev'&&v.environment==='DEV'&&v.database==='(default)'&&v.scriptId==='1M3GpuWwyRkV24nbHnopzj1q4J-Ea7zHv5eCG-ngOw0qR5lWDMvvv8UfD'&&id(v.uid)&&v.idPessoa==='PES-000009'&&v.technicalAccount===false&&id(v.runId)&&id(v.approvalRef)&&id(v.rollbackRef));list(v.capabilities,c=>caps.slice(2,9).includes(c),true);need(v.capabilities.length===7);list(v.actorScopes,s=>['PRESIDENCIA','ADMINISTRACAO','COMUNICACAO','EXTENSAO'].includes(s),true);need(v.actorScopes.length===4);targets(v.targets);list(v.resources,resource,true);list(v.paths,resource,true);need(timestamp(v.issuedAt)<timestamp(v.expiresAt));exact(v.revisions,['sources','binding','subject','grant','control']);for(const k of ['binding','subject','grant','control'])need(positive(v.revisions[k]));need(v.revisions.sources&&Object.keys(v.revisions.sources).length>0);for(const [p,n]of Object.entries(v.revisions.sources))need(resource(p)&&positive(n));need(v.planHash===digest(Object.fromEntries(Object.entries(v).filter(([k])=>k!=='planHash'))));return v}
 return {need,exact,id,positive,scopes,caps,fields,selfFields,list,resource,timestamp,live,targets,masks,policy,grant,subject,target,request,match,subset,roleCapability,digest,normalized,rehearsal};
}
