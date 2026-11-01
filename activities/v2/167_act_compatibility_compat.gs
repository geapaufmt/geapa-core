function f4ActivitiesCompatibilityCompat_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),db=f4Firestore_().connect(options.config),ports=options.ports,clock=options.clock||Date.now,calendar=f4ActivitiesCalendarCompat_(options);
 function context(id){const z=calendar.context(id),session=z.assignment?.sessionRef&&db.get(M.path(z.assignment.sessionRef));return {...z,session};}
 function read(input){
  C.exact(input,['idApresentacao',...('cursor'in input?['cursor']:[])]);const z=context(input.idApresentacao),s=z.session,enabled=!!(z.cycle.status==='OPEN'&&z.cfg.parameters.shared&&s?.slots.length===2),cache=new Map(),get=p=>{if(!cache.has(p))cache.set(p,db.get(p));return cache.get(p);},basis=enabled?f4ActivitiesCompatibilityBasis_(s,get):null;
  let ref=input.cursor?M.ref('compatibility',input.cursor):s?.compatibilityRef;const history=[],seen=new Set();
  while(ref&&seen.size<20){N(ref.kind==='compatibility'&&!seen.has(ref.id),'UNRESOLVED_DEPENDENCY');seen.add(ref.id);const v=get(M.path(ref));N(v?.sessionRef?.id===s?.id,'AUTHORIZATION_DENIED');if(z.board||v.units?.some(u=>u.id===z.u.id))history.push({id:v.id,resultado:v.result,motivo:v.reason,registradaEm:M.ms(v.recordedAt),atual:enabled&&f4ActivitiesCompatibilityCurrent_(v,s,get)});ref=v.predecessorRef;}
  const request=z.u.currentCompatibilityRequestRef&&get(M.path(z.u.currentCompatibilityRequestRef)),guide=z.board&&basis?get(M.path(basis.contextRef))?.text||'':'';calendar.recheck(z);
  if(s)N(get(M.path(M.ref('session',s.id)))?.revision===s.revision,'STALE_REVISION');
  // Reuse the same basis across history entries, then validate all dependencies
  // in one fresh batch. Caching must not disguise a concurrent thematic change.
  const paths=[...cache.keys()],fresh=db.getMany(paths),F=f4Firestore_();
  paths.forEach((p,i)=>N(C.hash(F.encode(fresh[i]||null))===C.hash(F.encode(cache.get(p)||null)),'STALE_REVISION','COMPATIBILITY_READ_CHANGED'));
  return {ok:true,data:{guia:guide,unitRevision:z.u.revision,sessionRevision:s?.revision||0,habilitada:enabled,podeSolicitar:enabled&&!z.board,podeRegistrar:enabled&&z.board,basisDigest:basis?f4ActivitiesCompatibilityDigest_(basis):'',conteudos:z.board&&basis?basis.units.map(b=>({unitId:b.unitRef.id,titulo:b.content?.titleOrObject||'',resumo:b.content?.summary||'',estado:b.state})):[],historico:history,proximaPagina:ref?.id||null,solicitacao:request?{estado:request.state,motivo:request.reason}:null}};
 }
 function prepare(input){
  C.exact(input,['operationId','route','payload',...('basis'in input?['basis']:[])]);N(input.route==='/v2/apresentacoes/compatibilidade','VALIDATION_ERROR');const p=input.payload;C.exact(p,['idApresentacao','unitRevision','sessionRevision','action','reason',...(p.action==='DECIDE'?['outcome','basisDigest']:[])]);
  const z=context(p.idApresentacao),s=z.session;N(z.u.revision===p.unitRevision&&s?.revision===p.sessionRevision,'STALE_REVISION','DISPLAYED_COMPATIBILITY_CHANGED');N(z.cycle.status==='OPEN'&&z.cfg.parameters.shared&&s.slots.length===2,'CONFIG_CAPABILITY_DISABLED');let command;
  if(p.action==='REQUEST'){N(!z.board,'AUTHORIZATION_DENIED');command={commandType:'ACT.COMPATIBILITY_REQUEST',payload:{unitRef:M.ref('unit',z.u.id),reason:p.reason}};}
  else{N(p.action==='DECIDE'&&z.board,'AUTHORIZATION_DENIED');const b=f4ActivitiesCompatibilityBasis_(s,path=>db.get(path));N(p.basisDigest===f4ActivitiesCompatibilityDigest_(b),'STALE_REVISION','DISPLAYED_COMPATIBILITY_CHANGED');command={commandType:'ACT.COMPATIBILITY_DECIDE',payload:{sessionRef:M.ref('session',s.id),twoUnitRefs:b.units.map(v=>v.unitRef),thematicRevisionRefs:b.units.map(v=>v.thematicRevision||{kind:'NONE'}),outcome:p.outcome,reason:p.reason,basisDigest:p.basisDigest}};}
  calendar.recheck(z);const request=ports.prepare({operationId:input.operationId,...command});N(db.get('presentationSessions/'+s.id)?.revision===s.revision,'STALE_REVISION');return request;
 }
 return {read,prepare};
}
