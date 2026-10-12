// Administrative confirmation on the existing cycle dialog; no browser-supplied
// eligibility, identity state, population manifest or grant is trusted.
function f4ActivitiesIncidenceCompat_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),db=f4Firestore_().connect(options.config),ports=options.ports,clock=options.clock||Date.now;
 function context(id){N(M.id(id),'VALIDATION_ERROR');const cycleRef=M.ref('cycle',id),auth=ports.readAuthority([M.path(cycleRef)],false,'PRESENTATION_BOARD_READ','PRESC_'+id);const snapshot=f4ActivitiesIncidenceSource_({db,config:options.config,clock}).scan(cycleRef);return {cycleRef,auth,...snapshot};}
 function recheck(z,get=p=>db.get(p)){f4SharedAuthz_().recheck(get,z.auth.q,z.auth.proof,clock());for(const b of z.basis)N(get(M.path(b.ref))?.revision===b.revision,'STALE_REVISION','INCIDENCE_SOURCE_CHANGED');N(clock()<z.validUntil,'STALE_REVISION','INCIDENCE_SOURCE_EXPIRED');}
 function read(input){C.exact(input,['cycleId']);const z=context(input.cycleId);const items=z.items.map(i=>({...i,name:db.get(M.path(i.personRef))?.displayName||i.personRef.id}));recheck(z);return {ok:true,data:{cycleRevision:z.cycle.revision,items,ready:items.every(i=>i.disposition!=='REVIEW_REQUIRED'&&(i.disposition!=='INCIDENT'||i.obligationRef))}};}
 function prepare(input){
  C.exact(input,['operationId','route','payload']);N(input.route==='/admin/apresentacoes/incidencia/salvar','VALIDATION_ERROR');const p=input.payload;
  N(['EVALUATE_PERSON','COVERAGE_APPROVE'].includes(p.kind),'VALIDATION_ERROR');C.exact(p,['kind','cycleId','cycleRevision','reason',...(p.kind==='EVALUATE_PERSON'?['personId','personRevision']:[])]);N(typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'INCOMPLETE_REQUEST');
  const z=context(p.cycleId);N(z.cycle.revision===p.cycleRevision,'STALE_REVISION');let items=z.items,command;
  const ref=M.ref('assessment',M.hash({operationId:input.operationId,kind:p.kind}).slice(0,32));
  if(p.kind==='EVALUATE_PERSON'){
   const item=items.find(i=>i.personRef.id===p.personId);N(item&&item.personRevision===p.personRevision,'STALE_REVISION','INCIDENCE_PERSON_CHANGED');N(item.disposition==='INCIDENT'&&!item.obligationRef,'INVALID_STATE_TRANSITION','INCIDENCE_NOT_PENDING');items=[item];
   command={commandType:'ACT.INCIDENCE_EVALUATE',payload:{cycleRef:z.cycleRef,personRef:item.personRef,membersEvidenceRef:M.ref('memberState',p.personId),capacityAssessmentRef:ref,decisionReason:p.reason.trim()}};
  }else{
   N(items.every(i=>i.disposition!=='REVIEW_REQUIRED'&&(i.disposition!=='INCIDENT'||i.obligationRef)),'HUMAN_DECISION_REQUIRED','INCIDENCE_EVALUATION_REQUIRED');
   command={commandType:'ACT.INCIDENCE_COVERAGE_APPROVE',payload:{cycleRef:z.cycleRef,sourceManifestRef:ref,reason:p.reason.trim()}};
  }
  const document={id:ref.id,revision:1,schemaVersion:'act-incidence-candidate-v1',producer:'CORE_INCIDENCE_SOURCE',cycleRef:z.cycleRef,configRef:z.cycle.currentConfigRef,state:'SEALED',complete:true,coverageReady:p.kind==='COVERAGE_APPROVE',items,digest:f4ActivitiesCapacityDigest_(items),basisRevisionVector:z.basis,sourceHeadsDigest:z.headsDigest,recordedAt:M.stamp(clock()),validUntil:M.stamp(z.validUntil)};
  const authorized=ports.prepare({operationId:input.operationId,...command});ports.authorizePrepared(p=>db.get(p),authorized);
  db.transaction(tx=>{const get=p=>tx.get(p).data();ports.authorizePrepared(get,authorized);recheck(z,get);const old=get(M.path(ref));if(old)N(old.digest===document.digest&&old.sourceHeadsDigest===document.sourceHeadsDigest,'CONFLICT');else tx.set(M.path(ref),document);});
  return ports.prepare({operationId:input.operationId,...command});
 }
 return {read,prepare};
}
