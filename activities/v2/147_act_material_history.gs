// User-facing material history uses domain facts, never the technical command ledger.
function f4ActivitiesMaterialHistory_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),db=F.connect(options.config),clock=options.clock||Date.now;
 const collections={submission:'presentationMaterialSubmissions',review:'presentationMaterialReviews',receipt:'presentationMaterialReceipts',opportunity:'presentationCorrectionOpportunities',regularization:'presentationMaterialRegularizations'};
 function read(input){
  C.exact(input,['idApresentacao',...('cursors'in input?['cursors']:[])]);N(typeof input.idApresentacao==='string'&&input.idApresentacao.startsWith('APU_')&&M.id(input.idApresentacao.slice(4)),'VALIDATION_ERROR');
  const path='presentationUnitHeads/'+input.idApresentacao.slice(4),u=db.get(path);N(u?.presentationRef,'UNRESOLVED_DEPENDENCY');let auth;
  try{auth=options.ports.readAuthority([path],true,'PRESENTATION_MATERIAL_READ_SELF');}catch(e){if(!['AUTHORIZATION_DENIED','AUTH_BINDING_INVALID','ACCESS_REVOKED','M13_AUTHZ_DENIED'].includes(e.code))throw e;auth=options.ports.readAuthority([path],false,'PRESENTATION_BOARD_READ','PRESC_'+u.cycleRef.id);}
  const pr=db.get(M.path(u.presentationRef)),rq=pr?.materialRequirementRef&&db.get(M.path(pr.materialRequirementRef));N(pr?.personRef?.id===u.personRef.id&&pr.cycleRef?.id===u.cycleRef.id&&rq?.presentationRef?.id===pr.id&&rq.personRef?.id===u.personRef.id,'UNRESOLVED_DEPENDENCY');
  const cursors=input.cursors||{};N(cursors&&Object.getPrototypeOf(cursors)===Object.prototype&&Object.keys(cursors).every(k=>Object.hasOwn(collections,k)&&cursors[k]!==undefined&&(cursors[k]===null||M.id(cursors[k]))),'VALIDATION_ERROR');
  const items=[],nextCursors={},date=t=>Number.isFinite(M.ms(t))?new Date(M.ms(t)).toISOString():'';
  for(const [kind,collection] of Object.entries(collections)){
   if(Object.hasOwn(cursors,kind)&&cursors[kind]===null){nextCursors[kind]=null;continue;}
   const prior=cursors[kind],rows=db.query({collection,filters:[{field:'presentationRef',op:'EQUAL',value:u.presentationRef}],order:[{field:'__name__',direction:'ASCENDING'}],limit:20,...(prior?{cursor:[new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+collection+'/'+prior)]}:{})});
   for(const {data:d} of rows){N(d.presentationRef?.id===pr.id&&d.personRef?.id===u.personRef.id&&d.cycleRef?.id===u.cycleRef.id,'UNRESOLVED_DEPENDENCY');
    const when=kind==='submission'?d.submittedAt:kind==='review'?d.reviewedAt:kind==='receipt'?d.receivedAt:kind==='regularization'?d.regularizedAt:d.createdAt||d.recordedAt||d.updatedAt;
    items.push({id:kind+'_'+d.id,tipo:{submission:d.predecessorSubmissionRef?'MATERIAL_REENVIADO':'MATERIAL_ENVIADO',review:'MATERIAL_ANALISADO',receipt:'RECEBIMENTO_DEFINITIVO',opportunity:'OPORTUNIDADE_CORRECAO',regularization:'MATERIAL_REGULARIZADO'}[kind],data:date(when),descricao:kind==='submission'?d.notes||'':kind==='review'?d.reason||'':kind==='opportunity'?d.reason||'':kind==='regularization'?d.reason||'':'',...(kind==='review'?{resultado:d.outcome}:{}),...(kind==='opportunity'?{estado:d.status,prazo:date(d.dueAt),cienciaEm:date(d.notifiedAt)}:{}),...(kind==='receipt'?{basePrazo:d.deadlineComplianceBasis?.kind||''}:{}),...(kind==='regularization'?{fundamento:d.basisKind||''}:{})});
   }
   nextCursors[kind]=rows.length===20?rows[rows.length-1].data.id:null;
  }
  f4SharedAuthz_().recheck(p=>db.get(p),auth.q,auth.proof,clock());const fresh=db.getMany([path,M.path(u.presentationRef),M.path(pr.materialRequirementRef)]);[u,pr,rq].forEach((d,i)=>N(C.hash(F.encode(d))===C.hash(F.encode(fresh[i])),'STALE_REVISION','MATERIAL_HISTORY_CHANGED'));
  items.sort((a,b)=>b.data.localeCompare(a.data)||a.id.localeCompare(b.id));return {ok:true,data:{itens:items,proximos:Object.values(nextCursors).some(Boolean)?nextCursors:null}};
 }
 return {read};
}
