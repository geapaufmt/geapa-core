// Sanitized Portal DTOs and private proofs. No service calls; authority is the M1.1 derivation.
function f4MemberPortalViews_(){const M=f4MemberModel_(true),C=f4Core_();
 function build(d,id,rev,now,dep,derived){const out={},state=derived['memberCurrentStates/'+id],functions=derived['memberCurrentFunctions/'+id];out['portalAccessStatus/'+id+'/state/current']={schemaVersion:'portal-access-status-v1',status:state.blocked?'REVOKED':state.state==='EGRESSO'?'EGRESS':'ACTIVE',updatedAt:state.asOf,validUntil:state.validUntil};
 for(const kind of ['memberCurrentFunctions','memberCurrentDirectors'])out[kind+'/'+id+'/index/current']={schemaVersion:'member-function-index-v12',idPessoa:id,kind,ids:derived[kind+'/'+id].records.map(f=>f.idVigencia),sourceRevision:rev,asOf:functions.asOf,validUntil:functions.validUntil};
 for(const f of functions.records){const source=d['memberFunctions/'+f.idVigencia],board=d['memberGovernanceBoards/'+f.idDiretoria],term=d['memberGovernanceTerms/'+f.idTermo],role=d['memberRoleCatalog/'+f.idCargo],relation=M.rows(d,'memberGovernanceBoardTerms').find(r=>r.idDiretoria===f.idDiretoria&&r.idTermo===f.idTermo&&M.interval(r.startDate,r.endDate).a<=now&&now<M.interval(r.startDate,r.endDate).z);const dto={schemaVersion:'safe-member-function-v12',idPessoa:id,idVigencia:f.idVigencia,roleKey:f.roleKey,group:f.group,scopeKey:f.scopeKey,idDiretoria:f.idDiretoria,idTermo:f.idTermo,startDate:f.startDate,endDatePlanned:f.endDatePlanned,endDateReal:f.endDateReal};
 const end=Math.min(C.ms(state.validUntil),M.interval(source.startDate,source.endDateReal||source.endDatePlanned).z,M.interval(board.startDate,board.endDate).z,M.interval(term.startDate,term.endDate).z,M.interval(relation.startDate,relation.endDate).z);
 out['memberFunctionProofs/'+id+'/items/'+f.idVigencia]={schemaVersion:'member-function-proof-v12',personRevision:rev,functionRevision:source.revision,boardId:f.idDiretoria,boardEpoch:dep.boards[f.idDiretoria].revision,termId:f.idTermo,termRevision:term.revision,roleId:f.idCargo,roleRevision:role.revision,startsAt:M.stamp(now),endsAtExclusive:M.stamp(end),dto};out['memberCurrentFunctions/'+id+'/items/'+f.idVigencia]=dto;if(f.group==='DIRETORIA')out['memberCurrentDirectors/'+id+'/items/'+f.idVigencia]=dto;}
 return out;}
 function project(d,id,rev,now,dep,derived){
   M.need(!M.rows(d,'memberFunctions').some(f=>f.schemaVersion==='m1-memberFunctions-v2'),'M13_EXPLICIT_VIEW_REQUIRED');
   return build(d,id,rev,now,dep,derived);
 }
 // Pure successor; computes its own derivation and never calls repository/policyPath/grantPath.
 function projectM13(d,id,rev,now,dep,control){
   const K=f4MemberContracts_(),A=f4MemberAuthority_().validate(control),N=f4MemberModel_(true,A);
   const r=N.derive(d,id,rev,now),derived={};
   derived['memberCurrentStates/'+id]=r.state;
   for(const kind of ['memberCurrentFunctions','memberCurrentDirectors']) derived[kind+'/'+id]={records:kind==='memberCurrentFunctions'?r.currentFunctions:r.currentDirectors,asOf:r.state.asOf,validUntil:r.state.validUntil};
   K.need(dep&&dep.boards&&Number.isSafeInteger(now));
   for(const f of r.currentFunctions) K.need(K.positive(dep.boards[f.idDiretoria]?.revision));
   const out=build(d,id,rev,now,dep,derived);
   for(const [path,v] of Object.entries(out)){
     if(v.schemaVersion==='member-function-index-v12'){v.schemaVersion='member-function-index-v13';v.authorityEpoch=A.epoch;K.index(v);}
     if(v.schemaVersion==='safe-member-function-v12')v.schemaVersion='safe-member-function-v13';
     if(v.schemaVersion==='member-function-proof-v12'){
       v.schemaVersion='member-function-proof-v13';v.authorityEpoch=A.epoch;
       v.dto.schemaVersion='safe-member-function-v13';
       v.nullableApprovalRef=v.dto.scopeKey===null?A.nullableFunctionApprovals[v.dto.idVigencia].approvalRef:null;
       K.proof(v,A);
     }
     K.need(!path.split('/').includes('null'));
   }
   return out;
 }
 return {project,projectM13};}
