// Synthetic IDs only. CaseRefs are the three sanitized institutional approval references.
function m13Fixture_(){
 const X=fixture,K=f4MemberContracts_(),T=f4Firestore_().Timestamp;
 const d=X.role(X.base());delete d['memberAuthBindings/SYN_UID'];delete d['memberRoleCatalog/M1_POLICY'];
 delete d['memberFunctions/VIG-001'];
 d['memberRoleCatalog/CARGO-001'].roleKey='CONSELHEIRO_CONSULTIVO';d['memberRoleCatalog/CARGO-001'].group='CONSELHO';
 const control={schemaVersion:'member-authority-control-v13',id:'MEMBERS',revision:1,mode:'PAUSED',epoch:1,policyEpoch:1,readAuthzEpoch:1,activeWriterVersion:'NONE',approvalRef:'M621-D1',updatedAt:T.fromMillis(X.NOW),nullableFunctionApprovals:{},selfReadPolicy:{policyId:'SYN_SELF',policyRevision:1,capability:'MEMBERS_SELF_READ',eligibleStates:['MEMBRO_EFETIVO'],fieldMasks:{MEMBERS_SELF_READ:[...K.dtoFields]},issuedAt:T.fromMillis(X.NOW-60000),expiresAt:T.fromMillis(X.NOW+60000)}};
 for(let i=1;i<=3;i++){
   const person='PES-00000'+i,link='VINC-00'+i,id='SYN_CONS_'+i;
   d['memberIdentities/'+person]=X.entity('memberIdentities',person,{idPessoa:person,displayName:'Synthetic council '+i,technicalAccount:false,active:true});
   d['memberLinks/'+link]=X.entity('memberLinks',link,{idPessoa:person,type:'MEMBRO_EFETIVO',startDate:'2026-01-01',endDate:null,exitDate:null});
   d['memberFunctions/'+id]=X.entity('memberFunctions',id,{schemaVersion:'m1-memberFunctions-v2',idPessoa:person,idVinculo:link,idCargo:'CARGO-001',scopeKey:null,idDiretoria:'DIR-001',idTermo:'TERM-001',startDate:'2026-01-01',endDatePlanned:'2026-12-31',endDateReal:null});
   control.nullableFunctionApprovals[id]={caseRef:K.cases[i-1],approvalRef:'M621-D1',evidenceHash:'a'.repeat(64)};
 }
 return {d,control,now:X.NOW,dep:{boards:{'DIR-001':{revision:1}}}};
}
