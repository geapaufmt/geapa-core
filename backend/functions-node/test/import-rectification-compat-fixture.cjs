'use strict';
const {setup}=require('./people-fixture.cjs'),{Timestamp}=require('firebase-admin/firestore'),E=require('../src/environment.cjs');
function fixture(){
 const h=setup('DIRETOR_ADMINISTRATIVO',1),personId='PERSON_000',operationId='a'.repeat(32),approvalId='SYN_IMPORT_APPROVAL',p=h.policy;
 p.capabilities=['MEMBERS_PROFILE_CORRECT'];p.fieldMasks={MEMBERS_PROFILE_CORRECT:['displayName']};p.targets={...p.targets,allowedTargetScopes:['ADMINISTRACAO'],allowedPersons:[],allowedResources:[],commandScopes:['PROFILE_CORRECTION','PROFILE_CORRECTION_CURRENT_MEMBERS']};
 const person=h.docs['memberIdentities/'+personId];person.origin='PESSOAS_BASE';person.active=false;person.revision=1;
 const hash=path=>JSON.parse(h.domain.run('(p)=>f4Core_().hash(f4Firestore_().encode(f4Firestore_().connect(nodeConfig).get(p)))',path));
 const item={personId,operationId,personRevision:1,identityRevision:1,identityHash:hash('memberIdentities/'+personId),aggregateHash:hash('memberAggregateHeads/'+personId),sourceSnapshotHash:'a'.repeat(64),sourceRow:2,sourceActive:null,sourceStatus:null,cause:'INCOMPLETE_SOURCE_ACTIVE',evidenceId:'SYN_IMPORT_EVIDENCE',eligibleState:'MEMBRO_EFETIVO'};
 const now=h.now,batch={schemaVersion:'member-import-rectification-approval-v1',id:approvalId,revision:1,active:true,projectId:E.PROJECT,environment:E.profile.environment,approvalRef:'SYN_DOCUMENTARY_APPROVAL',approverUid:'CF_ADMIN',policyId:p.id,issuedAt:Timestamp.fromMillis(now-1000),expiresAt:Timestamp.fromMillis(now+3600000),cases:[item]};
 h.docs['memberImportRectificationApprovals/'+approvalId]=batch;
 const itemHash=JSON.parse(h.domain.run('(x)=>f4Core_().hash(f4Firestore_().encode(x))',item));
 h.docs['memberEvidenceAttestations/'+item.evidenceId]={schemaVersion:'member-evidence-attestation-v13',id:item.evidenceId,revision:1,active:true,kind:'CADASTRAL_EVIDENCE',holderId:personId,linkId:null,targetResourceIds:['memberIdentities/'+personId],effect:'PROFILE_CORRECTION',effectiveDate:new Intl.DateTimeFormat('en-CA',{timeZone:'America/Cuiaba',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date(now-1000)),issuerRef:'SYN_AUTHORITY',approvalRef:batch.approvalRef,contentHash:itemHash,verifiedAt:Timestamp.fromMillis(now-1000),expiresAt:batch.expiresAt};
 require('./production-fixture-profile.cjs')(h.docs);
 h.input={approvalId,personId,operationId};h.resolve=()=>JSON.parse(h.domain.run('(x)=>f4MemberImportRectification_().resolve(p=>f4Firestore_().connect(nodeConfig).get(p),x.input,x.now)',{input:h.input,now}));return {...h,batch,item,person};
}
module.exports={fixture};
