// E K2 pure saga evaluator. It never invokes a child, Auth, HTTP or a second transaction.
function f4MemberOnboarding_(){
 if(f4MemberOnboarding_.value_)return f4MemberOnboarding_.value_;
 const C=f4Core_(),F=f4Firestore_(),K=f4MemberCommandCatalog_(),L=f4MemberLedgerM13_(),N=C.need,T=F.Timestamp.fromMillis;
 const stages=['INGRESS_PENDING','BINDING_PENDING','COMPLETED','RECONCILIATION_REQUIRED'];
 const childId=(parent,stage)=>{N(C.op(parent)&&['C','J'].includes(stage),'COMMAND_PAYLOAD_INVALID');return C.hash({parentOperationId:parent,stage}).slice(0,32)};
 function validate(r){N(r.commandType==='members.onboarding.complete','COMMAND_TYPE_INVALID');C.exact(r.target,['personId']);N(C.id(r.target.personId)&&r.expectedRevisions.length===0&&r.decisionRef===null&&r.evidenceRefs.length===0,'COMMAND_PAYLOAD_INVALID');C.exact(r.payload,['childOperations']);C.exact(r.payload.childOperations,['C','J']);
  for(const stage of ['C','J']){const v=r.payload.childOperations[stage];C.exact(v,['operationId','digest','receiptRef']);N(v.operationId===childId(r.operationId,stage)&&K.hex(v.digest)&&v.receiptRef==='commandReceipts/'+v.operationId,'COMMAND_PAYLOAD_INVALID');}return true;
 }
 function run(v,r){C.exact(v,['schemaVersion','parentOperationId','ownerUid','targetPersonId','revision','childOperations','stage','createdAt','updatedAt']);N(v.schemaVersion==='member-onboarding-run-v13'&&v.parentOperationId===r.operationId&&v.ownerUid===r.ownerUid&&v.targetPersonId===r.target.personId&&C.rev(v.revision)&&v.revision>0&&K.equal(v.childOperations,r.payload.childOperations)&&stages.includes(v.stage)&&Number.isSafeInteger(C.ms(v.createdAt))&&C.ms(v.createdAt)<=C.ms(v.updatedAt),'OPERATION_CONFLICT');return v;}
 function evaluate(get,r,authRequests,authProofs,now,config){validate(r);const projectId=config.projectId;N(authRequests.length===2&&authProofs.length===2,'AUTH_BINDING_INVALID');const old=get('memberOnboardingRuns/'+r.operationId);if(old)run(old,r);
  const children={},outcomes={};let inconsistent=false;
  for(const [stage,type]of [['C','members.ingress.execute'],['J','members.binding.associate']]){const ref=r.payload.childOperations[stage],key={operationId:ref.operationId,digest:ref.digest,commandType:type};try{outcomes[stage]=L.recover(get,key,r.ownerUid,now);}catch(e){if(e.code!=='RECONCILIATION_REQUIRED')throw e;inconsistent=true;}
   const request=get('commandRequests/'+ref.operationId);if(request){const validator=stage==='C'?f4MemberIngressCommands_().validate:f4MemberBindingCommands_(config).validate;K.envelope(projectId,request,validator);N(request.ownerUid===r.ownerUid&&request.digest===ref.digest&&request.target.personId===r.target.personId,'AUTH_BINDING_INVALID');children[stage]=request;}
   else N(old&&outcomes[stage],'RECONCILIATION_REQUIRED');
  }
  const A=f4MemberAuthz_(),proofs=[];
  for(let i=0;i<2;i++){const q=authRequests[i],proof=authProofs[i];N(q.uid===r.ownerUid&&q.target.personId===r.target.personId&&q.capability===K.types[r.commandType][i]&&proof.capability===q.capability,'AUTH_BINDING_INVALID');let next;
   if(i===0)next=get('memberLinks/'+q.target.linkId)?A.recheck(get,q,proof,now):A.recheckCommand(get,q,proof,now,children.C);
   else{N(children.J||outcomes.J,'AUTH_BINDING_INVALID');if(children.J&&!outcomes.J)f4MemberActContent_().association(get,children.J,now,projectId);if(get('memberAggregateHeads/'+q.target.personId))next=A.recheck(get,q,proof,now);else{N(children.C&&children.J,'AUTH_BINDING_INVALID');next=A.authorizeOnboardingBinding(get,q,children.C,children.J,now).proof;const stable=v=>Object.fromEntries(Object.entries(v).filter(([key])=>!['issuedAt','expiresAt'].includes(key)));N(C.ms(proof.issuedAt)<=now&&now<C.ms(proof.expiresAt)&&K.equal(stable(proof),stable(next))&&C.ms(next.expiresAt)>=C.ms(proof.expiresAt),'AUTH_BINDING_INVALID');}}
   proofs.push(next);
  }
  // No child is run here, and a completed child can never be returned to pending.
  let stage=inconsistent||Object.values(outcomes).some(v=>v&&v.state!=='COMPLETED')?'RECONCILIATION_REQUIRED':!outcomes.C?'INGRESS_PENDING':!outcomes.J?'BINDING_PENDING':'COMPLETED';
  if(outcomes.J&&!outcomes.C||old?.stage==='BINDING_PENDING'&&stage==='INGRESS_PENDING'||['COMPLETED','RECONCILIATION_REQUIRED'].includes(old?.stage)&&stage!==old.stage)stage='RECONCILIATION_REQUIRED';
  const value=old?.stage===stage?old:{schemaVersion:'member-onboarding-run-v13',parentOperationId:r.operationId,ownerUid:r.ownerUid,targetPersonId:r.target.personId,revision:(old?.revision||0)+1,childOperations:r.payload.childOperations,stage,createdAt:old?.createdAt||T(now),updatedAt:T(now)};
  const changed=!old||!K.equal(old,value);let stageAudit=null;
  if(changed){const path='memberAudit/'+C.hash({parentOperationId:r.operationId,revision:value.revision}).slice(0,32);N(!get(path),'OPERATION_CONFLICT');stageAudit={path,kind:'set',data:{schemaVersion:'member-onboarding-stage-audit-v13',event:'ONBOARDING_STAGE_CHANGED',parentOperationId:r.operationId,ownerUid:r.ownerUid,revision:value.revision,fromStage:old?.stage||null,toStage:stage,capabilities:K.types[r.commandType],proofDigest:C.hash(F.encode(proofs)),createdAt:T(now)}};}
  return {run:run(value,r),changed,stageAudit,proofs,terminal:['COMPLETED','RECONCILIATION_REQUIRED'].includes(stage)};
 }
 function resolve(){C.fail('COMMAND_TYPE_INVALID');}
 return (f4MemberOnboarding_.value_={childId,validate,run,evaluate,resolve});
}
