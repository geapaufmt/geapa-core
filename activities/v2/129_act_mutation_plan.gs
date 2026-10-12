// ACT owns path validation and source effects. Shared Core seals and commits the final plan.
function f4ActivitiesMutationPlan_(tx,id,writes,response,now,deadline){
 const C=f4Core_(),F=f4Firestore_(),N=C.need,M=f4ActivitiesModel_(),control=f4MemberAuthority_().validate(tx.get('memberAuthorityControl/MEMBERS').data());
 N((F.inspectTransaction(tx).environment==='SYNTHETIC_LOCAL' || f4DevBoundary_().permit(F.inspectTransaction(tx)))&&control.mode==='CANONICAL'&&control.activeWriterVersion==='M13'&&C.id(id)&&Array.isArray(writes)&&Number.isSafeInteger(now)&&now<deadline,'M13_PLAN_INVALID');
 const plan={schemaVersion:'core-mutation-plan-v13',aggregateId:id,kind:'ACT',expectedRevisions:[],readSet:[],sourceWrites:[],headWrites:[],claimWrites:[],projectionWrites:[],technicalWrites:[],deletes:[],revisionBefore:[],revisionAfter:[],fences:{authorityEpoch:control.epoch,readAuthzEpoch:control.readAuthzEpoch},auditFacts:{sourcePaths:[],mode:'M13'},footprint:{}};
 for(const w of writes){N(['set','delete'].includes(w.kind)&&M.validPath(w.path)&&!w.path.startsWith('command')&&!w.path.startsWith('member'),'M13_PLAN_INVALID');(w.kind==='delete'?plan.deletes:plan.sourceWrites).push(w);plan.auditFacts.sourcePaths.push(w.path);}
 plan.commandType=response.commandType;
 if(response.businessState==='COMMIT_APPLIED')plan.projectionWrites.push(...f4DevReadTiming_().stage('invalidation',()=>f4PortalProjectionInvalidation_(tx,writes,response.operationId,now)));
 return f4TransactionPlans_().finalize(tx,plan,response,now,deadline);
}
