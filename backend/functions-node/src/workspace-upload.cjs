'use strict';
// Async orchestration only. Each phase delegates validation/state decisions to
// f4ActivitiesWorkspaceUpload_.phases; provider I/O never runs in a transaction.
// This factory is server-internal. No HTTP body can supply any of these ports.
const {fail}=require('./environment.cjs');
function coordinator({phase,system,providerFor,assertOwnership,clock=Date.now}){
 for(const port of [phase,system,providerFor,assertOwnership])if(typeof port!=='function')fail('CONFIG_CAPABILITY_DISABLED');
 const run=async(name,...args)=>{await assertOwnership(args[0]);return phase(name,...args);};
 const dispatch=async(ref,plan)=>{await assertOwnership(ref);await system(plan);};
 const final=async ref=>{const {out}=await run('source',ref);return {state:out.status,effectRef:ref,...(out.artifactRef?{artifactRef:out.artifactRef}:{})};};
 async function observe(ref,out,result){const plan=await run('observePlan',ref,out,result);if(plan.result)return plan.result;await dispatch(ref,plan.command);return final(ref);}
 async function reconcile(ref){
  const {out,policy}=await run('source',ref);
  if(out.status==='CONFIRMED')return final(ref);
  if(!['CLAIMED','UNKNOWN'].includes(out.status))fail('INVALID_STATE_TRANSITION');
  const plan=await run('reservedPlan',ref,out),provider=await providerFor(policy.workspace);
  await assertOwnership(ref);const result=await provider.inspect(plan);
  const until=out.leaseUntil?.toMillis?.()??Date.parse(out.leaseUntil?.timestampValue);
  if(result.outcome==='UNKNOWN'&&out.status==='CLAIMED'&&until>clock())return {state:'PENDING',effectRef:ref};
  if(out.status==='CLAIMED')await run('attempt',ref);
  return observe(ref,out,result);
 }
 async function upload(ref,bytes){
  const input=await run('reservationInput',ref,bytes),provider=await providerFor(input.first.policy.workspace);
  await assertOwnership(ref);
  const fileId=input.first.out.workspaceReservation?.fileId||await provider.reserveId();
  const out=await run('reservationCommit',ref,input,fileId);
  if(out.status==='CONFIRMED')return final(ref);
  if(out.status==='UNKNOWN')return reconcile(ref);
  if(['PENDING','KNOWN_NO_EFFECT'].includes(out.status))await dispatch(ref,await run('commandPlan',ref,'OUTBOX_CLAIM',{effectRef:ref,expectedAttemptState:out.status}));
  const attempt=await run('attempt',ref);if(!attempt.execute)return reconcile(ref);
  const fresh=await run('source',ref),currentProvider=await providerFor(fresh.policy.workspace),plan=await run('reservedPlan',ref,attempt.out);
  await assertOwnership(ref);let result;
  try{result=await currentProvider.upload(plan,bytes);}catch(e){
   if(['DESTINATION_INVALID','VALIDATION_ERROR','CONFIG_CAPABILITY_DISABLED'].includes(e.code))throw e;
   result={outcome:'UNKNOWN',reason:'PROVIDER_RESULT_UNRESOLVED'};
  }
  return observe(ref,attempt.out,result);
 }
 return Object.freeze({upload,reconcile});
}
module.exports={coordinator};
