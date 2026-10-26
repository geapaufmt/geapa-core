'use strict';
const {PROJECT,EXECUTOR,fail}=require('./environment.cjs'),G=require('./workspace-storage.cjs');
const CONTROL='actStorageControl/current';
function check(get,uid,expiresAt){
 const s=get(CONTROL),ids=(v,max)=>Array.isArray(v)&&v.length<=max&&v.every(x=>typeof x==='string'&&/^[A-Za-z0-9_-]{1,128}$/.test(x));
 if(!Number.isSafeInteger(expiresAt)||expiresAt<=Date.now()||!s||s.projectId!==PROJECT||s.enabled!==true||s.syntheticOnly!==!require('./environment.cjs').PRODUCTION||s.provider!=='GOOGLE_CLOUD_STORAGE'||s.bucket!==G.BUCKET||!ids(s.actorUids,16)||!s.actorUids.includes(uid)||!ids(s.effectIds,60)||!ids(s.scopeIds||[],16)||!ids(s.noticeUnitIds||[],30)||!require('./execution-window.cjs').windowLive(s,Date.now())||s.expiresAt>expiresAt)fail('PILOT_DISABLED');return s;
}
function effect(get,s,id,uid,write,contains=require('./act-cycle-scope.cjs').contains){
 const out=get('jobOutbox/'+id),owner=out&&get('commandExecutorOwners/'+out.createdByOperation),receipt=out&&get('commandReceipts/'+out.createdByOperation),route=owner?.routeId&&get('commandExecutorRoutes/'+owner.routeId),scope=route&&get('actFlowExecutorScopes/'+route.scopeId);
 const at=out?.createdAt?.toMillis?.();
 if(!out||!s.actorUids.includes(out.actor?.uid)||owner?.ownerUid!==out.actor.uid||write&&out.actor.uid!==uid||owner?.executor!==EXECUTOR||receipt?.state!=='COMPLETED'||out.effectKind!=='FILE_CREATE'||scope?.synthetic!==!require('./environment.cjs').PRODUCTION||scope.projectId!==PROJECT||!contains(get,scope,'jobOutbox/'+id)||!(s.effectIds.includes(id)||(s.scopeIds||[]).includes(route.scopeId)&&Number.isFinite(at)&&at>=s.startsAt))fail('AUTHORIZATION_DENIED');
 return {out,scope,scopePath:'actFlowExecutorScopes/'+route.scopeId};
}
module.exports={CONTROL,check,effect};
