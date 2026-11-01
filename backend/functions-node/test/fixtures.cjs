'use strict';
const {ACTIVITY,ROOT,PROJECT,EXECUTOR}=require('../src/environment.cjs');
function seed(){const base=require('./fixtures/core-edit.json'),now=Date.now(),delta=now-base.now;
 const replace=v=>typeof v==='string'?v.replaceAll('dev_meeting',ACTIVITY).replaceAll('ACT_POLICY','CF_DEV_ACT_POLICY'):v&&typeof v==='object'?v.timestampValue?{timestampValue:v.timestampValue.startsWith('9999-')?v.timestampValue:new Date(Date.parse(v.timestampValue)+delta).toISOString()}:Array.isArray(v)?v.map(replace):Object.fromEntries(Object.entries(v).map(([k,x])=>[replace(k),replace(x)])):v;
 const docs=replace(base.docs);
 // Advancing synthetic timestamps must also advance their civil date. Without
 // this, a title-only edit repairs a stale fixture date and the pilot correctly
 // rejects that extra field change after the original fixture's calendar day.
 for(const [path,doc]of Object.entries(docs))if(path.startsWith('canonicalActivities/')&&doc.fields.startsAt?.timestampValue&&doc.fields.civilDate){doc.fields.civilDate={stringValue:new Intl.DateTimeFormat('en-CA',{timeZone:'America/Cuiaba',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date(doc.fields.startsAt.timestampValue))};}
 return {...base,docs,now};}
function pilotDocs(){const now=Date.now();return {
 'pilotControl/runtime':{schemaVersion:'geapa-functions-edit-pilot-v1',projectId:PROJECT,namespace:ROOT,activityId:ACTIVITY,revision:1,state:'ACTIVE',actorUids:['CF_ADMIN','CF_NONE','CF_MEMBER'],policyIds:['CF_DEV_ACT_POLICY'],startsAt:now-1000,expiresAt:now+3600000,maxOperations:100,browserEnabled:false,appCheck:{mode:'OBSERVE',appIds:[]}},
 'pilotExecutorOwnership/current':{schemaVersion:'geapa-executor-ownership-v1',executor:EXECUTOR,namespace:ROOT,commandType:'ACT.ACTIVITY_SAVE',mode:'ISOLATED',sharedCutover:false,epoch:1},
 ['pilotResources/'+ACTIVITY]:{schemaVersion:'geapa-pilot-resource-v1',projectId:PROJECT,namespace:ROOT,activityId:ACTIVITY,synthetic:true}
 };}
module.exports={seed,pilotDocs};
