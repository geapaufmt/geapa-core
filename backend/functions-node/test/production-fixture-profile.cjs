'use strict';
// Emulator-only transformation of synthetic controls, never a migration utility.
const E=require('../src/environment.cjs');
module.exports=function(docs){
 if(!E.PRODUCTION)return docs;
 if(process.env.GEAPA_RUNTIME!=='EMULATOR'||!process.env.FIRESTORE_EMULATOR_HOST?.startsWith('127.0.0.1:'))throw Error('LOCAL_ONLY');
 for(const [p,v]of Object.entries(docs)){
  if(v.projectId==='geapa-convergencia-dev')v.projectId=E.PROJECT;
  if(v.environment==='DEV')v.environment='PROD';
  if(p.startsWith('coreDevAdmins/')){delete docs[p];continue;}
  if(p.startsWith('commandExecutorRoutes/')){v.syntheticOnly=false;v.executor=E.EXECUTOR;v.mode='PROD_SCOPED_OPERATIONAL';}
  if(/^(activityExecutorScopes|actFlowExecutorScopes|memberExecutorScopes)\//.test(p))v.synthetic=false;
  if(p==='activityExecutorControl/current'||p==='memberExecutorControl/current')v.mode='PROD_SCOPED_OPERATIONAL';
  if(p==='projectionRecoveryControl/current'){v.mode='PROD_READERS_OPERATIONAL';v.actorUids??=['CF_ADMIN'];}
  if(p.startsWith('actHistoricalFileCopies/'))v.receipt.objectName='act-history/'+v.receipt.migrationId;
 }
 return docs;
};
