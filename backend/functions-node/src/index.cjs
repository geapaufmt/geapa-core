'use strict';
if(process.env.GEAPA_RUNTIME==='EMULATOR')require('./local-network.cjs');
const {onRequest}=require('firebase-functions/v2/https');
const {handler}=require('./http.cjs');
const {runtime}=require('./runtime.cjs');
const {execute}=require('./executor.cjs');
const {SERVICE_ACCOUNT}=require('./environment.cjs');
const {recoverProjection}=require('./projection-recovery.cjs');
const {readAct}=require('./act-readers.cjs');
const projectionEnabled=process.env.GEAPA_PROJECTION_RECOVERY_ENABLED==='true';
const actions=[...(process.env.GEAPA_ACT_FLOWS_ENABLED==='true'?['flowRoute','flowPreview','flowPrepare','flowPortalPrepare']:[]),'prepare','execute','status','recover',...(process.env.GEAPA_ACTIVITY_COMMANDS_ENABLED==='true'?['preview','activityRoute']:[]),...(projectionEnabled?['projectionRecover','projectionStatus']:[])];
if(process.env.GEAPA_ACT_READERS_ENABLED==='true')actions.push('actRead');
if(process.env.GEAPA_ACT_WORKSPACE_ENABLED==='true')actions.push('workspaceDownload');
if(process.env.GEAPA_ACT_STORAGE_ENABLED==='true')actions.push('storageStart','storageStatus','storageChunk','storageFinalize','storageDownload','storageNotice');
if(process.env.GEAPA_MEMBERS_ENABLED==='true')actions.push('memberPrepare','memberExecute','memberStatus','memberRecover','memberEnrollmentOpen','memberEnrollmentRead','memberEnrollmentPrepare');
// Endpoint may be deployed inert; the canonical control and distinct policy gate every call.
actions.push('institutionalGovernanceAttachment','institutionalGovernanceDocument','institutionalGovernancePreview','institutionalGovernanceExecute','institutionalGovernanceStatus','institutionalParameterPreview','institutionalParameterExecute','institutionalParameterStatus','institutionalPreview','institutionalExecute','institutionalStatus','institutionalVisibilityPreview','institutionalVisibilityExecute','institutionalVisibilityStatus');
if(process.env.GEAPA_MEMBERS_ENABLED==='true')actions.push(...require('./member-binding-admin.cjs').actions);
const dispatch=options=>options.body.action.startsWith('memberBinding')?require('./member-binding-admin.cjs').bindingAdmin(options):options.body.action.startsWith('institutionalGovernance')?require('./institutional-governance.cjs').editGovernance(options):options.body.action.startsWith('institutionalParameter')?require('./institutional-parameters.cjs').changeInstitutionalParameter(options):options.body.action.startsWith('institutionalVisibility')?require('./institutional-visibility.cjs').changeInstitutionalVisibility(options):options.body.action.startsWith('institutional')?require('./institutional-edit.cjs').editInstitutional(options):options.body.action.startsWith('member')?require('./member-executor.cjs').executeMember(options):options.body.action==='storageNotice'?require('./storage-notice.cjs').acknowledge(options):options.body.action==='storageDownload'?require('./storage-download.cjs').download(options):options.body.action.startsWith('storage')?require('./storage-transfer.cjs').transfer(options):options.body.action==='workspaceDownload'?require('./workspace-download.cjs').download(options):options.body.action==='actRead'?readAct(options):options.body.action.startsWith('projection')?recoverProjection(options):execute(options);
exports[require('./runtime-profile.cjs').profile.functionName]=onRequest({region:'southamerica-east1',memory:'512MiB',cpu:1,minInstances:0,maxInstances:2,concurrency:2,timeoutSeconds:120,serviceAccount:SERVICE_ACCOUNT,invoker:'public',cors:false},handler({runtime,execute:dispatch,actions}));

// Separate authenticated scheduler, absent unless explicitly enabled in the
// deployment manifest. Browser credentials cannot invoke this system executor.
if(process.env.GEAPA_BOARD_AUTOMATION_ENABLED==='true'){
 const {onSchedule}=require('firebase-functions/v2/scheduler');
 exports[require('./runtime-profile.cjs').profile.functionName+'BoardConsequences']=onSchedule({schedule:'every 1 minutes',region:'southamerica-east1',memory:'512MiB',cpu:1,maxInstances:1,concurrency:1,timeoutSeconds:120,serviceAccount:SERVICE_ACCOUNT,retryCount:0},async()=>{
  const results=await require('./board-automation.cjs').runDue(await runtime());
  console.info(JSON.stringify({event:'BOARD_AUTOMATIC_CYCLE',checked:results.length,blocked:results.filter(r=>r.state==='BLOCKED').map(r=>r.code)}));
 });
}

