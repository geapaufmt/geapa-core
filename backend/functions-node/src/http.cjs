'use strict';
const {fence,fail}=require('./environment.cjs');
const ORIGIN=require('./runtime-profile.cjs').profile.origin;
function payload(value,depth=0){
 if(depth>20)fail('COMMAND_PAYLOAD_INVALID');
 if(value&&typeof value==='object'){
  if(!Array.isArray(value)&&Object.getPrototypeOf(value)!==Object.prototype)fail('COMMAND_PAYLOAD_INVALID');
  for(const [k,v]of Object.entries(value)){if(['__proto__','constructor','prototype'].includes(k))fail('COMMAND_PAYLOAD_INVALID');payload(v,depth+1);}
 }else if(!['string','number','boolean'].includes(typeof value)&&value!==null||typeof value==='number'&&!Number.isFinite(value))fail('COMMAND_PAYLOAD_INVALID');
}
const codes=new Set(['DOMAIN_STATE_INVALID','WORKSPACE_IDENTITY_NOT_CONFIGURED','NOT_FOUND','UNRESOLVED_DEPENDENCY','CONFIG_CAPABILITY_DISABLED','SCOPE_TOO_LARGE','AUTH_BINDING_INVALID','ACCESS_REVOKED','AUTHORIZATION_DENIED','M13_AUTHZ_DENIED','CAPABILITY_REQUIRED','SCOPE_INVALID','REVISION_CONFLICT','STALE_REVISION','OPERATION_CONFLICT','VALIDATION_ERROR','SCHEMA_INVALID','COMMAND_TYPE_INVALID','COMMAND_PAYLOAD_INVALID','RECONCILIATION_REQUIRED','DESTINATION_INVALID','PILOT_DISABLED','PILOT_DRAINING','PILOT_RESOURCE_DENIED','PILOT_FIELD_DENIED','PILOT_WRITE_DENIED','EXECUTOR_FENCE_CONFLICT','PILOT_LIMIT_REACHED','APP_CHECK_REQUIRED','APP_CHECK_INVALID','ORIGIN_DENIED']);
function handler({runtime,execute,log=console.info,actions=['prepare','execute','status','recover']}){
 return async(req,res)=>{
  const started=performance.now();let code='OK',internalCode=null,metrics={};
  res.set('Cache-Control','no-store');res.set('Vary','Origin');res.set('X-Content-Type-Options','nosniff');
  try{
   fence();const origin=req.headers.origin;
   if(origin!==undefined&&origin!==ORIGIN)fail('ORIGIN_DENIED');
   if(origin)res.set('Access-Control-Allow-Origin',ORIGIN);
   if(req.method==='OPTIONS'){
    const headers=(req.headers['access-control-request-headers']||'').toLowerCase().split(',').map(s=>s.trim()).filter(Boolean);
    if(!origin||req.headers['access-control-request-method']!=='POST'||headers.some(x=>!['authorization','content-type','x-firebase-appcheck'].includes(x)))fail('ORIGIN_DENIED');
    res.set('Access-Control-Allow-Methods','POST');res.set('Access-Control-Allow-Headers','Authorization, Content-Type, X-Firebase-AppCheck');res.set('Access-Control-Max-Age','300');return res.status(204).send('');
   }
   if(req.method!=='POST')return res.status(405).json({error:{code:'COMMAND_TYPE_INVALID'}});
   const maximum=req.body?.action==='storageChunk'&&actions.includes('storageChunk')?Math.ceil(4*1024*1024/3)*4+4096:65536;
   if(!/^application\/json(?:\s*;\s*charset=utf-8)?$/i.test(req.headers['content-type']||'')||req.headers['content-encoding']&&req.headers['content-encoding']!=='identity'||(req.rawBody?.length||0)>maximum||Buffer.byteLength(JSON.stringify(req.body||{}))>maximum)fail('COMMAND_PAYLOAD_INVALID');
   payload(req.body);if(!req.body||Array.isArray(req.body)||!actions.includes(req.body.action)||!req.body.request||Array.isArray(req.body.request))fail('COMMAND_PAYLOAD_INVALID');
   const m=/^Bearer ([^\s]+)$/.exec(req.headers.authorization||'');if(!m)fail('AUTH_BINDING_INVALID');
   if(req.headers['x-firebase-appcheck']&&String(req.headers['x-firebase-appcheck']).length>16000)fail('APP_CHECK_INVALID');
   const deps=await runtime();const result=await execute({...deps,token:m[1],body:req.body,origin,appCheckToken:req.headers['x-firebase-appcheck']});
   metrics=Object.fromEntries(Object.entries(result.metrics||{}).filter(([k,v])=>['reads','writes','sdkCalls','transactionAttempts','totalMs','authMs','prepareMs','admissionMs','reserveMs','claimMs','finalizeMs','projectionMs'].includes(k)&&Number.isFinite(v)));return res.json(result);
  }catch(e){internalCode=['READ_DISCOVERY_LIMIT','QUERY_NOT_IN_PILOT','M1_SOURCE_MISSING'].includes(e.code)?e.code:Number.isInteger(e.code)&&e.code>=0&&e.code<=16?'SDK_'+e.code:null;code=codes.has(e.code)?e.code:'RECONCILIATION_REQUIRED';return res.status(code==='NOT_FOUND'?404:code==='UNRESOLVED_DEPENDENCY'?422:code==='AUTH_BINDING_INVALID'?401:code.includes('CONFLICT')||code==='STALE_REVISION'?409:code==='RECONCILIATION_REQUIRED'?503:403).json({error:{code}});
  }finally{
   // No request payloads, raw errors, headers, credentials, names or email in logs.
   log(JSON.stringify({component:'geapa-command-dev-pilot',code,internalCode,action:actions.includes(req.body?.action)?req.body.action:null,metrics,durationMs:Math.round(performance.now()-started),operationId:/^[a-f0-9]{32}$/.test(req.body?.request?.operationId)?req.body.request.operationId:null,revision:process.env.K_REVISION||'emulator'}));
  }
 };
}
module.exports={handler,ORIGIN,payload};
