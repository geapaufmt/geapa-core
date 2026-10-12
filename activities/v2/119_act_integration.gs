// Compatibility facade over the one Core endpoint. No direct ACT mutation or second transport.
function f4ActivitiesIntegratedCore_(options){
 const M=f4ActivitiesModel_(),pipeline=options.pipeline;
 M.need(pipeline?.endpoint&&(options.config?.environment==='SYNTHETIC_LOCAL'||f4DevBoundary_().permit(options.config)),'UNRESOLVED_DEPENDENCY','SHARED_CORE_REQUIRED');
 function post(action,token,request){return pipeline.endpoint.post({method:'POST',projectId:options.config.projectId,scriptId:options.config.environment==='SYNTHETIC_LOCAL'?'LOCAL_M13_G':f4DevBoundary_().script,body:JSON.stringify({action,idToken:token,request})});}
 function command(token,request){M.need(request?.schemaVersion==='geapa-act-command-v13'&&request.commandType?.startsWith('ACT.'),'VALIDATION_ERROR','SCHEMA_VERSION');return post('core.admit.v13',token,request);}
 function receipt(token,key){M.need(key?.commandType?.startsWith('ACT.'),'VALIDATION_ERROR');return post('core.status.v13',token,key);}
 function read(principal,id,args){M.need(options.cursor&&options.documentRoot,'UNRESOLVED_DEPENDENCY','READ_SERVICES');return f4ActivitiesReadService_(options).read(principal,id,args);}
 return Object.freeze({command,receipt,read,activities:f4ActivitiesSharedSystemCommands_(pipeline,options.principals)});
}

// Scheduler/Workspace compatibility only. Principal handles come from trusted Core composition.
function f4ActivitiesSharedSystemCommands_(pipeline,principals){
 const M=f4ActivitiesModel_();
 function execute(handle,q){const p=principals.resolve(handle);M.need(p&&['SYSTEM_JOB','WORKSPACE_EXECUTOR'].includes(p.kind),'AUTHORIZATION_DENIED');
  const authorization={kind:p.kind,servicePrincipal:p.servicePrincipal,scope:p.scope,ruleRef:p.ruleRef,jobKind:p.jobKind||'WORKSPACE',runId:p.runId};
  const request={schemaVersion:'geapa-act-command-v13',operationId:q.operationId,commandType:q.commandType,target:{scope:q.scope,expected:q.expected,authorization},payload:q.payload,evidenceRefs:[],decisionRef:null};
  const value=pipeline.system(request).attempt(p.servicePrincipal),d=value.state==='RECOVERED'?value.result:value;
  if(d.state==='COMPLETED')return {...d.domainResult,replay:value.state==='RECOVERED'};
  M.fail(d.domainResult?.error?.code||d.code,d.domainResult?.error?.reasonCode);
 }
 return Object.freeze({execute});
}

function f4ActivitiesCursor_(secret){
 const M=f4ActivitiesModel_();M.need(typeof secret==='string'&&secret.length>=32,'UNRESOLVED_DEPENDENCY','CURSOR_KEY');
 function signature(s){return Utilities.computeHmacSha256Signature(s,secret,Utilities.Charset.UTF_8).map(x=>(x&255).toString(16).padStart(2,'0')).join('');}
 function seal(v){const raw=JSON.stringify(M.json(v));return Utilities.base64EncodeWebSafe(raw)+'.'+signature(raw);}
 function open(s){M.need(typeof s==='string'&&s.length<=8192,'VALIDATION_ERROR','CURSOR');const parts=s.split('.');M.need(parts.length===2,'VALIDATION_ERROR','CURSOR');let raw;try{raw=Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();}catch{M.fail('VALIDATION_ERROR','CURSOR');}const expected=signature(raw);let diff=parts[1].length^expected.length;for(let i=0;i<expected.length;i++)diff|=expected.charCodeAt(i)^(parts[1].charCodeAt(i)||0);M.need(diff===0,'VALIDATION_ERROR','CURSOR_SIGNATURE');try{return JSON.parse(raw);}catch{M.fail('VALIDATION_ERROR','CURSOR');}}
 return {seal,open};
}
