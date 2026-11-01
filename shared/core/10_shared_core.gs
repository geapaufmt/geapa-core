// F4A: native synchronous port of F3.5 shared/core/index.cjs. No external module loading.
function f4Core_() {
  if (f4Core_.value_) return f4Core_.value_;
  return (f4Core_.value_ = (function () {
'use strict';

const codes=['AUTH_BINDING_INVALID','ACCESS_REVOKED','CAPABILITY_REQUIRED','SCOPE_INVALID','FUNCTION_EXPIRED','REVISION_CONFLICT','COMMAND_TYPE_INVALID','COMMAND_PAYLOAD_INVALID','OPERATION_CONFLICT','LEASE_LOST','DOMAIN_STATE_INVALID','DESTINATION_INVALID','SCHEMA_INVALID','RECONCILIATION_REQUIRED'];
function fail(code){const e=new Error(codes.includes(code)?code:'DOMAIN_STATE_INVALID');e.code=e.message;throw e}
function need(v,c='DOMAIN_STATE_INVALID'){if(!v)fail(c)}
const id=v=>typeof v==='string'&&/^[A-Za-z0-9_-]{1,48}$/.test(v),op=v=>typeof v==='string'&&/^[a-f0-9]{32}$/.test(v),rev=v=>Number.isSafeInteger(v)&&v>=0;
const ms=v=>v?.toMillis?v.toMillis():NaN;
const time=(a,z,n)=>Number.isFinite(n)&&ms(a)<ms(z)&&ms(a)<=n&&n<ms(z);
function exact(v,keys){need(v&&Object.getPrototypeOf(v)===Object.prototype&&Object.keys(v).length===keys.length&&keys.every(k=>Object.prototype.hasOwnProperty.call(v,k)),'SCHEMA_INVALID')}
function canonical(v){if(Array.isArray(v))return v.map(canonical);if(v&&typeof v==='object'){need(Object.getPrototypeOf(v)===Object.prototype,'COMMAND_PAYLOAD_INVALID');return Object.fromEntries(Object.keys(v).sort().map(k=>[k,canonical(v[k])]))}need(v===null||['string','boolean'].includes(typeof v)||(typeof v==='number'&&Number.isFinite(v)),'COMMAND_PAYLOAD_INVALID');return v}
const hash=v=>Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(canonical(v)),Utilities.Charset.UTF_8).map(x=>(x&255).toString(16).padStart(2,'0')).join('');
const digest=r=>hash({operationId:r.operationId,schemaVersion:r.schemaVersion,uid:r.uid,commandType:r.commandType,payload:r.payload,expectedRevision:r.expectedRevision});
const catalog={justificativaEnviar:['activityId','reason'],justificativaAnalisar:['activityId','justificationId','decision'],portalPublicAdminPublishPage:['pageId']};
function envelope(r,operationId){need(catalog[r?.commandType],'COMMAND_TYPE_INVALID');try{exact(r,['schemaVersion','operationId','uid','commandType','payload','expectedRevision','admissionSlot','grantId','receivedAt','state']);exact(r.payload,catalog[r.commandType])}catch{fail('COMMAND_PAYLOAD_INVALID')}
need(op(operationId)&&r.operationId===operationId&&r.schemaVersion==='geapa-command-v1'&&id(r.uid)&&r.state==='RECEIVED'&&rev(r.expectedRevision)&&[0,1].includes(r.admissionSlot)&&typeof r.grantId==='string'&&/^[A-Za-z0-9_-]{8,100}$/.test(r.grantId)&&Number.isFinite(ms(r.receivedAt)),'COMMAND_PAYLOAD_INVALID');
const p=r.payload;need(r.commandType==='portalPublicAdminPublishPage'?id(p.pageId):id(p.activityId),'COMMAND_PAYLOAD_INVALID');if(r.commandType==='justificativaEnviar')need(typeof p.reason==='string'&&p.reason.trim().length>0&&p.reason.length<=1024,'COMMAND_PAYLOAD_INVALID');if(r.commandType==='justificativaAnalisar')need(id(p.justificationId)&&['DEFERIR','INDEFERIR'].includes(p.decision),'COMMAND_PAYLOAD_INVALID');return r}
const scope=r=>r.commandType==='portalPublicAdminPublishPage'?'EDITORIAL':'ACT_'+r.payload.activityId;
const capability=r=>({justificativaEnviar:'JUSTIFICATION_SUBMIT_SELF',justificativaAnalisar:'JUSTIFICATION_REVIEW',portalPublicAdminPublishPage:'PUBLIC_CONTENT_MANAGE'})[r.commandType];
function destination(c){exact(c,['environment','projectId','database','host','port']);need(((c.environment==='SYNTHETIC_LOCAL'&&c.projectId==='demo-geapa-m04-local'&&c.host==='127.0.0.1'&&c.port===8184)||((c.environment==='DEV'&&c.projectId==='geapa-convergencia-dev'||c.environment==='PROD'&&c.projectId==='portal-geapa'&&f4ExecutionConfig_().environment==='PROD')&&c.host==='firestore.googleapis.com'&&c.port===443))&&c.database==='(default)','DESTINATION_INVALID')}
// Execution-local memo of pure date conversion; never caches authority or data.
const civilStarts=new Map();
function civilStart(s){need(typeof s==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(s),'DOMAIN_STATE_INVALID');if(civilStarts.has(s))return civilStarts.get(s);const [y,m,d]=s.split('-').map(Number);let t=Date.UTC(y,m-1,d);need(new Date(t).toISOString().slice(0,10)===s);const ref=t;for(let i=0;i<3;i++){const v=Utilities.formatDate(new Date(t),'America/Cuiaba',"yyyy-MM-dd'T'HH:mm:ss").match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/);need(v);t+=ref-Date.UTC(+v[1],+v[2]-1,+v[3],+v[4],+v[5],+v[6])}if(civilStarts.size>=512)civilStarts.clear();civilStarts.set(s,t);return t}
function nextDay(s){civilStart(s);const[y,m,d]=s.split('-').map(Number);return new Date(Date.UTC(y,m-1,d+1)).toISOString().slice(0,10)}
function expectedRevision(expected,current){need(rev(current));if(expected!==current){try{fail("REVISION_CONFLICT")}catch(e){e.canonicalRevision=current;throw e}}}
return {expectedRevision,codes,fail,need,id,op,rev,ms,time,exact,canonical,hash,digest,envelope,scope,capability,destination,civilStart,nextDay};


  })());
}

// Adapter-owned destination. Node supplies this immutable configuration; the
// Apps Script composition remains DEV. No request can select a destination.
function f4ExecutionConfig_(){return {environment:'DEV',projectId:'geapa-convergencia-dev',database:'(default)',host:'firestore.googleapis.com',port:443};}
function f4ExecutionEnvironment_(){
 const C=f4Core_(),N=C.need,c=f4ExecutionConfig_(),prod=c.environment==='PROD';
 N(c.projectId===(prod?'portal-geapa':'geapa-convergencia-dev')&&['DEV','PROD'].includes(c.environment),'DESTINATION_INVALID');
 return {config:c,project:c.projectId,environment:c.environment,production:prod,
  destination:q=>q?.projectId===c.projectId&&q.database==='(default)'&&(q.environment===c.environment||!prod&&q.environment==='SYNTHETIC_LOCAL'),
  scope:s=>s?.projectId===c.projectId&&s.synthetic===!prod,
  route:v=>v?.projectId===c.projectId&&v.syntheticOnly===!prod,
  executors:prod?['CF_GEN2_PROD_V1']:['APPS_SCRIPT_DEV','CF_GEN2_DEV_EDIT_V1'],
  prefix:prod?'act-prod/':'act-dev/'};
}

function f4OperationalDestination_(c){return c?.environment==='PROD'?f4ExecutionEnvironment_().destination(c):f4DevBoundary_().permit(c);}
