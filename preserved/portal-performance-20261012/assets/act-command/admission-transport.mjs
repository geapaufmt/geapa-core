import {fail,id,operation,PortalError} from './errors.mjs';
import {SCHEMA,TYPES,ACT_TYPES,input,exact,need,canonical,clone,frozen,hex,revisions} from './command-contract.mjs';
const PROJECT='demo-geapa-m13-local';
export function sanitized(e){const codes={M13_AUTHZ_DENIED:'FORBIDDEN',HUMAN_DECISION_REQUIRED:'HUMAN_DECISION_REQUIRED',UNRESOLVED_DEPENDENCY:'UNRESOLVED_DEPENDENCY',VALIDATION_ERROR:'VALIDATION_ERROR',INCOMPLETE_REQUEST:'INCOMPLETE_REQUEST',STALE_REVISION:'STALE_REVISION',DEADLINE_VIOLATION:'DEADLINE_VIOLATION',INVALID_STATE_TRANSITION:'INVALID_STATE_TRANSITION',CAPACITY_EXHAUSTED:'CAPACITY_EXHAUSTED',SCOPE_TOO_LARGE:'SCOPE_TOO_LARGE',AUTH_BINDING_INVALID:'STALE_AUTHORIZATION',AUTHORIZATION_DENIED:'FORBIDDEN',ACCESS_REVOKED:'ACCESS_REVOKED',OPERATION_CONFLICT:'COMMAND_NOT_ADMITTED',CONFIG_CAPABILITY_DISABLED:'CONTRACT_NOT_AVAILABLE','permission-denied':'FORBIDDEN'};return new PortalError(e instanceof PortalError?e.code:codes[e?.code]||'COMMAND_RECONCILIATION_REQUIRED');}
async function digest(r,projectId){const keys=['schemaVersion','operationId','ownerUid','commandType','target','expectedRevisions','payload','evidenceRefs','decisionRef'];const bytes=new TextEncoder().encode(canonical({projectId,...Object.fromEntries(keys.map(k=>[k,r[k]]))}));return [...new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))].map(x=>x.toString(16).padStart(2,'0')).join('');}
// Only a locally composed port; no URL override, endpoint, issuer or verifier.
export class AdmissionTransport {
 #auth; #ports; #clock; #project;
 constructor(auth,ports,config){exact(config,['environment','projectId']);need(config.environment==='SYNTHETIC_LOCAL'&&[PROJECT,'demo-geapa-m04-local'].includes(config.projectId)||config.environment==='DEV'&&config.projectId==='geapa-convergencia-dev');exact(ports,['admit','commitRequestUse','status','clock',...('process'in ports?['process']:[])]);need(Object.values(ports).every(v=>typeof v==='function'));this.#auth=auth;this.#ports=Object.freeze({...ports});this.#clock=ports.clock;this.#project=config.projectId;}
 async submit(raw,onPrepared=()=>{}){const v=input(raw),t=this.#auth.capture();try{
  const token=await this.#auth.token(t);this.#auth.assert(t);
  if(v.schemaVersion==='geapa-act-command-v13'){const key=frozen({operationId:v.operationId,commandType:v.commandType,digest:await digest({...v,ownerUid:t.uid,expectedRevisions:[]},this.#project)});this.#auth.assert(t);onPrepared(key);}
  // The protected issuer supplies preconditions; no proof/ticket document is returned.
  const a=await this.#ports.admit('core.admit.v13',v,token);this.#auth.assert(t);
  exact(a,['requestEnvelope','ticketId','expiresAt']);need(id(a.ticketId)&&Number.isSafeInteger(a.expiresAt)&&this.#clock()<a.expiresAt&&a.expiresAt<=this.#clock()+120000);
  const r=a.requestEnvelope;exact(r,['schemaVersion','operationId','ownerUid','commandType','target','expectedRevisions','payload','evidenceRefs','decisionRef','digest']);
  need(r.ownerUid===t.uid&&r.schemaVersion===v.schemaVersion&&hex(r.digest));v.schemaVersion===SCHEMA?revisions(r.expectedRevisions):need(Array.isArray(r.expectedRevisions)&&r.expectedRevisions.length===0);
  need(canonical(v)===canonical(Object.fromEntries(Object.keys(v).map(k=>[k,r[k]])))&&r.digest===await digest(r,this.#project));this.#auth.assert(t);need(this.#clock()<a.expiresAt);
  // Atomic create-only request + use; the port must supply server timestamps and Rules.
  const pair=frozen({request:{...clone(r),ticketId:a.ticketId,state:'RECEIVED'},use:{schemaVersion:'member-admission-use-v13',ticketId:a.ticketId,operationId:r.operationId,ownerUid:t.uid}});
  const key=frozen({operationId:r.operationId,digest:r.digest,commandType:r.commandType});onPrepared(key);
  await this.#ports.commitRequestUse(pair,token);this.#auth.assert(t);return key;
 }catch(e){throw sanitized(e)}}
 async status(key){exact(key,['operationId','digest','commandType']);need(operation(key.operationId)&&hex(key.digest)&&[...TYPES,...ACT_TYPES,'members.binding.associate','members.binding.revoke'].includes(key.commandType));const t=this.#auth.capture();try{const token=await this.#auth.token(t);const out=await this.#ports.status(clone(key),token);this.#auth.assert(t);return out;}catch(e){throw sanitized(e)}}
 async process(key){need(typeof this.#ports.process==='function');exact(key,['operationId','digest','commandType']);need(operation(key.operationId)&&hex(key.digest)&&ACT_TYPES.includes(key.commandType));const t=this.#auth.capture();try{const token=await this.#auth.token(t);const out=await this.#ports.process(clone(key),token);this.#auth.assert(t);return out;}catch(e){throw sanitized(e)}}
}
