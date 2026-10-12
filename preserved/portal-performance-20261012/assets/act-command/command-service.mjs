import {input,canonical,frozen,clone} from './command-contract.mjs';
import {fail} from './errors.mjs';
import {result} from './result-dtos.mjs';
const terminal=v=>!!v.businessState||['COMPLETED','FAILED'].includes(v.receipt?.state);
// The authenticated process response follows the committed transaction. Reuse
// its validated receipt; a missing/uncertain response still needs status recovery.
const committed=(v,id)=>{const d=v?.state==='RECOVERED'?v.result:v;return d?.domainResult?result(d,id):null;};
// In-memory operation retention only. An uncertain submit never gets a fresh key.
export class CommandService {
 #auth; #transport; #operations=new Map(); #generation=-1; #cancel=()=>this.#operations.clear();
 constructor(auth,transport){this.#auth=auth;this.#transport=transport;this.#auth.listeners.add(this.#cancel);}
 #current(){const t=this.#auth.capture();if(this.#generation!==t.generation){this.#operations.clear();this.#generation=t.generation;}this.#auth.listeners.add(this.#cancel);return t;}
 async submit(raw){const v=input(raw),t=this.#current(),old=this.#operations.get(v.operationId),mark=canonical(v);if(old&&old.mark!==mark)fail('COMMAND_NOT_ADMITTED');if(old?.key)return this.status(old.key);if(old?.pending)fail('COMMAND_RECONCILIATION_REQUIRED');
  const entry={mark,pending:true};this.#operations.set(v.operationId,entry);
  try{const key=await this.#transport.submit(v,key=>{this.#auth.assert(t);entry.key=key;});this.#auth.assert(t);entry.key=key;entry.pending=false;return frozen({operationId:key.operationId,state:'RECEIVED'});}catch(e){this.#auth.assert(t);throw e;}
 }
 async status(key){const t=this.#current(),value=result(await this.#transport.status(key),key.operationId);this.#auth.assert(t);return value;}
 // Recovery selectors contain no token/proof/third-party UID and may be retained by UI.
 selector(operationId){this.#current();const key=this.#operations.get(operationId)?.key;if(!key)fail('COMMAND_RECONCILIATION_REQUIRED');return frozen(clone(key));}
 // Hosted ACT recovery keeps the original input and operationId throughout uncertainty.
 async execute(raw){const v=input(raw),t=this.#current(),mark=canonical(v);let entry=this.#operations.get(v.operationId);if(entry&&entry.mark!==mark)fail('COMMAND_NOT_ADMITTED');if(entry?.busy)fail('COMMAND_PENDING');if(!entry){entry={mark,raw:v};this.#operations.set(v.operationId,entry)}entry.busy=true;
  try{
   if(entry.key){try{const state=await this.status(entry.key);if(terminal(state))return state;}catch(e){if(['FORBIDDEN','ACCESS_REVOKED','STALE_AUTHORIZATION','AUTH_REQUIRED'].includes(e.code))throw e;}}
   let submitError;
   if(!entry.accepted){try{entry.key=await this.#transport.submit(v,key=>{this.#auth.assert(t);entry.key=key});entry.accepted=true}catch(e){submitError=e;if(['FORBIDDEN','ACCESS_REVOKED','STALE_AUTHORIZATION','AUTH_REQUIRED','CONTRACT_NOT_AVAILABLE'].includes(e.code))throw e;}}
   if(!entry.key)throw submitError;
   try{const response=await this.#transport.process(entry.key);this.#auth.assert(t);const done=committed(response,entry.key.operationId);if(done)return done;}catch(error){
    if(['FORBIDDEN','ACCESS_REVOKED','STALE_AUTHORIZATION','AUTH_REQUIRED','CONTRACT_NOT_AVAILABLE'].includes(error.code))throw error;
    // A lost response may follow a successful commit. Query this same selector
    // once before asking the operator to retry; never create a second command.
    const recovered=await this.status(entry.key);this.#auth.assert(t);if(terminal(recovered))return recovered;throw error;
   }
   this.#auth.assert(t);return await this.status(entry.key);
  }finally{entry.busy=false}
 }
 async recover(key){const t=this.#current();try{const state=await this.status(key);if(terminal(state))return state;}catch(e){if(['FORBIDDEN','ACCESS_REVOKED','STALE_AUTHORIZATION','AUTH_REQUIRED'].includes(e.code))throw e;}const response=await this.#transport.process(key);this.#auth.assert(t);return committed(response,key.operationId)||this.status(key);}
 clear(){this.#operations.clear();}
 dispose(){this.clear();this.#auth.listeners.delete(this.#cancel);}
}
