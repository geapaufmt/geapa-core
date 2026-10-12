import {dataPrefix,fixtureUid} from './dev-scope.mjs';
import {decodeOwnRead} from './own-reader.mjs';
import {AdmissionTransport,sanitized} from './admission-transport.mjs';
import {CommandService} from './command-service.mjs';
import {canonical,exact,need,frozen} from './command-contract.mjs';
import {fail} from './errors.mjs';
import {createDirectReaders} from './direct-readers.mjs';
import {createActivityReader} from './activity-reader.mjs';

const PROJECT='geapa-convergencia-dev',ROOT='';
const ENDPOINT='https://script.google.com/macros/s/AKfycbzYMS3ZpS38--VgUNh8vlC56688Lc9Dd1do8OMDBMRMu7W5AJjUunLWAaf6EZEL1YxA/exec';
const hash=async value=>[...new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(value)))].map(x=>x.toString(16).padStart(2,'0')).join('');
const proofMimeTypes={pdf:'application/pdf',jpg:'image/jpeg',jpeg:'image/jpeg',png:'image/png',doc:'application/msword',docx:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'};
const operation=()=>crypto.randomUUID().replaceAll('-','');
const terminal=v=>!!v.businessState||['COMPLETED','FAILED'].includes(v.receipt?.state);
const reasonMessages={ADMISSION_EXPIRED:'O envio anterior expirou sem gravar a alteração. Confira os dados e envie novamente.',NORMATIVE_CHANGE_NOT_OPERATIONAL:'Esta configuração alteraria uma regra institucional. Use os critérios já aprovados.',APPROVED_OPERATIONAL_TEMPLATE_REQUIRED:'Não há modelo canônico aprovado para esses critérios. A configuração permanece desabilitada.',"INITIAL_QUOTA_SUM":"A soma dos limites iniciais precisa corresponder às obrigações incidentes aprovadas para o ciclo.","QUOTA_EXISTING_RIGHTS":"O limite informado é menor que as ocupações existentes. Confira as reservas antes de reduzir a cota.","INITIAL_QUOTA_BEFORE_OPEN":"A configuração inicial de cotas deve ser registrada antes da abertura do ciclo.","WAIT_ADMISSION_POLICY_REQUIRED":"A admissão da espera ainda não foi habilitada na configuração do ciclo.","PRIORITY_TIE":"Há pedidos com a mesma prioridade. Registre a decisão de desempate antes de admitir.","PRIORITY_TIE_CHANGED":"A fila mudou ou não há empate pendente. Atualize os pedidos e confira a decisão.","EARLIER_PRIORITY":"Há pedido anterior com precedência nesta cota.","PRIORITY_UNRESOLVED":"A prioridade de uma alteração temática ainda depende de decisão da Diretoria.","WAITING_MATERIALITY":"A alteração temática precisa de decisão sobre sua natureza antes da admissão.","QUOTA_NOT_CONFIGURED":"Os limites temáticos deste ciclo ainda precisam ser configurados.","THEMATIC_QUOTA_EXHAUSTED":"Não há cota temática disponível para esta admissão.",SESSION_HAS_COMMITMENTS:'Esta atividade tem apresentações agendadas. Remarque ou cancele os agendamentos antes de cancelar a atividade.',SESSION_RESIDUAL_FEASIBILITY:'A alteração deixaria participantes sem vaga elegível. Revise a agenda antes de confirmar.',DRAFT_SESSION_OCCUPIED:'Há apresentações vinculadas a esta sessão. Confira o planejamento antes de alterar a data.',D30:'A nova data precisa respeitar a antecedência mínima configurada para apresentações.',ACTIVITY_ALREADY_COLLECTING:'A chamada ou a realização já tem registros. A data e o modelo da atividade não podem ser alterados por este formulário.',ACTIVITY_MODEL_SNAPSHOT_REQUIRED:'O modelo original desta atividade precisa ser conferido pela administração antes da edição.'};
export function actDomainFailure(error){return {ok:false,errorCode:error?.code||'COMMAND_FAILED',reasonCode:error?.reasonCode||error?.code||'COMMAND_FAILED',message:reasonMessages[error?.reasonCode]||'Não foi possível concluir a ação. Atualize os dados e confira as condições informadas.'};}

// The HTTP port uses the already published shared Core endpoint. The only direct
// browser write is its existing, Rules-protected atomic request/use pair.
export function createActBrowserClient({auth,db,sdk,storage,fetchImpl=fetch,enabled=false}){
 need(db.app.options.projectId===PROJECT);const active=new Set(),generations=new Map(),presentationGenerations=new Map(),materialGenerations=new Map();
 async function http(action,request,token){
  if(!enabled)fail('CONTRACT_NOT_AVAILABLE');const controller=new AbortController(),timeout=['core.act.material-review.v13','core.act.workspace-upload.v13','core.act.workspace-reconcile.v13'].includes(action)?120000:55000,timer=setTimeout(()=>controller.abort(),timeout);auth.pending?.add(controller);
  try{const response=await fetchImpl(ENDPOINT,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action,idToken:token,request}),credentials:'omit',redirect:'follow',cache:'no-store',signal:controller.signal});
   if(!response.ok)fail('COMMAND_RECONCILIATION_REQUIRED');if(response.url&&!['script.google.com','script.googleusercontent.com'].includes(new URL(response.url).hostname))fail('COMMAND_RECONCILIATION_REQUIRED');
   const value=await response.json();if(value?.error)throw Object.assign(Error(value.error.code),{code:value.error.code});return value;
  }finally{clearTimeout(timer);auth.pending?.delete(controller)}
 }
 async function request(action,value){const ticket=auth.capture();try{const token=await auth.token(ticket),result=await http(action,value,token);auth.assert(ticket);return result;}catch(e){throw sanitized(e)}}
 const transport=new AdmissionTransport(auth,{clock:Date.now,admit:(action,v,token)=>http(action,v,token),status:(key,token)=>http('core.status.v13',key,token),process:(key,token)=>http('core.act.process.v13',key,token),commitRequestUse:async pair=>{
  const ticket=auth.capture(),batch=sdk.writeBatch(db);batch.set(sdk.doc(db,dataPrefix(ticket.uid)+'commandRequests/'+pair.request.operationId),{...pair.request,receivedAt:sdk.serverTimestamp()});batch.set(sdk.doc(db,dataPrefix(ticket.uid)+'memberAdmissionUses/'+pair.use.ticketId),{...pair.use,reservedAt:sdk.serverTimestamp()});auth.assert(ticket);let timer;try{await Promise.race([batch.commit(),new Promise((_,reject)=>{timer=setTimeout(()=>reject(Object.assign(Error('COMMAND_PENDING'),{code:'COMMAND_PENDING'})),20000)})]);}finally{clearTimeout(timer)}auth.assert(ticket);
 }},{environment:'DEV',projectId:PROJECT});
 const commands=new CommandService(auth,transport);
 async function journalName(slot){const t=auth.capture();const name='geapa.act.pending.v1.'+await hash(PROJECT+'|'+t.uid+'|'+slot);auth.assert(t);return {name,t};}
 function load(name){const raw=storage.getItem(name);if(!raw)return null;let value;try{value=JSON.parse(raw);exact(value,['operationId','intentHash','selector']);need(/^[a-f0-9]{32}$/.test(value.operationId)&&/^[a-f0-9]{64}$/.test(value.intentHash));if(value.selector){exact(value.selector,['operationId','commandType','digest']);need(value.selector.operationId===value.operationId&&/^[a-f0-9]{64}$/.test(value.selector.digest)&&value.selector.commandType.startsWith('ACT.'));}}catch{fail('COMMAND_RECONCILIATION_REQUIRED')}return value;}
 function save(name,value){try{storage.setItem(name,JSON.stringify(value));}catch{fail('COMMAND_RECONCILIATION_REQUIRED')}}
 async function recover(entry){try{const state=await commands.recover(entry.selector);if(terminal(state))return state;}catch(e){if(['FORBIDDEN','ACCESS_REVOKED','STALE_AUTHORIZATION','AUTH_REQUIRED'].includes(e.code))throw e;}
  // Retrieve the original protected intent, never recalculate a stale operation.
  const original=await request('core.act.resume.v13',entry.selector);return commands.execute(original);
 }
 async function execute(slot,intent,prepare,retain=false){
  if(!enabled)fail('CONTRACT_NOT_AVAILABLE');const {name,t}=await journalName(slot);if(active.has(name))fail('COMMAND_PENDING');active.add(name);
  try{const mark=await hash(canonical(intent));auth.assert(t);let entry=load(name);
   if(entry&&entry.intentHash!==mark){if(!entry.selector){storage.removeItem(name);entry=null;}else{const prior=await recover(entry);auth.assert(t);if(!terminal(prior))fail('COMMAND_PENDING');storage.removeItem(name);fail('COMMAND_NOT_ADMITTED');}}
   if(entry?.selector){
    try{const answer=await recover(entry);auth.assert(t);if(terminal(answer)&&!retain)storage.removeItem(name);return answer;}
    catch(error){if(error.code!=='COMMAND_NOT_ADMITTED')throw error;}
    // The selector is saved before admission. A failed admission can leave no
    // protected request to resume. Re-prepare the SAME ID and require the exact
    // original digest before retrying admission; never silently create a new ID.
    const raw=await prepare(entry.operationId);auth.assert(t);
    const retryTransport={submit:(v,prepared)=>transport.submit(v,key=>{if(canonical(key)!==canonical(entry.selector))fail('STALE_REVISION');prepared(key)}),status:key=>transport.status(key),process:key=>transport.process(key)};
    const retry=new CommandService(auth,retryTransport);
    try{const answer=await retry.execute(raw);auth.assert(t);if(terminal(answer)&&!retain)storage.removeItem(name);return answer;}finally{retry.dispose();}
   }
   // Without a persisted selector, admission has not started. A fresh attempt
   // must not reuse an expired preparation left by a reload or validation error.
   entry={operationId:operation(),intentHash:mark,selector:null};save(name,entry);
   let raw;try{raw=await prepare(entry.operationId);auth.assert(t);}catch(e){storage.removeItem(name);throw e;}
   // Store the opaque recovery selector before admission can reach the Core.
   const originalSubmit=transport.submit.bind(transport);
   const wrapped={submit:(v,prepared)=>originalSubmit(v,key=>{auth.assert(t);entry.selector=key;save(name,entry);prepared(key)}),status:key=>transport.status(key),process:key=>transport.process(key)};
   const single=new CommandService(auth,wrapped);
   try{const answer=await single.execute(raw);auth.assert(t);if(terminal(answer)&&!retain)storage.removeItem(name);return answer;}catch(e){if(!entry.selector)storage.removeItem(name);throw e;}finally{single.dispose();}
  }finally{active.delete(name)}
 }
 async function readChamada(activityId){let result=await request('core.act.portal-read.v13',{route:'/atividades/chamada',activityId});if(result.needsRoster){need(result.activityId===activityId&&Number.isInteger(result.membersEpoch));const applied=await execute('roster:'+activityId,{activityId,membersEpoch:result.membersEpoch},operationId=>request('core.act.roster-prepare.v13',{operationId,activityRef:{kind:'activity',id:activityId}}));if(applied.businessState!=='COMMIT_APPLIED')fail('COMMAND_PENDING');result=await request('core.act.portal-read.v13',{route:'/atividades/chamada',activityId});}need(result.ok&&Number.isInteger(result.canonical?.collectionRevision));const t=auth.capture();generations.set(activityId,{revision:result.canonical.collectionRevision,uid:t.uid,generation:t.generation});return result;}
 async function saveChamada(payload){const t=auth.capture(),v=generations.get(payload.idAtividade);if(!v||v.uid!==t.uid||v.generation!==t.generation)fail('STALE_AUTHORIZATION');const answer=await execute('chamada:'+payload.idAtividade,payload,operationId=>request('core.act.portal-prepare.v13',{operationId,payload,collectionRevision:v.revision}));auth.assert(t);if(answer.businessState==='COMMIT_APPLIED'&&Number.isInteger(answer.result?.collectionRevision))generations.set(payload.idAtividade,{...v,revision:answer.result.collectionRevision});return answer;}
 function legacyResult(answer){if(answer.businessState==='REJECTED')return actDomainFailure(answer.error);if(answer.businessState!=='COMMIT_APPLIED')return {ok:false,errorCode:'COMMAND_PENDING',message:'Aguardando confirmação da operação. Tente novamente para consultar o mesmo envio.'};return {ok:true,data:answer.result};}
 const directReaders=createDirectReaders({auth,db,sdk,record:v=>{if(typeof window!=='undefined'){const rows=window.__geapaReadTimings||(window.__geapaReadTimings=[]);rows.push({...v,at:Date.now()});if(rows.length>30)rows.shift();}}});
 const activityReader=createActivityReader({auth,db,sdk,root:ROOT,request});
 async function apiGet(route,params={},context){const direct=enabled&&!fixtureUid(auth.uid)?await directReaders.route(route,params).catch(e=>{if(['permission-denied','not-found','SCHEMA_UNSUPPORTED'].includes(e.code))return null;throw e;}):null;if(direct)return direct;if(['/v2/apresentacoes/material/arquivo','/v2/justificativas/documento/arquivo','/v2/apresentacoes/foto/arquivo'].includes(route)){const proof=route==='/v2/justificativas/documento/arquivo',photo=route==='/v2/apresentacoes/foto/arquivo';exact(params,[proof?'idJustificativa':'idApresentacao']);const ticket=auth.capture(),answer=await request(proof||photo?'core.act.portal-read.v13':'core.act.material-download.v13',proof||photo?{route,params}:params);auth.assert(ticket);need(answer.ok&&answer.data);const d=answer.data;exact(d,['name','mime','size','sha256','contentBase64']);need(typeof d.name==='string'&&d.name.length<=240&&['application/pdf','application/vnd.ms-powerpoint','application/vnd.openxmlformats-officedocument.presentationml.presentation','application/vnd.oasis.opendocument.presentation','image/jpeg','image/png','image/webp','application/msword','application/vnd.openxmlformats-officedocument.wordprocessingml.document'].includes(d.mime)&&Number.isSafeInteger(d.size)&&d.size>0&&d.size<=25*1024*1024&&typeof d.contentBase64==='string'&&d.contentBase64.length===Math.ceil(d.size/3)*4);need(proof?proofMimeTypes[d.name.split('.').pop().toLowerCase()]===d.mime:photo?['image/jpeg','image/png','image/webp'].includes(d.mime):!['application/msword','application/vnd.openxmlformats-officedocument.wordprocessingml.document','image/webp'].includes(d.mime));const bytes=Uint8Array.from(atob(d.contentBase64),c=>c.charCodeAt(0));need(bytes.length===d.size&&[...new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))].map(x=>x.toString(16).padStart(2,'0')).join('')===d.sha256);auth.assert(ticket);return answer;}if(route==='/atividades/chamada')return readChamada(params.idAtividade);const ticket=auth.capture(),answer=await request('core.act.portal-read.v13',{route,params});auth.assert(ticket);if(['/v2/minhas-apresentacoes','/v2/apresentacoes/pendencias'].includes(route)&&answer.ok)for(const row of answer.data.apresentacoes||answer.data.pendencias||[]){exact(row.canonicalBasis,['unitRevision','thematicRevision']);need(Number.isInteger(row.canonicalBasis.unitRevision)&&Number.isInteger(row.canonicalBasis.thematicRevision));presentationGenerations.set(row.idApresentacao,{basis:row.canonicalBasis,photoRevision:row.canonicalPhotoRevision,uid:ticket.uid,generation:ticket.generation});if(row.canonicalMaterialBasis){exact(row.canonicalMaterialBasis,['unitRevision','requirementRevision']);materialGenerations.set(row.idApresentacao,{basis:row.canonicalMaterialBasis,opportunityId:row.canonicalMaterialOpportunityId||null,termsRevision:row.canonicalCorrectionTermsRevision||null,uid:ticket.uid,generation:ticket.generation});}}return answer;}
 async function photoAction(route,payload){
  delete payload.requestId;const ticket=auth.capture(),shown=presentationGenerations.get(payload.idApresentacao);need(shown&&shown.uid===ticket.uid&&shown.generation===ticket.generation&&Number.isSafeInteger(shown.photoRevision));const basis={photoRevision:shown.photoRevision};
  if(route==='/v2/apresentacoes/foto/revisar')return legacyResult(await execute('photo-review:'+payload.idApresentacao,payload,operationId=>request('core.act.portal-prepare.v13',{operationId,route,payload,basis})));
  exact(payload,['idApresentacao','nomeArquivoOriginal','mimeType','conteudoBase64','linkArquivo','reenvio','observacoes']);need(typeof payload.conteudoBase64==='string'&&payload.conteudoBase64.length<=34952536);const submit={idApresentacao:payload.idApresentacao,observacoes:payload.observacoes};
  if(!payload.conteudoBase64)return legacyResult(await execute('photo-link:'+payload.idApresentacao,{...submit,linkArquivo:payload.linkArquivo},operationId=>request('core.act.portal-prepare.v13',{operationId,route,payload:{...submit,linkArquivo:payload.linkArquivo},basis})));
  const mime={jpg:'image/jpeg',jpeg:'image/jpeg',png:'image/png',webp:'image/webp'}[String(payload.nomeArquivoOriginal).split('.').pop().toLowerCase()];need(mime&&(!payload.mimeType||payload.mimeType===mime));let bytes;try{bytes=Uint8Array.from(atob(payload.conteudoBase64),c=>c.charCodeAt(0))}catch{fail('VALIDATION_ERROR')}need(bytes.length>0&&bytes.length<=25*1024*1024);const sha256=[...new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))].map(x=>x.toString(16).padStart(2,'0')).join('');auth.assert(ticket);
  const fileDescriptor={name:payload.nomeArquivoOriginal,mime,size:bytes.length,sha256},intent={...submit,fileDescriptor},slot='photo:'+payload.idApresentacao,uploadSlot=slot+':upload',submitSlot=slot+':submit',lock=(await journalName('flow:'+slot)).name;if(active.has(lock))fail('COMMAND_PENDING');active.add(lock);
  try{let answer;const pending=load((await journalName(submitSlot)).name);if(pending?.selector)answer=await execute(submitSlot,intent,()=>fail('COMMAND_RECONCILIATION_REQUIRED'),true);else{
   const started=await execute(uploadSlot,intent,operationId=>request('core.act.portal-prepare.v13',{operationId,route:'/v2/apresentacoes/foto/iniciar',payload:{idApresentacao:payload.idApresentacao,fileDescriptor},basis}),true);
   if(started.businessState!=='COMMIT_APPLIED'){if(terminal(started))storage.removeItem((await journalName(uploadSlot)).name);return legacyResult(started);}const effectRef=started.result.effectRef;need(effectRef?.kind==='outbox');const uploaded=await request('core.act.workspace-upload.v13',{effectRef,contentBase64:payload.conteudoBase64});if(uploaded.state!=='CONFIRMED')return {ok:false,errorCode:'COMMAND_PENDING',message:'A foto está sendo confirmada. Repita o envio para consultar a mesma operação.'};
   answer=await execute(submitSlot,intent,operationId=>request('core.act.portal-prepare.v13',{operationId,route,payload:{...submit,effectRef},basis}),true);
  }auth.assert(ticket);if(terminal(answer))for(const stage of [uploadSlot,submitSlot])storage.removeItem((await journalName(stage)).name);return legacyResult(answer);}finally{active.delete(lock)}
 }
 async function sendJustificationProof(payload){
  delete payload.requestId;const file=payload.documentoComprobatorio;exact(file,['nomeArquivo','mimeType','conteudoBase64']);
  const ticket=auth.capture(),mime=proofMimeTypes[String(file.nomeArquivo).split('.').pop().toLowerCase()];need(typeof file.nomeArquivo==='string'&&file.nomeArquivo.length<=240&&mime&&(!file.mimeType||file.mimeType===mime)&&typeof file.conteudoBase64==='string'&&file.conteudoBase64.length<=34952536);
  let binary;try{binary=atob(file.conteudoBase64);}catch{fail('VALIDATION_ERROR')}need(binary.length>0&&binary.length<=25*1024*1024);const bytes=Uint8Array.from(binary,c=>c.charCodeAt(0)),sha256=[...new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))].map(x=>x.toString(16).padStart(2,'0')).join('');auth.assert(ticket);
  const fileDescriptor={name:file.nomeArquivo,mime,size:bytes.length,sha256},finalPayload={...payload,linkDocumentoComprobatorio:''};delete finalPayload.documentoComprobatorio;
  const intent={payload:finalPayload,fileDescriptor},slot='justification-proof:'+payload.idAtividade,uploadSlot=slot+':upload',submitSlot=slot+':submit',lock=(await journalName('flow:'+slot)).name;
  if(active.has(lock))fail('COMMAND_PENDING');active.add(lock);
  try{let answer;const prior=load((await journalName(submitSlot)).name);
   if(prior?.selector)answer=await execute(submitSlot,intent,()=>fail('COMMAND_RECONCILIATION_REQUIRED'),true);
   else{const started=await execute(uploadSlot,intent,operationId=>request('core.act.portal-prepare.v13',{operationId,route:'/v2/justificativas/documento/iniciar',payload:{idAtividade:payload.idAtividade,idRegistroPresenca:payload.idRegistroPresenca||'',fileDescriptor,confirmouCienciaForaPrazo:payload.confirmouCienciaForaPrazo}}),true);
    if(started.businessState!=='COMMIT_APPLIED')return legacyResult(started);need(started.result?.effectRef?.kind==='outbox');const proofUploadRef=started.result.effectRef;
    const uploaded=await request('core.act.workspace-upload.v13',{effectRef:proofUploadRef,contentBase64:file.conteudoBase64});if(uploaded.state!=='CONFIRMED')return {ok:false,errorCode:'COMMAND_PENDING',message:'O comprovante está sendo confirmado. Repita o envio para consultar a mesma operação.'};
    answer=await execute(submitSlot,intent,operationId=>request('core.act.portal-prepare.v13',{operationId,route:'/v2/justificativas/enviar',payload:{...finalPayload,proofUploadRef}}),true);
   }
   auth.assert(ticket);if(terminal(answer))for(const key of [uploadSlot,submitSlot]){const journal=await journalName(key);storage.removeItem(journal.name);}return legacyResult(answer);
  }finally{active.delete(lock)}
 }
 async function sendMaterial(payload){
  delete payload.requestId;exact(payload,['idApresentacao','nomeArquivoOriginal','mimeType','conteudoBase64','observacoes',...('correctionNotice'in payload?['correctionNotice']:[])]);
  const ticket=auth.capture(),shown=materialGenerations.get(payload.idApresentacao);
  if(!shown||shown.uid!==ticket.uid||shown.generation!==ticket.generation)fail('STALE_AUTHORIZATION');
  need(typeof payload.conteudoBase64==='string'&&payload.conteudoBase64.length<=34952536&&typeof payload.nomeArquivoOriginal==='string'&&payload.nomeArquivoOriginal.length<=240);
  let binary;try{binary=atob(payload.conteudoBase64);}catch{fail('VALIDATION_ERROR')}
  need(binary.length>0&&binary.length<=25*1024*1024);const bytes=Uint8Array.from(binary,c=>c.charCodeAt(0));
  const sha256=[...new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))].map(x=>x.toString(16).padStart(2,'0')).join('');auth.assert(ticket);
  const mimeByExtension={pdf:'application/pdf',ppt:'application/vnd.ms-powerpoint',pptx:'application/vnd.openxmlformats-officedocument.presentationml.presentation',odp:'application/vnd.oasis.opendocument.presentation'};
  const mime=mimeByExtension[payload.nomeArquivoOriginal.split('.').pop().toLowerCase()];need(mime&&(!payload.mimeType||payload.mimeType===mime));
  const intent={idApresentacao:payload.idApresentacao,opportunityId:shown.opportunityId||null,fileDescriptor:{name:payload.nomeArquivoOriginal,mime,size:bytes.length,sha256},notes:payload.observacoes||''},slot='material:'+payload.idApresentacao,{name:lockName}=await journalName('flow:'+slot);
  if(active.has(lockName))fail('COMMAND_PENDING');active.add(lockName);
  const uploadSlot=slot+':upload',submitSlot=slot+':submit';
  try{
   const submitJournal=await journalName(submitSlot),prior=load(submitJournal.name);let answer;
   // A lost final response is recovered before rechecking a now-consumed upload
   // or an answered correction opportunity. Only opaque selectors are stored.
   if(prior?.selector)answer=await execute(submitSlot,intent,()=>fail('COMMAND_RECONCILIATION_REQUIRED'),true);
   else{
    let basis=shown.basis;const pendingUpload=load((await journalName(uploadSlot)).name);if(payload.correctionNotice&&!pendingUpload?.selector){exact(payload.correctionNotice,['opportunityId','opportunityRevision','dueAt','acknowledged']);need(payload.correctionNotice.acknowledged===true);const notice=await request('core.act.material-notice.v13',{idApresentacao:payload.idApresentacao,...payload.correctionNotice});need(notice.ok&&notice.opportunityId===payload.correctionNotice.opportunityId);basis=notice.basis;auth.assert(ticket);}
    const started=await execute(uploadSlot,intent,operationId=>request('core.act.material-begin.v13',{operationId,idApresentacao:intent.idApresentacao,basis,fileDescriptor:intent.fileDescriptor}),true);
    if(started.businessState!=='COMMIT_APPLIED')return legacyResult(started);
    need(started.result?.effectRef?.kind==='outbox');
    const effectRef=started.result.effectRef,result=await request('core.act.workspace-upload.v13',{effectRef,contentBase64:payload.conteudoBase64});
    if(result.state!=='CONFIRMED')return {ok:false,errorCode:'COMMAND_PENDING',message:'O arquivo está sendo confirmado. Repita o envio para consultar a mesma operação.'};
    answer=await execute(submitSlot,intent,operationId=>request('core.act.material-submit.v13',{operationId,effectRef,notes:intent.notes}),true);
   }
   auth.assert(ticket);if(terminal(answer)){for(const key of [uploadSlot,submitSlot]){const journal=await journalName(key);auth.assert(ticket);storage.removeItem(journal.name);}}
   return legacyResult(answer);
  }finally{active.delete(lockName)}
 }
 async function reviewMaterial(payload){
  delete payload.requestId;exact(payload,['idApresentacao','decisao','observacaoPublica','observacaoInterna',...('prazoCorrecao'in payload?['prazoCorrecao']:[])]);need(['APROVAR','DISPENSAR','SOLICITAR_AJUSTE'].includes(payload.decisao));if(payload.decisao==='SOLICITAR_AJUSTE')need(typeof payload.prazoCorrecao==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(payload.prazoCorrecao));
  const ticket=auth.capture(),shown=materialGenerations.get(payload.idApresentacao);if(!shown||shown.uid!==ticket.uid||shown.generation!==ticket.generation)fail('STALE_AUTHORIZATION');
  const root='material-review:'+payload.idApresentacao,{name:lockName}=await journalName('flow:'+root);if(active.has(lockName))fail('COMMAND_PENDING');active.add(lockName);
  const stages=payload.decisao==='DISPENSAR'?['REVIEW']:payload.decisao==='SOLICITAR_AJUSTE'?(shown.termsRevision?['TERMS']:['REVIEW','REQUEST','OPPORTUNITY']):['REVIEW','RECEIVE','EVALUATE'];let priorSelector,answer;
  try{for(const stage of stages){
    answer=await execute(root+':'+stage,payload,operationId=>request('core.act.material-review.v13',{operationId,idApresentacao:payload.idApresentacao,basis:shown.basis,decision:payload.decisao,publicReason:payload.observacaoPublica,internalReason:payload.observacaoInterna,stage,...(payload.decisao==='SOLICITAR_AJUSTE'?{dueCivilDate:payload.prazoCorrecao,...(shown.termsRevision?{opportunityRevision:shown.termsRevision}:{})}:{}),...(priorSelector?{priorSelector}:{})}),true);
    if(answer.businessState!=='COMMIT_APPLIED')break;
    const journal=await journalName(root+':'+stage);priorSelector=load(journal.name)?.selector;need(priorSelector);
   }
   auth.assert(ticket);if(terminal(answer)){for(const stage of stages){const journal=await journalName(root+':'+stage);auth.assert(ticket);storage.removeItem(journal.name);}}
   return legacyResult(answer);
  }finally{active.delete(lockName)}
 }
 async function apiPost(route,params={}){
  let payload;try{payload=JSON.parse(params.payload);}catch{fail('SCHEMA_UNSUPPORTED')}
  if(['/v2/apresentacoes/foto/registrar','/v2/apresentacoes/foto/revisar'].includes(route))return photoAction(route,payload);
  if(route==='/v2/justificativas/enviar'&&payload.documentoComprobatorio)return sendJustificationProof(payload);
  if(route==='/v2/apresentacoes/material/registrar')return sendMaterial(payload);
  if(route==='/v2/apresentacoes/material/revisar')return reviewMaterial(payload);
  if(route==='/atividades/modelo/validar')return request('core.act.portal-preview.v13',{route,payload});
  if(route==='/atividades/chamada/salvar')return legacyResult(await saveChamada(payload));
  // Legacy requestId is a UI correlation hint. The shared journal owns recovery,
  // so a reload-generated hint must not change the protected business intent.
  let basis;if(route.startsWith('/v2/apresentacoes/')){delete payload.requestId;const ticket=auth.capture(),prior=presentationGenerations.get(payload.idApresentacao);if(!prior||prior.uid!==ticket.uid||prior.generation!==ticket.generation)fail('STALE_AUTHORIZATION');basis=prior.basis;}
  const answer=await execute((route.startsWith('/v2/apresentacoes/')?'presentation:':route.startsWith('/v2/justificativas/')?'justification:':'activity:')+(payload.idApresentacao||payload.idJustificativa||payload.idAtividade||'new'),{route,payload},operationId=>request('core.act.portal-prepare.v13',{operationId,route,payload,...(basis?{basis}:{})}));
  return legacyResult(answer);
 }
 function clear(){directReaders.clear();generations.clear();presentationGenerations.clear();materialGenerations.clear();commands.clear();}
 return frozen({invalidateReads:()=>directReaders.clear(),memberOwn:async view=>{if(fixtureUid(auth.uid))return decodeOwnRead(await request('core.members.own-read.v13',{view}),view);try{return decodeOwnRead((await directReaders.read('self'))[view],view)}catch(e){if(['permission-denied','not-found','SCHEMA_UNSUPPORTED'].includes(e.code))return decodeOwnRead(await request('core.members.own-read.v13',{view}),view);throw e;}},context:async force=>{if(fixtureUid(auth.uid))return request('core.act.context.v13',{});try{return await directReaders.read('context',force)}catch(e){if(['permission-denied','not-found','SCHEMA_UNSUPPORTED'].includes(e.code))return request('core.act.context.v13',{});throw e;}},apiGet,apiPost:async(...args)=>{try{return await apiPost(...args)}finally{directReaders.clear()}},read:input=>request('core.act.read.v13',input),command:input=>execute(input.commandType+':'+canonical(input.payload),input,operationId=>request('core.act.prepare.v13',{...input,operationId})),readChamada,saveChamada,clear,dispose:()=>{clear();directReaders.dispose();commands.dispose();}});
}

