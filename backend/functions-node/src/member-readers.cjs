'use strict';
// Transport and derived-write limits only; private DTOs and policies remain Core.
const routes=Object.freeze(['/admin/frequencia-rh/providencias','/admin/pessoas','/admin/pessoas/detalhe','/members/own','/members/operational-context','/meu-vinculo/acompanhamento','/meu-perfil/correcoes','/admin/correcoes-cadastrais','/admin/correcoes-cadastrais/detalhe','/admin/membros','/admin/ingressos-membros','/admin/ingressos-membros/catalogos','/meu-vinculo/opcoes','/meu-vinculo/solicitacoes','/admin/solicitacoes-vinculo','/admin/solicitacoes-vinculo/detalhe','/meu-vinculo/notificacoes','/meu-vinculo/notificacoes/leitura']);
const projections={'/members/own':'self','/members/operational-context':'members_context','/admin/ingressos-membros/catalogos':'members_catalog','/meu-vinculo/opcoes':'members_self','/meu-vinculo/solicitacoes':'members_self'};
// Validate the exact receipt transitions against documents read by the same transaction.
function validateNotificationWrites(s,params,uid){
 const {fail}=require('./environment.cjs'),{decode}=require('./firestore.cjs'),{isDeepStrictEqual:equal}=require('node:util');
 const need=ok=>{if(!ok)fail('PILOT_WRITE_DENIED');};
 if(!s.writes.length)return;
 const before=path=>{const encoded=s.cache.get(path)?.encoded;need(encoded);return decode(encoded);};
 const binding=before('memberAuthBindings/'+uid),person=binding.idPessoa;
 need(binding.active===true&&binding.uid===uid&&s.writes.length<=101);
 const path='memberNotificationIndexes/'+person,old=before(path);
 need(old.personId===person&&old.id===person&&old.notificationIds.length<=100);
 const rows=old.notificationIds.map(id=>before('memberNotifications/'+id));
 need(rows.every((row,i)=>row.id===old.notificationIds[i]&&row.personId===person));
 const all=params.action==='MARK_ALL_READ';
 need(all||params.action==='MARK_READ'&&old.notificationIds.includes(params.notificationId));
 const pending=rows.filter(row=>!row.read),selected=pending.filter(row=>all||row.id===params.notificationId);
 const writes=new Map(s.writes.map(w=>[w.path,w]));
 const after=path=>{const w=writes.get(path);need(w?.mode==='set');return decode(w.data);};
 need(writes.size===selected.length+1&&writes.size===s.writes.length);
 const ix=after(path),now=ix.updatedAt;
 need(Number.isSafeInteger(now)&&equal(ix,{...old,revision:old.revision+1,updatedAt:now,unreadCount:pending.length-selected.length}));
 for(const row of selected)need(equal(after('memberNotifications/'+row.id),{...row,revision:row.revision+1,read:true,readAt:now}));
}
function validateWrites(s,request,uid){
 const {fail}=require('./environment.cjs'),{decode}=require('./firestore.cjs');
 if(request.route==='/meu-vinculo/notificacoes/leitura')return validateNotificationWrites(s,request.params,uid);
 if(s.writes.length>1)fail('PILOT_WRITE_DENIED');
 for(const w of s.writes){const d=decode(w.data),id=projections[request.route];
  if(w.mode!=='set'||Buffer.byteLength(JSON.stringify(w.data))>=700000)fail('PILOT_WRITE_DENIED');
  if(id&&w.path==='portalReaderViews/'+uid+'/items/'+id&&d.schemaVersion==='portal-reader-view-v1'&&d.viewId===id&&d.basis?.ownerUid===uid)continue;
  // Revealing a correction's private fields must retain its canonical read audit.
  if(request.route==='/admin/correcoes-cadastrais/detalhe'&&request.params?.revelarDados===true&&/^memberAudit\/[a-f0-9]{32}$/.test(w.path)&&d.schemaVersion==='member-profile-correction-read-v1'&&d.actorUid===uid&&d.requestId===request.params.idSolicitacao)continue;
  fail('PILOT_WRITE_DENIED');
 }
}
module.exports={routes,projections,validateWrites};
