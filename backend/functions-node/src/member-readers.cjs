'use strict';
// Transport and derived-write limits only; private DTOs and policies remain Core.
const routes=Object.freeze(['/members/own','/members/operational-context','/meu-vinculo/acompanhamento','/meu-perfil/correcoes','/admin/correcoes-cadastrais','/admin/correcoes-cadastrais/detalhe','/admin/membros','/admin/ingressos-membros','/admin/ingressos-membros/catalogos','/meu-vinculo/opcoes','/meu-vinculo/solicitacoes','/admin/solicitacoes-vinculo','/admin/solicitacoes-vinculo/detalhe']);
const projections={'/members/own':'self','/members/operational-context':'members_context','/admin/ingressos-membros/catalogos':'members_catalog','/meu-vinculo/opcoes':'members_self','/meu-vinculo/solicitacoes':'members_self'};
function validateWrites(s,request,uid){
 const {fail}=require('./environment.cjs'),{decode}=require('./firestore.cjs');
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
