'use strict';
// Transport admission for a bounded synthetic SELF rehearsal, never a policy.
// Canonical producers still validate binding, membership and every requested DTO.
const routes=Object.freeze(['/act/context','/members/own','/members/operational-context','/meu-perfil/correcoes','/meu-vinculo/opcoes','/meu-vinculo/solicitacoes','/v2/minha-frequencia','/v2/minhas-apresentacoes','/v2/justificativas/pendencias','/v2/apresentacoes/pendencias','/atividades/listar','/atividades/bundle','/atividades/detalhes-preload','/atividades/detalhe']);
function admitted(read,uid,route,now,production){
 if(production||!routes.includes(route))return false;
 const g=read('coreDevLinkedReaderGrants/'+uid);
 if(g?.schemaVersion!=='core-dev-linked-reader-grant-v1'||g.projectId!=='geapa-convergencia-dev'||g.active!==true||g.uid!==uid||!/^DEVADM_[a-f0-9]{12}_[12]$/.test(g.personId)||!Number.isSafeInteger(g.revision)||g.revision<1||!Number.isSafeInteger(g.startsAt)||!Number.isSafeInteger(g.expiresAt)||!(g.startsAt<=now&&now<g.expiresAt)||g.expiresAt-g.startsAt>4*3600000)return false;
 const b=read('memberAuthBindings/'+uid),p=read('memberIdentities/'+g.personId);
 return b?.active===true&&b.idPessoa===g.personId&&b.bindingVersion===g.bindingVersion&&p?.idPessoa===g.personId&&p.active===true&&p.origin==='DEV_ADMIN_SYNTHETIC_V1';
}
module.exports={admitted,routes};
