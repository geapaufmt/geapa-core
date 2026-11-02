// Informational catalog only; never accepted as proof or used to grant a command.
function f4MemberAccessCatalog_(board,control,policies,now){
 const N=f4Core_().need;N(board?.ok&&board.data?.schemaVersion==='board-competencies-v1','ACCESS_REVOKED');
 const relations=[['MEMBRO_EFETIVO','Membro efetivo'],['MEMBRO_INGRESSANTE','Membro ingressante'],['MEMBRO_EM_ESPERA','Membro em espera'],['SUSPENSO','Membro suspenso'],['EGRESSO','Egresso'],['ACADEMIC_COLLABORATOR','Colaborador acadêmico'],['EXTERNAL_PARTICIPANT','Participante externo']].map(([id,label])=>({id,label,condition:'Depende de relação canônica documentada; não concede automaticamente função ou acesso privado.'}));
 const historical=['ADMIN','DIRETORIA','SECRETARIA','COMUNICACAO','CONSELHO','COLABORADOR','MEMBRO','EGRESSO','EXTERNO','VISITANTE'];
 const profiles=historical.map(id=>({id,label:id.replaceAll('_',' '),source:'Catálogo documental histórico',status:'DOCUMENTARY_ONLY',permissions:[],condition:'O rótulo histórico não é uma autorização operacional. Nenhum direito é inferido deste perfil.'}));
 const selfProfiles={MEMBRO:['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE'],EGRESSO:['EGRESSO']};
 for(const [key,states] of Object.entries(selfProfiles)){
  const profile=profiles.find(p=>p.id===key),self=control.selfReadPolicy,C=f4Core_();
  const applicable=policies.filter(p=>p.active&&p.proofKind==='SELF'&&C.ms(p.issuedAt)<=now&&now<C.ms(p.expiresAt));
  const caps=key==='MEMBRO'?[...new Set(applicable.flatMap(p=>p.capabilities))]:[];
  if(C.ms(self.issuedAt)<=now&&now<C.ms(self.expiresAt)&&states.some(s=>self.eligibleStates.includes(s)))caps.push(self.capability);
  profile.status='CONDITIONAL_SELF_CATALOG';profile.source='Contratos SELF e policies vigentes';profile.condition='Consulta de catálogo: cada operação exige o próprio titular, estado elegível, binding, recurso e policy aplicáveis. Não comprova acesso de uma pessoa específica.';
  profile.permissions=[...new Set(caps)].map(cap=>({capability:cap,label:f4BoardCapabilityLabel_(cap),module:cap.startsWith('ACTIVITY_')?'Atividades':cap.startsWith('PRESENTATION_')?'Apresentações':'Membros',action:cap.endsWith('_READ')?'Consulta':'Operação própria',status:'NOT_ENABLED',reason:'Policy SELF configurada; elegibilidade individual e habilitação da rota devem ser verificadas na operação.',limits:{cycleIds:[],nominalPersons:true,nominalResources:true,policyValidUntil:null}}));
 }
 // Keep each incumbent separate: nominal policies cannot be unioned into a
 // fictitious global DIRETORIA grant or inherited by another office holder.
 const roles=board.data.items.map(x=>({roleKey:x.roleKey,name:x.name,personId:x.personId,boardId:x.boardId,startDate:x.startDate,endDate:x.endDate,duties:x.duties,permissions:x.permissions}));
 for(const role of roles)profiles.push({id:'ROLE_'+role.roleKey+'_'+role.personId,label:role.roleKey.replaceAll('_',' '),source:'Core e policies vigentes',status:'RESOLVED_ROLE',holder:role.name,permissions:role.permissions,condition:'Função vigente, vínculo elegível, binding, recurso, estado, campos, policy e rota são revalidados em cada operação.'});
 return {ok:true,data:{schemaVersion:'people-access-catalog-v1',asOf:board.data.asOf,validUntil:board.data.validUntil,relations,roles,profiles,notice:board.data.notice,otherFunctions:{groups:['ASSESSORIA','CONSELHO'],condition:'Exigem nomeação canônica vigente e policy específica. Não herdam as competências da Diretoria. Colaborações acadêmicas são relações documentais, não cargos diretivos.'}}};
}
