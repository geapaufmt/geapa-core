// Active Portal consumers only. This catalog selects adapters, never authority.
function f4ActRemainingConsumers_(){
 const commands=['CYCLE_THEME_SETTINGS_RECORD','CYCLE_MATERIAL_SETTINGS_RECORD','ACTIVITY_TYPE_SAVE','ACTIVITY_POLICY_SAVE','ACTIVITY_SEMESTER_SAVE','CYCLE_CREATE','CYCLE_GOVERNANCE_EDIT','CONFIG_RECORD','CONFIG_ACTIVATE','DYNAMIC_SUBMIT','DYNAMIC_DECIDE','DYNAMIC_CANCEL','PUBLICATION_APPROVE','PUBLICATION_WITHDRAW','ARTIFACT_UPLOAD_REQUEST','MATERIAL_SUBMIT','MATERIAL_CORRECTION_SUBMIT','MATERIAL_REVIEW','MATERIAL_RECEIVE','MATERIAL_FINAL_EVALUATE','MATERIAL_EXCEPTION','MATERIAL_CORRECTION_REQUEST','CORRECTION_OPPORTUNITY_OPEN','CORRECTION_TERMS_REVISE','PHOTO_UPLOAD_REQUEST','PHOTO_SUBMIT','PHOTO_REVIEW','ATTENDANCE_PROOF_UPLOAD_REQUEST'];
 const routes=['/admin/apresentacoes/regras/salvar','/v2/apresentacoes/material/iniciar','/v2/apresentacoes/material/submeter','/v2/apresentacoes/material/revisar','/admin/atividades/configuracoes/salvar','/admin/ciclos-semestres/semestre','/admin/ciclos-semestres/ciclo','/admin/apresentacoes/configuracao/salvar','/admin/apresentacoes/entrega-publicacao/salvar','/v2/apresentacoes/dinamica','/v2/apresentacoes/foto/iniciar','/v2/apresentacoes/foto/registrar','/v2/apresentacoes/foto/revisar','/v2/justificativas/documento/iniciar'];
 function type(input,read){const p=input.payload||{},r=input.route;
  if(r==='/v2/apresentacoes/material/iniciar')return 'ARTIFACT_UPLOAD_REQUEST';
  if(r==='/v2/apresentacoes/material/submeter'){
   if(typeof read!=='function')return null;
   f4Core_().need(p.effectRef?.kind==='outbox'&&f4ActivitiesModel_().id(p.effectRef.id),'VALIDATION_ERROR');
   const out=read(f4ActivitiesModel_().path(p.effectRef)),prov=out&&read('commandProvenance/'+out.createdByOperation),ticket=prov&&read('memberAdmissionTickets/'+prov.ticketId);
   f4Core_().need(out&&ticket,'UNRESOLVED_DEPENDENCY');
   return ticket.requestEnvelope.payload.opportunityRef?'MATERIAL_CORRECTION_SUBMIT':'MATERIAL_SUBMIT';
  }
  if(r==='/v2/apresentacoes/material/revisar')return p.stage==='REVIEW'?(p.decision==='DISPENSAR'?'MATERIAL_EXCEPTION':'MATERIAL_REVIEW'):{RECEIVE:'MATERIAL_RECEIVE',EVALUATE:'MATERIAL_FINAL_EVALUATE',REQUEST:'MATERIAL_CORRECTION_REQUEST',OPPORTUNITY:'CORRECTION_OPPORTUNITY_OPEN',TERMS:'CORRECTION_TERMS_REVISE'}[p.stage];
  if(r==='/admin/atividades/configuracoes/salvar')return {models:'ACTIVITY_TYPE_SAVE',windows:'ACTIVITY_POLICY_SAVE'}[p.kind];
  if(r==='/admin/ciclos-semestres/semestre')return 'ACTIVITY_SEMESTER_SAVE';
  if(r==='/admin/ciclos-semestres/ciclo')return p.id?'CYCLE_GOVERNANCE_EDIT':'CYCLE_CREATE';
  if(r==='/admin/apresentacoes/configuracao/salvar')return {RECORD:'CONFIG_RECORD',ACTIVATE:'CONFIG_ACTIVATE'}[p.kind];
  if(r==='/admin/apresentacoes/regras/salvar')return 'CYCLE_THEME_SETTINGS_RECORD';
  if(r==='/admin/apresentacoes/entrega-publicacao/salvar')return p.section==='MATERIAL'?'CYCLE_MATERIAL_SETTINGS_RECORD':'CONFIG_RECORD';
  if(r==='/v2/apresentacoes/dinamica')return {SOLICITAR:'DYNAMIC_SUBMIT',DECIDIR:'DYNAMIC_DECIDE',DESISTIR:'DYNAMIC_CANCEL',REVOGAR:'DYNAMIC_CANCEL'}[p.acao];
  if(r==='/v2/apresentacoes/finalizacao')return {PUBLICAR:'PUBLICATION_APPROVE',RETIRAR:'PUBLICATION_WITHDRAW'}[p.acao];
  return {'/v2/apresentacoes/foto/iniciar':'PHOTO_UPLOAD_REQUEST','/v2/apresentacoes/foto/registrar':'PHOTO_SUBMIT','/v2/apresentacoes/foto/revisar':'PHOTO_REVIEW','/v2/justificativas/documento/iniciar':'ATTENDANCE_PROOF_UPLOAD_REQUEST'}[r]||null;
 }
 function prepare(options,input){const r=input.route;
  if(r.startsWith('/v2/apresentacoes/material/')){
   f4Core_().exact(input,['operationId','route','payload']);const value={...input.payload,operationId:input.operationId};
   if(r==='/v2/apresentacoes/material/iniciar')return f4ActivitiesMaterialCompat_(options).begin(value);
   if(r==='/v2/apresentacoes/material/submeter')return f4ActivitiesMaterialCompat_(options).submit(value);
   if(r==='/v2/apresentacoes/material/revisar')return f4ActivitiesMaterialReviewCompat_(options).prepare(value);
  }
  if(r==='/admin/atividades/configuracoes/salvar')return f4ActivitiesActivityConfigCompat_(options).prepare(input);
  if(r.startsWith('/admin/ciclos-semestres/'))return f4ActivitiesGovernanceCompat_(options).prepare(input);
  if(r==='/admin/apresentacoes/configuracao/salvar')return f4ActivitiesConfigCompat_(options).prepare(input);
  if(r==='/admin/apresentacoes/regras/salvar')return f4ActivitiesCyclePoliciesCompat_(options).prepare(input);
  if(r==='/admin/apresentacoes/entrega-publicacao/salvar')return f4ActivitiesDeliveryPoliciesCompat_(options).prepare(input);
  if(r==='/v2/apresentacoes/dinamica')return f4ActivitiesDynamicCompat_(options).prepare(input);
  if(r.startsWith('/v2/apresentacoes/foto/'))return f4ActivitiesPhotoCompat_(options).prepare(input);
  if(r==='/v2/justificativas/documento/iniciar')return f4ActivitiesAttendanceProofCompat_(options).begin(input);
  f4Core_().fail('COMMAND_TYPE_INVALID');
 }
 return {commands,routes,type,prepare};
}
