// Fixed ACT policy forms. Values are explicit decisions for a draft cycle;
// existing operational facts are not rewritten when a new revision is recorded.
function f4ActivitiesCyclePolicyCommands_(x){
 const {M,N,p,load,save}=x,C=f4Core_();
 function context(){
  const cycle=load(p.cycleRef,true);N(cycle.status==='DRAFT','HUMAN_DECISION_REQUIRED','CYCLE_POLICY_DRAFT_REQUIRED');
  N(typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000&&Number.isSafeInteger(p.effectiveFrom)&&p.effectiveFrom>=x.now,'VALIDATION_ERROR');
  const predecessor=cycle.scheduledConfigRef?.kind==='config'?cycle.scheduledConfigRef:cycle.currentConfigRef,prior=predecessor&&load(predecessor,true),policy=load(cycle.policyVersionRef);
  const parameters=prior?.parameters||policy.initialParameters;N(parameters,'CONFIG_CAPABILITY_DISABLED','INITIAL_PARAMETERS_REQUIRED');return {cycle,predecessor,parameters};
 }
 function policy(key,values){const ref=M.ref('policy',M.hash({operationId:x.operationId,key}).slice(0,32));N(!load(ref,false,true),'CONFLICT');save(ref,{...values,cycleRef:p.cycleRef,approved:true,active:true,approvalReason:p.reason.trim(),approvedAt:M.stamp(x.now),approvedBy:x.actor,actorAuthorityBasis:x.authorityBasis});return ref;}
 function schedule(z,parameters){
  const ref=M.ref('assessment',M.hash({operationId:x.operationId,kind:'CYCLE_POLICY_CONFIG'}).slice(0,32)),items=Object.entries(parameters).map(([key,value])=>({key,value}));N(items.length<=32,'SCOPE_TOO_LARGE');
  save(ref,{cycleRef:p.cycleRef,state:'SEALED',complete:true,items,digest:M.hash(items)});
  return f4ActivitiesPlanning_({...x,p:{cycleRef:p.cycleRef,parameters,parameterClasses:'STRUCTURAL',effectiveFrom:p.effectiveFrom,predecessorRef:z.predecessor||{kind:'NONE',id:'INITIAL'},approvalBasis:{reasonCode:'PORTAL_CYCLE_POLICY_CONFIGURATION',reason:p.reason.trim()},itemManifestRef:ref}}).CONFIG_RECORD();
 }
 return {
  CYCLE_THEME_SETTINGS_RECORD(){
   const z=context();N(Array.isArray(p.axes)&&p.axes.length>=1&&p.axes.length<=12,'VALIDATION_ERROR');const codes=new Set(),oldAxes=(z.parameters.axisRefs||[]).map(ref=>({ref,value:load(ref)}));
   const axes=p.axes.map(axis=>{C.exact(axis,['code','label','description']);N(typeof axis.code==='string'&&/^[A-Za-z][A-Za-z0-9_-]{0,31}$/.test(axis.code)&&!codes.has(axis.code)&&typeof axis.label==='string'&&axis.label.trim()&&axis.label.length<=120&&typeof axis.description==='string'&&axis.description.length<=2000,'VALIDATION_ERROR','AXIS_DEFINITION');codes.add(axis.code);const values={usage:'THEMATIC_AXIS',code:axis.code,label:axis.label.trim(),description:axis.description.trim()},same=oldAxes.find(a=>(a.value.code||a.value.id)===values.code&&a.value.label===values.label&&(a.value.description||'')===values.description&&a.value.approved&&a.value.active);return same?same.ref:policy('AXIS_'+axis.code,values);});
   C.exact(p.analysisClock,['unit','amount','dayRule','pauseDayRule']);const clock=p.analysisClock;N(['HOURS','DAYS'].includes(clock.unit)&&Number.isSafeInteger(clock.amount)&&clock.amount>=1&&clock.amount<=366,'VALIDATION_ERROR','ANALYSIS_CLOCK_REQUIRED');
   N(clock.unit==='DAYS'?clock.dayRule==='CIVIL_END_INCLUSIVE'&&clock.pauseDayRule==='PRESERVE_REMAINING_CIVIL_DAYS':clock.dayRule===''&&clock.pauseDayRule==='','VALIDATION_ERROR','CLOCK_DAY_RULE');
   N(typeof p.compatibilityGuide==='string'&&p.compatibilityGuide.trim()&&p.compatibilityGuide.length<=4000,'INCOMPLETE_REQUEST','COMPATIBILITY_GUIDE_REQUIRED');
   const clockRef=policy('THEMATIC_CLOCK',{usage:'THEMATIC_CLOCK',unit:clock.unit,amount:clock.amount,...(clock.unit==='DAYS'?{dayRule:clock.dayRule,pauseDayRule:clock.pauseDayRule}:{})}),guideRef=policy('COMPATIBILITY_GUIDE',{usage:'COMPATIBILITY_GUIDE',text:p.compatibilityGuide.trim()});
   return schedule(z,{...z.parameters,axisRefs:axes,clockPolicyRef:clockRef,compatibilityGuideRef:guideRef});
  },
  CYCLE_PUBLICATION_SETTINGS_RECORD(){
   const z=context();N(Array.isArray(p.checklist)&&p.checklist.length>=1&&p.checklist.length<=12,'INCOMPLETE_REQUEST','OPERATIONAL_CHECKLIST_REQUIRED');
   const checks={approved:true,items:p.checklist};f4ActivitiesOperationalChecklist_({operationalChecklist:checks});
   const fields=['title','civilDate','approvedFormat','authorizedAxis'];N(Array.isArray(p.publicFields)&&p.publicFields.length>=1&&new Set(p.publicFields).size===p.publicFields.length&&p.publicFields.every(v=>fields.includes(v)),'VALIDATION_ERROR','PUBLIC_FIELD_SELECTION');
   const publicationPolicyRef=policy('PUBLICATION',{usage:'PUBLICATION',publicScope:'PUBLIC_PORTAL',publicFieldSelection:p.publicFields});
   return schedule(z,{...z.parameters,operationalChecklist:checks,publicationPolicyRef});
  },
  CYCLE_MATERIAL_SETTINGS_RECORD(){
   const z=context();N(Number.isSafeInteger(p.operatingHours)&&p.operatingHours>=1&&p.operatingHours<=8760,'VALIDATION_ERROR','OPERATING_MATERIAL_TARGET');
   M.civil(p.finalCivilDate);N(Date.parse(p.finalCivilDate+'T00:00:00-04:00')+86400000>p.effectiveFrom,'VALIDATION_ERROR','FINAL_MATERIAL_DEADLINE');
   N(p.providerPolicyRef?.kind==='policy','VALIDATION_ERROR');const provider=load(p.providerPolicyRef,true);
   N(provider.approved&&provider.active&&provider.usage==='MATERIAL_PROVIDER'&&provider.cycleRef?.id===p.cycleRef.id&&provider.workspace?.enabled===true,'CONFIG_CAPABILITY_DISABLED','MATERIAL_PROVIDER_CONFIGURATION');
   const executor=provider.executorRuleRef&&load(provider.executorRuleRef);N(executor?.approved&&executor.active&&executor.cycleRef?.id===p.cycleRef.id&&['ACT.OUTBOX_CLAIM','ACT.OUTBOX_OBSERVE','ACT.OUTBOX_RECONCILE'].every(cmd=>executor.commands?.includes(cmd)),'CONFIG_CAPABILITY_DISABLED','MATERIAL_PROVIDER_EXECUTOR');
   if(p.photoPolicyRef){const photo=load(p.photoPolicyRef,true);N(photo.approved&&photo.active&&photo.usage==='PRESENTATION_PHOTO'&&photo.cycleRef?.id===p.cycleRef.id&&photo.workspace?.enabled,'CONFIG_CAPABILITY_DISABLED','PHOTO_CONFIGURATION');}
   N(typeof p.explicitPortalNotice==='boolean','VALIDATION_ERROR');
   const cutoff=policy('MATERIAL_CUTOFF',{usage:'MATERIAL_CUTOFF',finalCivilDate:p.finalCivilDate}),notice=p.explicitPortalNotice?policy('MATERIAL_NOTICE',{usage:'MATERIAL_NOTICE',commands:['ACT.CORRECTION_NOTICE_RECORD']}):null;
   function requirement(key,items){
    N(Array.isArray(items)&&items.length>=1&&items.length<=12,'INCOMPLETE_REQUEST','MATERIAL_ITEMS');const ids=new Set();for(const item of items){C.exact(item,['itemId','label']);N(/^[A-Za-z][A-Za-z0-9_-]{0,39}$/.test(item.itemId)&&!ids.has(item.itemId)&&typeof item.label==='string'&&item.label.trim()&&item.label.length<=300,'VALIDATION_ERROR','MATERIAL_ITEMS');ids.add(item.itemId);}
    const uploadPolicyRef=policy(key+'_UPLOAD',{usage:'MATERIAL_UPLOAD',workspace:provider.workspace,executorRuleRef:provider.executorRuleRef,...(p.photoPolicyRef?{photoPolicyRef:p.photoPolicyRef}:{}),providerPolicyRef:p.providerPolicyRef,providerPolicyRevision:provider.revision,portalReview:{approved:true,mode:'COMPLETE_FILE_REVIEW',coveredItemIds:[...ids],evaluationCutoffRef:cutoff},...(notice?{portalCorrectionNotice:{approved:true,channel:'PORTAL_EXPLICIT_ACK',executorRuleRef:notice}}:{})});
    const values={usage:'MATERIAL_REQUIREMENT',state:'SEALED',complete:true,items,digest:M.hash(items),finalEndExclusive:M.stamp(Date.parse(p.finalCivilDate+'T00:00:00-04:00')+86400000),finalCivilDate:p.finalCivilDate,operatingHours:p.operatingHours,uploadPolicyRef};return policy(key,values);
   }
   N(Array.isArray(p.dynamicItems),'VALIDATION_ERROR','MATERIAL_ITEMS');const ordinaryMaterialRequirementRef=requirement('ORDINARY_MATERIAL',p.ordinaryItems);let dynamicMaterialRequirementRef;
   if(p.dynamicItems?.length)dynamicMaterialRequirementRef=requirement('DYNAMIC_MATERIAL',p.dynamicItems);else N(!z.parameters.dynamic,'INCOMPLETE_REQUEST','DYNAMIC_MATERIAL_REQUIRED');
   const parameters={...z.parameters,ordinaryMaterialRequirementRef};if(dynamicMaterialRequirementRef)parameters.dynamicMaterialRequirementRef=dynamicMaterialRequirementRef;else delete parameters.dynamicMaterialRequirementRef;
   return schedule(z,parameters);
  }
 };
}
