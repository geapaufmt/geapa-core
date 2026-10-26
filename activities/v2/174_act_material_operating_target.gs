// A reminder milestone only. It neither changes the final deadline nor decides
// receipt, fulfilment, or sanctions. Date-only occurrence stays date-only.
function f4ActivitiesMaterialOperatingTarget_(requirement,recognition){
 const M=f4ActivitiesModel_(),hours=requirement.operatingHours;M.need(Number.isSafeInteger(hours)&&hours>=1&&hours<=8760,'CONFIG_CAPABILITY_DISABLED','OPERATING_MATERIAL_TARGET');
 const base={hours,realizationRef:recognition.ref,occurrence:recognition.occurrence};
 if(recognition.occurrence.precision!=='INSTANT')return {...base,state:'OCCURRENCE_INSTANT_UNRESOLVED'};
 const instant=recognition.occurrence.instant;M.need(Number.isSafeInteger(instant),'VALIDATION_ERROR');const dueAt=instant+hours*3600000;return {...base,state:'DEFINED',dueAt:M.stamp(dueAt),exceedsFinalDeadline:dueAt>=M.ms(requirement.finalEndExclusive)};
}
