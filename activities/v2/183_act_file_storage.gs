// Canonical file-provider contract. Provider choice is server policy, never a
// browser field. Legacy Drive artifacts remain immutable and readable as before.
function f4ActivitiesFileStorage_(){
 const N=f4Core_().need,project=f4ExecutionEnvironment_().project,bucket=project+'.firebasestorage.app';
 const providers=['GOOGLE_DRIVE_V3','GOOGLE_CLOUD_STORAGE'];
 function configured(policy){
  const w=policy?.workspace;N(w?.enabled===true,'CONFIG_CAPABILITY_DISABLED');
  if(w.provider==='GOOGLE_CLOUD_STORAGE'){
   N(w.projectId===project&&w.bucket===bucket&&w.prefix===f4ExecutionEnvironment_().prefix&&Number.isSafeInteger(w.maximumBytes)&&w.maximumBytes>0&&w.maximumBytes<=50*1024*1024,'DESTINATION_INVALID');
   return 'GOOGLE_CLOUD_STORAGE';
  }
  N(!w.provider||w.provider==='GOOGLE_DRIVE_V3','CONFIG_CAPABILITY_DISABLED');return 'GOOGLE_DRIVE_V3';
 }
 function effect(policy){return policy?.workspace?.provider==='GOOGLE_CLOUD_STORAGE'&&configured(policy)==='GOOGLE_CLOUD_STORAGE'?'FILE_CREATE':'DRIVE_CREATE';}
 const isCreate=out=>['DRIVE_CREATE','FILE_CREATE'].includes(out?.effectKind);
 const isArtifact=a=>providers.includes(a?.provider);
 function reserved(out,evidence,externalId){
  const r=out.workspaceReservation;N(r&&r.fileId===externalId&&evidence.externalId===externalId&&evidence.sha256===r.descriptor.sha256&&evidence.size===r.descriptor.size&&evidence.mime===r.descriptor.mime,'UNRESOLVED_DEPENDENCY');
  if(evidence.provider==='GOOGLE_CLOUD_STORAGE')N(out.effectKind==='FILE_CREATE'&&r.provider===evidence.provider&&r.bucket===bucket&&evidence.bucket===bucket&&externalId===f4ExecutionEnvironment_().prefix+out.id&&/^[0-9]+$/.test(evidence.versionId)&&/^[a-f0-9]{32}$/.test(evidence.md5),'UNRESOLVED_DEPENDENCY');
  else N(evidence.md5===r.descriptor.md5,'UNRESOLVED_DEPENDENCY');
 }
 return {configured,effect,isCreate,isArtifact,reserved,project,bucket};
}
