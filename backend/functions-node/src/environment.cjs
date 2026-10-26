'use strict';
const {profile}=require('./runtime-profile.cjs');
const PROJECT=profile.projectId,DEMO='demo-geapa-functions-local',PRODUCTION=profile.environment==='PROD';
const SHARED=PRODUCTION||process.env.GEAPA_PILOT_MODE==='SHARED_DEV';
const ROOT=SHARED?'':'coreDevValidation/CF_DEV_EDIT_PILOT_V1',ACTIVITY=SHARED?'CF_DEV_EDIT_ACTIVITY_V2':'CF_DEV_EDIT_ACTIVITY_V1',EXECUTOR=profile.executor;
const SERVICE_ACCOUNT=profile.serviceAccount;
function fail(code){throw Object.assign(new Error(code),{code});}
function fence(env=process.env){
 if(require('./runtime-profile.cjs').select(env)!==profile)fail('DESTINATION_INVALID');
 const emulator=env.GEAPA_RUNTIME==='EMULATOR';
 if(PRODUCTION&&env.GEAPA_PILOT_MODE)fail('DESTINATION_INVALID');
 if(env.GEAPA_PILOT_MODE&&!['SHARED_DEV','ISOLATED'].includes(env.GEAPA_PILOT_MODE))fail('DESTINATION_INVALID');
 if(env.GOOGLE_APPLICATION_CREDENTIALS||env.CLOUDSDK_AUTH_CREDENTIAL_FILE_OVERRIDE||env.GOOGLE_GHA_CREDS_PATH||env.GCE_METADATA_HOST||env.GCE_METADATA_IP||env.GCE_METADATA_ROOT)fail('DESTINATION_INVALID');
 if(emulator){
  if(env.GCLOUD_PROJECT!==DEMO||env.FUNCTIONS_EMULATOR!=='true'||env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8187'||env.FIREBASE_AUTH_EMULATOR_HOST!=='127.0.0.1:9197')fail('DESTINATION_INVALID');
 }else{
  if(env.GEAPA_RUNTIME!==profile.environment||env.GCLOUD_PROJECT!==PROJECT||env.GOOGLE_CLOUD_PROJECT&&env.GOOGLE_CLOUD_PROJECT!==PROJECT||!env.K_SERVICE||!env.K_REVISION)fail('DESTINATION_INVALID');
  if(Object.keys(env).some(k=>/EMULATOR/.test(k)&&env[k]))fail('DESTINATION_INVALID');
  if(env.FIREBASE_CONFIG){let c;try{c=JSON.parse(env.FIREBASE_CONFIG);}catch{fail('DESTINATION_INVALID');}if(c.projectId&&c.projectId!==PROJECT)fail('DESTINATION_INVALID');}
 }
 return Object.freeze({emulator,projectId:emulator?DEMO:PROJECT,root:ROOT,activityId:ACTIVITY,executor:EXECUTOR});
}
async function metadataIdentity(fetcher=fetch){
 const get=async path=>{const r=await fetcher('http://169.254.169.254/computeMetadata/v1/'+path,{headers:{'Metadata-Flavor':'Google'},redirect:'error',signal:AbortSignal.timeout(3000)});if(!r.ok||r.headers.get('metadata-flavor')!=='Google')fail('DESTINATION_INVALID');return (await r.text()).trim();};
 const [project,email]=await Promise.all([get('project/project-id'),get('instance/service-accounts/default/email')]);
 if(project!==PROJECT||email!==SERVICE_ACCOUNT)fail('DESTINATION_INVALID');return {project,email};
}
module.exports={SHARED,PROJECT,DEMO,ROOT,ACTIVITY,EXECUTOR,SERVICE_ACCOUNT,PRODUCTION,profile,fence,fail,metadataIdentity};
