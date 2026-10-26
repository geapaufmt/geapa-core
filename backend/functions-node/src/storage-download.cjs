'use strict';
const {fence,fail}=require('./environment.cjs'),{readAct}=require('./act-readers.cjs'),{routes}=require('./workspace-download.cjs'),GCS=require('./workspace-storage.cjs');
async function download(options){
 const {body,hooks={}}=options,env=fence(),q=body?.request;
 if(process.env.GEAPA_ACT_STORAGE_ENABLED!=='true'||!env.emulator&&Object.keys(hooks).length)fail('CONFIG_CAPABILITY_DISABLED');
 if(body.action!=='storageDownload'||Object.keys(body).sort().join()!=='action,request'||!q||Object.keys(q).sort().join()!=='length,offset,params,route'||!routes.includes(q.route))fail('COMMAND_PAYLOAD_INVALID');
 const args={...options,hooks:{},body:{action:'actRead',request:{route:q.route,params:q.params}},workspacePhase:{storage:true}};
 const first=await readAct(args),p=first.result.plan;
 if(p.settings.provider!=='GOOGLE_CLOUD_STORAGE'){if(q.offset===0&&q.length===0)return {state:'LEGACY_DRIVE',metrics:{authorization:first.metrics}};fail('CONFIG_CAPABILITY_DISABLED');}
 if(p.settings.bucket!==GCS.BUCKET)fail('CONFIG_CAPABILITY_DISABLED');
 const {name,mime,size,sha256}=p.request.descriptor,plan={effectId:p.request.effectId,objectName:p.request.fileId,descriptor:{name,mime,size,sha256}};
 if(q.offset===0&&q.length===0)return {state:'CURRENT',data:{name,mime,size,sha256,offset:0,contentBase64:''},metrics:{authorization:first.metrics}};
 const provider=await (env.emulator&&hooks.provider?hooks.provider:GCS.runtimeProvider()),bytes=await provider.downloadRange(plan,p.version,q.offset,q.length);
 await hooks.afterDownload?.();
 // Re-resolve authorization, provenance, policy and immutable generation after
 // fetching each bounded chunk. No raw URL or bucket credential reaches Portal.
 const checked=await readAct({...args,workspacePhase:{storage:true,expected:p}});
 return {state:'CURRENT',data:{name,mime,size,sha256,offset:q.offset,contentBase64:bytes.toString('base64')},metrics:{authorization:first.metrics,recheck:checked.metrics,bytes:bytes.length}};
}
module.exports={download};
