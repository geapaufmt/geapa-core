'use strict';
const {fence,fail,ROOT}=require('./environment.cjs');
const {readAct}=require('./act-readers.cjs');
const {runtimeProvider}=require('./workspace-drive.cjs');
const routes=Object.freeze(['/v2/apresentacoes/material/arquivo','/v2/apresentacoes/foto/arquivo','/v2/justificativas/documento/arquivo']);
async function download(options){
 const env=fence(),{body,hooks={}}=options;
 if(!env.emulator&&Object.keys(hooks).length)fail('DESTINATION_INVALID');
 if(process.env.GEAPA_ACT_WORKSPACE_ENABLED!=='true')fail('CONFIG_CAPABILITY_DISABLED');
 if(body?.action!=='workspaceDownload'||!routes.includes(body.request?.route)||Object.keys(body).sort().join()!=='action,request')fail('COMMAND_PAYLOAD_INVALID');
 const start=performance.now(),readOptions={...options,body:{action:'actRead',request:body.request},hooks:{}};
 const first=await readAct({...readOptions,workspacePhase:{}});
 const plan=first.result.plan;
 // Bound inline JSON/VM memory; larger files need a separately reviewed stream.
 if(plan.request.descriptor.size>8*1024*1024)fail('SCOPE_TOO_LARGE');
 const provider=await (env.emulator&&hooks.providerFor?hooks.providerFor(plan.settings):runtimeProvider({settings:plan.settings,root:ROOT}));
 const at=performance.now(),bytes=await provider.download(plan.request,plan.version),providerMs=performance.now()-at;
 if(!Array.isArray(bytes)||bytes.length!==plan.request.descriptor.size||bytes.length>25*1024*1024||bytes.some(b=>!Number.isInteger(b)||b< -128||b>255))fail('UNRESOLVED_DEPENDENCY');
 const hash=require('node:crypto').createHash('sha256').update(Buffer.from(bytes.map(b=>b&255))).digest('hex');
 if(hash!==plan.request.descriptor.sha256)fail('UNRESOLVED_DEPENDENCY');
 await hooks.afterDownload?.();
 // A fresh canonical resolution checks current membership, policies, file
 // association, version and provenance. Failed/revoked reads discard bytes.
 const last=await readAct({...readOptions,workspacePhase:{expected:plan,bytes}});
 return {route:body.request.route,result:last.result.result,metrics:{authorization:first.metrics,recheck:last.metrics,providerMs,totalMs:performance.now()-start}};
}
module.exports={download,routes};
