'use strict';
// Async transport adaptation of the canonical Core provider, not a second
// implementation of file provenance, privacy, digest or replay policy.
const fs=require('node:fs'),path=require('node:path'),vm=require('node:vm');
const {createDomain}=require('./domain.cjs'),{fail}=require('./environment.cjs');
const methods=['call','folder','reserveId','inspect','upload','download'];
function asynchronous(source){
 for(const name of methods){
  source=source.replace(new RegExp('function '+name+'\\(', 'g'),'async function '+name+'(');
  source=source.replace(new RegExp('(?<!function )\\b'+name+'\\(', 'g'),'await '+name+'(');
 }
 return source.replace(/(?<![.\w])fetch\(/g,'await fetch(').replace(/\btoken\(\)/g,'await token()');
}
function provider({settings,root,fetchImpl,token}){
 if(typeof fetchImpl!=='function'||typeof token!=='function')fail('CONFIG_CAPABILITY_DISABLED');
 const domain=createDomain(()=>{throw Error('UNEXPECTED_DOMAIN_IO');});
 const source=fs.readFileSync(path.join(__dirname,'../generated/workspace-drive.gs'),'utf8');
 domain.context.workspaceOptions={settings,root,config:{environment:'DEV',projectId:'geapa-convergencia-dev',database:'(default)',host:'firestore.googleapis.com',port:443},fetch:fetchImpl,token};
 vm.runInContext('f4DevBoundary_=()=>({permit:c=>c.projectId==="geapa-convergencia-dev"&&c.environment==="DEV",root:workspaceOptions.root,check:()=>{if(workspaceOptions.config.projectId!=="geapa-convergencia-dev")throw Error("DESTINATION_INVALID");}});',domain.context);
 vm.runInContext(asynchronous(source),domain.context);
 // Host functions cannot be JSON serialized; inject only the two transport ports.
 vm.runInContext('var nodeDrive=f4ActivitiesWorkspaceDrive_({...JSON.parse(JSON.stringify({settings:workspaceOptions.settings,root:workspaceOptions.root,config:workspaceOptions.config})),fetch:workspaceOptions.fetch,token:workspaceOptions.token});',domain.context);
 return Object.fromEntries(['reserveId','inspect','upload','download'].map(method=>[method,async(...args)=>{
  // A provider instance is request-local. Concurrent operations use distinct VMs.
  const fn=vm.runInContext('(a)=>nodeDrive.'+method+'(...JSON.parse(a))',domain.context);
  return fn(JSON.stringify(args));
 }]));
}
function transport({fetchImpl=fetch,maximumBytes=25*1024*1024,timeoutMs=30000}={}){
 return async(url,request)=>{
  const u=new URL(url);
  if(u.origin!=='https://www.googleapis.com'||u.username||u.password||u.hash||!/^\/(drive\/v3\/files|upload\/drive\/v3\/files)(\/|\?|$)/.test(u.pathname+u.search))fail('DESTINATION_INVALID');
  const body=Array.isArray(request.payload)?Buffer.from(request.payload.map(b=>b&255)):request.payload;
  const response=await fetchImpl(url,{method:request.method.toUpperCase(),headers:{...request.headers,...(request.contentType?{'Content-Type':request.contentType}:{})},body,redirect:'error',signal:AbortSignal.timeout(timeoutMs)});
  if(Number(response.headers.get('content-length'))>maximumBytes+65536)fail('SCOPE_TOO_LARGE');
  const chunks=[];let size=0;for await(const chunk of response.body||[]){size+=chunk.length;if(size>maximumBytes+65536)fail('SCOPE_TOO_LARGE');chunks.push(Buffer.from(chunk));}
  const bytes=Buffer.concat(chunks);return {getResponseCode:()=>response.status,getContentText:()=>bytes.toString('utf8'),getAllHeaders:()=>Object.fromEntries(response.headers),getBlob:()=>({getBytes:()=>[...bytes]})};
 };
}
// Explicitly unavailable until a Workspace identity/storage model is approved.
// The existing canonical provider requires owner-private My Drive objects;
// runtime service accounts cannot own those files. Never fall back to user ADC.
async function runtimeProvider(){fail('WORKSPACE_IDENTITY_NOT_CONFIGURED');}
module.exports={provider,transport,asynchronous,runtimeProvider};
