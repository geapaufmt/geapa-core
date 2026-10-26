'use strict';
// Transport only. Authorization, ownership and durable effect transitions belong
// to the canonical Core. Session URLs are secrets and must remain server-side.
const crypto=require('node:crypto');
const {fail,PROJECT,SERVICE_ACCOUNT}=require('./environment.cjs');
const BUCKET=PROJECT+'.firebasestorage.app',PREFIX=require('./runtime-profile.cjs').profile.prefix,MAXIMUM=50*1024*1024,CHUNK=4*1024*1024;
const ORIGIN='https://storage.googleapis.com';
const TYPES=new Set(['application/pdf','application/vnd.ms-powerpoint','application/vnd.openxmlformats-officedocument.presentationml.presentation','application/vnd.oasis.opendocument.presentation','image/jpeg','image/png','image/webp','application/msword','application/vnd.openxmlformats-officedocument.wordprocessingml.document']);
function plan(p,readOnly=false){
 if(!p||Object.keys(p).sort().join()!=='descriptor,effectId,objectName'||!/^[a-f0-9]{32}$/.test(p.effectId)||p.objectName!==PREFIX+p.effectId&&!(readOnly&&require('./environment.cjs').PRODUCTION&&p.objectName==='act-history/'+p.effectId))fail('DESTINATION_INVALID');
 const d=p.descriptor;
 if(!d||Object.keys(d).sort().join()!=='mime,name,sha256,size'||typeof d.name!=='string'||!d.name.trim()||d.name.length>240||/[\x00-\x1f\x7f]/.test(d.name)||!TYPES.has(d.mime)||!Number.isSafeInteger(d.size)||d.size<1||d.size>MAXIMUM||!/^[a-f0-9]{64}$/.test(d.sha256))fail('VALIDATION_ERROR');
 return p;
}
function session(uri){
 let u;try{u=new URL(uri);}catch{fail('DESTINATION_INVALID');}
 if(u.origin!==ORIGIN||u.username||u.password||u.hash||u.pathname!=='/upload/storage/v1/b/'+BUCKET+'/o'||u.searchParams.get('uploadType')!=='resumable'||!u.searchParams.get('upload_id'))fail('DESTINATION_INVALID');
 return u.href;
}
function provider({token,fetchImpl=fetch,timeoutMs=45000}){
 if(typeof token!=='function')fail('CONFIG_CAPABILITY_DISABLED');
 async function call(url,method,headers={},body){
  return fetchImpl(url,{method,headers:{Authorization:'Bearer '+await token(),...headers},body,redirect:'manual',signal:AbortSignal.timeout(timeoutMs)});
 }
 const resource=(p,readOnly=false)=>ORIGIN+'/storage/v1/b/'+BUCKET+'/o/'+encodeURIComponent(plan(p,readOnly).objectName);
 async function initiate(p){
  plan(p);
  // A precondition makes even an abandoned/duplicate upload session unable to
  // overwrite a completed object. The fixed length bounds each session.
  // Bucket IAM and the authenticated Core enforce access. Object ACLs are
  // unnecessary and rejected by uniform-access buckets.
  const r=await call(ORIGIN+'/upload/storage/v1/b/'+BUCKET+'/o?uploadType=resumable&ifGenerationMatch=0','POST',{'Content-Type':'application/json','X-Upload-Content-Type':p.descriptor.mime,'X-Upload-Content-Length':String(p.descriptor.size)},JSON.stringify({name:p.objectName,contentType:p.descriptor.mime,cacheControl:'private, no-store',metadata:{environment:require('./runtime-profile.cjs').profile.environment,effectId:p.effectId,sha256:p.descriptor.sha256}}));
  if(r.status!==200)fail('RECONCILIATION_REQUIRED');
  return {sessionUri:session(r.headers.get('location'))};
 }
 async function progress(p,uri){
  plan(p);const r=await call(session(uri),'PUT',{'Content-Length':'0','Content-Range':'bytes */'+p.descriptor.size});
  if(r.status===200||r.status===201)return {state:'UPLOADED',offset:p.descriptor.size};
  if(r.status===404||r.status===410)return {state:'UNKNOWN',offset:null};
  if(r.status!==308)fail('RECONCILIATION_REQUIRED');
  const range=r.headers.get('range');if(range&&!/^bytes=0-\d+$/.test(range))fail('UNRESOLVED_DEPENDENCY');
  const offset=range?Number(range.slice(8))+1:0;
  if(!Number.isSafeInteger(offset)||offset<0||offset>=p.descriptor.size)fail('UNRESOLVED_DEPENDENCY');
  return {state:'UPLOADING',offset};
 }
 async function chunk(p,uri,offset,bytes){
  plan(p);if(!Buffer.isBuffer(bytes)||!bytes.length||bytes.length>CHUNK||!Number.isSafeInteger(offset)||offset<0||offset%262144||offset+bytes.length>p.descriptor.size||(offset+bytes.length<p.descriptor.size&&bytes.length%262144))fail('VALIDATION_ERROR');
  const r=await call(session(uri),'PUT',{'Content-Type':p.descriptor.mime,'Content-Length':String(bytes.length),'Content-Range':`bytes ${offset}-${offset+bytes.length-1}/${p.descriptor.size}`},bytes);
  if(![200,201,308].includes(r.status))return {state:'UNKNOWN'};
  // The coordinator always resolves the actual offset; an HTTP response cannot
  // by itself attest that the hash or institutional receipt is valid.
  return {state:r.status===308?'UPLOADING':'UPLOADED'};
 }
 function metadata(p,d){
  if(d?.bucket!==BUCKET||d.name!==p.objectName||d.metadata?.environment!==require('./runtime-profile.cjs').profile.environment||d.metadata.effectId!==p.effectId||d.metadata.sha256!==p.descriptor.sha256||String(d.size)!==String(p.descriptor.size)||d.contentType!==p.descriptor.mime||!/^\d+$/.test(d.generation)||d.metadata.firebaseStorageDownloadTokens||!d.md5Hash)fail('UNRESOLVED_DEPENDENCY');
  return d;
 }
 async function inspect(p){
  const r=await call(resource(p),'GET');
  if(r.status===404)return {outcome:'UNKNOWN',reason:'OBJECT_NOT_OBSERVED'};
  if(r.status!==200)fail('RECONCILIATION_REQUIRED');
  const d=metadata(p,await r.json());
  // Never trust the client-provided SHA stored in object metadata. Read and hash
  // the immutable generation once before producing a trusted upload receipt.
  const media=await call(resource(p)+'?alt=media&generation='+d.generation,'GET');
  if(media.status!==200)fail('RECONCILIATION_REQUIRED');
  const sha=crypto.createHash('sha256'),md5=crypto.createHash('md5');let size=0;
  for await(const bytes of media.body||[]){size+=bytes.length;if(size>p.descriptor.size)fail('UNRESOLVED_DEPENDENCY');sha.update(bytes);md5.update(bytes);}
  const digest=sha.digest('hex'),sum=md5.digest('base64');
  if(size!==p.descriptor.size||digest!==p.descriptor.sha256||sum!==d.md5Hash)fail('UNRESOLVED_DEPENDENCY');
  return {outcome:'CONFIRMED',provider:'GOOGLE_CLOUD_STORAGE',externalId:p.objectName,versionId:d.generation,providerVersion:d.metageneration,sha256:digest,md5:Buffer.from(sum,'base64').toString('hex'),size,mime:d.contentType,bucket:BUCKET};
 }
 async function downloadRange(p,generation,offset,length){
  plan(p,true);if(!/^\d+$/.test(generation)||!Number.isSafeInteger(offset)||offset<0||!Number.isSafeInteger(length)||length<1||length>CHUNK||offset+length>p.descriptor.size)fail('VALIDATION_ERROR');
  const r=await call(resource(p,true)+'?alt=media&generation='+generation,'GET',{Range:`bytes=${offset}-${offset+length-1}`});
  if(r.status!==206&&!(r.status===200&&offset===0&&length===p.descriptor.size))fail('UNRESOLVED_DEPENDENCY');
  const parts=[];let count=0;for await(const b of r.body||[]){count+=b.length;if(count>length)fail('UNRESOLVED_DEPENDENCY');parts.push(Buffer.from(b));}
  if(count!==length)fail('UNRESOLVED_DEPENDENCY');return Buffer.concat(parts);
 }
 return Object.freeze({initiate,progress,chunk,inspect,downloadRange});
}
async function runtimeProvider(){
 const {fence,metadataIdentity}=require('./environment.cjs');if(fence().emulator)fail('DESTINATION_INVALID');await metadataIdentity();
 const {GoogleAuth,Compute}=require('google-auth-library'),auth=new GoogleAuth({projectId:PROJECT,scopes:['https://www.googleapis.com/auth/devstorage.read_write']}),client=await auth.getClient();
 if(!(client instanceof Compute))fail('DESTINATION_INVALID');
 return provider({token:async()=>{const t=await client.getAccessToken();if(!t.token)fail('DESTINATION_INVALID');return t.token;}});
}
module.exports={BUCKET,PREFIX,MAXIMUM,CHUNK,plan,session,provider,runtimeProvider};
