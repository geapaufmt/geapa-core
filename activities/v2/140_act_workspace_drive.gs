// Drive provider for explicitly configured, owner-private DEV files only.
// The coordinator must durably reserve fileId before upload. No provider call
// changes sharing, sends mail, deletes content, or decides material acceptance.
// https://developers.google.com/workspace/drive/api/guides/manage-uploads
function f4ActivitiesWorkspaceDrive_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),cfg=options.settings;
 N(f4DevBoundary_().permit(options.config)&&options.root===f4DevBoundary_().root,'DESTINATION_INVALID');
 C.exact(cfg,['enabled','folderId','maximumBytes','mimeTypes']);
 const fileId=id=>typeof id==='string'&&/^[A-Za-z0-9_-]{10,160}$/.test(id),mimeTypes=['application/pdf','application/vnd.ms-powerpoint','application/vnd.openxmlformats-officedocument.presentationml.presentation','application/vnd.oasis.opendocument.presentation','image/jpeg','image/png','image/webp','application/msword','application/vnd.openxmlformats-officedocument.wordprocessingml.document'];
 N(cfg.enabled===true&&fileId(cfg.folderId)&&Number.isSafeInteger(cfg.maximumBytes)&&cfg.maximumBytes>0&&cfg.maximumBytes<=25*1024*1024&&Array.isArray(cfg.mimeTypes)&&cfg.mimeTypes.length>0&&new Set(cfg.mimeTypes).size===cfg.mimeTypes.length&&cfg.mimeTypes.every(m=>mimeTypes.includes(m)),'CONFIG_CAPABILITY_DISABLED');
 const origin='https://www.googleapis.com',fetch=options.fetch||((url,request)=>UrlFetchApp.fetch(url,request)),token=options.token||(()=>ScriptApp.getOAuthToken());
 const fields='id,name,mimeType,parents,ownedByMe,trashed,permissions(id,type,role,deleted),appProperties,size,md5Checksum,headRevisionId,version';
 const privateOwner=d=>d?.ownedByMe===true&&Array.isArray(d.permissions)&&d.permissions.length===1&&d.permissions[0].type==='user'&&d.permissions[0].role==='owner'&&d.permissions[0].deleted!==true;
 function call(path,method,body,headers={},binary=false){
  f4DevBoundary_().check();
  const request={method,headers:{Authorization:'Bearer '+token(),...headers},muteHttpExceptions:true,followRedirects:false};
  if(body!==undefined){request.payload=binary?body:JSON.stringify(body);request.contentType=binary?headers['Content-Type']:'application/json';delete request.headers['Content-Type'];}
  const r=fetch(origin+path,request);let data;try{data=JSON.parse(r.getContentText()||'null');}catch{data=null;}return {status:r.getResponseCode(),data,headers:r.getAllHeaders?.()||{}};
 }
 function folder(){const r=call('/drive/v3/files/'+encodeURIComponent(cfg.folderId)+'?fields='+encodeURIComponent(fields),'get');N(r.status===200&&r.data?.id===cfg.folderId&&r.data.mimeType==='application/vnd.google-apps.folder'&&privateOwner(r.data)&&r.data.trashed===false&&r.data.name?.startsWith('[GEAPA DEV ACT]')&&r.data.appProperties?.environment==='DEV'&&r.data.appProperties?.root===options.root,'DESTINATION_INVALID','DEV_FOLDER_NOT_VERIFIED');return r.data;}
 function plan(v){C.exact(v,['effectId','fileId','descriptor']);N(M.id(v.effectId)&&fileId(v.fileId),'VALIDATION_ERROR');C.exact(v.descriptor,['name','mime','size','sha256','md5']);const d=v.descriptor;N(typeof d.name==='string'&&d.name.trim()&&d.name.length<=240&&!/[\x00-\x1f]/.test(d.name)&&cfg.mimeTypes.includes(d.mime)&&Number.isSafeInteger(d.size)&&d.size>0&&d.size<=cfg.maximumBytes&&/^[a-f0-9]{64}$/.test(d.sha256)&&/^[a-f0-9]{32}$/.test(d.md5),'VALIDATION_ERROR');return v;}
 function reserveId(){folder();const r=call('/drive/v3/files/generateIds?count=1&space=drive&type=files','get');N(r.status===200&&Array.isArray(r.data?.ids)&&r.data.ids.length===1&&fileId(r.data.ids[0]),'UNRESOLVED_DEPENDENCY','DRIVE_ID_UNAVAILABLE');return r.data.ids[0];}
 function inspect(v){plan(v);const r=call('/drive/v3/files/'+encodeURIComponent(v.fileId)+'?fields='+encodeURIComponent(fields),'get');if(r.status===404)return {outcome:'UNKNOWN',reason:'FILE_NOT_OBSERVED'};if(r.status!==200)return {outcome:'UNKNOWN',reason:'DRIVE_READ_UNAVAILABLE'};const d=r.data,p=d?.appProperties;
  N(d?.id===v.fileId&&privateOwner(d)&&d.trashed===false&&d.parents?.length===1&&d.parents[0]===cfg.folderId&&p?.environment==='DEV'&&p.root===options.root&&p.effectId===v.effectId&&p.sha256===v.descriptor.sha256&&d.mimeType===v.descriptor.mime&&String(d.size)===String(v.descriptor.size)&&d.md5Checksum===v.descriptor.md5&&typeof d.headRevisionId==='string'&&d.headRevisionId&&typeof d.version==='string','UNRESOLVED_DEPENDENCY','DRIVE_OBJECT_MISMATCH');
  return {outcome:'CONFIRMED',provider:'GOOGLE_DRIVE_V3',externalId:d.id,versionId:d.headRevisionId,providerVersion:d.version,sha256:p.sha256,md5:d.md5Checksum,size:v.descriptor.size,mime:d.mimeType};
 }
 function digest(algorithm,bytes){return Utilities.computeDigest(algorithm,bytes).map(b=>(b&255).toString(16).padStart(2,'0')).join('');}
 function upload(v,bytes){plan(v);N(Array.isArray(bytes)&&bytes.length===v.descriptor.size&&bytes.every(b=>Number.isInteger(b)&&b>=-128&&b<=255),'VALIDATION_ERROR');N(digest(Utilities.DigestAlgorithm.SHA_256,bytes)===v.descriptor.sha256&&digest(Utilities.DigestAlgorithm.MD5,bytes)===v.descriptor.md5,'VALIDATION_ERROR','UPLOAD_CONTENT_CHANGED');folder();
  let before;try{before=inspect(v);}catch(e){if(e.code)throw e;return {outcome:'UNKNOWN',reason:'DRIVE_READ_UNAVAILABLE'};}if(before.outcome==='CONFIRMED')return {...before,replay:true};if(before.reason!=='FILE_NOT_OBSERVED')return before;
  try{
   const metadata={id:v.fileId,name:'[GEAPA DEV ACT] '+v.descriptor.name,mimeType:v.descriptor.mime,parents:[cfg.folderId],appProperties:{environment:'DEV',root:options.root,effectId:v.effectId,sha256:v.descriptor.sha256}};
   const initiated=call('/upload/drive/v3/files?uploadType=resumable','post',metadata,{'X-Upload-Content-Type':v.descriptor.mime,'X-Upload-Content-Length':String(v.descriptor.size)});
   if(initiated.status===200){const location=Object.entries(initiated.headers).find(([k])=>k.toLowerCase()==='location')?.[1];N(typeof location==='string'&&/^https:\/\/www\.googleapis\.com\/upload\/drive\/v3\/files\?/.test(location)&&/[?&]upload_id=[A-Za-z0-9_-]+(?:&|$)/.test(location),'DESTINATION_INVALID','UPLOAD_SESSION_ORIGIN');
    call(location.slice(origin.length),'put',bytes,{'Content-Type':v.descriptor.mime,'Content-Range':'bytes 0-'+(bytes.length-1)+'/'+bytes.length},true);
   }
  }catch(e){if(e.code==='DESTINATION_INVALID')throw e;/* An uncertain upload is reconciled by its already reserved file ID. */}
  try{return inspect(v);}catch(e){if(e.code)throw e;return {outcome:'UNKNOWN',reason:'DRIVE_READ_UNAVAILABLE'};}
 }
 // https://developers.google.com/workspace/drive/api/guides/manage-downloads
 function download(v,versionId){
  plan(v);folder();const before=inspect(v);N(before.outcome==='CONFIRMED'&&before.versionId===versionId,'STALE_REVISION','FILE_VERSION_CHANGED');
  const capabilities=call('/drive/v3/files/'+encodeURIComponent(v.fileId)+'?fields=capabilities(canDownload)','get');N(capabilities.status===200&&capabilities.data?.capabilities?.canDownload===true,'AUTHORIZATION_DENIED','FILE_DOWNLOAD_DENIED');
  f4DevBoundary_().check();const response=fetch(origin+'/drive/v3/files/'+encodeURIComponent(v.fileId)+'?alt=media',{method:'get',headers:{Authorization:'Bearer '+token()},muteHttpExceptions:true,followRedirects:false});
  N(response.getResponseCode()===200,'UNRESOLVED_DEPENDENCY','FILE_DOWNLOAD_UNAVAILABLE');const bytes=response.getBlob().getBytes();
  N(bytes.length===v.descriptor.size&&digest(Utilities.DigestAlgorithm.SHA_256,bytes)===v.descriptor.sha256&&digest(Utilities.DigestAlgorithm.MD5,bytes)===v.descriptor.md5,'UNRESOLVED_DEPENDENCY','FILE_CONTENT_CHANGED');
  const after=inspect(v);N(after.outcome==='CONFIRMED'&&after.versionId===versionId&&after.providerVersion===before.providerVersion,'STALE_REVISION','FILE_VERSION_CHANGED');return bytes;
 }
 return Object.freeze({reserveId,inspect,upload,download});
}
