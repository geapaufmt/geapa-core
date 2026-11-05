// Attachments are versioned evidence; Storage is never an approval authority.
function f4InstitutionalDocumentFiles_(){
 const C=f4Core_(),F=f4Firestore_(),N=C.need,MAX=3145728;
 function prepare(input,old){
  if(!Object.hasOwn(input,'attachments'))return (old?.attachments||[]).map(f=>({existing:f,caption:f.caption}));
  N(Array.isArray(input.attachments)&&input.attachments.length<=6,'DOCUMENT_FILE_LIMIT');
  let bytes=0;const files=input.attachments.map(f=>{
   N(typeof f.caption==='string'&&f.caption===f.caption.trim()&&f.caption.length<=500,'COMMAND_PAYLOAD_INVALID');
   if(Object.hasOwn(f,'fileId')){C.exact(f,['fileId','caption']);const existing=old?.attachments?.find(x=>x.fileId===f.fileId);N(existing,'DOCUMENT_FILE_SCOPE_INVALID');return {existing,caption:f.caption};}
   C.exact(f,['name','mime','base64','sha256','origin','caption']);
   N(typeof f.name==='string'&&f.name.length>0&&f.name.length<=180&&!/[\x00-\x1f\\/]/.test(f.name)&&['application/pdf','image/jpeg','image/png','image/webp'].includes(f.mime),'DOCUMENT_FILE_INVALID');
   N(typeof f.base64==='string'&&f.base64.length<=4194304&&/^[A-Za-z0-9+/]+={0,2}$/.test(f.base64)&&/^[a-f0-9]{64}$/.test(f.sha256),'DOCUMENT_FILE_INVALID');
   N(f.origin==='LOCAL_UPLOAD'||/^https:\/\/drive\.google\.com\/file\/d\/[A-Za-z0-9_-]+\/view$/.test(f.origin),'DOCUMENT_FILE_INVALID');
   const data=Utilities.base64Decode(f.base64);bytes+=data.length;
   N(data.length>0&&bytes<=MAX&&Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,data).map(x=>((x+256)%256).toString(16).padStart(2,'0')).join('')===f.sha256,'DOCUMENT_FILE_INVALID');
   if(f.mime.startsWith('image/'))N(f.caption.length>=3,'DOCUMENT_IMAGE_CAPTION_REQUIRED');
   return {upload:f,caption:f.caption};
  });
  N(new Set(files.map(f=>f.existing?.sha256||f.upload.sha256)).size===files.length,'DOCUMENT_FILE_DUPLICATE');return files;
 }
 function materialize(files,o,recordId,version,personId){return files.map((f,order)=>{
  if(f.existing)return {...f.existing,caption:f.caption,order};
  N(typeof o.stageAttachment==='function','CONFIG_CAPABILITY_DISABLED');
  const d=o.stageAttachment(f.upload,recordId,version,personId);
  N(d.documentId===recordId&&d.documentVersion===version&&d.sha256===f.upload.sha256&&d.mime===f.upload.mime&&d.name===f.upload.name&&d.actorPersonId===personId&&d.size>0&&d.size<=MAX&&d.storage==='PRIVATE_STORAGE','SCHEMA_INVALID');
  return {...d,caption:f.caption,order};
 });}
 return {prepare,materialize,MAX};
}
