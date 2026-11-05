'use strict';
const crypto=require('node:crypto'),G=require('./workspace-storage.cjs'),{PROJECT,fail}=require('./environment.cjs');
const sha=x=>crypto.createHash('sha256').update(x).digest('hex');
function stage(f,documentId,documentVersion,actorPersonId,now){
 const bytes=Buffer.from(f.base64,'base64');
 if(bytes.length<12||bytes.length>3145728||sha(bytes)!==f.sha256)fail('DOCUMENT_FILE_INVALID');
 const valid=f.mime==='application/pdf'?bytes.subarray(0,5).toString()==='%PDF-':f.mime==='image/png'?bytes.subarray(0,8).equals(Buffer.from([137,80,78,71,13,10,26,10])):f.mime==='image/jpeg'?bytes[0]===255&&bytes[1]===216&&bytes[2]===255:f.mime==='image/webp'?bytes.subarray(0,4).toString()==='RIFF'&&bytes.subarray(8,12).toString()==='WEBP':false;
 if(!valid)fail('DOCUMENT_FILE_INVALID');
 if(f.mime.startsWith('image/')){let p;try{p=new (require('jspdf').jsPDF)().getImageProperties(bytes);}catch{fail('DOCUMENT_FILE_INVALID');}if(!p.width||!p.height||p.width*p.height>12000000)fail('DOCUMENT_FILE_INVALID');}
 const fileId=sha(PROJECT+'|institutional-document|'+documentId+'|'+f.sha256).slice(0,32);
 const descriptor={fileId,documentId,documentVersion,name:f.name,mime:f.mime,size:bytes.length,sha256:f.sha256,origin:f.origin,actorPersonId,addedAt:now,objectName:G.PREFIX+fileId,bucket:G.BUCKET,storage:'PRIVATE_STORAGE'};
 return {bytes,descriptor,plan:plan(descriptor)};
}
function plan(d){if(d.bucket!==G.BUCKET||d.objectName!==G.PREFIX+d.fileId)fail('DOCUMENT_FILE_SCOPE_INVALID');return {effectId:d.fileId,objectName:d.objectName,descriptor:{name:d.name,mime:d.mime,size:d.size,sha256:d.sha256}};}
async function persist(g,entry){let found=await g.inspect(entry.plan);if(found.outcome!=='CONFIRMED'){const started=await g.initiate(entry.plan);await g.chunk(entry.plan,started.sessionUri,0,entry.bytes);found=await g.inspect(entry.plan);}if(found.outcome!=='CONFIRMED'||found.sha256!==entry.descriptor.sha256)fail('RECONCILIATION_REQUIRED');return found;}
module.exports={stage,plan,persist,sha};
