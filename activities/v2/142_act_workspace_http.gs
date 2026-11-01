// Binary upload travels through the same authenticated Core endpoint. It is a
// continuation of a protected ARTIFACT_UPLOAD_REQUEST, not a second command bus.
function f4ActivitiesHttpBody_(event){
 const C=f4Core_(),text=event?.postData?.contents,maxBinary=25*1024*1024;
 C.need(typeof text==='string'&&text.length<=Math.ceil(maxBinary/3)*4+16384,'COMMAND_PAYLOAD_INVALID');
 // The shared client writes action first. Reject oversized ordinary commands
 // before parsing or allocating a large JSON object.
 if(text.length>65536)C.need(/^\s*\{\s*"action"\s*:\s*"core\.act\.workspace-upload\.v13"\s*,/.test(text.slice(0,128)),'COMMAND_PAYLOAD_INVALID');
 let body;try{body=JSON.parse(text);}catch{C.fail('COMMAND_PAYLOAD_INVALID');}
 C.need(body&&typeof body==='object'&&!Array.isArray(body),'COMMAND_PAYLOAD_INVALID');
 const limit=body?.action==='core.act.workspace-upload.v13'?Math.ceil(maxBinary/3)*4+16384:65536;
 C.need(text.length<=limit&&Utilities.newBlob(text).getBytes().length<=limit,'COMMAND_PAYLOAD_INVALID');
 return body;
}
function f4ActivitiesWorkspaceHttp_(options){
 const C=f4Core_(),M=f4ActivitiesModel_(),N=M.need;
 function handle(action,request){
  N(['SYNTHETIC_FUNCTIONAL','DEV_OPERATIONAL'].includes(options.mode)&&options.provider==='GOOGLE_DRIVE_V3','CONFIG_CAPABILITY_DISABLED');
  N(['core.act.workspace-upload.v13','core.act.workspace-reconcile.v13'].includes(action),'COMMAND_TYPE_INVALID');
  const upload=action==='core.act.workspace-upload.v13';C.exact(request,['effectRef',...(upload?['contentBase64']:[])]);
  N(request.effectRef?.kind==='outbox'&&M.id(request.effectRef.id),'COMMAND_PAYLOAD_INVALID');
  // Recheck current Membros authority before allocating/decoding binary data.
  const db=f4Firestore_().connect(options.config),out=db.get(M.path(request.effectRef));
  N(out,'COMMAND_PAYLOAD_INVALID');options.ports.authorizeUpload(path=>db.get(path),out);
  const coordinator=f4ActivitiesWorkspaceUpload_({config:options.config,pipeline:options.pipeline,clock:options.clock,authorize:options.ports.authorizeUpload,providerFor:settings=>f4ActivitiesWorkspaceDrive_({config:options.config,root:f4DevBoundary_().root,settings})});
  if(!upload)return coordinator.reconcile(request.effectRef);
  const text=request.contentBase64,expected=out.details?.fileDescriptor?.size;
  N(Number.isSafeInteger(expected)&&expected>0&&expected<=25*1024*1024&&typeof text==='string'&&text.length===Math.ceil(expected/3)*4&&!/[^A-Za-z0-9+/=]/.test(text)&&(text.indexOf('=')===-1||text.indexOf('=')>=text.length-2),'COMMAND_PAYLOAD_INVALID');
  let bytes;try{bytes=Utilities.base64Decode(text);}catch{C.fail('COMMAND_PAYLOAD_INVALID');}
  N(bytes.length===expected&&Utilities.base64Encode(bytes)===text,'COMMAND_PAYLOAD_INVALID');
  return coordinator.upload(request.effectRef,bytes);
 }
 return {handle};
}
