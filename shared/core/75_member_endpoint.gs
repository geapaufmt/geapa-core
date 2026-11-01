// Local composition only. Deliberately not wired to doPost, RPC or a scheduler.
function f4MemberEndpoint_(){
 const C=f4Core_(),N=C.need;
 function createLocal(options){
  C.exact(options,['config','scriptId','admission','recovery','consumers']);C.destination(options.config);
  N((options.config.environment==='SYNTHETIC_LOCAL' || f4DevBoundary_().permit(options.config))&&(options.config.environment==='SYNTHETIC_LOCAL'?options.scriptId==='LOCAL_M13_G':options.scriptId===f4DevBoundary_().script)&&typeof options.admission?.authenticate==='function'&&typeof options.recovery?.authenticate==='function'&&typeof options.consumers?.status==='function','DESTINATION_INVALID');
  const config={...options.config};
  function post(request){
   C.exact(request,['method','projectId','scriptId','body']);
   N(request.method==='POST'&&request.projectId===config.projectId&&request.scriptId===options.scriptId,'DESTINATION_INVALID');
   N(typeof request.body==='string'&&request.body.length<=65536&&encodeURIComponent(request.body).replace(/%[A-F\d]{2}/gi,'x').length<=65536,'COMMAND_PAYLOAD_INVALID');
   let body;try{body=JSON.parse(request.body);}catch{C.fail('COMMAND_PAYLOAD_INVALID');}
   C.exact(body,['action','idToken','request']);N(typeof body.idToken==='string','AUTH_BINDING_INVALID');
   N(['members.admit.v13','members.status.v13','core.admit.v13','core.status.v13'].includes(body.action),'COMMAND_TYPE_INVALID');
   if(body.action.startsWith('members.'))N(!f4SharedCommandCatalog_().act(body.request?.commandType),'COMMAND_TYPE_INVALID');
   if(['members.admit.v13','core.admit.v13'].includes(body.action)){
    N(!['members.binding.associate','members.binding.revoke'].includes(body.request?.commandType),'ACCESS_REVOKED');
    const principal=options.admission.authenticate(body.idToken); // before transaction; no HTTP in callback
    return f4Firestore_().connect(config).transaction(tx=>options.admission.issueBrowser(tx,principal,body.request));
   }
   const principal=options.recovery.authenticate(body.idToken);
   return f4Firestore_().connect(config).transaction(tx=>options.consumers.status(tx,principal,body.request));
  }
  // Only trusted internal code can call attempt; it is absent from the POST dispatcher.
  function attempt(token,key,workerId){
   const principal=options.recovery.authenticate(token),db=f4Firestore_().connect(config);
   const claim=db.transaction(tx=>options.recovery.claim(tx,principal,key,workerId));
   if(claim.state!=='CLAIMED')return claim;
   const current=options.recovery.authenticate(token); // recheck external Auth immediately before final attempt
   return db.transaction(tx=>options.recovery.finalize(tx,current,key,claim.handle));
  }
  return Object.freeze({post,attempt});
 }
 return Object.freeze({createLocal});
}
