// Internal factories only. No web app/RPC/trigger surface; no automatic initialization.
function f4aTarget_() {
 return Object.freeze({environment:'DEV',projectId:'geapa-convergencia-dev',database:'(default)',host:'firestore.googleapis.com',port:443});
}
function f4aCreateFixtureBackend_(publicKey,systemPublicKey,clock) {
 // Keys belong to the future trusted fixture issuer, never to a browser request.
 const config=f4aTarget_();
 if(ScriptApp.getScriptId()!=='1M3GpuWwyRkV24nbHnopzj1q4J-Ea7zHv5eCG-ngOw0qR5lWDMvvv8UfD')f4Core_().fail('DESTINATION_INVALID');
 const identity=f4Identity_().createIdentityVerifier({config,publicKey,systemPublicKey,clock:clock||Date.now});
 const db=f4Firestore_().connect(config);
 return f4Lifecycle_().createOperational({config,db,identity,clock:clock||Date.now});
}
function f4Workspace_() {
 const denied=()=>f4Core_().fail('DESTINATION_INVALID');
 return Object.freeze({drive:denied,gmail:denied,forms:denied,sheets:denied});
}
function f4Jobs_(backend) {
 return Object.freeze({processCommands:(actor,operationId,workerId)=>backend.process(actor,operationId,workerId),maintenance:Object.freeze({releaseSlot:(actor,handle)=>backend.releaseSlot(actor,handle),purge:(actor,operationId)=>backend.purge(actor,operationId),expireResult:(actor,operationId)=>backend.expireResult(actor,operationId),resume:(actor,operationId,workerId)=>backend.resume(actor,operationId,workerId)}),outboxTransition:(actor,operationId,transition)=>backend.outboxTransition(actor,operationId,transition),reconcileOutbox:(actor,operationId,observed)=>backend.resolveOutbox(actor,operationId,observed)});
}
