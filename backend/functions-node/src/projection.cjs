'use strict';
// Explicit awaited maintenance after commit, never detached work or a new HTTP command.
// Reuses the Core producer and CAS; no parallel implementation of the reader schema.
const {Store}=require('./firestore.cjs'),{createDomain}=require('./domain.cjs'),{authenticate}=require('./auth.cjs');
const {ACTIVITY,fail}=require('./environment.cjs');
async function rebuildActivityViews({db,auth,token,viewIds,hooks={}}){
 if(!Array.isArray(viewIds)||!viewIds.length||viewIds.length>2||new Set(viewIds).size!==viewIds.length||viewIds.some(id=>!['activities','activity_'+ACTIVITY].includes(id)))fail('PILOT_RESOURCE_DENIED');
 const identity=await authenticate(auth,token),store=new Store(db,identity),domain=createDomain(x=>store.io(x));
 if(!store.environment.emulator&&Object.keys(hooks).length)fail('DESTINATION_INVALID');
 const run=(fn,x)=>{const value=domain.run(fn,x);return value===undefined?undefined:JSON.parse(value);};run('(x)=>nodeIdentity(x)',identity);
 store.beforeCommit=hooks.beforeCommit;
 store.projectionViews=viewIds;
 const started=performance.now();
 const docs=await store.transaction(()=>run('(ids)=>nodeProject(ids)',viewIds));
 store.metrics.totalMs=performance.now()-started;
 return {state:'CURRENT',viewIds:Object.values(docs).map(x=>x.viewId),metrics:store.metrics};
}
const rebuildDetail=options=>rebuildActivityViews({...options,viewIds:['activity_'+ACTIVITY]});
// Internal maintenance adapter only. No new HTTP route: the existing executor
// uses it after a successful commit/recovery, for the authenticated actor only.
module.exports={rebuildDetail,rebuildActivityViews};
