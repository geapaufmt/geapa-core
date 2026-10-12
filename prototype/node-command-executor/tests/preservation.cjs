'use strict';
// No credentials, SDK open, remote calls or generated fixture snapshot files.
const a=require('node:assert/strict'),fs=require('node:fs'),p=require('node:path'),cp=require('node:child_process');
const root=p.resolve(__dirname,'..'),{materialize}=require('../scripts/materialize-baseline.cjs');
async function main(){
 const checks=[];const check=(name,fn)=>{fn();checks.push(name);};
 check('194 immutable v43 sources reconstructed from pinned Git commit',()=>a.equal(materialize().files,194));
 const S=require('../sdk.cjs'),gate=require('../private-session.cjs');delete process.env.B0_ALLOW_REMOTE;
 check('remote benchmark disabled by default',()=>a.throws(gate.remoteGate,/REMOTE_BENCHMARK_DISABLED/));
 process.env.B0_ALLOW_REMOTE='1';process.env.B0_NAMESPACE='B0_NODE_POC_20261012';
 check('original closed campaign cannot be rerun',()=>a.throws(gate.remoteGate,/ORIGINAL_CAMPAIGN_CLOSED/));
 delete process.env.B0_ALLOW_REMOTE;delete process.env.B0_NAMESPACE;
 check('non DEV project and environment fail closed',()=>{a.throws(()=>S.fence('geapa-prod','DEV'));a.throws(()=>S.fence(S.PROJECT,'PROD'));});
 const uid='B0_OFFLINE_SYNTHETIC',seed=require('../fixtures.cjs').seed(uid),memoryDocs={};
 for(const [path,d]of Object.entries(seed))memoryDocs[path]={...d,name:'projects/'+S.PROJECT+'/databases/(default)/documents/'+S.ROOT+'/'+path,updateTime:'2026-01-01T00:00:00.001Z'};
 check('synthetic fixture generator has no source snapshot dependency',()=>a(memoryDocs['canonicalActivities/dev_meeting']));
 process.env.B0_ACTOR_UIDS=[uid,'B0_OFFLINE_SELF','B0_OFFLINE_NONE'].join(',');
 const {Host}=require('../sdk-host.cjs'),host=new Host({projectId:S.PROJECT,databaseId:'(default)'},async()=>({uid,exp:Math.floor(Date.now()/1000)+3600}));
 const revision=Number(seed['canonicalActivities/dev_meeting'].fields.revision.integerValue);
 const first=await host.execute({memoryDocs,token:'OFFLINE_ONLY',mode:'edit',activityId:'dev_meeting',operationId:'b'.repeat(32),expectedRevision:revision,title:'B0 portability only'});
 a.equal(first.result.state,'COMPLETED');checks.push('original command contract executes through relocated worker in memory');
 const replay=await host.execute({memoryDocs:first.memoryAfter,token:'OFFLINE_ONLY',request:first.request});a.equal(replay.result.state,'RECOVERED');checks.push('same immutable request recovers');
 await a.rejects(()=>host.execute({memoryDocs:first.memoryAfter,token:'OFFLINE_ONLY',request:{...first.request,payload:{...first.request.payload,title:'Changed'}}}),/OPERATION_CONFLICT/);checks.push('changed payload cannot reuse operation ID');
 a.equal(host.touchedAll.size,0);a(host.metrics.every(x=>x.kind==='verify'));checks.push('all worker checks used zero Firestore calls/writes');
 const files=fs.readdirSync(root,{recursive:true}).filter(x=>!x.startsWith('node_modules')&&!x.startsWith('base')&&!x.startsWith('evidence')&&x.endsWith('.cjs'));
 for(const file of files){const r=cp.spawnSync(process.execPath,['--check',p.join(root,file)],{encoding:'utf8',windowsHide:true});a.equal(r.status,0,file+': '+r.stderr);}
 checks.push('all relocated CJS files pass syntax');
 console.log(JSON.stringify({passed:checks.length,failed:0,syntaxFiles:files.length,remoteCalls:0,checks},null,2));
}
main().catch(e=>{console.error({code:e.code||e.message,detail:e.detail,stack:e.stack});process.exitCode=1;});
