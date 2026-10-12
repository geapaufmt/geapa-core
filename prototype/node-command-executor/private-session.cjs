'use strict';
// Credentials remain outside Git. No account, key, grant or IAM is created here.
const a=require('node:assert/strict'),p=require('node:path');
function remoteGate(){
 a.equal(process.env.B0_ALLOW_REMOTE,'1','REMOTE_BENCHMARK_DISABLED');
 a(/^B0_NODE_POC_[A-Za-z0-9_-]+$/.test(process.env.B0_NAMESPACE||''),'NEW_B0_NAMESPACE_REQUIRED');
 a.notEqual(process.env.B0_NAMESPACE,'B0_NODE_POC_20261012','ORIGINAL_CAMPAIGN_CLOSED');
 a(process.env.B0_ACTOR_UIDS?.split(',').filter(Boolean).length>=3,'EXPLICIT_SYNTHETIC_ACTORS_REQUIRED');
}
let provider;
function load(){
 remoteGate();if(provider)return provider;
 const file=process.env.B0_SESSION_PROVIDER;a(file&&p.isAbsolute(file),'PRIVATE_SESSION_PROVIDER_REQUIRED');
 const repo=p.resolve(__dirname,'../..');a(!p.resolve(file).toLowerCase().startsWith(repo.toLowerCase()+p.sep),'CREDENTIAL_PROVIDER_MUST_STAY_OUTSIDE_REPO');
 provider=require(file);a.equal(typeof provider.secrets,'function');a.equal(typeof provider.authApi,'function');return provider;
}
function secrets(){const s=load().secrets();a.equal(s.project,'geapa-convergencia-dev');const allowed=process.env.B0_ACTOR_UIDS.split(',');for(const kind of ['BOARD','SELF','NONE'])a(s.accounts[kind]?.uid&&allowed.includes(s.accounts[kind].uid),'EXPLICIT_ACTOR_MISMATCH');return s;}
async function authApi(action,body){a.equal(action,'signInWithPassword','AUTH_MUTATION_NOT_ALLOWED');return load().authApi(action,body);}
module.exports={remoteGate,secrets,authApi};
