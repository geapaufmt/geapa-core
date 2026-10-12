'use strict';
const assert=require('node:assert/strict');
const {Firestore}=require('@google-cloud/firestore');
const {OAuth2Client}=require('google-auth-library');
const PROJECT='geapa-convergencia-dev',ROOT='coreDevValidation/'+(process.env.B0_NAMESPACE||'B0_NODE_POC_20261012');
function fence(project=PROJECT,environment='DEV'){
 assert.equal(project,PROJECT,'B0_PROJECT_DENIED');assert.equal(environment,'DEV','B0_ENVIRONMENT_DENIED');
 assert(!process.env.FIRESTORE_EMULATOR_HOST,'REAL_DEV_BENCHMARK_REQUIRED');
 assert(!process.env.GOOGLE_APPLICATION_CREDENTIALS,'EXPLICIT_EXISTING_DEV_OAUTH_ONLY');
}
async function open(){
 fence();require('./private-session.cjs').remoteGate();const lib=require('node:path').join(process.env.B0_FIREBASE_TOOLS_ROOT||require('node:path').dirname(require.resolve('firebase-tools/package.json')),'lib');
 const auth=require(lib+'/auth.js'),api=require(lib+'/api.js');
 const accounts=auth.getAllAccounts().filter(a=>a.user.email==='geapadev@gmail.com'&&a.user.email_verified===true);
 assert.equal(accounts.length,1,'DEV_ACCOUNT_UNIQUE');
 const credentials={type:'authorized_user',client_id:api.clientId(),client_secret:api.clientSecret(),refresh_token:accounts[0].tokens.refresh_token};
 const identity=new OAuth2Client(credentials.client_id,credentials.client_secret);identity.setCredentials({refresh_token:credentials.refresh_token});
 const verification=await identity.request({url:'https://www.googleapis.com/oauth2/v2/userinfo',timeout:20000,retry:true,retryConfig:{retry:2,httpMethodsToRetry:['GET']}});
 assert.equal(verification.data.email,'geapadev@gmail.com','DEV_IDENTITY_DENIED');assert.equal(verification.data.verified_email,true);
 const db=new Firestore({projectId:PROJECT,databaseId:'(default)',credentials,preferRest:false});
 db.b0RpcEvents=[];for(const method of ['request','requestStream']){const original=db[method].bind(db);db[method]=(...args)=>{db.b0RpcEvents.push({method:args[0],at:performance.now()});return original(...args);};}
 return {db,identityEmail:verification.data.email,credential:{getAccessToken:async()=>{const {token}=await identity.getAccessToken();assert(token);return {access_token:token,expires_in:Math.max(1,Math.floor(((identity.credentials.expiry_date||Date.now()+300000)-Date.now())/1000))};}},ref(path){assert(typeof path==='string'&&path&&!path.startsWith('/')&&!path.includes('..'),'INVALID_RELATIVE_PATH');return db.doc(ROOT+'/'+path);}};
}
module.exports={open,fence,PROJECT,ROOT};
if(require.main===module)(async()=>{const start=performance.now(),{db,ref,identityEmail}=await open(),ready=performance.now(),d=await ref('b0Control/probe').get();console.log({project:PROJECT,namespace:ROOT,identity:identityEmail,exists:d.exists,credentialMs:ready-start,sdkReadMs:performance.now()-ready,remoteWrites:0});await db.terminate();})().catch(e=>{console.error({code:e.code||e.name,message:e.code==='ERR_ASSERTION'?e.message:'SDK_READ_FAILED'});process.exitCode=1;});
