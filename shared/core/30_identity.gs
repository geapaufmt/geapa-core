// F4A: native synchronous port of F3.5 identity/index.cjs. No external module loading.
function f4Identity_() {
  if (f4Identity_.value_) return f4Identity_.value_;
  return (f4Identity_.value_ = (function () {
'use strict';
const C=f4Core_();
const EXPECTED={environment:'SYNTHETIC_LOCAL',projectId:'demo-geapa-m04-local',database:'(default)',host:'127.0.0.1',port:8184};
function error(code){const e=new Error(code);e.code=code;throw e}
function need(v,code='IDENTITY_INVALID'){if(!v)error(code)}
function createIdentityVerifier(options){const expected=options.config||EXPECTED;C.destination(expected);const {publicKey,systemPublicKey,clock=Date.now}=options;need(publicKey&&systemPublicKey);const actorSet=new WeakSet();
 function decode(token,key){need(typeof token==='string'&&token.length<12000,'TOKEN_MALFORMED');try{const parts=token.split('.');need(parts.length===3&&parts.every(x=>/^[A-Za-z0-9_-]+$/.test(x)),'TOKEN_MALFORMED');const header=JSON.parse(f4Ed25519_().decodeText(parts[0]));C.exact(header,['alg','typ']);need(header.alg==='EdDSA'&&header.typ==='GEAPA_LOCAL_FIXTURE','TOKEN_MALFORMED');need(f4Ed25519_().verify(parts[0]+'.'+parts[1],key,parts[2]),'TOKEN_SIGNATURE_INVALID');return JSON.parse(f4Ed25519_().decodeText(parts[1]))}catch(e){if(e.code&&['TOKEN_MALFORMED','TOKEN_SIGNATURE_INVALID'].includes(e.code))throw e;error('TOKEN_MALFORMED')}}
 function verifyActor(context){try{need(context&&typeof context==='object');const portal=context.type==='PORTAL';C.exact(context,portal?['type','token']:['type','attestation']);need(['PORTAL','SYSTEM_JOB','SYSTEM_RECONCILIATION'].includes(context.type));const p=decode(portal?context.token:context.attestation,portal?publicKey:systemPublicKey);const now=clock();
 if(portal){C.exact(p,['iss','aud','sub','iat','exp','auth_time','projectId','environment','provenanceId','claims']);need(p.iss==='https://securetoken.google.com/'+expected.projectId,'ISSUER_INVALID');need(p.aud===expected.projectId,'AUDIENCE_INVALID');need(C.id(p.sub),'UID_INVALID');need(p.claims&&typeof p.claims==='object'&&!Array.isArray(p.claims),'TOKEN_MALFORMED')}
 else{C.exact(p,['type','jobName','executionId','projectId','environment','iat','exp']);need(p.type===context.type&&C.id(p.executionId),'SYSTEM_PROVENANCE_INVALID');need(p.jobName===(p.type==='SYSTEM_JOB'?'outbox-fixture':'maintenance-fixture'),'SYSTEM_PROVENANCE_INVALID')}
 need(p.projectId===expected.projectId,'PROJECT_INVALID');need(p.environment===expected.environment,'ENVIRONMENT_INVALID');need(Number.isSafeInteger(p.iat)&&Number.isSafeInteger(p.exp)&&p.iat<=now&&p.iat<p.exp&&now<p.exp,'TOKEN_TIME_INVALID');let actor;
 if(portal){need(Number.isSafeInteger(p.auth_time)&&p.auth_time<=p.iat&&p.auth_time>=0&&C.id(p.provenanceId),'PROVENANCE_INVALID');actor={uid:p.sub,authTime:p.auth_time,issuedAt:p.iat,expiresAt:p.exp,projectId:p.projectId,environment:p.environment,provenanceType:'PORTAL_AUTH',provenanceId:p.provenanceId}}
 else actor={provenanceType:p.type,jobName:p.jobName,environment:p.environment,projectId:p.projectId,executionId:p.executionId,issuedAt:p.iat,expiresAt:p.exp};
 Object.freeze(actor);actorSet.add(actor);return actor;
 }catch(e){if(e.code==='SCHEMA_INVALID')error('IDENTITY_INVALID');if(e.code)throw e;error('IDENTITY_INVALID')}}
 function trusted(actor,type){need(actorSet.has(actor)&&actor.provenanceType===type&&actor.expiresAt>clock(),'IDENTITY_INVALID');return actor}
 return {verifyActor,trusted};
}
return {createIdentityVerifier,EXPECTED,error,need};

  })());
}
