'use strict';
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GEAPA_PILOT_MODE:'SHARED_DEV',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {test}=require('node:test'),a=require('node:assert/strict'),{Firestore}=require('@google-cloud/firestore'),{setup}=require('./people-fixture.cjs'),{readAct}=require('../src/act-readers.cjs'),{PROJECT,DEMO}=require('../src/environment.cjs');
test('linked SELF read uses real Firestore transactions and rejects grant revocation before response',{skip:process.env.GEAPA_LINKED_READER_EMULATOR_TEST!=='1'},async()=>{
 const h=setup('DIRETOR_ADMINISTRATIVO',0),uid='CF_ADMIN',person='DEVADM_0123456789ab_1',old=h.docs['memberAuthBindings/'+uid].idPessoa,now=Date.now();
 const replace=v=>typeof v==='string'?v.replaceAll(old,person):Array.isArray(v)?v.map(replace):v&&typeof v==='object'&&Object.getPrototypeOf(v)===Object.prototype?Object.fromEntries(Object.entries(v).map(([k,x])=>[replace(k),replace(x)])):v;
 const docs=replace(h.docs),b=docs['memberAuthBindings/'+uid];docs['memberIdentities/'+person].origin='DEV_ADMIN_SYNTHETIC_V1';
 const sp=docs['memberAuthorityControl/MEMBERS'].selfReadPolicy;
 sp.fieldMasks.MEMBERS_SELF_READ=JSON.parse(h.domain.run('()=>f4MemberAuthzContracts_().selfFields'));
 docs['memberAuthzPolicies/'+sp.policyId]={schemaVersion:'member-authz-policy-v13',id:sp.policyId,revision:sp.policyRevision,active:true,origin:'F4_M621_INSTITUTIONAL',approvalRef:'SYN_APPROVAL',issuedAt:sp.issuedAt,expiresAt:sp.expiresAt,proofKind:'SELF',capabilities:['MEMBERS_SELF_READ'],actorSelector:{kind:'SELF'},targets:{ownPerson:true,allowedTargetScopes:[],allowedBoards:[],allowedRoles:[],allowedPersons:[],allowedResources:['memberCurrentStates/'+person]},fieldMasks:sp.fieldMasks};
 docs['memberAuthzPolicies/'+sp.policyId].targets.commandScopes=['OWN_MEMBER_RESOURCES'];
 docs['projectionRecoveryControl/current']={schemaVersion:'activity-projection-recovery-v1',projectId:PROJECT,namespace:'',mode:'DEV_READERS_OPERATIONAL',enabled:true,revision:1,startsAt:now-1000,expiresAt:now+3600000,actorUids:['OTHER'],policyIds:['PEOPLE_POLICY'],readModelCommands:[],browserEnabled:true,appCheck:{mode:'REQUIRED',appIds:['SYN_APP']},specializedReaders:{enabled:true,actorUids:['OTHER'],routes:['/members/own','/admin/pessoas']}};
 const grantPath='coreDevLinkedReaderGrants/'+uid;docs[grantPath]={schemaVersion:'core-dev-linked-reader-grant-v1',projectId:PROJECT,uid,personId:person,bindingVersion:b.bindingVersion,revision:1,active:true,startsAt:now-1000,expiresAt:now+3600000};
 const db=new Firestore({projectId:DEMO,databaseId:'linked-reader-test'}),auth={verifyIdToken:async(token,revoked)=>{a.equal(revoked,true);return {uid:token,sub:token,aud:DEMO,iss:'https://securetoken.google.com/'+DEMO,iat:Math.floor(now/1000),auth_time:Math.floor(now/1000),exp:Math.floor(now/1000)+3600}}},appCheck={verifyToken:async()=>({appId:'SYN_APP'})};
 const read=(hooks={},route='/members/own',token=uid,appCheckToken='emulator')=>readAct({db,auth,appCheck,token,appCheckToken,origin:'https://geapa-convergencia-dev.web.app',body:{action:'actRead',request:{route,params:route==='/members/own'?{view:'profile'}:{}}},hooks});
 try{await fetch('http://127.0.0.1:8187/emulator/v1/projects/'+DEMO+'/databases/linked-reader-test/documents',{method:'DELETE'});const batch=db.batch();for(const [p,v]of Object.entries(docs))batch.set(db.doc(p),v);await batch.commit();
 const answer=await read();a.equal(answer.result.data.idPessoa,person);a(answer.result.data.profileDetails);
 await a.rejects(read({},'/admin/pessoas'),/PILOT_DISABLED/);await a.rejects(read({},'/members/own','NO_GRANT'),/PILOT_DISABLED/);await a.rejects(read({},'/members/own',uid,null),/APP_CHECK_REQUIRED/);
 await a.rejects(read({beforeReturn:()=>db.doc(grantPath).update({active:false})}),/ACCESS_REVOKED/);
 await a.rejects(read(),/PILOT_DISABLED/);
 }finally{await db.terminate();}
});


