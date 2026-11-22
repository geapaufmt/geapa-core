'use strict';
// Operator-only Admin SDK command. No HTTP endpoint and no client credentials.
const fs=require('node:fs'),crypto=require('node:crypto'),a=require('node:assert/strict');
const PROJECT='geapa-convergencia-dev',ORIGIN='DEV_ADMIN_SYNTHETIC_V1';
const hash=v=>crypto.createHash('sha256').update(JSON.stringify(v)).digest('hex');
const fields={PEOPLE_DIRECTORY_READ:['idPessoa','displayLabel','state','categories','functions','relations','history'],MEMBERS_PROFILE_CORRECT:['displayName','cpf','rga','birthDate','countryCode','courseId'],INSTITUTIONAL_DIRECTORY_READ:['summary','fields','provenance','issues','canonical'],INSTITUTIONAL_GOVERNANCE_DRAFT:['draft']};
function fixtures(uid,now){const suffix=hash(uid).slice(0,12),docs={},persons=[1,2].map(n=>'DEVADM_'+suffix+'_'+n);
 for(const [n,id]of persons.entries()){const link=id+'_LINK',request=hash({uid,id,purpose:'DEV_ADMIN_CORRECTION_FIXTURE'}).slice(0,32),name='Pessoa sintética de desenvolvimento '+(n+1);
  docs['memberIdentities/'+id]={schemaVersion:'m1-memberIdentities-v1',id,idPessoa:id,revision:1,origin:ORIGIN,displayName:name,technicalAccount:false,active:true};
  docs['memberPrivateProfiles/'+id]={schemaVersion:'m1-memberPrivateProfiles-v1',id,idPessoa:id,revision:1,origin:ORIGIN,contactEmail:id.toLowerCase()+'@example.invalid',academicReference:'SYNTHETIC'};
  docs['memberLinks/'+link]={schemaVersion:'m1-memberLinks-v1',id:link,idPessoa:id,revision:1,origin:ORIGIN,type:'MEMBRO_EFETIVO',startDate:'2026-01-01',endDate:null,exitDate:null};
  docs['memberAggregateHeads/'+id]={schemaVersion:'m11-head-v1',id,kind:'PERSON',revision:1,paths:['memberIdentities/'+id,'memberPrivateProfiles/'+id,'memberLinks/'+link]};
  docs['memberProfileCorrections/'+request]={schemaVersion:'member-profile-correction-v1',id:request,personId:id,revision:1,field:'displayName',previousValue:name,proposedValue:name+' revisada',baseValueHash:hash(name),status:'PENDENTE',createdAt:now,updatedAt:now,history:[{operationId:request,actorUid:uid,action:'SUBMIT',at:now,reason:'Fixture sintética autorizada para homologação DEV'}]};
  docs['memberProfileCorrectionIndexes/'+id]={schemaVersion:'member-profile-corrections-index-v1',id,revision:1,requestIds:[request],activeByField:{displayName:request}};
 }return {persons,docs};
}
async function administer({db,uid,personId,approvalRef,operator,days=30,expectedGrantRevision,expectedRouteRevision,expectedGovernanceRevision,apply=false,revoke=false,clock=Date.now}){
 a.equal(db.projectId,PROJECT,'DEV_ONLY');for(const v of [uid,personId,approvalRef])a.match(v,/^[A-Za-z0-9_-]{1,48}$/);a(days>0&&days<=30&&Number.isInteger(days));a(Number.isInteger(expectedGrantRevision));
 const now=clock(),end=now+days*86400000,fixture=fixtures(uid,now),grantPath='coreDevAdministrativeGrants/'+uid,routePath='commandExecutorRoutes/MEMBERS_PROFILE_CORRECT',govPath='institutionalGovernanceControl/current',gatePath='portalReaderControl/current',registryPath='coreDevAdministrativeFixtures/'+uid;
 return db.runTransaction(async tx=>{const reads=new Map(),get=async p=>{if(!reads.has(p)){const s=await tx.get(db.doc(p));reads.set(p,s.exists?s.data():null);}return reads.get(p);};
  const old=await get(grantPath),binding=await get('memberAuthBindings/'+uid),subject=await get('memberAuthzSubjects/'+personId),gate=await get(gatePath),route=await get(routePath),gov=await get(govPath),registry=await get(registryPath);
  a.equal(old?.revision||0,expectedGrantRevision,'GRANT_CHANGED');a(binding?.active&&binding.idPessoa===personId&&subject?.active,'BINDING_REQUIRED');a.equal(gate?.projectId,PROJECT);a.equal(route?.projectId,PROJECT);a.equal(gov?.projectId,PROJECT);a.equal(route.revision,expectedRouteRevision,'ROUTE_CHANGED');a.equal(gov.revision,expectedGovernanceRevision,'GOVERNANCE_CHANGED');
  const grant=revoke?{...old,active:false,revision:old.revision+1}:{schemaVersion:'core-dev-administrative-grant-v1',id:uid,uid,personId,bindingVersion:binding.bindingVersion,subjectEpoch:subject.authzEpoch,projectId:PROJECT,environment:'DEV',exportable:false,active:true,revision:(old?.revision||0)+1,issuedAt:new Date(now),expiresAt:new Date(end),approvalRef,capabilities:Object.keys(fields),fieldMasks:fields,syntheticPersonIds:fixture.persons,sections:['identity','parameters','norms','governance']};
  a(!revoke||old,'GRANT_MISSING');const writes={[grantPath]:grant,[gatePath]:{...gate,generation:gate.generation+1}};
  if(!revoke){
   // Never reopen a shared window or displace recoverable admissions.
   a(route.actorUids.length===1&&route.actorUids[0]===uid||route.expiresAt<=now&&route.operationIds.length===0,'ROUTE_REQUIRES_RECONCILIATION');
   a(gov.actorUids.length===1&&gov.actorUids[0]===uid||gov.enabled===false,'GOVERNANCE_REQUIRES_RECONCILIATION');
   const scopePath='memberExecutorScopes/'+route.scopeId,scope=await get(scopePath);a(scope?.synthetic===true&&scope.projectId===PROJECT);a(scope.personIds.length+fixture.persons.filter(id=>!scope.personIds.includes(id)).length<=100);
   writes[scopePath]={...scope,revision:scope.revision+1,personIds:[...new Set([...scope.personIds,...fixture.persons])]};
   writes[routePath]={...route,revision:route.revision+1,actorUids:[uid],startsAt:now,expiresAt:end};
   writes[govPath]={...gov,revision:gov.revision+1,enabled:true,actorUids:[uid],availabilityMode:'TIMEBOXED',startsAt:now,expiresAt:end};
   if(registry){a.equal(registry.origin,ORIGIN);a.deepEqual(registry.personIds,fixture.persons);for(const id of fixture.persons)a.equal((await get('memberIdentities/'+id))?.origin,ORIGIN);}
   else{for(const [p,v]of Object.entries(fixture.docs)){a.equal(await get(p),null,'FIXTURE_COLLISION');writes[p]=v;}writes[registryPath]={schemaVersion:'core-dev-administrative-fixtures-v1',uid,origin:ORIGIN,personIds:fixture.persons,approvalRef,createdAt:new Date(now)};}
  }
  const auditPath='coreDevAdministrativeAudit/'+hash({uid,revision:grant.revision,approvalRef}).slice(0,32);a.equal(await get(auditPath),null,'AUDIT_EXISTS');
  writes[auditPath]={schemaVersion:'core-dev-administrative-audit-v1',uid,operator,approvalRef,action:revoke?'REVOKE':'ISSUE_OR_RENEW',at:new Date(now),grantRevision:grant.revision,before:Object.fromEntries([...reads].filter(([p])=>Object.hasOwn(writes,p))),changedPaths:Object.keys(writes),afterGrant:grant};
  if(apply)for(const [p,v]of Object.entries(writes))tx.set(db.doc(p),v);
  return {projectId:PROJECT,applied:apply,uid,grantRevision:grant.revision,expiresAt:grant.expiresAt,paths:Object.keys(writes),fixturePersons:fixture.persons};
 });
}
async function main(){const file=process.argv[2];a(file,'INPUT_FILE_REQUIRED');const input=JSON.parse(fs.readFileSync(file,'utf8'));a.deepEqual(Object.keys(input).sort(),['approvalRef','days','expectedGrantRevision','expectedGovernanceRevision','expectedRouteRevision','operator','personId','uid'].sort());
 const token=process.env.GEAPA_ADMIN_ACCESS_TOKEN;a(token,'ADMIN_TOKEN_REQUIRED');const info=await fetch('https://www.googleapis.com/oauth2/v2/userinfo',{headers:{Authorization:'Bearer '+token},redirect:'error'});a(info.ok);a.equal((await info.json()).email,input.operator);
 const {OAuth2Client,GoogleAuth}=require('google-auth-library'),{Firestore}=require('@google-cloud/firestore'),client=new OAuth2Client();client.setCredentials({access_token:token});const db=new Firestore({projectId:PROJECT,auth:new GoogleAuth({authClient:client,projectId:PROJECT,quotaProjectId:PROJECT})});
 try{console.log(JSON.stringify(await administer({db,...input,apply:process.argv.includes('--apply'),revoke:process.argv.includes('--revoke')})));}finally{await db.terminate();}
}
if(require.main===module)main().catch(e=>{console.error(e.code||e.message);process.exitCode=1});module.exports={administer,fixtures};
