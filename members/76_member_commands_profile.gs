// M64-E A/B: pure family producers; no IO outside the supplied transaction reader.
function f4MemberProfileCommands_(){
 const C=f4Core_(),K=f4MemberCommandCatalog_(),M=f4MemberModel_(true),N=C.need;
 const copy=v=>f4Firestore_().decode(f4Firestore_().encode(v));
 function text(v,max){N(typeof v==='string'&&v.length<=max&&!/[\x00-\x1f<>]/.test(v),'COMMAND_PAYLOAD_INVALID')}
 function validate(r){
  if(f4MemberProfileCorrections_().is(r))return f4MemberProfileCorrections_().validate(r);
  N(['members.self.update','members.profile.correct'].includes(r.commandType),'COMMAND_TYPE_INVALID');
  C.exact(r.target,['personId']);N(C.id(r.target.personId),'COMMAND_PAYLOAD_INVALID');
  if(r.commandType==='members.self.update'){
   C.exact(r.payload,['patch']);f4MemberProfileContract_().patch(r.payload.patch,true,'9999-12-31');N(r.decisionRef===null&&r.evidenceRefs.length===0,'COMMAND_PAYLOAD_INVALID');
  }else{
   C.exact(r.payload,['corrections','justification']);text(r.payload.justification,120);N(r.payload.justification.trim().length>0,'COMMAND_PAYLOAD_INVALID');
   const patch=r.payload.corrections;N(patch&&Object.getPrototypeOf(patch)===Object.prototype&&Object.keys(patch).length>0,'COMMAND_PAYLOAD_INVALID');
   f4MemberProfileContract_().patch(patch,false,'9999-12-31');
   N(r.evidenceRefs.length>0,'COMMAND_PAYLOAD_INVALID');
  }
  return true;
 }
 function resolve(get,r,now){
  validate(r);const {request:q,proof}=f4MemberFamilyInputs_().authorization(get,r);
  N(q.target.personId===r.target.personId,'AUTH_BINDING_INVALID');
  const self=r.commandType==='members.self.update',W=f4MemberProfileCorrections_(),workflow=W.is(r),patch=workflow?null:f4MemberProfileContract_().patch(self?r.payload.patch:r.payload.corrections,self,Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd'));
  N(q.capability===(self?'MEMBERS_SELF_UPDATE':'MEMBERS_PROFILE_CORRECT')&&(!self||['SELF','DEV_ADMIN'].includes(q.kind)),'AUTH_BINDING_INVALID');
  if(!self)f4MemberActContent_().empty(get,r,now);
  N(K.equal([...q.target.fieldMask].sort(),workflow?(self?[]:[r.payload.field]):Object.keys(patch).sort()),'AUTH_BINDING_INVALID');
  f4MemberAuthz_().recheck(get,q,proof,now);
  const changes=workflow?W.changes(get,r,now,profileChanges):profileChanges(get,r,patch,now);
  N(K.equal([...q.target.resources].sort(),changes.map(c=>c.path).sort()),'AUTH_BINDING_INVALID');
  return {id:r.target.personId,expected:f4MemberFamilyInputs_().revision(r,'memberAggregateHeads/'+r.target.personId),changes,now,boardExpected:{},mode:'M13',authorization:{request:q,proof},deadline:C.ms(proof.expiresAt)};
 }
 function profileChanges(get,r,patch,now){
  for(const field of ['courseId','municipalityCode','interestAxisIds'])if(Object.hasOwn(patch,field)){if(field==='courseId'&&patch[field])N(get('memberCourseCatalog/'+patch[field])?.active===true,'COMMAND_PAYLOAD_INVALID');if(field==='municipalityCode'&&patch[field]){const city=get('memberMunicipalityCatalog/'+patch[field]);N(city?.active===true,'COMMAND_PAYLOAD_INVALID');}if(field==='interestAxisIds')for(const id of patch[field])N(get('presentationPolicyReferences/'+id)?.active===true&&get('presentationPolicyReferences/'+id)?.approved===true,'COMMAND_PAYLOAD_INVALID');}
  const changes=[];
  for(const [col,fields]of [['memberIdentities',['displayName']],['memberPrivateProfiles',['contactEmail','academicReference',...f4MemberProfileContract_().extra]]]){
   const applied=fields.filter(f=>Object.hasOwn(patch,f));if(!applied.length)continue;
   const path=col+'/'+r.target.personId,old=get(path);N(old?old.idPessoa===r.target.personId:col==='memberPrivateProfiles','PROFILE_MISSING');
   const revision=f4MemberFamilyInputs_().revision(r,path);N((old?.revision||0)===revision,'REVISION_CONFLICT');
   const data={...(old?copy(old):{schemaVersion:'m1-memberPrivateProfiles-v2',id:r.target.personId,idPessoa:r.target.personId,origin:'M13_COMMAND',contactEmail:'',academicReference:''}),revision:revision+1};if(col==='memberPrivateProfiles'&&applied.some(f=>f4MemberProfileContract_().extra.includes(f)))data.schemaVersion='m1-memberPrivateProfiles-v2';for(const f of applied)data[f]=patch[f];M.validate(col,data);
   // Validate the merged origin, including an existing municipality code. A
   // city/UF edit must not leave that stable code pointing to another place.
   if(col==='memberPrivateProfiles'&&['countryCode','stateCode','city','municipalityCode'].some(k=>Object.hasOwn(patch,k))&&data.municipalityCode){const city=get('memberMunicipalityCatalog/'+data.municipalityCode);N(city?.active===true&&city.countryCode===data.countryCode&&city.stateCode===data.stateCode&&city.name===data.city,'COMMAND_PAYLOAD_INVALID');}
   changes.push({path,expectedRevision:revision,data});
  }
  return changes;
 }
 return {validate,resolve};
}

// Protected inputs, reused by E producers. Client selectors never contain these values.
function f4MemberFamilyInputs_(){
 const C=f4Core_(),K=f4MemberCommandCatalog_(),N=C.need;
 function revision(r,path){const v=r.expectedRevisions.find(x=>x.resourceType+'/'+x.id===path);N(v&&C.rev(v.revision),'REVISION_CONFLICT');return v.revision}
 function authorization(get,r){
  const saved=get('commandRequests/'+r.operationId),p=get('commandProvenance/'+r.operationId),t=get('memberAdmissionTickets/'+r.ticketId);
  N(K.equal(saved,r)&&p&&t&&p.operationId===r.operationId&&p.ownerUid===r.ownerUid&&p.digest===r.digest&&p.ticketId===r.ticketId&&t.digest===r.digest&&t.ownerUid===r.ownerUid&&K.equal(t.requestEnvelope,K.template(r)),'AUTH_BINDING_INVALID');
  N(p.authRequests.length===1&&t.authProofs.length===1,'AUTH_BINDING_INVALID');const q=p.authRequests[0],proof=t.authProofs[0];
  const binding=['members.binding.associate','members.binding.revoke'].includes(r.commandType);
  N(q.uid===r.ownerUid&&q.capability===f4SharedCommandCatalog_().capabilities(r)[0]&&(binding?q.evidenceRefs.length===0:K.equal(q.evidenceRefs,r.evidenceRefs)),'AUTH_BINDING_INVALID');
  if(r.decisionRef!==null)for(const id of r.evidenceRefs){const e=get('memberEvidenceAttestations/'+id);N(e&&e.approvalRef===r.decisionRef,'AUTH_BINDING_INVALID');}
  return {request:q,proof};
 }
 return {revision,authorization};
}
