// Documentary custody and human verification only; never normative approval.
function f4InstitutionalDocumentManagement_(){
 const C=f4Core_(),N=f4MemberAuthzContracts_().need,L=f4InstitutionalDocumentaryLinks_();
 const caps={NORM_FILES_SAVE:'INSTITUTIONAL_NORM_DOCUMENT_MANAGE',DOCUMENTARY_LINK_REVIEW:'INSTITUTIONAL_DOCUMENTARY_VERIFY'};
 const hash=L.hash,text=(v,min,max)=>typeof v==='string'&&v===v.trim()&&v.length>=min&&v.length<=max;
 function norm(get,o,c,ref){
  C.exact(ref,['id','revision','hash']);N(C.id(ref.id)&&C.rev(ref.revision)&&/^[a-f0-9]{64}$/.test(ref.hash),'COMMAND_PAYLOAD_INVALID');
  N(c.documentaryNormVersionIds?.includes(ref.id),'SCOPE_INVALID');
  const auth=f4InstitutionalPolicies_(get,o.identity,o.clock(),get('institutionalDirectoryControl/current').policyIds,'INSTITUTIONAL_DIRECTORY_READ','institutionalDirectory/norms');
  N(auth.mask.includes('fields')&&auth.mask.includes('canonical'),'AUTHORIZATION_DENIED');
  N(get('institutionalDirectory/norms')?.records.some(r=>r.visibility==='BOARD'&&r.canonical?.path==='institutionalNormVersions/'+ref.id),'AUTHORIZATION_DENIED');
  const n=get('institutionalNormVersions/'+ref.id);N(n&&n.id===ref.id&&n.revision===ref.revision&&hash(n)===ref.hash,'NORMATIVE_VERSION_CONFLICT');return n;
 }
 function prepare(get,o,input,v,c){
  N(caps[input.action]&&v.revision===input.expectedRevision&&text(input.reason,10,1000),'REVISION_CONFLICT');
  const n=norm(get,o,c,input.norm),now=o.clock();
  if(input.action==='NORM_FILES_SAVE'){
   C.exact(input,['action','id','expectedRevision','expectedRecordRevision','norm','attachments','officialSha256','reason']);
   N(c.normFilesEnabled===true&&c.documentFilesEnabled===true,'PILOT_DISABLED');
   const old=v.records.find(r=>r.kind==='NORM_DOCUMENT_FILES'&&r.normativeFiles.norm.id===n.id);
   N(old?input.id===old.id&&input.expectedRecordRevision===old.revision:input.id===''&&input.expectedRecordRevision===0,'REVISION_CONFLICT');
   const files=f4InstitutionalDocumentFiles_().prepare(input,old);N(files.length>0,'DOCUMENT_FILE_INVALID');
   const official=files.find(f=>(f.existing||f.upload).sha256===input.officialSha256);N(official&&(official.existing||official.upload).mime==='application/pdf','DOCUMENT_FILE_INVALID');
   if(old)N(hash(old.attachments.map(f=>({fileId:f.fileId,caption:f.caption})))!==hash(input.attachments)||old.normativeFiles.officialSha256!==input.officialSha256,'DOCUMENTARY_NO_CHANGE');
   return {old,attachments:files,next:{kind:'NORM_DOCUMENT_FILES',title:'Arquivos da versão normativa '+n.id,meetingDate:null,body:'Custódia documental. O arquivo não altera aprovação ou vigência.',documentaryState:'DOCUMENTED',sourceRecordId:null,references:[],formalApproval:null,publication:'CLOSED',effectiveFrom:null,createdAt:old?.createdAt||now,normativeFiles:{norm:input.norm,officialSha256:input.officialSha256,meaning:'DOCUMENT_CUSTODY_ONLY'},versions:[...(old?.versions||[]),...(old?[{contentRevision:old.revision,record:Object.fromEntries(Object.entries(old).filter(([k])=>k!=='versions'))}]:[])]}};
  }
  C.exact(input,['action','id','expectedRevision','expectedRecordRevision','norm','targetId','targetRevision','targetHash','relation','decision','evidence','reason']);
  N(c.documentaryVerificationEnabled===true,'PILOT_DISABLED');N(input.id===''&&input.expectedRecordRevision===0,'COMMAND_PAYLOAD_INVALID');
  N(['VERIFIED','REJECTED','REVOKED'].includes(input.decision)&&text(input.evidence,10,2000),'COMMAND_PAYLOAD_INVALID');
  const target=v.records.find(r=>r.id===input.targetId);N(target&&target.revision===input.targetRevision&&hash(target)===input.targetHash,'REVISION_CONFLICT');
  const ra=f4InstitutionalPolicies_(get,o.identity,now,get('institutionalDirectoryControl/current').policyIds,'INSTITUTIONAL_DIRECTORY_READ','institutionalDirectory/governance');N(!target.minutes||f4InstitutionalMinutes_().accessible(ra,target),'AUTHORIZATION_DENIED');
  const relation=target.documentaryEvidence?.relations.find(r=>r.normVersionId===n.id&&r.relation===input.relation);N(relation&&relation.expectedNormRevision===n.revision&&relation.expectedNormHash===input.norm.hash,'NORMATIVE_VERSION_CONFLICT');
  N(target.attachments?.some(f=>f.sha256===target.documentaryEvidence.sourceSha256),'DOCUMENTARY_SOURCE_FILE_REQUIRED');
  const previous=reviews(v.records,target,n.id).find(r=>r.review.relation===input.relation);
  N(input.decision!=='REVOKED'||previous?.review.decision==='VERIFIED','DOCUMENTARY_REVIEW_INVALID');N(previous?.review.decision!==input.decision,'DOCUMENTARY_NO_CHANGE');
  return {old:null,attachments:[],next:{kind:'DOCUMENTARY_LINK_REVIEW',title:'Conferência documental — '+target.title,meetingDate:null,body:input.evidence,documentaryState:'DOCUMENTED',sourceRecordId:null,references:[],formalApproval:null,publication:'CLOSED',effectiveFrom:null,createdAt:now,review:{targetId:target.id,targetRevision:target.revision,targetHash:input.targetHash,norm:input.norm,relation:input.relation,decision:input.decision,evidence:input.evidence,meaning:'DOCUMENTARY_VERIFICATION_NOT_NORM_APPROVAL',previousReviewId:previous?.id||null}}};
 }
 function reviews(records,target,normId){return records.filter(r=>r.kind==='DOCUMENTARY_LINK_REVIEW'&&r.review.targetId===target.id&&r.review.targetRevision===target.revision&&r.review.targetHash===hash(target)&&(!normId||r.review.norm.id===normId)).sort((a,b)=>b.updatedAt-a.updatedAt||b.id.localeCompare(a.id)).filter((r,i,all)=>all.findIndex(t=>t.review.norm.id===r.review.norm.id&&t.review.relation===r.review.relation)===i);}
 function readable(get,o,r,auth){
  if(r.minutes&&!f4InstitutionalMinutes_().accessible(auth,r))return false;
  if(!r.kind)return true;
  if(r.kind==='BOARD_PROCEDURE')return auth.actor.derived.currentDirectors.some(f=>f.idDiretoria===r.procedure.boardId);
  const ref=r.normativeFiles?.norm||r.review?.norm;if(!ref)return false;
  try{const a=f4InstitutionalPolicies_(get,o.identity,o.clock(),get('institutionalDirectoryControl/current').policyIds,'INSTITUTIONAL_DIRECTORY_READ','institutionalDirectory/norms');if(!a.mask.includes('fields')||!a.mask.includes('canonical'))return false;
   if(!get('institutionalDirectory/norms')?.records.some(n=>n.visibility==='BOARD'&&n.canonical?.path==='institutionalNormVersions/'+ref.id))return false;
   if(r.review){const target=get('institutionalGovernanceDrafts/current')?.records.find(t=>t.id===r.review.targetId);return !!target&&!target.kind&&(!target.minutes||f4InstitutionalMinutes_().accessible(auth,target));}return true;
  }catch(e){if(['AUTHORIZATION_DENIED','M13_AUTHZ_DENIED'].includes(e.code))return false;throw e;}
 }
 return {caps,prepare,reviews,hash,readable};
}
