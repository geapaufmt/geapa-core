'use strict';
const {sample}=require('./flow-fixture.cjs'),{decode}=require('../src/firestore.cjs');
const extra=['ATTENDANCE_PROOF_UPLOAD_REQUEST','ACTIVITY_TYPE_SAVE','ACTIVITY_POLICY_SAVE','ACTIVITY_SEMESTER_SAVE','CYCLE_GOVERNANCE_EDIT','CORRECTION_TERMS_REVISE','PHOTO_UPLOAD_REQUEST','PHOTO_SUBMIT','PHOTO_REVIEW'];
function remaining(command){
 if(!extra.includes(command))return sample(command);
 const s=sample(command==='CORRECTION_TERMS_REVISE'?'MATERIAL_CORRECTION_SUBMIT':'MATERIAL_REVIEW');
 const base=Object.fromEntries(Object.entries(require('./fixtures.cjs').seed().docs).map(([p,v])=>[p,decode({mapValue:{fields:v.fields}})]));
 let payload,self=false;
 const decision={reasonCode:'DEV_FIXTURE',reason:'Synthetic operational test'},c='c'.repeat(32);
 if(command==='ACTIVITY_TYPE_SAVE'){
  const type=base['activityTypes/meeting'];payload={typeRef:{kind:'activityType',id:'meeting'},...Object.fromEntries(['countsAttendance','countsAbsence','durationMinutes','visibility','allowsJustification','status','attendancePolicyRef'].map(k=>[k,type[k]])),name:'Synthetic renamed model',decision};
 }
 if(command==='ACTIVITY_POLICY_SAVE'){
  const policy=base['activityPolicies/fn_policy'];payload={attendancePolicyRef:{kind:'attendancePolicy',id:'fn_policy'},...Object.fromEntries(['cycleRef','frequencyBasis','opensBeforeMinutes','closesAfterMinutes','previousJustificationDays','justificationAfterDays','acceptLateJustification'].map(k=>[k,policy[k]])),decision};
 }
 if(command==='ACTIVITY_SEMESTER_SAVE')payload={semesterRef:{kind:'academicSemester',id:'synthetic_semester'},label:'Synthetic semester',civilStart:'2027-03-01',civilEnd:'2027-07-01',status:'PLANNED',reason:'Synthetic semester fixture'};
 if(command==='CYCLE_GOVERNANCE_EDIT'){
  const cycle=s.docs['presentationCycles/'+c];payload={cycleRef:{kind:'cycle',id:c},label:'Synthetic relabel',civilStart:cycle.civilStart,civilEnd:cycle.civilEnd,academicSemesterRefs:[],confirmAssociations:true,reason:'Synthetic cycle fixture'};
 }
 if(command==='CORRECTION_TERMS_REVISE'){
  const o=s.docs['presentationCorrectionOpportunities/opportunity'];o.status='AWAITING_NOTICE';delete o.notifiedAt;delete o.resultingSubmissionRef;o.dueAt=require('firebase-admin/firestore').Timestamp.fromMillis(Date.now()-1000);
  const rq=s.docs['presentationMaterialRequirements/requirement'];rq.openCorrectionCount=1;rq.currentSubmissionRef=o.originalSubmissionRef;rq.currentReviewRef=o.sourceReviewRef;
  payload={opportunityRef:{kind:'opportunity',id:o.id},basisRevision:o.revision,dueAt:Date.now()+86400000,reason:'Synthetic extension',authorityDecision:{kind:'REVISE_UNNOTIFIED_TERMS',reason:'Synthetic extension'}};
 }
 if(command.startsWith('PHOTO_')){
  const pr=s.docs['canonicalPresentations/presentation'];pr.photoPolicyRef={kind:'policy',id:'photo_policy'};
  s.docs['presentationPolicyReferences/photo_policy']={id:'photo_policy',revision:1,cycleRef:pr.cycleRef,approved:true,active:true,usage:'PRESENTATION_PHOTO',allowBoardReview:true,allowBoardUpload:true,allowMemberUpload:true,allowDriveLink:true,workspace:{enabled:true,folderId:'synthetic_folder_123',maximumBytes:1000,mimeTypes:['image/png']}};
  payload={presentationRef:{kind:'presentation',id:pr.id}};
  if(command==='PHOTO_UPLOAD_REQUEST'){self=true;payload.fileDescriptor={name:'synthetic.png',mime:'image/png',size:3,sha256:'a'.repeat(64)};}
  if(command==='PHOTO_SUBMIT'){self=true;payload={...payload,link:'https://drive.google.com/file/d/synthetic_file_123',notes:'Synthetic link only; no external access'};}
  if(command==='PHOTO_REVIEW'){pr.photo={revision:1,state:'SUBMITTED',link:'https://drive.google.com/file/d/synthetic_file_123'};payload={...payload,decision:'APPROVE',reason:'Synthetic review',internalReason:''};}
 }
 if(command==='ACTIVITY_SEMESTER_SAVE')s.request.target.scope.targetRefs.push(payload.semesterRef);
 if(command==='ATTENDANCE_PROOF_UPLOAD_REQUEST'){
  self=true;const aid=require('../src/environment.cjs').ACTIVITY,entry=Object.entries(base).find(([k,v])=>k.startsWith('activityAttendance/')&&k.split('/').length===2&&v.personRef?.id==='PES-000001'),fact=entry[1],activity=base['canonicalActivities/'+aid],collection=base['activityAttendanceCollections/'+aid],T=require('firebase-admin/firestore').Timestamp;
  activity.startsAt=T.fromMillis(Date.now()-3600000);activity.endsAt=T.fromMillis(Date.now()-1000);fact.status='ABSENT';collection.justificationEndsAt=T.fromMillis(Date.now()+86400000);collection.proofUploadPolicyRef={kind:'attachmentPolicy',id:'proof_policy'};
  s.docs[entry[0]]=fact;s.docs['canonicalActivities/'+aid]=activity;s.docs['activityAttendanceCollections/'+aid]=collection;
  s.docs[f4ActivitiesModel_().path(collection.proofUploadPolicyRef)]={id:'proof_policy',revision:1,approved:true,active:true,usage:'ATTENDANCE_PROOF',workspace:{enabled:true,folderId:'synthetic_folder_123',maximumBytes:1000,mimeTypes:['application/pdf']}};
  payload={activityRef:{kind:'activity',id:aid},attendanceRef:{kind:'attendance',id:fact.id},fileDescriptor:{name:'synthetic.pdf',mime:'application/pdf',size:3,sha256:'a'.repeat(64)},lateAcknowledged:false};
  s.request.target.scope.targetRefs=[payload.activityRef,payload.attendanceRef];
 }
 s.request.commandType='ACT.'+command;s.request.payload=payload;
 s.request.target.authorization={kind:self?'SELF':'ROLE',sourceId:'FLOW_POLICY',functionId:self?null:'VIG-001',actorScope:self?null:'PRESIDENCIA'};
 s.capability=f4ActivitiesPipeline_().capability(s.request);return s;
}
module.exports={remaining,extra};
