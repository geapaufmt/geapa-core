'use strict';
const {FieldPath,Timestamp}=require('firebase-admin/firestore');
const {PROJECT,ROOT,fail}=require('./environment.cjs');
// Only plans emitted by the Core reader producers. Never exposed in HTTP.
const material=['presentationMaterialSubmissions','presentationMaterialReviews','presentationMaterialReceipts','presentationCorrectionOpportunities','presentationMaterialRegularizations'];
const shapes={
 canonicalPresentations:[['personRef.ref.id']],canonicalActivities:[[],['cycleRef']],memberAggregateHeads:[[]],activityTypes:[[]],activityPolicies:[[]],activityAttachmentPolicies:[[]],
 academicSemesters:[[]],activityExceptionalJustificationWindows:[[]],presentationCycles:[[]],
 presentationPolicyReferences:[[],['cycleRef'],['cycleRef','usage']],
 presentationSessions:[['cycleRef']],presentationObligations:[['cycleRef']],
 presentationUnitHeads:[[],['personRef'],['cycleRef']],presentationFlowCases:[['cycleRef']],
 presentationThematicRequests:[['cycleRef','workflowState']],calendarAssignments:[['unitRef']],
 activityAttendance:[['ownerPersonRef']],activityJustifications:[['state']],presentationArtifacts:[['activityRef']]
};
for(const c of material)shapes[c]=[['presentationRef']];
function readerQuery(db,p,decode){
 const members=require('./member-query.cjs').memberQuery(db,p,decode);if(members)return members;
 if(!p||Object.keys(p).some(k=>!['collection','filters','limit','order','cursor'].includes(k))||!shapes[p.collection]||
 !Number.isInteger(p.limit)||p.limit<1||p.limit>(['canonicalActivities','presentationArtifacts'].includes(p.collection)?100:20)||!Array.isArray(p.filters)||
 !shapes[p.collection].some(shape=>JSON.stringify(shape)===JSON.stringify(p.filters.map(f=>f.field))))fail('QUERY_NOT_IN_PILOT');
 const calendar=p.collection==='calendarAssignments',order=calendar?[{field:'createdAt',direction:'DESCENDING'},{field:'__name__',direction:'ASCENDING'}]:[{field:'__name__',direction:'ASCENDING'}];
 if(JSON.stringify(p.order)!==JSON.stringify(order))fail('QUERY_NOT_IN_PILOT');
 let q=db.collection((ROOT?ROOT+'/':'')+p.collection);
 for(const f of p.filters){
  if(Object.keys(f).sort().join()!=='field,op,value'||f.op!=='EQUAL')fail('QUERY_NOT_IN_PILOT');
  const v=decode(f.value),kind={cycleRef:'cycle',personRef:'person',unitRef:'unit',presentationRef:'presentation'}[f.field];
  if(kind){if(v?.kind!==kind||Object.keys(v).length!==2||!/^[A-Za-z0-9_-]{1,48}$/.test(v.id)||kind==='person'&&!/^PES-[0-9]{6}$/.test(v.id))fail('QUERY_NOT_IN_PILOT');}
  else if(f.field==='personRef.ref.id'){if(typeof v!=='string'||!/^[A-Za-z0-9_-]{1,48}$/.test(v))fail('QUERY_NOT_IN_PILOT');}
  else if(f.field==='ownerPersonRef'){if(typeof v!=='string'||!/^PES-[0-9]{6}$/.test(v))fail('QUERY_NOT_IN_PILOT');}
  else if(f.field==='state'){if(!['PREVIOUS','PENDING','CORRECTION_REQUESTED'].includes(v))fail('QUERY_NOT_IN_PILOT');}
  else if(f.field==='workflowState'){if(v!=='WAITING')fail('QUERY_NOT_IN_PILOT');}
  else if(f.field==='usage'){if(!['MATERIAL_PROVIDER','PRESENTATION_PHOTO'].includes(v))fail('QUERY_NOT_IN_PILOT');}
  else if(f.field==='activityRef'){if(v?.resolution!=='RESOLVED'||v.ref?.kind!=='activity'||Object.keys(v).length!==2||Object.keys(v.ref).length!==2||!/^[A-Za-z0-9_-]{1,128}$/.test(v.ref.id))fail('QUERY_NOT_IN_PILOT');}
  else fail('QUERY_NOT_IN_PILOT');
  q=q.where(f.field,'==',v);
 }
 for(const o of order)q=q.orderBy(o.field==='__name__'?FieldPath.documentId():o.field,o.direction==='ASCENDING'?'asc':'desc');
 q=q.limit(p.limit);
 if(p.cursor){
  const prefix='projects/'+PROJECT+'/databases/(default)/documents/'+(ROOT?ROOT+'/':'')+p.collection+'/',n=calendar?2:1;
  if(!Array.isArray(p.cursor)||p.cursor.length!==n)fail('DESTINATION_INVALID');
  const tail=p.cursor[n-1];if(!tail||Object.keys(tail).join()!=='referenceValue'||!tail.referenceValue.startsWith(prefix)||!/^[A-Za-z0-9_-]{1,128}$/.test(tail.referenceValue.slice(prefix.length)))fail('DESTINATION_INVALID');
  const values=[];if(calendar){const t=decode(p.cursor[0]);if(!(t instanceof Timestamp))fail('DESTINATION_INVALID');values.push(t);}
  values.push(db.doc((ROOT?ROOT+'/':'')+p.collection+'/'+tail.referenceValue.slice(prefix.length)));q=q.startAfter(...values);
 }
 return q;
}
module.exports={readerQuery};
