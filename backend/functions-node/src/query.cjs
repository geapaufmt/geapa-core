'use strict';
const {FieldPath}=require('firebase-admin/firestore');
const {PROJECT,ROOT,fail}=require('./environment.cjs');
// Server-created plans only. Bounded list queries used by the original producer;
// no collection name, filter, cursor or SDK reference comes from HTTP input.
function queryPlan(db,plan,decode,specialized=false){
 if(specialized)return require('./reader-query.cjs').readerQuery(db,plan,decode);
 if(!plan||Object.keys(plan).some(k=>!['collection','filters','limit','order','cursor'].includes(k))||
  !['calendarAssignments','canonicalActivities','presentationArtifacts','activityAttendance','memberAggregateHeads','presentationSessions','presentationObligations','activityTypes','presentationCycles','presentationUnitHeads','presentationCorrectionOpportunities'].includes(plan.collection)||
  !Number.isInteger(plan.limit)||plan.limit<1||plan.limit>(['canonicalActivities','presentationArtifacts'].includes(plan.collection)?100:20)||
  !Array.isArray(plan.filters)||plan.filters.length>1||
  JSON.stringify(plan.order)!==JSON.stringify([{field:'__name__',direction:'ASCENDING'}]))fail('QUERY_NOT_IN_PILOT');
 if(plan.collection==='calendarAssignments'&&(plan.limit!==1||plan.cursor))fail('QUERY_NOT_IN_PILOT');
 let query=db.collection((ROOT?ROOT+'/':'')+plan.collection);
 if(['calendarAssignments','activityAttendance','presentationSessions','presentationObligations','presentationCorrectionOpportunities'].includes(plan.collection)&&plan.filters.length!==1)fail('QUERY_NOT_IN_PILOT');
 for(const f of plan.filters){
  if(plan.collection==='presentationCorrectionOpportunities'){const v=decode(f.value);if(f.field!=='presentationRef'||f.op!=='EQUAL'||Object.keys(f).length!==3||v?.kind!=='presentation'||!/^[-A-Za-z0-9_]{1,48}$/.test(v.id)||Object.keys(v).length!==2)fail('QUERY_NOT_IN_PILOT');query=query.where(f.field,'==',v);continue;}
  if(['calendarAssignments','presentationSessions','presentationObligations'].includes(plan.collection)){const v=decode(f.value);if(f.field!=='cycleRef'||f.op!=='EQUAL'||Object.keys(f).length!==3||v?.kind!=='cycle'||!/^[-A-Za-z0-9_]{1,48}$/.test(v.id)||Object.keys(v).length!==2)fail('QUERY_NOT_IN_PILOT');query=query.where(f.field,'==',v);continue;}
  if(plan.collection==='activityAttendance'){const v=decode(f.value);if(f.field!=='ownerPersonRef'||f.op!=='EQUAL'||Object.keys(f).length!==3||typeof v!=='string'||!/^PES-[0-9]{6}$/.test(v))fail('QUERY_NOT_IN_PILOT');query=query.where(f.field,'==',v);continue;}
  if(plan.collection!=='presentationArtifacts'||f.field!=='activityRef'||f.op!=='EQUAL'||Object.keys(f).length!==3)fail('QUERY_NOT_IN_PILOT');
  const value=decode(f.value);
  if(value?.resolution!=='RESOLVED'||value.ref?.kind!=='activity'||!/^[A-Za-z0-9_-]{1,128}$/.test(value.ref.id)||Object.keys(value).length!==2||Object.keys(value.ref).length!==2)fail('QUERY_NOT_IN_PILOT');
  query=query.where(f.field,'==',value);
 }
 query=query.orderBy(FieldPath.documentId(),'asc').limit(plan.limit);
 if(plan.cursor){
  const prefix='projects/'+PROJECT+'/databases/(default)/documents/'+(ROOT?ROOT+'/':'')+plan.collection+'/';
  const value=plan.cursor[0]?.referenceValue;
  if(!Array.isArray(plan.cursor)||plan.cursor.length!==1||typeof value!=='string'||!value.startsWith(prefix)||!/^[A-Za-z0-9_-]+$/.test(value.slice(prefix.length)))fail('DESTINATION_INVALID');
  query=query.startAfter(db.doc((ROOT?ROOT+'/':'')+plan.collection+'/'+value.slice(prefix.length)));
 }
 return query;
}
module.exports={queryPlan};
