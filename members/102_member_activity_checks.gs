// Bounded ACT checks used by voluntary requests. The query result is protected
// by the existing global ACT fence; every source is re-read in the command
// transaction. A browser assertion is never evidence of an empty calendar.
function f4MemberActivityChecks_(){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),M=f4ActivitiesModel_();
 const fencePath='portalReaderResourceFences/activities';
 function capture(get,db,personId,now){
  N(C.id(personId),'COMMAND_PAYLOAD_INVALID');
  const gate=get('portalReaderControl/current');N(gate?.enabled&&gate.operationalWrites===true,'ACCESS_REVOKED');
  const fence=get(fencePath),rows=[];let cursor;
  for(let page=0;page<5;page++){
   const part=db.query({collection:'presentationUnitHeads',filters:[{field:'personRef',op:'EQUAL',value:M.ref('person',personId)}],order:[{field:'id',direction:'ASCENDING'}],limit:20,...(cursor?{cursor:[cursor]}:{})});rows.push(...part);if(part.length<20)break;cursor=part.at(-1).data.id;
  }
  N(rows.length<100,'SCOPE_TOO_LARGE');
  const paths=new Set(),seen=new Set();
  function visit(path,depth=0){
   if(seen.has(path))return;seen.add(path);N(depth<=6&&seen.size<=180,'SCOPE_TOO_LARGE');
   const d=get(path);N(d,'RECONCILIATION_REQUIRED');paths.add(path);
   for(const k of ['assignmentRef','sessionRef','presentationRef','realizationRef','materialRequirementRef','currentRegularizationRef','currentReceiptRef'])if(d[k]?.kind&&d[k].kind!=='NONE')visit(M.path(d[k]),depth+1);
  }
  for(const row of rows){N(row.data.personRef?.id===personId,'RECONCILIATION_REQUIRED');visit('presentationUnitHeads/'+row.data.id);}
  const sources=[...paths].sort().map(path=>({path,hash:C.hash(F.encode(get(path)))}));
  return {schemaVersion:'member-activity-check-v1',personId,at:now,generation:gate.generation,fenceRevision:fence?.revision||0,unitIds:rows.map(row=>row.data.id).sort(),sources};
 }
 function read(get,check,personId){
  N(check?.schemaVersion==='member-activity-check-v1'&&check.personId===personId,'AUTH_BINDING_INVALID');
  const gate=get('portalReaderControl/current');N(gate?.enabled&&gate.operationalWrites===true&&gate.generation===check.generation&&(get(fencePath)?.revision||0)===check.fenceRevision,'REVISION_CONFLICT');
  N(Array.isArray(check.sources)&&check.sources.length<=180&&Array.isArray(check.unitIds)&&check.unitIds.length<100,'COMMAND_PAYLOAD_INVALID');
  const docs={};for(const s of check.sources){N(M.validPath(s.path),'COMMAND_PAYLOAD_INVALID');const d=get(s.path);N(d&&C.hash(F.encode(d))===s.hash,'REVISION_CONFLICT');docs[s.path]=d;}
  const doc=ref=>{if(!ref||ref.kind==='NONE')return null;const d=docs[M.path(ref)];N(d,'RECONCILIATION_REQUIRED');return d;},appointments=[],pendingMaterials=[];
  for(const id of check.unitIds){const u=docs['presentationUnitHeads/'+id];N(u?.personRef?.id===personId,'AUTH_BINDING_INVALID');const a=doc(u.assignmentRef),p=doc(u.presentationRef),rz=doc(p?.realizationRef),rq=doc(p?.materialRequirementRef);
   if(a&&['ACTIVE','CONFIRMED'].includes(a.status)&&!rz?.recognized){const s=doc(a.sessionRef);N(s?.civilDate,'RECONCILIATION_REQUIRED');C.civilStart(s.civilDate);appointments.push({unitId:id,civilDate:s.civilDate});}
   // Only an existing, applicable material requirement creates a pending fact.
   // Missing historical fields do not invent a new obligation.
   if(rz?.recognized&&rq&&!rq.currentRegularizationRef&&!rq.currentReceiptRef)pendingMaterials.push(id);
  }
  return {appointments,pendingMaterials};
 }
 function exit(get,check,personId,today,days){
  const facts=read(get,check,personId);N(Number.isInteger(days)&&days>=0,'COMMAND_PAYLOAD_INVALID');
  N(!facts.appointments.some(v=>{const gap=M.gap(today,v.civilDate);return gap>=0&&gap<days;}),'DOMAIN_STATE_INVALID');
  N(facts.pendingMaterials.length===0,'DOMAIN_STATE_INVALID');return facts;
 }
 function suspension(get,check,personId,startDate,endDate){const facts=read(get,check,personId);N(!facts.appointments.some(v=>v.civilDate>=startDate&&v.civilDate<=endDate),'DOMAIN_STATE_INVALID');return facts;}
 return {capture,read,exit,suspension};
}
