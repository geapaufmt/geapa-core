'use strict';
// Bounded query shapes emitted by canonical member readers, never client SDK plans.
const {fail,ROOT}=require('./environment.cjs');
const collections=['memberAggregateHeads','memberEnrollments','memberProfileCorrections','memberVoluntaryRequests','presentationUnitHeads'];
function memberQuery(db,p,decode){
 if(!collections.includes(p?.collection))return null;
 const voluntary=p.collection==='memberVoluntaryRequests',order=voluntary?[{field:'createdAt',direction:'DESCENDING'},{field:'id',direction:'ASCENDING'}]:[{field:'id',direction:'ASCENDING'}];
 if(JSON.stringify(p.order)!==JSON.stringify(order))return null;
 if(Object.keys(p).some(k=>!['collection','filters','limit','order','cursor'].includes(k))||!Number.isInteger(p.limit)||p.limit<1||p.limit>20||!Array.isArray(p.filters)||p.filters.length>(voluntary?6:3))fail('QUERY_NOT_IN_PILOT');
 let q=db.collection((ROOT?ROOT+'/':'')+p.collection),seen=new Set();
 for(const f of p.filters){const v=decode(f.value);if(Object.keys(f).sort().join()!=='field,op,value'||seen.has(f.field+':'+f.op))fail('QUERY_NOT_IN_PILOT');seen.add(f.field+':'+f.op);
  const allowed=p.collection==='presentationUnitHeads'&&f.field==='personRef'&&f.op==='EQUAL'&&v?.kind==='person'&&Object.keys(v).length===2&&/^[A-Za-z0-9_-]{1,48}$/.test(v.id)||f.field==='id'&&f.op==='GREATER_THAN'&&typeof v==='string'&&/^[A-Za-z0-9_-]{1,48}$/.test(v)||f.op==='EQUAL'&&typeof v==='string'&&v.length<=100&&((p.collection==='memberProfileCorrections'&&['status','field'].includes(f.field))||(voluntary&&['status','modality','personId','semesterId'].includes(f.field)));
  const range=voluntary&&f.field==='createdAt'&&['GREATER_THAN_OR_EQUAL','LESS_THAN'].includes(f.op)&&typeof v?.toMillis==='function';if(!allowed&&!range)fail('QUERY_NOT_IN_PILOT');q=q.where(f.field,({EQUAL:'==',GREATER_THAN:'>',GREATER_THAN_OR_EQUAL:'>=',LESS_THAN:'<'})[f.op],v);
 }
 for(const o of order)q=q.orderBy(o.field,o.direction==='ASCENDING'?'asc':'desc');q=q.limit(p.limit);
 if(p.collection==='presentationUnitHeads'&&(p.filters.length!==1||p.filters[0].field!=='personRef'))fail('QUERY_NOT_IN_PILOT');
 if(p.cursor){if(!Array.isArray(p.cursor))fail('QUERY_NOT_IN_PILOT');const values=p.cursor.map(decode);if(voluntary){if(values.length!==2||typeof values[0]?.toMillis!=='function'||!/^[a-f0-9]{32}$/.test(values[1]))fail('QUERY_NOT_IN_PILOT');}else if(p.collection!=='presentationUnitHeads'||values.length!==1||typeof values[0]!=='string'||!/^[A-Za-z0-9_-]{1,48}$/.test(values[0]))fail('QUERY_NOT_IN_PILOT');q=q.startAfter(...values);}
 return q;
}
module.exports={memberQuery};
