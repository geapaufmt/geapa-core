'use strict';
const fs=require('node:fs'),a=require('node:assert/strict'),W=require('./wire.cjs'),S=require('./sdk.cjs'),{save}=require('./session.cjs');
async function createFixtureDocs(db,docs){S.fence(db.projectId,'DEV');a.equal(db.databaseId,'(default)');const j=JSON.parse(fs.readFileSync(__dirname+'/evidence/seed.json'));a(!j.cleaned);const paths=Object.keys(docs);a(paths.length>0&&paths.length<=400);const old=await db.getAll(...paths.map(p=>db.doc(S.ROOT+'/'+p)));a(old.every(d=>!d.exists),'FIXTURE_PATH_ALREADY_EXISTS');j.paths=[...new Set([...j.paths,...paths])];save('seed.json',j);const b=db.batch();for(const [path,fields]of Object.entries(docs))b.create(db.doc(S.ROOT+'/'+path),W.decode({mapValue:{fields}},db));await b.commit();}
const replace=(value,map)=>{if(typeof value==='string'){for(const [a,b]of Object.entries(map))value=value.split(a).join(b);return value;}return Array.isArray(value)?value.map(v=>replace(v,map)):value&&typeof value==='object'?Object.fromEntries(Object.entries(value).map(([k,v])=>[replace(k,map),replace(v,map)])):value;};
async function attendance(db,uid){
 const seed=require('./fixtures.cjs').seed(uid),docs={};
 for(let i=3;i<=24;i++){const person='PES-'+String(i).padStart(6,'0'),link='VINC-'+String(i).padStart(3,'0'),map={'PES-000002':person,'VINC-002':link};for(const old of ['memberIdentities/PES-000002','memberLinks/VINC-002','memberPrivateProfiles/PES-000002','memberAggregateHeads/PES-000002']){a(seed[old]);docs[replace(old,map)]=replace(seed[old].fields,map);}}
 for(let i=0;i<4;i++)for(const old of ['canonicalActivities/dev_meeting','activityAttendanceCollections/dev_meeting'])docs[replace(old,{dev_meeting:'B0_CALL_'+i})]=replace(seed[old].fields,{dev_meeting:'B0_CALL_'+i,'Reunião sintética DEV':'B0_NODE_POC_CALL_'+i});
 await createFixtureDocs(db,docs);return {people:24,activities:['B0_CALL_0','B0_CALL_1','B0_CALL_2','B0_CALL_3']};
}
module.exports={createFixtureDocs,replace,attendance};
