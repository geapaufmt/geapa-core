'use strict';
// Read-only audit arithmetic. Never grants authority, changes a function or ends a mandate.
const a=require('assert/strict'),DAY=86400000;
function day(s){a.equal(typeof s,'string');a(/^\d{4}-\d{2}-\d{2}$/.test(s));const n=Date.parse(s+'T00:00:00Z');a.equal(new Date(n).toISOString().slice(0,10),s);return n/DAY;}
const date=n=>new Date(n*DAY).toISOString().slice(0,10);
function interval(v){const start=day(v.startDate),end=day(v.endDateReal||v.endDatePlanned||v.endDate)+1;a(end>start);return [start,end];}
function union(xs){const out=[];for(const [s,e]of xs.filter(([s,e])=>e>s).sort((a,b)=>a[0]-b[0]||a[1]-b[1])){const last=out.at(-1);if(last&&s<=last[1])last[1]=Math.max(e,last[1]);else out.push([s,e]);}return out;}
function continuity(functions){const groups=new Map();for(const f of functions){const key=JSON.stringify([f.idPessoa,f.idVinculo,f.idDiretoria,f.idCargo,f.scopeKey]);if(!groups.has(key))groups.set(key,[]);groups.get(key).push(f);}
 return [...groups.values()].flatMap(rows=>union(rows.map(interval)).map(([s,e])=>({personId:rows[0].idPessoa,linkId:rows[0].idVinculo,boardId:rows[0].idDiretoria,role:rows[0].idCargo,scope:rows[0].scopeKey,startDate:date(s),endDate:date(e-1),sourceIds:rows.filter(f=>{const [a,z]=interval(f);return a<e&&z>s;}).map(f=>f.id).sort()})));}
function proportionalService(functions,terms,asOf){
 const end=day(asOf)+1,service=union(functions.map(interval).map(([a,z])=>[a,Math.min(z,end)]));
 const unique=new Map();for(const t of terms){const iv=interval(t),key=iv.join('/');if(!unique.has(key))unique.set(key,{...t,iv});}
 const sorted=[...unique.values()].sort((a,b)=>a.iv[0]-b.iv[0]);for(let i=1;i<sorted.length;i++)a(sorted[i-1].iv[1]<=sorted[i].iv[0],'AMBIGUOUS_BOARD_TERMS');
 const parts=sorted.map(t=>{const [a,z]=t.iv,days=service.reduce((n,[s,e])=>n+Math.max(0,Math.min(z,e)-Math.max(a,s)),0);return {termId:t.id,days,totalDays:z-a,equivalentSemesters:days/(z-a)};}).filter(x=>x.days);
 const coveredDays=parts.reduce((s,x)=>s+x.days,0),days=service.reduce((s,[a,z])=>s+z-a,0);
 return {asOf,days,coveredDays,unmappedDays:days-coveredDays,equivalentSemesters:parts.reduce((s,x)=>s+x.equivalentSemesters,0),parts,automaticRevocation:false,decisionEligible:days===coveredDays};
}
module.exports={continuity,proportionalService,union};
