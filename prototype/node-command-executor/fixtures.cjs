'use strict';
const fs=require('node:fs'),p=require('node:path'),Module=require('node:module'),assert=require('node:assert/strict');
function fixture(){
 const original=p.resolve(__dirname,'test-support/current/operational-harness.cjs'),m=new Module(original,module);m.filename=original;m.paths=Module._nodeModulePaths(p.dirname(original));
 let s=fs.readFileSync(original,'utf8').replace("p.join(__dirname,'local.cjs')",JSON.stringify(p.join(__dirname,'local.cjs'))).replace("p.join(__dirname,'../ops/api.cjs')",JSON.stringify(p.join(__dirname,'fixture-config.cjs')));
 m._compile(s,original);const F=m.exports,h=F.setup({withoutRoster:true});F.H.queries(h.r);F.H.authorize(h.r);
 h.r.run(()=>{const db=f4Firestore_().connect(cfg),T=f4Firestore_().Timestamp.fromMillis;db.transaction(tx=>{const p=tx.get('memberAuthzPolicies/ACT_POLICY').data();tx.set('memberAuthzPolicies/ACT_POLICY',{...p,expiresAt:T(tick+86400000)});});});
 return {h,F};
}
function seed(uid){
 const {h}=fixture(),now=Date.now(),delta=now-Number(h.r.run(()=>tick)),source=[...h.r.documents];
 const shift=v=>typeof v==='string'?v.replaceAll('SYN_UID',uid):v&&typeof v==='object'?v.timestampValue?{timestampValue:new Date(Date.parse(v.timestampValue)+delta).toISOString()}:Array.isArray(v)?v.map(shift):Object.fromEntries(Object.entries(v).map(([k,x])=>[k.replaceAll('SYN_UID',uid),shift(x)])):v;
 const rows={};for(const [path,d]of source){if(/^(command|memberAdmission|memberAudit|presentationAudit|activityHistory)/.test(path))continue;rows[path.replaceAll('SYN_UID',uid)]={fields:shift(d.fields)};}
 rows['portalReaderControl/current']={fields:require('./wire.cjs').encode({schemaVersion:'portal-reader-control-v1',enabled:true,generation:1,operationalWrites:true,startsAt:new Date(now-1000),expiresAt:new Date(now+86400000)}).mapValue.fields};
 assert(rows['canonicalActivities/dev_meeting']);return rows;
}
module.exports={fixture,seed};
if(require.main===module){const docs=seed('B0_FAKE_UID');console.log({documents:Object.keys(docs).length,collections:[...new Set(Object.keys(docs).map(p=>p.split('/')[0]))]});}
