'use strict';
// Local-only synthetic President and approved configuration templates. Existing
// fixture documents are transformed before loading; no real identity is changed.
const fs=require('node:fs'),p=require('node:path'),Module=require('node:module');
const base=p.resolve(__dirname,'../functional');
const change=s=>s.replaceAll('DIRETOR_ADMINISTRATIVO','PRESIDENTE').replaceAll('ADMINISTRACAO','PRESIDENCIA');
function load(name,transform){const file=p.join(base,name),m=new Module(file,module);m.filename=file;m.paths=Module._nodeModulePaths(base);m._compile(transform(fs.readFileSync(file,'utf8')),file);return m.exports;}
const H=load('compat-harness.cjs',change);
const attendance=load('attendance.cjs',s=>change(s)
 .replace('   if(payload.activityRef)implicit.push',"   for(const a of payload.activityRefs||[])implicit.push(M.ref('attendanceCollection',a.id));\n   if(payload.activityRef)implicit.push")
 .replace("require('./local.cjs')",`require(${JSON.stringify(p.join(__dirname,'local.cjs'))})`)
 .replace("require('../ops/api.cjs')",`require(${JSON.stringify(p.join(__dirname,'../../fixture-config.cjs'))})`)
 .replace("sample=L.sample('CYCLE_CREATE')",`sample=JSON.parse(JSON.stringify(L.sample('CYCLE_CREATE')).replaceAll('DIRETOR_ADMINISTRATIVO','PRESIDENTE').replaceAll('ADMINISTRACAO','PRESIDENCIA'))`)
 .replace('   function collect(v)',`   if(['ACTIVITY_TYPE_SAVE','FREQUENCY_PERIOD_SAVE','ACTIVITY_POLICY_SAVE'].includes(cmd)&&!payload.operationalTemplateRef){
    const ref=payload.typeRef||payload.periodRef&&cmd==='FREQUENCY_PERIOD_SAVE'&&payload.periodRef||payload.attendancePolicyRef;
    if(!db.get(M.path(ref))){const template=M.ref(ref.kind,'approved_fixture_'+ref.kind);db.transaction(tx=>{tx.get(M.path(template));tx.set(M.path(template),{...payload,id:template.id,schemaVersion:'act-v2-candidate',recordMode:'NEW_V2',revision:1,approved:true,usedByCount:0,decision:{reasonCode:'DEV_FIXTURE',reason:'Approved synthetic criteria'},createdAt:M.stamp(tick),updatedAt:M.stamp(tick)});});payload.operationalTemplateRef=template;}
   }
   function collect(v)`));
module.exports={...attendance,H};
