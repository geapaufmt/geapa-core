const {test}=require('node:test'),a=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),p=require('node:path'),{fixture}=require('./institutional-fixture.cjs');
test('read-only director write-discovery denial remains a handled authorization result',()=>{
 const h=fixture('DIRETOR_ADMINISTRATIVO');vm.runInContext(fs.readFileSync(p.resolve(__dirname,'../../../institutional/206_institutional_governance.gs'),'utf8'),h.domain.context);
 const run=(allowed,enabled)=>h.domain.run('(x)=>{const c={schemaVersion:"institutional-governance-control-v1",projectId:nodeConfig.projectId,executor:f4CommandExecutorId_(),resource:"institutionalGovernanceDrafts/current",commandType:"institutional.governance.draft.save",publicationEnabled:false,formalApprovalEnabled:false,appCheck:{mode:"REQUIRED"},revision:1,epoch:1,actorUids:x.allowed?["CF_ADMIN"]:["OTHER"],enabled:x.enabled};return f4InstitutionalGovernance_().control(()=>c,nodeConfig,{uid:"CF_ADMIN"},Date.now(),true);}',{allowed,enabled});
 a.throws(()=>run(false,false),e=>e.code==='AUTHORIZATION_DENIED');
 a.throws(()=>run(true,false),e=>e.code==='PILOT_DISABLED');
});
