const fs=require('fs'),vm=require('vm'),p=require('path');
function fixture(role='DIRETOR_RECURSOS_HUMANOS'){
 const h=require('./people-fixture.cjs').setup(role,1);for(const file of ['members/66_member_contracts.gs','institutional/200_institutional_directory.gs'])vm.runInContext(fs.readFileSync(p.resolve(__dirname,'../../..',file),'utf8'),h.domain.context);
 const pol=h.docs['memberAuthzPolicies/PEOPLE_POLICY'];pol.capabilities=['INSTITUTIONAL_DIRECTORY_READ'];pol.fieldMasks={INSTITUTIONAL_DIRECTORY_READ:['summary','fields','provenance','issues','canonical']};pol.targets.allowedPersons=['PES-000001'];pol.targets.allowedResources=['institutionalDirectory/identity'];
 h.docs['institutionalDirectoryControl/current']={schemaVersion:'institutional-directory-control-v1',enabled:true,policyIds:['PEOPLE_POLICY']};
 const record={id:'EMAIL_OFICIAL',title:'Contato institucional',kind:'Contato',status:'DOCUMENTADO',classification:'EDITORIAL',visibility:'BOARD',fields:[{label:'Valor',value:'test@example.invalid'}],sources:[{document:'Fonte sintética',sheet:'Config_GEAPA',row:2,hash:'a'.repeat(64)}],issues:[],canonical:null};
 h.doc=h.docs['institutionalDirectory/identity']={schemaVersion:'institutional-directory-v1',id:'identity',section:'identity',revision:1,nature:'DOCUMENTARY_ONLY',sourceHash:'a'.repeat(64),records:[record,{...record,id:'LINK_GRUPO_MEMBROS_WHATSAPP',visibility:'RESTRICTED',fields:[{label:'Valor',value:'PRIVATE_INVITE_SENTINEL'}]}]};
 h.read=(params={})=>JSON.parse(h.domain.run('(x)=>f4InstitutionalDirectoryRead_({config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:x.now+60000},clock:()=>x.now},x.params)',{now:h.now,params:{section:'identity',query:'',kind:'',page:1,...params}}));h.model=(method,input)=>JSON.parse(h.domain.run('(x)=>f4InstitutionalDirectoryModel_()[x.method](x.doc,x.input)',{method,doc:h.doc,input}));return h;
}

module.exports={fixture};
